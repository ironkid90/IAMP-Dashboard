// Minimal Microsoft Graph helpers for Vercel Functions (Node.js runtime)
// - Uses OAuth2 client credentials to get an access token
// - Supports Drive Item lookup by driveId/itemId OR by Share link

let tokenCache = {
  accessToken: null,
  expiresAtMs: 0,
};

function readJsonConfig() {
  const raw = (process.env.IAMP_GRAPH_CONFIG || process.env.IAMP_CONFIG || "").trim();
  const b64 = (process.env.IAMP_GRAPH_CONFIG_B64 || process.env.IAMP_CONFIG_B64 || "").trim();

  try {
    if (raw) return JSON.parse(raw);
  } catch (_) {}

  try {
    if (b64) {
      const decoded = Buffer.from(b64, "base64").toString("utf8");
      return JSON.parse(decoded);
    }
  } catch (_) {}

  return null;
}

const CONFIG = readJsonConfig();

function env(name, fallback = "") {
  const v = (process.env[name] ?? (CONFIG ? (CONFIG[name] ?? CONFIG[name.toLowerCase()] ?? "") : "") ?? fallback);
  return String(v ?? "").trim();
}

export function requireEnv(name) {
  const v = env(name);
  if (!v) throw new Error(`Missing environment variable: ${name}`);
  return v;
}

export async function getAccessToken() {
  const tenantId = requireEnv("MS_TENANT_ID");
  const clientId = requireEnv("MS_CLIENT_ID");
  const clientSecret = requireEnv("MS_CLIENT_SECRET");

  // Reuse token across invocations when the function instance is warm.
  const now = Date.now();
  if (tokenCache.accessToken && now < tokenCache.expiresAtMs - 60_000) {
    return tokenCache.accessToken;
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded",
    },
    body,
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Token request failed (HTTP ${res.status}). ${txt ? "Details: " + txt : ""}`.trim());
  }

  const json = await res.json();
  if (!json?.access_token) throw new Error("Token response missing access_token.");

  tokenCache.accessToken = json.access_token;
  tokenCache.expiresAtMs = now + (Number(json.expires_in || 3600) * 1000);
  return tokenCache.accessToken;
}

// Microsoft Graph expects a "sharing token" in the form: u!<base64url(url)>
// See: https://learn.microsoft.com/graph/api/shares-get
export function encodeShareLinkToShareId(shareLink) {
  const raw = String(shareLink || "").trim();
  if (!raw) return "";
  const b64 = Buffer.from(raw, "utf8").toString("base64");
  // Base64Url encode
  const b64url = b64.replace(/=+$/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  return `u!${b64url}`;
}

export async function resolveDriveItem(accessToken) {
  const driveId = env("SP_DRIVE_ID");
  const itemId = env("SP_ITEM_ID");
  const shareLink = env("SP_SHARE_LINK");

  if (driveId && itemId) {
    return { driveId, itemId, name: env("SP_FILE_NAME") || null };
  }

  if (shareLink) {
    const shareId = encodeShareLinkToShareId(shareLink);
    if (!shareId) throw new Error("SP_SHARE_LINK is set but could not be encoded.");

    const url = `https://graph.microsoft.com/v1.0/shares/${encodeURIComponent(shareId)}/driveItem`;
    const res = await fetch(url, {
      headers: {
        authorization: `Bearer ${accessToken}`,
        accept: "application/json",
      },
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      throw new Error(`Share link lookup failed (HTTP ${res.status}). ${txt ? "Details: " + txt : ""}`.trim());
    }

    const di = await res.json();
    const resolvedDriveId = di?.parentReference?.driveId;
    const resolvedItemId = di?.id;
    const name = di?.name || null;

    if (!resolvedDriveId || !resolvedItemId) {
      throw new Error("Share link lookup succeeded but did not return driveId/itemId.");
    }

    return { driveId: resolvedDriveId, itemId: resolvedItemId, name };
  }

  throw new Error(
    "No SharePoint file configured. Provide either (SP_DRIVE_ID + SP_ITEM_ID) OR SP_SHARE_LINK in Vercel env vars."
  );
}

export async function getDriveItemMeta(accessToken, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=name,lastModifiedDateTime,size,webUrl`;
  const res = await fetch(url, {
    headers: {
      authorization: `Bearer ${accessToken}`,
      accept: "application/json",
    },
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Metadata request failed (HTTP ${res.status}). ${txt ? "Details: " + txt : ""}`.trim());
  }

  return await res.json();
}

export async function downloadDriveItemContent(accessToken, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`;
  const res = await fetch(url, {
    headers: {
      authorization: `Bearer ${accessToken}`,
    },
    redirect: "follow",
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Download failed (HTTP ${res.status}). ${txt ? "Details: " + txt : ""}`.trim());
  }

  const buf = await res.arrayBuffer();
  return buf;
}
