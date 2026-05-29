import { getAccessToken, resolveDriveItem, getDriveItemMeta } from "./_lib/graph.js";

function jsonResponse(obj, status = 200, headers = {}) {
  return new Response(JSON.stringify(obj, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      ...headers,
    },
  });
}

function getConfiguredSourceMode() {
  const directUrl = (process.env.LIVE_XLSX_URL || "").trim();
  if (directUrl) {
    return "direct-url";
  }

  const hasDriveCoordinates = Boolean((process.env.SP_DRIVE_ID || "").trim() && (process.env.SP_ITEM_ID || "").trim());
  const hasShareLink = Boolean((process.env.SP_SHARE_LINK || "").trim());
  if (hasDriveCoordinates || hasShareLink) {
    return "graph";
  }

  return null;
}

export default async function handler(request) {
  if (request.method !== "GET") {
    return jsonResponse({ ok: false, error: "Method not allowed" }, 405, {
      "allow": "GET",
    });
  }

  try {
    const directUrl = (process.env.LIVE_XLSX_URL || "").trim();
    if (directUrl) {
      // Use HEAD when possible to avoid downloading the whole file.
      const head = await fetch(directUrl, { method: "HEAD", redirect: "follow" }).catch(() => null);
      const lastModifiedDateTime = head ? (head.headers.get("last-modified") || null) : null;
      const size = head ? (head.headers.get("content-length") || null) : null;
      return jsonResponse({
        ok: true,
        mode: "direct-url",
        sourceConfigured: true,
        workbookDownloadExposed: false,
        name: "iamp_sites_mapping.xlsx",
        lastModifiedDateTime,
        size,
      });
    }

    const token = await getAccessToken();
    const { driveId, itemId } = await resolveDriveItem(token);
    const meta = await getDriveItemMeta(token, driveId, itemId);

    return jsonResponse({
      ok: true,
      mode: "graph",
      sourceConfigured: true,
      workbookDownloadExposed: false,
      name: meta?.name || null,
      lastModifiedDateTime: meta?.lastModifiedDateTime || null,
      size: meta?.size || null,
    });
  } catch (err) {
    const configuredMode = getConfiguredSourceMode();
    return jsonResponse({
      ok: false,
      sourceConfigured: configuredMode !== null,
      mode: configuredMode,
      workbookDownloadExposed: false,
      error: configuredMode ? "Source metadata is unavailable." : (err?.message || String(err)),
    }, configuredMode ? 502 : 500);
  }
}
