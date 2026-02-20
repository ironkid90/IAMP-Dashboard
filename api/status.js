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
        url: directUrl,
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
      name: meta?.name || null,
      lastModifiedDateTime: meta?.lastModifiedDateTime || null,
      size: meta?.size || null,
      webUrl: meta?.webUrl || null,
      driveId,
      itemId,
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err?.message || String(err) }, 500);
  }
}
