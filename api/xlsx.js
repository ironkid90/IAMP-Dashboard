import { getAccessToken, resolveDriveItem, getDriveItemMeta, downloadDriveItemContent } from "./_lib/graph.js";

export default async function handler(request) {
  if (request.method !== "GET") {
    return new Response("Method not allowed", {
      status: 405,
      headers: { "allow": "GET" },
    });
  }

  try {
    // Simplest mode: proxy a direct XLSX URL (publicly downloadable)
    // Set LIVE_XLSX_URL in Vercel env vars.
    const directUrl = (process.env.LIVE_XLSX_URL || "").trim();
    if (directUrl) {
      const res = await fetch(directUrl, { redirect: "follow" });
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(`Direct URL fetch failed (HTTP ${res.status}). ${txt ? "Details: " + txt : ""}`.trim());
      }
      const buf = await res.arrayBuffer();
      const lastMod = res.headers.get("last-modified") || "";
      return new Response(buf, {
        status: 200,
        headers: {
          "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "content-disposition": `inline; filename="iamp_sites_mapping.xlsx"`,
          "cache-control": "no-store",
          "x-iamp-last-modified": lastMod,
        },
      });
    }

    const token = await getAccessToken();
    const resolved = await resolveDriveItem(token);
    const driveId = resolved.driveId;
    const itemId = resolved.itemId;

    // Grab metadata so the UI can display last-modified and we can name the file.
    const meta = await getDriveItemMeta(token, driveId, itemId).catch(() => ({}));

    const buf = await downloadDriveItemContent(token, driveId, itemId);

    const filename = meta?.name || resolved?.name || "iamp_sites_mapping.xlsx";
    const lastMod = meta?.lastModifiedDateTime || "";

    return new Response(buf, {
      status: 200,
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `inline; filename="${String(filename).replace(/"/g, "")}"`,
        "cache-control": "no-store",
        "x-iamp-last-modified": lastMod,
      },
    });
  } catch (err) {
    const message = err?.message || String(err);
    return new Response(
      JSON.stringify({ ok: false, error: message }, null, 2),
      {
        status: 500,
        headers: {
          "content-type": "application/json; charset=utf-8",
          "cache-control": "no-store",
        },
      }
    );
  }
}
