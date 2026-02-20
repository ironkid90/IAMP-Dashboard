# IAMP Sites Mapping Dashboard (Redesigned + Map-ready)

This is a **static** (HTML/CSS/JS) dashboard for the **IAMP Informal Settlements Sites Mapping** spreadsheet.

It runs entirely in the browser and provides:
- Progress KPIs (assessed vs not assessed)
- Phone-call outcomes + site-status mix
- Active-site totals (structures / households / individuals / latrines + composition)
- QC monitoring (issue rate + issues by type + quick remediation tips)
- Records table (search/sort/export) with **PII hidden by default**
- **Interactive Lebanon map** (Leaflet) — supports point markers now, choropleths later

---

## Quick start

1. Open: `hand_sharepoint_dashboard/index.html`
2. Click **Data → Load File** and select your latest exported `.xlsx`.

A small **redacted sample** is included and auto-loaded on first open:
- `hand_sharepoint_dashboard/assets/data/IAMP_sites_mapping_SAMPLE_REDACTED.xlsx`
Disable auto-load via: `?noSample=1`

---

## Map: how it works

### What the map needs
The map reads coordinates **directly from your main spreadsheet**.

The dashboard automatically detects the best Latitude/Longitude columns (supports common names like `Latitude` / `Longitude`, `Lat` / `Lng`, etc.).
If no usable coordinates are found, the Map tab will show a clear “No coordinates found” hint.

### Optional boundaries
You can upload a **GeoJSON** boundary file to show admin boundaries (Governorate/District/Cadaster):
- Map tab → **Boundaries (optional)**

This is the placeholder step needed before we add **choropleths + click-to-filter** using ACS codes.

---

## Deploy

### Netlify
- Drag-and-drop the folder, or connect a repo.
- Publish directory: the project root (so `/hand_sharepoint_dashboard` works).

### Vercel
- Deploy as a static project.
- Ensure `hand_sharepoint_dashboard/` is included in the output.

### Vercel (recommended) — Live SharePoint mode (no manual uploads)
This repo includes optional **Vercel Functions** that can securely fetch the latest XLSX from SharePoint/OneDrive via **Microsoft Graph**.

Once configured, the dashboard can run in **Live mode** and load the latest spreadsheet from:
- `GET /api/xlsx` (binary XLSX)
- `GET /api/status` (metadata)

> Why: avoids CORS/auth issues because the browser only talks to your own Vercel domain.

#### Minimal Environment Variables

**Option 1 (simplest):** if you can get a **direct download URL** to the XLSX (e.g., an “Anyone with the link” file that returns the `.xlsx` bytes), set just:
- `LIVE_XLSX_URL`

**Option 2 (private SharePoint):** use Microsoft Graph (app-only). To reduce env vars, set **one** variable:
- `IAMP_GRAPH_CONFIG` (JSON) **or** `IAMP_GRAPH_CONFIG_B64` (base64 of JSON)

Example JSON:
```json
{
  "MS_TENANT_ID": "...",
  "MS_CLIENT_ID": "...",
  "MS_CLIENT_SECRET": "...",
  "SP_SHARE_LINK": "https://..."
}
```

(You can also set the individual `MS_*` and `SP_*` vars as a classic setup — see `.env.example`.)

After deploy, open the dashboard → **Data panel → Live mode**.

---

## Notes about SharePoint embedding
Some SharePoint/OneDrive preview pages block JavaScript in iframes. For reliable results:
- Host this dashboard on Vercel/Netlify
- Embed using an iframe that allows scripts

---

## Privacy
PII columns (names/phone numbers) are **hidden by default** in the Records tab.

**Important:** Do NOT ship the real spreadsheet inside `assets/data/` if this will be publicly accessible, because anyone could download it from the browser.
