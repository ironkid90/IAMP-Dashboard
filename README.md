# IAMP Sites Mapping Dashboard (Static Build)

This package is a **Phase 1 + Phase 2 structure** build:

- **Phase 1 (UI polish)**: consistent design tokens, cleaner background/card hierarchy, smaller charts, click-to-filter interactions, improved table UX.
- **Phase 2 (Geo structure)**: points + boundaries + choropleth + boundary click-to-filter, using **PCODEs** as the join key.

## What’s inside

### `/hand_sharepoint_dashboard/`
Static dashboard (HTML/CSS/JS):

- Auto-loads `assets/data/sites.json` (already merged, no browser Excel parsing).
- Filters by Governorate/District/Cadaster (PCODE), Site Status, Phone call status, QC.
- KPIs + charts + table + map (Leaflet).
- CSV export of the filtered view.

### `/hand_sharepoint_dashboard/assets/data/`
Generated data artifacts:

- `sites.json` – canonical points dataset (redacted; no phone numbers)
- `boundaries_admin1.geojson` – trimmed admin1
- `boundaries_admin2.geojson` – trimmed admin2
- `boundaries_admin3_subset.geojson` – trimmed admin3 **subset** (fast load)
- `boundaries_admin3_full.geojson` – trimmed admin3 **full** (optional; heavier)

### `/api/data.js`
A Vercel Serverless Function stub.

Current behavior: serves the bundled `sites.json` with caching + ETag.

Next step: replace it with logic that fetches the latest XLSX + ArcGIS JSON, merges server-side, and returns clean JSON.

### `/tools/preprocess_iamp.py`
Offline preprocessing script to regenerate `sites.json` + trimmed boundaries.

## Deploy (Vercel)

1. Deploy the folder as a static site.
2. The dashboard entrypoint is:
   - `/` (redirects to `/hand_sharepoint_dashboard/`)

## Regenerate the JSON (local)

Example:

```bash
python tools/preprocess_iamp.py \
  --assessment_csv "IAMP_sites_mapping_YYYYMMDD.csv" \
  --master_xlsx "IAMP-133_ListofInformalSettlements_29_August_2025.xlsx" \
  --boundaries_dir "./boundaries" \
  --out_dir "./hand_sharepoint_dashboard/assets/data" \
  --subset_admin3 \
  --write_full_admin3
```

Where `./boundaries` contains:

- `lbn_admin1.geojson`
- `lbn_admin2.geojson`
- `lbn_admin3.geojson`

## Data model (sites.json)

Each site record includes:

- `pcode`, `name`, `local_name`
- `lat`, `lon`
- `adm1_pcode`, `adm2_pcode`, `adm3_pcode`
- operational fields: `site_status`, `phone_status`, `assessed_date`
- metrics: totals + structure breakdown
- QC flags: `qc_any_issue`, `qc_issue_count`, and issue-specific booleans

PII (phone numbers, personal names) is intentionally not exported.

## Notes / next improvements

- Replace lightweight QC flags with your full QC pivot rules.
- Convert boundaries to TopoJSON + simplify for faster map loading.
- Switch frontend fetch to `/api/data` for “always latest” without redeploying static assets.
