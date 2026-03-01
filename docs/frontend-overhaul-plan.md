# Frontend Overhaul Plan (IAMP Dashboard)

This repo currently ships a static, public-facing dashboard under [`hand_sharepoint_dashboard/index.html`](hand_sharepoint_dashboard/index.html:1), built with Bootstrap, Chart.js, Leaflet, and Tabulator.

The requested overhaul focuses on:

1. A more professional, “snapshot report” visual language (taking cues from IOM DTM mobility snapshot layouts).
2. Better information architecture and clearer navigation (re-arranged submenus and sections).
3. A redesigned front page for dissemination, while keeping the existing interactive exploration tools.

## 1) UX goals

### Primary user journeys

- **Quick briefing** (1–2 minutes)
  - See “what changed” and key figures at a glance.
  - Understand coverage (assessed %, QC flags, missing coordinates).
  - Export or share a filtered link.

- **Exploration** (5–15 minutes)
  - Filter by location (Gov, District, Cadaster) and operational status.
  - Explore the map (markers + choropleth) and click-to-filter.
  - Inspect records and export filtered data.

- **Data quality review**
  - Find QC flagged records.
  - Export QC-only subset.

### Visual direction (inspired by IOM “snapshot” PDFs)

- Strong top header with dataset metadata (Round/date, last updated, source).
- Sections with clear titles, consistent spacing, and “report-style” cards.
- Limited palette, high contrast, and restrained decoration.

## 2) Proposed information architecture

### New top-level navigation

Replace the current tab row with a grouped navigation that reads like a report + tool suite:

- **Home**: key KPIs + headline charts + “what this shows” + quick actions.
- **Explore**
  - **Map**: map-first exploration with filters and legend.
  - **Insights**: charts that support interpretation.
- **Data Quality**: QC overview + guidance + QC-only preview table.
- **Records**: full table with export.
- **About / Methodology**: definitions, disclaimers, and data notes.

This can still be implemented with Bootstrap tabs under the hood to keep the app single-page and static.

### Filter panel restructure

Convert the current filter list (left sidebar) into an accordion with 3 groups:

1. **Location**: Gov, District, Cadaster
2. **Operational status**: Site status, Phone status
3. **Data quality**: QC filter, missing coords, etc.

Add a “Quick actions” area:

- Reset filters
- Copy share link
- Export filtered CSV
- Export QC-only CSV

## 3) Front page redesign (root)

The repo root currently redirects immediately via [`index.html`](index.html:1). Replace this with a real landing page:

- Brief description + CTA button (“Open Dashboard”)
- Links to documentation and methodology
- Optional thumbnails or “feature highlights” tiles

Keep a fallback link to the dashboard folder for reliability.

## 4) Dashboard front page (“Home” inside app)

Create a **Home** view that works as a shareable briefing screen:

- **Snapshot KPIs** (existing data points): total records, assessed, active sites, QC-any, missing coordinates.
- **Headline visuals** (lightweight, no new deps):
  - Assessment progress
  - Site status distribution
  - Phone outcome distribution
  - “Top districts/cadasters” by active individuals
- “How to use this dashboard” quick instructions.

## 5) Interactive tools to add/improve

### A) “Story-like” interactions

- Click bars or slices to apply filters (already present for several charts in [`assets/js/app.js`](hand_sharepoint_dashboard/assets/js/app.js:1)). Extend to any new charts.
- Persist current filter state in the share URL (already implemented via `?s=&g=&d=&c=...`).

### B) “Snapshot export” (phase 2)

If needed later, add a minimal “Download snapshot”:

- Export key charts as PNG (already supported via download buttons).
- Optionally implement an HTML-to-canvas capture for a single summary section (would add a dependency like html2canvas; defer unless requested).

## 6) Technical approach

### Keep the stack

- Static HTML/CSS/JS under [`hand_sharepoint_dashboard/`](hand_sharepoint_dashboard/index.html:1)
- Bootstrap 5 for layout
- Chart.js for charts
- Leaflet + markercluster for map
- Tabulator for tables

### Code structure refactor (incremental)

The JS is currently a single file [`assets/js/app.js`](hand_sharepoint_dashboard/assets/js/app.js:1). For maintainability, consider splitting (phase 2):

- `state.js` (state + helpers)
- `filters.js`
- `charts.js`
- `map.js`
- `tables.js`

For this iteration, keep changes contained to avoid regressions.

## 7) Implementation phases

### Phase 1 (this PR)

- Add root landing page redesign (replace redirect).
- Add “Home” view inside dashboard and reorganize nav/submenus.
- Restructure filter sidebar into accordion + quick actions.
- Update styling to match the snapshot-report feel.

### Phase 2 (follow-up)

- Add trend chart(s) if `assessed_date` coverage is strong.
- Add a “Methodology” page section with definitions and a glossary.
- Consider TopoJSON simplification and lazy-load boundaries for performance.

## 8) Acceptance criteria

- Front page is no longer an immediate redirect and looks professional.
- Dashboard has a clear Home view suitable for briefings.
- Navigation/submenus are reorganized and intuitive.
- Existing interactive capabilities remain: filtering, share links, exports, map click-to-filter.
- No new long-running build steps or server requirements.

