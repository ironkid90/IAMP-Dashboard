/* HAND Mapping Platform
   Static build: loads preprocessed JSON for speed + a map-ready structure.

   Key ideas:
   - Filter state lives in one place.
   - Everything (KPIs, charts, table, map) is derived from the filtered set.
   - Admin filtering uses PCODEs from the boundary join (robust for choropleths + click-to-filter).
*/

(() => {
  "use strict";

  // ---------- tiny helpers ----------
  const $ = (id) => document.getElementById(id);
  const fmtInt = new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 });
  const fmtPct = new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 });

  // Brand palette (HAND) – matched to the org logo
  const BRAND = {
    primary: "#2868a8",
    navy: "#1a4f8c",
    accent: "#e8832a",
    slate: "#64748b",
    noData: "#e2e8f0",
  };

  function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }

  function safeLower(s){
    return (s ?? "").toString().trim().toLowerCase();
  }

  function isBlank(s){
    const t = safeLower(s);
    return !t || t === "-" || t === "—" || t === "–" || t === "na" || t === "n/a" || t === "null" || t === "none";
  }

  function downloadText(filename, text){
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function toCSV(rows, columns){
    const esc = (v) => {
      if (v === null || v === undefined) return "";
      const s = String(v);
      if (s.includes('"') || s.includes(",") || s.includes("\n")) return `"${s.replaceAll('"', '""')}"`;
      return s;
    };
    const header = columns.map(c => esc(c.title)).join(",");
    const lines = rows.map(r => columns.map(c => esc(c.get(r))).join(","));
    return [header, ...lines].join("\n");
  }

  function setLoading(on, detail){
    const overlay = $("loadingOverlay");
    const detailEl = $("loadingDetail");
    if (!overlay) return;
    if (detailEl && detail) detailEl.textContent = detail;
    overlay.classList.toggle("d-none", !on);
  }

  function setHealth({ lastLoad, lastSuccess, errorCount, message }){
    $("healthLastLoad").textContent = lastLoad ?? "—";
    $("healthLastSuccess").textContent = lastSuccess ?? "—";
    $("healthErrors").textContent = String(errorCount ?? 0);
    $("healthMessage").textContent = message ?? "—";
  }

  function nowLocalString(){
    return new Date().toLocaleString();
  }

  function buildLabelFromMap(code, map){
    return map.get(code) || code;
  }

  // ---------- state ----------
  const state = {
    meta: null,
    raw: [],
    filtered: [],
    lookups: {
      adm1NameByCode: new Map(),
      adm2NameByCode: new Map(),
      adm3NameByCode: new Map(),
    },
    boundaries: {
      adm1: null,
      adm2: null,
      adm3: null,
      adm3_full: null,
    },
    filters: {
      search: "",
      gov: "all",
      district: "all",
      cadaster: "all",
      siteStatus: "all",
      phoneStatus: "all",
      qc: "all",
    },
    charts: {},
    table: null,
    qcMiniTable: null,
    map: {
      map: null,
      base: null,
      markerLegendHTML: "",
      cluster: null,
      plain: null,
      boundaryLayer: null,
      choroplethLayer: null,
      selectionLayer: null,
      selectedFeatureId: null,
      lastLevel: "adm2",
    },
  };

  // ---------- URL filters ----------
  function readFiltersFromURL(){
    const p = new URLSearchParams(window.location.search);
    const f = {};
    if (p.has("s")) f.search = p.get("s") || "";
    if (p.has("g")) f.gov = p.get("g") || "all";
    if (p.has("d")) f.district = p.get("d") || "all";
    if (p.has("c")) f.cadaster = p.get("c") || "all";
    if (p.has("ss")) f.siteStatus = p.get("ss") || "all";
    if (p.has("ps")) f.phoneStatus = p.get("ps") || "all";
    if (p.has("qc")) f.qc = p.get("qc") || "all";
    return f;
  }

  function buildShareURL(){
    const p = new URLSearchParams();
    if (state.filters.search) p.set("s", state.filters.search);
    if (state.filters.gov !== "all") p.set("g", state.filters.gov);
    if (state.filters.district !== "all") p.set("d", state.filters.district);
    if (state.filters.cadaster !== "all") p.set("c", state.filters.cadaster);
    if (state.filters.siteStatus !== "all") p.set("ss", state.filters.siteStatus);
    if (state.filters.phoneStatus !== "all") p.set("ps", state.filters.phoneStatus);
    if (state.filters.qc !== "all") p.set("qc", state.filters.qc);

    const url = new URL(window.location.href);
    url.search = p.toString();
    return url.toString();
  }

  async function copyShareLink(){
    const url = buildShareURL();
    try{
      await navigator.clipboard.writeText(url);
      toast(`Copied share link ✓`);
    }catch{
      downloadText("share_link.txt", url);
      toast(`Clipboard blocked – downloaded link instead`);
    }
  }

  // ---------- Toast ----------
  let toastTimer = null;
  function toast(msg){
    let node = document.querySelector(".app-toast");
    if (!node){
      node = document.createElement("div");
      node.className = "app-toast";
      node.style.position = "fixed";
      node.style.bottom = "16px";
      node.style.right = "16px";
      node.style.zIndex = "3000";
      node.style.padding = "10px 12px";
      node.style.borderRadius = "12px";
      node.style.border = "1px solid var(--app-border)";
      node.style.background = "var(--app-card)";
      node.style.boxShadow = "var(--app-shadow)";
      node.style.fontSize = "13px";
      document.body.appendChild(node);
    }
    node.textContent = msg;
    node.style.display = "block";
    if (toastTimer) window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => { node.style.display = "none"; }, 2200);
  }

  // ---------- data load ----------
  async function loadJSON(url){
    // Allow browser revalidation/caching (ETag) while still pulling fresh data when available.
    const res = await fetch(url, { cache: "no-cache" });
    if (!res.ok) throw new Error(`${url}: HTTP ${res.status}`);
    return await res.json();
  }

  function setDatasetBadge(){
    const badge = $("datasetBadge");
    const updated = $("lastUpdatedText");
    if (!state.meta){
      badge.textContent = "No data";
      updated.textContent = "—";
      return;
    }
    const c = state.meta.counts || {};

    const total = c.records || state.raw.length;
    const withCoords = c.with_coords ?? null;
    const qcAny = c.qc_any_issue ?? null;

    badge.textContent = `${fmtInt.format(total)} sites`;

    // Friendly last updated label
    const gen = state.meta.generated_at_utc;
    let updatedLabel = "—";
    if (gen){
      const d = new Date(gen);
      updatedLabel = isNaN(d.getTime()) ? `Updated ${gen}` : `Updated ${d.toLocaleString()}`;
    }
    updated.textContent = updatedLabel;

    // Data panel meta text
    const parts = [];
    if (gen) parts.push(`Generated: ${gen} (UTC)`);
    if (withCoords !== null) parts.push(`With coordinates: ${fmtInt.format(withCoords)} (${fmtPct.format((withCoords/Math.max(1,total))*100)}%)`);
    if (qcAny !== null) parts.push(`QC flagged: ${fmtInt.format(qcAny)} (${fmtPct.format((qcAny/Math.max(1,total))*100)}%)`);

    const sf = state.meta.source_files || {};
    const sources = [];
    if (sf.assessment_csv) sources.push(`Assessment: ${sf.assessment_csv}`);
    if (sf.master_xlsx) sources.push(`Site list: ${sf.master_xlsx}`);
    if (sf.boundaries_zip) sources.push(`Boundaries: ${sf.boundaries_zip}`);

    $("dataMetaText").textContent = [...parts, ...(sources.length ? ["Sources: " + sources.join(" • ")] : [])].join("\n");
    const hint = $("dataSourceHint");
    if (hint){
      hint.textContent = (state.meta._loaded_via === "bundle") ? "Bundled JSON" : "API: /api/data";
    }
  }

  function buildLookups(){
    // Admin1
    if (state.boundaries.adm1){
      for (const f of state.boundaries.adm1.features || []){
        const p = f.properties || {};
        if (p.adm1_pcode) state.lookups.adm1NameByCode.set(String(p.adm1_pcode), p.adm1_name || String(p.adm1_pcode));
      }
    }
    // Admin2
    if (state.boundaries.adm2){
      for (const f of state.boundaries.adm2.features || []){
        const p = f.properties || {};
        if (p.adm2_pcode) state.lookups.adm2NameByCode.set(String(p.adm2_pcode), p.adm2_name || String(p.adm2_pcode));
      }
    }
    // Admin3 (subset)
    if (state.boundaries.adm3){
      for (const f of state.boundaries.adm3.features || []){
        const p = f.properties || {};
        if (p.adm3_pcode) state.lookups.adm3NameByCode.set(String(p.adm3_pcode), p.adm3_name || String(p.adm3_pcode));
      }
    }
  }

  async function loadAllData(){
    setLoading(true, "Fetching JSON");
    setHealth({ lastLoad: nowLocalString(), lastSuccess: null, errorCount: 0, message: "Loading…" });

    try{
      // Prefer the serverless endpoint (Vercel). Fall back to the bundled JSON for offline/static hosting.
      let sitesData = null;
      let loadedVia = "api";
      try{
        sitesData = await loadJSON("/api/data");
      }catch{
        loadedVia = "bundle";
        sitesData = await loadJSON("assets/data/sites.json");
      }
      state.meta = sitesData.meta || null;
      if (state.meta) state.meta._loaded_via = loadedVia;
      state.raw = Array.isArray(sitesData.sites) ? sitesData.sites : [];

      // boundaries
      setLoading(true, "Loading boundaries");
      const [adm1, adm2, adm3, adm3full] = await Promise.all([
        loadJSON("assets/data/boundaries_admin1.geojson"),
        loadJSON("assets/data/boundaries_admin2.geojson"),
        loadJSON("assets/data/boundaries_admin3_subset.geojson"),
        loadJSON("assets/data/boundaries_admin3_full.geojson"),
      ]);

      state.boundaries.adm1 = adm1;
      state.boundaries.adm2 = adm2;
      state.boundaries.adm3 = adm3;
      state.boundaries.adm3_full = adm3full;

      buildLookups();
      setDatasetBadge();

      // Pull any filter values from URL (but only after lookups exist).
      const urlFilters = readFiltersFromURL();
      state.filters = { ...state.filters, ...urlFilters };

      initUIOnce();
      applyFilters();

      setHealth({ lastLoad: nowLocalString(), lastSuccess: nowLocalString(), errorCount: 0, message: "Loaded ✓" });
      setLoading(false);
    }catch(err){
      console.error(err);
      setLoading(false);
      setHealth({ lastLoad: nowLocalString(), lastSuccess: null, errorCount: 1, message: `Failed to load: ${err.message}` });
      $("datasetBadge").textContent = "Load failed";
      $("lastUpdatedText").textContent = "—";
      toast("Data load failed – see console");
    }
  }

  // ---------- filter option builders ----------
  function uniqueSorted(arr){
    return [...new Set(arr.filter(v => !isBlank(v)).map(v => String(v)))].sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
  }

  function populateSelect(selectEl, items, selectedValue, { allLabel="All", allValue="all", labelMap=null } = {}){
    const cur = selectedValue ?? allValue;
    selectEl.innerHTML = "";

    const optAll = document.createElement("option");
    optAll.value = allValue;
    optAll.textContent = allLabel;
    selectEl.appendChild(optAll);

    for (const v of items){
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = labelMap ? (labelMap.get(v) || v) : v;
      selectEl.appendChild(opt);
    }

    // If current selection is not available anymore, fall back to all.
    const exists = [...selectEl.options].some(o => o.value === cur);
    selectEl.value = exists ? cur : allValue;
  }

  function refreshAdminDropdowns(){
    const f = state.filters;

    // Use raw data for building cascading options based on current filter state
    const sitesForGov = state.raw;
    const govCodes = uniqueSorted(sitesForGov.map(s => s.adm1_pcode));
    populateSelect($("govFilter"), govCodes, f.gov, { allLabel: "All governorates", labelMap: state.lookups.adm1NameByCode });

    // District options cascade from governorate selection
    const sitesForDistrict = (f.gov !== "all") ? state.raw.filter(s => String(s.adm1_pcode) === String(f.gov)) : state.raw;
    const distCodes = uniqueSorted(sitesForDistrict.map(s => s.adm2_pcode));
    populateSelect($("districtFilter"), distCodes, f.district, { allLabel: "All districts", labelMap: state.lookups.adm2NameByCode });

    // Cadaster options cascade from both governorate and district selections
    const sitesForCadaster = (() => {
      let x = state.raw;
      if (f.gov !== "all") x = x.filter(s => String(s.adm1_pcode) === String(f.gov));
      if (f.district !== "all") x = x.filter(s => String(s.adm2_pcode) === String(f.district));
      return x;
    })();
    const cadCodes = uniqueSorted(sitesForCadaster.map(s => s.adm3_pcode));
    populateSelect($("cadasterFilter"), cadCodes, f.cadaster, { allLabel: "All cadasters", labelMap: state.lookups.adm3NameByCode });
  }

  function refreshOtherDropdowns(){
    const f = state.filters;
    // Site status options: include "Not recorded"
    const statusVals = uniqueSorted(state.raw.map(s => s.site_status || "Not recorded"));
    populateSelect($("siteStatusFilter"), statusVals, f.siteStatus, { allLabel: "All site statuses" });

    // Phone status options: include "Not assessed"
    const phoneVals = uniqueSorted(state.raw.map(s => s.phone_status || "Not assessed"));
    populateSelect($("phoneStatusFilter"), phoneVals, f.phoneStatus, { allLabel: "All phone statuses" });

    // QC options
    const qcSelect = $("qcFilter");
    qcSelect.innerHTML = "";
    const opts = [
      { v: "all", t: "All records" },
      { v: "any", t: "QC: Any issue" },
      { v: "missing_coords", t: "QC: Missing coordinates" },
      { v: "missing_site_status", t: "QC: Missing site status (assessed)" },
      { v: "missing_totals_active", t: "QC: Active sites missing totals" },
      { v: "inactive_missing_date", t: "QC: Inactive/demolished missing date" },
    ];
    for (const o of opts){
      const opt = document.createElement("option");
      opt.value = o.v;
      opt.textContent = o.t;
      qcSelect.appendChild(opt);
    }
    qcSelect.value = opts.some(o => o.v === f.qc) ? f.qc : "all";
  }

  // ---------- filtering ----------
  function applyFilters(){
    // Capture current UI selections BEFORE rebuilding dropdowns so
    // user-driven changes (especially cadaster) are not overwritten.
    state.filters.search = $("searchInput").value.trim();
    state.filters.gov = $("govFilter").value;
    state.filters.district = $("districtFilter").value;
    state.filters.cadaster = $("cadasterFilter").value;
    state.filters.siteStatus = $("siteStatusFilter").value;
    state.filters.phoneStatus = $("phoneStatusFilter").value;
    state.filters.qc = $("qcFilter").value;

    // Rebuild cascading dropdowns (may invalidate child selections when
    // a parent level changed, e.g. district/cadaster after gov change).
    refreshAdminDropdowns();
    refreshOtherDropdowns();

    // Re-read after refresh in case a selection was invalidated by cascade.
    state.filters.gov = $("govFilter").value;
    state.filters.district = $("districtFilter").value;
    state.filters.cadaster = $("cadasterFilter").value;
    state.filters.siteStatus = $("siteStatusFilter").value;
    state.filters.phoneStatus = $("phoneStatusFilter").value;
    state.filters.qc = $("qcFilter").value;

    const f = state.filters;

    let arr = state.raw.slice();

    if (f.search){
      const q = safeLower(f.search);
      arr = arr.filter(s => {
        const hay = [
          s.pcode, s.name, s.local_name,
          buildLabelFromMap(s.adm3_pcode, state.lookups.adm3NameByCode),
          buildLabelFromMap(s.adm2_pcode, state.lookups.adm2NameByCode),
        ].filter(Boolean).map(x => String(x).toLowerCase()).join(" | ");
        return hay.includes(q);
      });
    }

    if (f.gov !== "all") arr = arr.filter(s => String(s.adm1_pcode) === f.gov);
    if (f.district !== "all") arr = arr.filter(s => String(s.adm2_pcode) === f.district);
    if (f.cadaster !== "all") arr = arr.filter(s => String(s.adm3_pcode) === f.cadaster);

    if (f.siteStatus !== "all"){
      arr = arr.filter(s => (s.site_status || "Not recorded") === f.siteStatus);
    }

    if (f.phoneStatus !== "all"){
      arr = arr.filter(s => (s.phone_status || "Not assessed") === f.phoneStatus);
    }

    if (f.qc !== "all"){
      if (f.qc === "any"){
        arr = arr.filter(s => s.qc?.qc_any_issue);
      } else if (f.qc === "missing_coords"){
        arr = arr.filter(s => s.qc?.missing_coords);
      } else if (f.qc === "missing_site_status"){
        arr = arr.filter(s => s.qc?.missing_site_status_when_assessed);
      } else if (f.qc === "missing_totals_active"){
        arr = arr.filter(s => s.qc?.missing_totals_active);
      } else if (f.qc === "inactive_missing_date"){
        arr = arr.filter(s => s.qc?.inactive_missing_date);
      }
    }

    state.filtered = arr;

    updateKPIs();
    updateCharts();
    updateTables();
    updateMap();

    // Update note text
    updateAssessmentNote();
  }

  function resetFilters(){
    state.filters = {
      search: "",
      gov: "all",
      district: "all",
      cadaster: "all",
      siteStatus: "all",
      phoneStatus: "all",
      qc: "all",
    };
    $("searchInput").value = "";
    refreshAdminDropdowns();
    refreshOtherDropdowns();
    applyFilters();
    toast("Filters reset");
  }

  // ---------- KPIs ----------
  function computeAssessed(s){
    // In this build: phone_status present => assessed
    return !isBlank(s.phone_status);
  }

  function computeActive(s){
    return (s.site_status || "").trim() === "Active";
  }

  function sumMetric(arr, getter){
    let sum = 0;
    for (const s of arr){
      const v = getter(s);
      if (typeof v === "number" && Number.isFinite(v)) sum += v;
    }
    return sum;
  }

  function updateKPIs(){
    const total = state.filtered.length;
    const assessed = state.filtered.filter(computeAssessed).length;
    const activeSites = state.filtered.filter(computeActive);

    const qcAny = state.filtered.filter(s => s.qc?.qc_any_issue).length;

    const hh = sumMetric(activeSites, s => s.metrics?.households_total ?? null);
    const ind = sumMetric(activeSites, s => s.metrics?.individuals_total ?? null);
    const struct = sumMetric(activeSites, s => s.metrics?.structures_total ?? null);
    const lat = sumMetric(activeSites, s => s.metrics?.latrines_total ?? null);

    $("kpiTotal").textContent = fmtInt.format(total);
    $("kpiTotalSub").textContent = (total === state.raw.length) ? "All records" : `Filtered from ${fmtInt.format(state.raw.length)}`;

    $("kpiAssessed").textContent = fmtInt.format(assessed);
    $("kpiAssessedSub").textContent = total ? `${fmtPct.format(assessed * 100 / total)}% of filtered` : "—";

    $("kpiActive").textContent = fmtInt.format(activeSites.length);
    $("kpiActiveSub").textContent = total ? `${fmtPct.format(activeSites.length * 100 / total)}% of filtered` : "—";

    $("kpiQC").textContent = fmtInt.format(qcAny);
    $("kpiQCSub").textContent = total ? `${fmtPct.format(qcAny * 100 / total)}% QC any-issue` : "—";

    $("kpiHH").textContent = fmtInt.format(hh);
    $("kpiHHSub").textContent = `Sum across active sites`;

    $("kpiIND").textContent = fmtInt.format(ind);
    $("kpiINDSub").textContent = `Sum across active sites`;

    $("kpiStruct").textContent = fmtInt.format(struct);
    $("kpiStructSub").textContent = `Sum across active sites`;

    $("kpiLat").textContent = fmtInt.format(lat);
    $("kpiLatSub").textContent = `Sum across active sites`;

    // Update filter summary banner
    updateFilterSummary(total);
  }

  function updateFilterSummary(total){
    const el = $("filterSummaryText");
    if (!el) return;
    const f = state.filters;
    const parts = [];
    if (f.gov !== "all") parts.push(buildLabelFromMap(f.gov, state.lookups.adm1NameByCode));
    if (f.district !== "all") parts.push(buildLabelFromMap(f.district, state.lookups.adm2NameByCode));
    if (f.cadaster !== "all") parts.push(buildLabelFromMap(f.cadaster, state.lookups.adm3NameByCode));
    if (f.siteStatus !== "all") parts.push(f.siteStatus);
    if (f.phoneStatus !== "all") parts.push(f.phoneStatus);
    if (f.qc !== "all") parts.push("QC filter active");
    if (f.search) parts.push(`"${f.search}"`);

    if (!parts.length){
      el.textContent = `Showing all ${fmtInt.format(total)} records. Use the sidebar filters to refine.`;
    } else {
      el.textContent = `Filtered: ${parts.join(" · ")} — ${fmtInt.format(total)} records`;
    }
  }

  function updateAssessmentNote(){
    const total = state.filtered.length;
    const assessed = state.filtered.filter(computeAssessed).length;
    const pct = total ? (assessed * 100 / total) : 0;
    $("assessmentNote").textContent = total
      ? `${fmtInt.format(assessed)} of ${fmtInt.format(total)} assessed (${fmtPct.format(pct)}%). Click charts to filter.`
      : "No records in current filter.";
  }

  // ---------- charts ----------
  function chartDefaults(){
    const dark = document.documentElement.getAttribute("data-bs-theme") === "dark";
    const tick = dark ? "#cbd5e1" : "#475569";
    const grid = dark ? "rgba(148,163,184,0.16)" : "rgba(15,23,42,0.08)";
    const tooltipBg = dark ? "rgba(15,23,42,0.95)" : "rgba(255,255,255,0.98)";
    const tooltipBorder = dark ? "rgba(148,163,184,0.25)" : "rgba(15,23,42,0.12)";

    return {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: tooltipBg,
          borderColor: tooltipBorder,
          borderWidth: 1,
          titleColor: dark ? "#e5e7eb" : "#0f172a",
          bodyColor: dark ? "#e5e7eb" : "#0f172a",
        }
      },
      scales: {
        x: {
          ticks: { color: tick, font: { size: 11 } },
          grid: { color: grid },
          border: { display: false },
        },
        y: {
          ticks: { color: tick, font: { size: 11 } },
          grid: { display: false },
          border: { display: false },
        }
      }
    };
  }

  function destroyChart(id){
    const c = state.charts[id];
    if (c){ c.destroy(); state.charts[id] = null; }
  }

  function rebuildCharts(){
    // Recreate chart instances (needed for theme toggles)
    const keep = { ...state.charts };
    for (const k of Object.keys(keep)) destroyChart(k);
    initCharts();
    updateCharts();
  }

  function initCharts(){
    // Assessment progress (stacked single bar)
    const ctxA = $("chartAssessment").getContext("2d");
    state.charts.assessment = new Chart(ctxA, {
      type: "bar",
      data: {
        labels: ["Progress"],
        datasets: [
          { label: "Assessed", data: [0], backgroundColor: BRAND.primary },
          { label: "Not assessed", data: [0], backgroundColor: BRAND.noData },
        ]
      },
      options: {
        ...chartDefaults(),
        indexAxis: "y",
        scales: {
          x: { ...chartDefaults().scales.x, stacked: true, ticks: { ...chartDefaults().scales.x.ticks, callback: (v) => fmtInt.format(v) } },
          y: { ...chartDefaults().scales.y, stacked: true },
        },
        plugins: {
          ...chartDefaults().plugins,
          legend: { display: true, position: "bottom", labels: { boxWidth: 12, color: chartDefaults().scales.x.ticks.color } }
        },
        onClick: () => { /* no-op */ }
      }
    });

    // Site status mix
    const ctxS = $("chartSiteStatus").getContext("2d");
    state.charts.siteStatus = new Chart(ctxS, {
      type: "bar",
      data: { labels: [], datasets: [{ label: "Sites", data: [], backgroundColor: BRAND.primary }] },
      options: {
        ...chartDefaults(),
        indexAxis: "y",
        onClick: (evt, elements, chart) => {
          if (!elements?.length) return;
          const idx = elements[0].index;
          const label = chart.data.labels[idx];
          $("siteStatusFilter").value = label;
          applyFilters();
        }
      }
    });

    // Phone outcomes
    const ctxP = $("chartPhoneOutcomes").getContext("2d");
    state.charts.phone = new Chart(ctxP, {
      type: "bar",
      data: { labels: [], datasets: [{ label: "Records", data: [], backgroundColor: BRAND.accent }] },
      options: {
        ...chartDefaults(),
        indexAxis: "y",
        onClick: (evt, elements, chart) => {
          if (!elements?.length) return;
          const idx = elements[0].index;
          const label = chart.data.labels[idx];
          $("phoneStatusFilter").value = label;
          applyFilters();
        }
      }
    });

    // Structures mix (stacked single bar)
    const ctxSt = $("chartStructures").getContext("2d");
    state.charts.structures = new Chart(ctxSt, {
      type: "bar",
      data: {
        labels: ["Active structures"],
        datasets: [
          { label: "Tents", data: [0], backgroundColor: BRAND.accent },
          { label: "Self-built (non-concrete roof)", data: [0], backgroundColor: "#93c5fd" },
          { label: "Prefab", data: [0], backgroundColor: "#60a5fa" },
          { label: "Self-built (concrete roof)", data: [0], backgroundColor: BRAND.primary },
        ]
      },
      options: {
        ...chartDefaults(),
        indexAxis: "y",
        scales: {
          x: { ...chartDefaults().scales.x, stacked: true, ticks: { ...chartDefaults().scales.x.ticks, callback: (v) => fmtInt.format(v) } },
          y: { ...chartDefaults().scales.y, stacked: true },
        },
        plugins: {
          ...chartDefaults().plugins,
          legend: { display: true, position: "bottom", labels: { boxWidth: 12, color: chartDefaults().scales.x.ticks.color } }
        },
      }
    });

    // Top cadaster by active individuals
    const ctxT = $("chartTopCadaster").getContext("2d");
    state.charts.topCad = new Chart(ctxT, {
      type: "bar",
      data: { labels: [], datasets: [{ label: "Individuals (active)", data: [], backgroundColor: BRAND.primary }] },
      options: {
        ...chartDefaults(),
        indexAxis: "y",
        onClick: (evt, elements, chart) => {
          if (!elements?.length) return;
          const idx = elements[0].index;
          const code = chart.$codes?.[idx];
          if (!code) return;
          $("cadasterFilter").value = code;
          applyFilters();
        }
      }
    });

    // QC by type
    const ctxQC = $("chartQC").getContext("2d");
    state.charts.qc = new Chart(ctxQC, {
      type: "bar",
      data: { labels: [], datasets: [{ label: "Records", data: [], backgroundColor: "#7c3aed" }] },
      options: {
        ...chartDefaults(),
        indexAxis: "y",
        onClick: (evt, elements, chart) => {
          if (!elements?.length) return;
          const idx = elements[0].index;
          const key = chart.$keys?.[idx];
          if (!key) return;
          $("qcFilter").value = key;
          applyFilters();
        }
      }
    });
  }

  function updateCharts(){
    if (!state.charts.assessment) return;

    const total = state.filtered.length;
    const assessed = state.filtered.filter(computeAssessed).length;
    const notAssessed = total - assessed;

    // Assessment stacked bar
    state.charts.assessment.data.datasets[0].data = [assessed];
    state.charts.assessment.data.datasets[1].data = [notAssessed];
    state.charts.assessment.update();

    // Site status counts
    const statusCounts = new Map();
    for (const s of state.filtered){
      const k = s.site_status || "Not recorded";
      statusCounts.set(k, (statusCounts.get(k) || 0) + 1);
    }
    const statusLabels = [...statusCounts.keys()].sort((a,b)=>statusCounts.get(b)-statusCounts.get(a));
    state.charts.siteStatus.data.labels = statusLabels;
    state.charts.siteStatus.data.datasets[0].data = statusLabels.map(l => statusCounts.get(l));
    state.charts.siteStatus.update();

    // Phone outcomes (assessed only)
    const phoneCounts = new Map();
    for (const s of state.filtered){
      if (!computeAssessed(s)) continue;
      const k = s.phone_status || "Unknown";
      phoneCounts.set(k, (phoneCounts.get(k) || 0) + 1);
    }
    const phoneLabels = [...phoneCounts.keys()].sort((a,b)=>phoneCounts.get(b)-phoneCounts.get(a));
    state.charts.phone.data.labels = phoneLabels;
    state.charts.phone.data.datasets[0].data = phoneLabels.map(l => phoneCounts.get(l));
    state.charts.phone.update();

    // Structures mix (active)
    const activeSites = state.filtered.filter(computeActive);
    const tents = sumMetric(activeSites, s => s.metrics?.tents ?? null);
    const nonc = sumMetric(activeSites, s => s.metrics?.selfbuilt_nonconcrete ?? null);
    const prefab = sumMetric(activeSites, s => s.metrics?.prefab ?? null);
    const conc = sumMetric(activeSites, s => s.metrics?.selfbuilt_concrete ?? null);
    state.charts.structures.data.datasets[0].data = [tents];
    state.charts.structures.data.datasets[1].data = [nonc];
    state.charts.structures.data.datasets[2].data = [prefab];
    state.charts.structures.data.datasets[3].data = [conc];
    state.charts.structures.update();

    // Top cadasters by active individuals
    const cadAgg = new Map(); // code -> sum
    for (const s of activeSites){
      const code = String(s.adm3_pcode);
      const ind = s.metrics?.individuals_total;
      if (typeof ind === "number" && Number.isFinite(ind)){
        cadAgg.set(code, (cadAgg.get(code) || 0) + ind);
      }
    }
    const top = [...cadAgg.entries()].sort((a,b)=>b[1]-a[1]).slice(0,10);
    const labels = top.map(([code]) => buildLabelFromMap(code, state.lookups.adm3NameByCode));
    const values = top.map(([,v]) => v);
    state.charts.topCad.data.labels = labels;
    state.charts.topCad.data.datasets[0].data = values;
    state.charts.topCad.$codes = top.map(([code]) => code);
    state.charts.topCad.update();

    // QC by type
    const qcKeys = [
      { key: "missing_coords", label: "Missing coordinates" },
      { key: "missing_site_status", label: "Missing site status (assessed)" },
      { key: "missing_totals_active", label: "Active sites missing totals" },
      { key: "inactive_missing_date", label: "Inactive/demolished missing date" },
    ];
    const qcCounts = qcKeys.map(k => {
      let n = 0;
      for (const s of state.filtered){
        if (k.key === "missing_coords" && s.qc?.missing_coords) n++;
        if (k.key === "missing_site_status" && s.qc?.missing_site_status_when_assessed) n++;
        if (k.key === "missing_totals_active" && s.qc?.missing_totals_active) n++;
        if (k.key === "inactive_missing_date" && s.qc?.inactive_missing_date) n++;
      }
      return n;
    });

    // Sort by count
    const qcCombined = qcKeys.map((k, i) => ({ ...k, count: qcCounts[i] }))
      .filter(x => x.count > 0)
      .sort((a,b) => b.count - a.count);

    state.charts.qc.data.labels = qcCombined.map(x => x.label);
    state.charts.qc.data.datasets[0].data = qcCombined.map(x => x.count);
    // Use keys as values for filter select
    state.charts.qc.$keys = qcCombined.map(x => x.key);
    state.charts.qc.update();
  }

  // ---------- QC accordion ----------
  function initQCAccordion(){
    const acc = $("qcAccordion");
    if (!acc) return;

    const items = [
      {
        id: "qc1",
        title: "Missing coordinates",
        body: "These sites cannot be shown as point markers. Fix by ensuring each settlement PCode has a valid lat/lon (ideally from ArcGIS export) and regenerating the dashboard JSON."
      },
      {
        id: "qc2",
        title: "Missing site status (assessed)",
        body: "A phone assessment exists but site status is blank. Confirm the site status (Active / Inactive / Fully Demolished) and update the source sheet."
      },
      {
        id: "qc3",
        title: "Active sites missing totals",
        body: "Active sites should typically have totals (individuals/households/structures). Check whether totals are missing or recorded in split fields only."
      },
      {
        id: "qc4",
        title: "Inactive/demolished missing date",
        body: "If a site is marked Inactive or Fully Demolished, capture the relevant date (or at least month/year) to support timeline reporting."
      },
    ];

    acc.innerHTML = items.map((it, idx) => `
      <div class="accordion-item">
        <h2 class="accordion-header">
          <button class="accordion-button ${idx===0 ? "" : "collapsed"}" type="button" data-bs-toggle="collapse" data-bs-target="#${it.id}">
            ${it.title}
          </button>
        </h2>
        <div id="${it.id}" class="accordion-collapse collapse ${idx===0 ? "show" : ""}" data-bs-parent="#qcAccordion">
          <div class="accordion-body small text-body-secondary">
            ${it.body}
          </div>
        </div>
      </div>
    `).join("");
  }

  // ---------- tables ----------
  function initTables(){
    // Main records table (redacted)
    const cols = [
      { title: "PCode", field: "pcode", width: 140, headerSortStartingDir: "asc" },
      { title: "Name", field: "name", minWidth: 180 },
      { title: "District", field: "adm2", minWidth: 120 },
      { title: "Cadaster", field: "adm3", minWidth: 160 },
      { title: "Site Status", field: "site_status", width: 130 },
      { title: "Phone Status", field: "phone_status", width: 220 },
      { title: "Active IND", field: "ind", hozAlign: "right", width: 110 },
      { title: "Active HH", field: "hh", hozAlign: "right", width: 110 },
      { title: "Active Structures", field: "struct", hozAlign: "right", width: 140 },
      { title: "QC Any", field: "qc_any", width: 80, hozAlign: "center" },
      { title: "QC Count", field: "qc_count", width: 90, hozAlign: "right" },
      { title: "Has coords", field: "has_coords", width: 110, hozAlign: "center" },
    ];

    // Add a header menu (column picker) on the first column
    const makeColumnToggleMenu = () => {
      const menu = [];
      for (const c of state.table.getColumns()){
        const def = c.getDefinition();
        const title = def.title || def.field;
        menu.push({
          label: `${c.isVisible() ? "✓" : " "} ${title}`,
          action: function(e){
            e.stopPropagation();
            c.toggle();
          }
        });
      }
      return menu;
    };
    cols[0].headerMenu = makeColumnToggleMenu;

    state.table = new Tabulator("#recordsTable", {
      data: [],
      layout: "fitDataStretch",
      height: "620px",
      pagination: "local",
      paginationSize: 50,
      paginationSizeSelector: [25, 50, 100, 250],
      columns: cols,
    });

    // QC mini table
    state.qcMiniTable = new Tabulator("#qcMiniTable", {
      data: [],
      layout: "fitDataStretch",
      height: "340px",
      pagination: "local",
      paginationSize: 20,
      columns: [
        { title: "PCode", field: "pcode", width: 140 },
        { title: "Name", field: "name", minWidth: 180 },
        { title: "QC Count", field: "qc_count", width: 90, hozAlign: "right" },
        { title: "Issues", field: "qc_list", minWidth: 260 },
      ]
    });

    $("downloadTableBtn").addEventListener("click", () => {
      state.table.download("csv", "hand_records_filtered.csv");
    });

    $("downloadQCBtn").addEventListener("click", () => {
      const qcRows = buildTableRows(state.filtered.filter(s => s.qc?.qc_any_issue));
      const csv = toCSV(qcRows, tableCSVColumns());
      downloadText("hand_qc_records.csv", csv);
    });
  }

  function buildTableRows(sites){
    return sites.map(s => {
      const active = computeActive(s);
      return {
        pcode: s.pcode,
        name: s.name || "",
        adm2: buildLabelFromMap(String(s.adm2_pcode), state.lookups.adm2NameByCode),
        adm3: buildLabelFromMap(String(s.adm3_pcode), state.lookups.adm3NameByCode),
        site_status: s.site_status || "Not recorded",
        phone_status: s.phone_status || "Not assessed",
        ind: active ? (s.metrics?.individuals_total ?? null) : null,
        hh: active ? (s.metrics?.households_total ?? null) : null,
        struct: active ? (s.metrics?.structures_total ?? null) : null,
        qc_any: s.qc?.qc_any_issue ? "Yes" : "No",
        qc_count: s.qc?.qc_issue_count ?? 0,
        qc_list: qcIssueList(s),
        has_coords: (typeof s.lat === "number" && typeof s.lon === "number") ? "Yes" : "No",
      };
    });
  }

  function qcIssueList(s){
    const parts = [];
    if (s.qc?.missing_coords) parts.push("Missing coords");
    if (s.qc?.missing_site_status_when_assessed) parts.push("Missing site status");
    if (s.qc?.missing_totals_active) parts.push("Missing totals (active)");
    if (s.qc?.inactive_missing_date) parts.push("Missing inactive date");
    return parts.join("; ");
  }

  function tableCSVColumns(){
    return [
      { title: "PCode", get: r => r.pcode },
      { title: "Name", get: r => r.name },
      { title: "District", get: r => r.adm2 },
      { title: "Cadaster", get: r => r.adm3 },
      { title: "Site Status", get: r => r.site_status },
      { title: "Phone Status", get: r => r.phone_status },
      { title: "Active IND", get: r => r.ind ?? "" },
      { title: "Active HH", get: r => r.hh ?? "" },
      { title: "Active Structures", get: r => r.struct ?? "" },
      { title: "QC Any", get: r => r.qc_any },
      { title: "QC Count", get: r => r.qc_count },
      { title: "QC Issues", get: r => r.qc_list },
      { title: "Has coords", get: r => r.has_coords },
    ];
  }

  function updateTables(){
    if (!state.table || !state.qcMiniTable) return;

    const rows = buildTableRows(state.filtered);
    state.table.replaceData(rows);

    const qcRows = rows.filter(r => r.qc_any === "Yes").slice(0, 200);
    state.qcMiniTable.replaceData(qcRows);
  }

  // ---------- Export filtered CSV ----------
  function exportFilteredCSV(){
    const rows = buildTableRows(state.filtered);
    const csv = toCSV(rows, tableCSVColumns());
    downloadText("hand_filtered_export.csv", csv);
  }

  // ---------- map ----------
  function initMap(){
    const mapEl = $("lebanonMap");
    if (!mapEl) return;

    const map = L.map(mapEl, { preferCanvas: true, zoomSnap: 0.5, zoomControl: false });
    state.map.map = map;

    // Keep top corners clean for overlay cards.
    L.control.zoom({ position: "bottomleft" }).addTo(map);
    L.control.scale({ position: "bottomright", imperial: false }).addTo(map);

    // Clean light basemap for public-facing dashboards.
    state.map.base = L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
      subdomains: "abcd",
      maxZoom: 20,
      attribution: "&copy; OpenStreetMap contributors &copy; CARTO"
    }).addTo(map);

    // Marker layers
    state.map.cluster = L.markerClusterGroup({ showCoverageOnHover: false, maxClusterRadius: 40 });
    state.map.plain = L.layerGroup();

    // Default: cluster on
    state.map.cluster.addTo(map);

    // Fit initial bounds
    fitToLebanon();

    // Buttons
    $("fitLebanonBtn").addEventListener("click", () => fitToLebanon());
    $("clearMapBtn").addEventListener("click", () => clearMapSelection());

    $("toggleCluster").addEventListener("change", () => {
      const on = $("toggleCluster").checked;
      if (on){
        map.removeLayer(state.map.plain);
        map.addLayer(state.map.cluster);
      } else {
        map.removeLayer(state.map.cluster);
        map.addLayer(state.map.plain);
      }
      updateMap();
    });

    $("mapColorBy").addEventListener("change", () => updateMap());

    $("choroplethLevel").addEventListener("change", () => updateChoropleth());
    $("choroplethMetric").addEventListener("change", () => updateChoropleth());
    $("toggleChoropleth").addEventListener("change", () => updateChoropleth());

    // When switching tabs, Leaflet needs a size refresh.
    document.querySelectorAll('button[data-bs-toggle="tab"]').forEach(btn => {
      btn.addEventListener("shown.bs.tab", (e) => {
        const target = e.target?.getAttribute("data-bs-target");
        if (target === "#pane-map"){
          setTimeout(() => map.invalidateSize(), 50);
        }
      });
    });
  }

  function fitToLebanon(){
    const map = state.map.map;
    if (!map || !state.boundaries.adm1) return;
    const layer = L.geoJSON(state.boundaries.adm1);
    const b = layer.getBounds();
    if (b.isValid()) map.fitBounds(b.pad(0.03));
  }

  function clearMapSelection(){
    state.map.selectedFeatureId = null;
    if (state.map.selectionLayer){
      state.map.map.removeLayer(state.map.selectionLayer);
      state.map.selectionLayer = null;
    }
    toast("Map selection cleared");
  }

  function updateMap(){
    if (!state.map.map) return;
    updateMapMarkers();
    updateChoropleth();
  }

  function updateMapMarkers(){
    const map = state.map.map;
    if (!map) return;

    const useCluster = $("toggleCluster").checked;
    const layer = useCluster ? state.map.cluster : state.map.plain;
    const other = useCluster ? state.map.plain : state.map.cluster;

    // Clear both layers to be safe.
    layer.clearLayers();
    other.clearLayers();

    const colorBy = $("mapColorBy").value;

    const palette = markerPalette(colorBy, state.filtered);
    const withCoords = state.filtered.filter(s => typeof s.lat === "number" && typeof s.lon === "number");

    for (const s of withCoords){
      const cat = markerCategory(colorBy, s);
      const color = palette.get(cat) || BRAND.slate;

      const m = L.circleMarker([s.lat, s.lon], {
        radius: 4,
        weight: 1,
        color: "rgba(0,0,0,0.25)",
        fillColor: color,
        fillOpacity: 0.85,
      });

      const title = s.name || s.pcode;
      const district = buildLabelFromMap(String(s.adm2_pcode), state.lookups.adm2NameByCode);
      const cadaster = buildLabelFromMap(String(s.adm3_pcode), state.lookups.adm3NameByCode);

      m.bindPopup(`
        <div style="min-width:220px">
          <div style="font-weight:700">${escapeHtml(title)}</div>
          <div style="font-size:12px;color:var(--app-text-muted)">PCode: ${escapeHtml(s.pcode)}</div>
          <hr style="margin:8px 0"/>
          <div style="font-size:12px"><b>District:</b> ${escapeHtml(district)}</div>
          <div style="font-size:12px"><b>Cadaster:</b> ${escapeHtml(cadaster)}</div>
          <div style="font-size:12px"><b>Site status:</b> ${escapeHtml(s.site_status || "Not recorded")}</div>
          <div style="font-size:12px"><b>Phone status:</b> ${escapeHtml(s.phone_status || "Not assessed")}</div>
          <div style="font-size:12px"><b>QC any issue:</b> ${s.qc?.qc_any_issue ? "Yes" : "No"}</div>
        </div>
      `);

      layer.addLayer(m);
    }

    // Update legend for markers
    renderMarkerLegend(palette);

    // Map status
    const missing = state.filtered.length - withCoords.length;
    $("mapStatusText").textContent = `${fmtInt.format(withCoords.length)} sites with coordinates shown. ${missing ? fmtInt.format(missing) + " missing coords." : ""}`;
  }

  function markerCategory(colorBy, s){
    if (colorBy === "single"){
      return "Sites";
    }
    if (colorBy === "site_status"){
      return s.site_status || "Not recorded";
    }
    if (colorBy === "qc_any_issue"){
      return s.qc?.qc_any_issue ? "QC issue" : "No issue";
    }
    if (colorBy === "phone_status"){
      return s.phone_status || "Not assessed";
    }
    return "Other";
  }

  function markerPalette(colorBy, sites){
    if (colorBy === "single"){
      return new Map([["Sites", BRAND.accent]]);
    }
    // Deterministic category ordering
    const cats = uniqueSorted(sites.map(s => markerCategory(colorBy, s)));
    // Pick simple (and readable) colors. Not fancy, just functional.
    const base = [
      "#16a34a", "#0ea5e9", "#f59e0b", "#ef4444",
      "#7c3aed", "#14b8a6", "#a855f7", "#64748b",
      "#84cc16", "#f97316", "#06b6d4", "#d946ef",
    ];
    const m = new Map();
    cats.forEach((c, i) => m.set(c, base[i % base.length]));
    return m;
  }

  function renderMarkerLegend(palette){
    const node = $("mapLegend");
    if (!node) return;
    const rows = [];
    for (const [label, color] of palette.entries()){
      rows.push(`
        <div class="legend-row">
          <span class="legend-swatch" style="background:${color}"></span>
          <span>${escapeHtml(label)}</span>
        </div>
      `);
    }
    const html = rows.join("");
    state.map.markerLegendHTML = html;
    node.innerHTML = html;
  }

  function updateChoropleth(){
    const map = state.map.map;
    if (!map) return;

    // Remove existing layers
    if (state.map.choroplethLayer){
      map.removeLayer(state.map.choroplethLayer);
      state.map.choroplethLayer = null;
    }
    if (state.map.selectionLayer){
      map.removeLayer(state.map.selectionLayer);
      state.map.selectionLayer = null;
    }

    if (!$("toggleChoropleth").checked){
      const node = $("mapLegend");
      if (node) node.innerHTML = state.map.markerLegendHTML || "";
      return;
    }

    const level = $("choroplethLevel").value; // adm2 or adm3
    const metric = $("choroplethMetric").value;

    const fc = (level === "adm2") ? state.boundaries.adm2 : state.boundaries.adm3;
    if (!fc) return;

    const codeProp = (level === "adm2") ? "adm2_pcode" : "adm3_pcode";
    const nameProp = (level === "adm2") ? "adm2_name" : "adm3_name";

    const agg = aggregateBy(level, metric);

    const values = [];
    for (const f of fc.features || []){
      const code = String(f.properties?.[codeProp] ?? "");
      const g = agg.get(code);
      if (!g) continue;
      const v = g.value;
      if (typeof v === "number" && Number.isFinite(v)) values.push(v);
    }

    const scale = buildScale(values, metric);

    const layer = L.geoJSON(fc, {
      style: (feature) => {
        const code = String(feature.properties?.[codeProp] ?? "");
        const g = agg.get(code);
        if (!g){
          return {
            weight: 1,
            color: "rgba(255,255,255,0.85)",
            fillColor: scale.noDataColor,
            fillOpacity: 0.75,
          };
        }
        const v = g.value;
        const fill = scale.color(v);
        return {
          weight: 1,
          color: "rgba(255,255,255,0.85)",
          fillColor: fill,
          fillOpacity: 0.85,
        };
      },
      onEachFeature: (feature, lyr) => {
        const code = String(feature.properties?.[codeProp] ?? "");
        const name = String(feature.properties?.[nameProp] ?? code);
        const g = agg.get(code);

        const displayValue = g ? scale.format(g.value) : scale.noDataLabel;
        const displaySites = g ? fmtInt.format(g.count) : "0";

        lyr.bindTooltip(
          `${escapeHtml(name)}<br/><b>${escapeHtml(scale.label)}:</b> ${escapeHtml(displayValue)}<br/><b>Sites:</b> ${escapeHtml(displaySites)}`,
          { sticky: true }
        );

        lyr.on("click", () => {
          // Highlight clicked polygon (visual selection)
          clearMapSelection();
          state.map.selectionLayer = L.geoJSON(feature, { style: { weight: 3, color: BRAND.accent, fillOpacity: 0.08 } }).addTo(map);

          // Apply filter
          if (level === "adm2"){
            $("districtFilter").value = code;
          } else {
            $("cadasterFilter").value = code;
          }
          applyFilters();
        });
      }
    });

    state.map.choroplethLayer = layer.addTo(map);

    // Add choropleth legend *above* marker legend (simple)
    renderChoroplethLegend(scale);
  }

  function aggregateBy(level, metric){
    const keyField = (level === "adm2") ? "adm2_pcode" : "adm3_pcode";
    const groups = new Map();

    for (const s of state.filtered){
      const code = String(s[keyField] ?? "");
      if (!code) continue;

      if (!groups.has(code)) groups.set(code, { count: 0, assessed: 0, qcAny: 0, indActive: 0, hhActive: 0 });
      const g = groups.get(code);

      g.count += 1;
      if (computeAssessed(s)) g.assessed += 1;
      if (s.qc?.qc_any_issue) g.qcAny += 1;

      if (computeActive(s)){
        const ind = s.metrics?.individuals_total;
        const hh = s.metrics?.households_total;
        if (typeof ind === "number" && Number.isFinite(ind)) g.indActive += ind;
        if (typeof hh === "number" && Number.isFinite(hh)) g.hhActive += hh;
      }
    }

    const out = new Map();
    for (const [code, g] of groups.entries()){
      let value = 0;
      if (metric === "sites_count") value = g.count;
      if (metric === "individuals_active") value = g.indActive;
      if (metric === "households_active") value = g.hhActive;
      if (metric === "qc_rate") value = g.count ? (g.qcAny * 100 / g.count) : 0;
      if (metric === "assessed_rate") value = g.count ? (g.assessed * 100 / g.count) : 0;
      out.set(code, { value, ...g });
    }
    return out;
  }

  function buildScale(values, metric){
  const nonNa = values.filter(v => typeof v === "number" && Number.isFinite(v));

  const labelByMetric = {
    sites_count: "Sites",
    individuals_active: "Individuals (active)",
    households_active: "Households (active)",
    qc_rate: "QC any-issue %",
    assessed_rate: "Assessed %",
  };

  const isPct = (metric === "qc_rate" || metric === "assessed_rate");

  const format = (v) => {
    if (!Number.isFinite(v)) return "—";
    if (isPct) return `${fmtPct.format(v)}%`;
    return fmtInt.format(v);
  };

  // Classic cartographic "Blues" (close to the IOM/DTM look)
  const colors = ["#deebf7", "#9ecae1", "#6baed6", "#3182bd", "#08519c"];

  if (!nonNa.length){
    return {
      min: 0,
      max: 0,
      breaks: [],
      bins: [],
      color: () => BRAND.noData,
      format,
      label: labelByMetric[metric] || metric,
      isPct,
      noDataColor: BRAND.noData,
      noDataLabel: "No sites",
    };
  }

  const sorted = [...nonNa].sort((a,b)=>a-b);
  const min = sorted[0];
  const max = sorted[sorted.length-1];

  // Edge case: everything the same
  if (max === min){
    return {
      min,
      max,
      breaks: [],
      bins: [{ from: min, to: max, color: colors[2] }],
      color: () => colors[2],
      format,
      label: labelByMetric[metric] || metric,
      isPct,
      noDataColor: BRAND.noData,
      noDataLabel: "No sites",
    };
  }

  const quantile = (p) => {
    const idx = Math.floor((sorted.length - 1) * p);
    return sorted[idx];
  };

  // Quantile breaks (fallback to linear if too many ties)
  let breaks = [quantile(0.2), quantile(0.4), quantile(0.6), quantile(0.8)];
  const uniqueBreaks = Array.from(new Set(breaks));
  if (uniqueBreaks.length < breaks.length){
    breaks = [];
    for (let i=1; i<5; i++){
      breaks.push(min + (max - min) * (i/5));
    }
  }

  const idxFor = (v) => {
    for (let i=0; i<breaks.length; i++){
      if (v <= breaks[i]) return i;
    }
    return colors.length - 1;
  };

  const bins = [];
  for (let i=0; i<colors.length; i++){
    const from = (i === 0) ? min : breaks[i-1];
    const to = (i < breaks.length) ? breaks[i] : max;
    bins.push({ from, to, color: colors[i] });
  }

  const color = (v) => {
    if (!Number.isFinite(v)) return BRAND.noData;
    return colors[idxFor(v)];
  };

  return {
    min,
    max,
    breaks,
    bins,
    color,
    format,
    label: labelByMetric[metric] || metric,
    isPct,
    noDataColor: BRAND.noData,
    noDataLabel: "No sites",
  };
}

function renderChoroplethLegend(scale){
  const node = $("mapLegend");
  if (!node) return;

  const rows = [];

  // No-data row
  rows.push(`
    <div class="legend-row">
      <span class="legend-swatch" style="background:${scale.noDataColor}"></span>
      <span>${escapeHtml(scale.noDataLabel)}</span>
    </div>
  `);

  // Data bins
  for (const b of scale.bins){
    const label = (b.from === b.to)
      ? scale.format(b.from)
      : `${scale.format(b.from)} – ${scale.format(b.to)}`;

    rows.push(`
      <div class="legend-row">
        <span class="legend-swatch" style="background:${b.color}"></span>
        <span>${escapeHtml(label)}</span>
      </div>
    `);
  }

  const title = `<div class="fw-semibold mb-1">Choropleth: ${escapeHtml(scale.label)}</div>`;
  const divider = `<hr class="my-2" style="border-color: var(--app-border)">`;

  const markerLegend = state.map.markerLegendHTML
    ? `<div class="mt-2">${divider}<div class="fw-semibold mb-1">Markers</div>${state.map.markerLegendHTML}</div>`
    : "";

  node.innerHTML = title + rows.join("") + markerLegend;
}

function escapeHtml(s){

    return String(s ?? "").replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  // ---------- UI wiring ----------
  let uiInitialized = false;

  function initUIOnce(){
    if (uiInitialized) return;
    uiInitialized = true;

    initCharts();
    initQCAccordion();
    initTables();
    initMap();

    // Set initial filter UI values from state.filters (URL)
    $("searchInput").value = state.filters.search || "";

    refreshAdminDropdowns();
    refreshOtherDropdowns();

    // Apply initial selections if present
    $("govFilter").value = state.filters.gov;
    $("districtFilter").value = state.filters.district;
    $("cadasterFilter").value = state.filters.cadaster;
    $("siteStatusFilter").value = state.filters.siteStatus;
    $("phoneStatusFilter").value = state.filters.phoneStatus;
    $("qcFilter").value = state.filters.qc;

    // Listeners
    const debounce = (fn, ms=200) => {
      let t=null;
      return (...args) => { if (t) clearTimeout(t); t=setTimeout(()=>fn(...args), ms); };
    };

    $("searchInput").addEventListener("input", debounce(() => applyFilters(), 250));
    $("govFilter").addEventListener("change", () => applyFilters());
    $("districtFilter").addEventListener("change", () => applyFilters());
    $("cadasterFilter").addEventListener("change", () => applyFilters());
    $("siteStatusFilter").addEventListener("change", () => applyFilters());
    $("phoneStatusFilter").addEventListener("change", () => applyFilters());
    $("qcFilter").addEventListener("change", () => applyFilters());

    $("resetFiltersBtn").addEventListener("click", resetFilters);
    $("downloadFilteredBtn").addEventListener("click", exportFilteredCSV);

    $("copyShareLinkBtn").addEventListener("click", copyShareLink);

    $("reloadDataBtn").addEventListener("click", () => {
      toast("Reloading data…");
      loadAllData();
    });

    $("scrollTopLink").addEventListener("click", (e) => {
      e.preventDefault();
      window.scrollTo({ top: 0, behavior: "smooth" });
    });

    // Chart download buttons
    document.querySelectorAll("[data-download-chart]").forEach(btn => {
      btn.addEventListener("click", () => {
        const id = btn.getAttribute("data-download-chart");
        const canvas = $(id);
        if (!canvas) return;
        const url = canvas.toDataURL("image/png");
        const a = document.createElement("a");
        a.href = url;
        a.download = `${id}.png`;
        document.body.appendChild(a);
        a.click();
        a.remove();
      });
    });
  }

  // ---------- start ----------
  document.addEventListener("DOMContentLoaded", () => {
    loadAllData();
  });

})();
