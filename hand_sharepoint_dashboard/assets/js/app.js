/* IAMP Sites Mapping Dashboard
 * - Loads an XLSX (SheetJS)
 * - Computes KPIs + charts (Chart.js)
 * - Interactive table (Tabulator)
 * - Filters + shareable link
 *
 * No server required; everything runs in the browser.
 */

(function () {
  "use strict";

  // -----------------------------
  // Config
  // -----------------------------
  const DEFAULT_SAMPLE_URL = "assets/data/IAMP_sites_mapping_SAMPLE_REDACTED.xlsx";
  const PREFERRED_SHEET_NAME = "IAMP sites mapping";

  // Vercel backend (optional)
  // If configured, /api/xlsx proxies the latest SharePoint XLSX securely (no CORS / no manual uploads)
  const API_XLSX_ENDPOINT = "/api/xlsx";
  const API_STATUS_ENDPOINT = "/api/status";

  const NUM_FIELDS = [
    "A- Number of Tents",
    "A1- Number of Households in Tents",
    "A2- Number of Individuals in Tents",
    "B- Number of Self-built Structures with Non-Concrete Roof",
    "B1- Number of Households in Self-built Structures with Non-Concrete Roof",
    "B2- Number of Individuals in Self-built Structures with Non-Concrete Roof",
    "C- Number of Prefab Structure",
    "C1- Number of Households in Prefab Structure",
    "C2- Number of Individuals in Prefab Structure",
    "D- Number of Self-built Structures with Concrete Roof",
    "D1- Number of Households in Self-built Structures with Concrete Roof",
    "D2- Number of Individuals in Self-built Structures with Concrete Roof",
    "Total number of Structures",
    "Total number of Households",
    "Total number of Individuals",
    "Number of Latrines",
    "E1- Number of Households came from Syria",
    "E2- Number of Individuals came from Syria",
    "F1- Number of Households came from Lebanon",
    "F2- Number of Individuals came from Lebanon",
    "G1- Number of Households left to Syria",
    "G2- Number of Individuals left to Syria",
    "QC - Issue count",
  ];

  const QC_RULES = [
    {
      col: "QC - Missing assessment date",
      label: "Missing assessment date (status filled but date blank)",
      help: "Fill Date of phone assessment when Phone call status is set.",
    },
    {
      col: "QC - Current phone length",
      label: "Current Shawish phone length not 8 digits",
      help: "Fix phone numbers (often missing leading 0) or confirm number is correct.",
    },
    {
      col: "QC - Phone status/details mismatch",
      label: "Phone call status/details mismatch or invalid status value",
      help: "Ensure status is valid; if status is not Answer, No response details should be filled.",
    },
    {
      col: "QC - Living status/new focal point mismatch",
      label: "Living status/new focal point mismatch or invalid value",
      help: "If living=No, new focal point name & phone must be filled. If living=Yes, new focal point fields should be blank.",
    },
    {
      col: "QC - New focal point missing assessment date",
      label: "New focal point: status filled but assessment date blank",
      help: "Fill Date of phone assessment with New Focal point when New FP status is set.",
    },
    {
      col: "QC - New FP status/details mismatch",
      label: "New focal point: status/details mismatch or invalid status value",
      help: "Ensure New FP status is valid; if status is not Answer, No response details should be filled.",
    },
    {
      col: "QC - Record status mismatch",
      label: "Record status mismatch or invalid value",
      help: "Record status should be filled when phone call status is filled (Finish/Need Follow-up).",
    },
    {
      col: "QC - Totals mismatch",
      label: "Totals mismatch (structures/HH/IND vs components)",
      help: "Check totals (structures/HH/IND) match sum of components (Tents/Shelters/Prefab/etc).",
    },
    {
      col: "QC - HH size outlier (>10 ind/hh or ind<hh)",
      label: "Household size outlier (>10 ind/HH or ind < HH)",
      help: "Review Total households vs Total individuals for possible data entry errors.",
    },
    {
      col: "QC - Latrines missing (HH>0 & latrines 0/blank)",
      label: "Latrines missing (HH>0 but latrines = 0/blank)",
      help: "Fill Number of Latrines when households exist (HH>0).",
    },
    {
      col: "QC - Latrines > HH",
      label: "Latrines greater than households",
      help: "Review latrine count relative to households.",
    },
    {
      col: "QC - High population (>500 ind)",
      label: "High population outlier (>500 individuals)",
      help: "Confirm Total individuals — unusually high compared with most sites.",
    },
    {
      col: "QC - Invalid Site Status",
      label: "Invalid Site Status",
      help: "Site Status should be one of: Active / Inactive / Fully Demolished.",
    },
    {
      col: "QC - Invalid PCode format",
      label: "Invalid PCode format",
      help: "PCode should follow #####-##-###.",
    },
  ];

  const PII_COLS = new Set([
    "Current Shawish Name",
    "Current Shawish Phone",
    "Available contact in the sites",
    "Name of the new focal point in the site",
    "Phone number of the new focal point in the site",
    "Shawish Name",
    "Shawish Phone",
    "Shawish Name2",
    "Shawish Phone 2",
    "Landlord Name",
    "Landlord Phone",
  ]);

  // -----------------------------
  // DOM helpers
  // -----------------------------
  const $ = (id) => document.getElementById(id);
  const el = (tag, attrs = {}, children = []) => {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") n.className = v;
      else if (k === "html") n.innerHTML = v;
      else if (k.startsWith("data-")) n.setAttribute(k, v);
      else n.setAttribute(k, v);
    }
    for (const c of children) n.appendChild(c);
    return n;
  };

  const fmtInt = (n) => new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 }).format(n || 0);
  const fmtPct = (x) => {
    if (!isFinite(x)) return "—";
    return (x * 100).toFixed(1) + "%";
  };
  const nowStamp = () => new Date().toLocaleString();

  function toNum(v) {
    if (v === null || v === undefined || v === "") return 0;
    const n = Number(v);
    return isFinite(n) ? n : 0;
  }

  function normText(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  function truthy(v) {
    if (v === true) return true;
    if (v === 1) return true;
    const s = normText(v).toLowerCase();
    return s === "yes" || s === "true" || s === "1";
  }

  function downloadText(filename, text, mime = "text/plain") {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function toCSV(rows, columns) {
    const esc = (s) => {
      const str = s === null || s === undefined ? "" : String(s);
      if (/["\n,]/.test(str)) return '"' + str.replace(/"/g, '""') + '"';
      return str;
    };

    const header = columns.map(esc).join(",");
    const lines = rows.map((r) => columns.map((c) => esc(r[c])).join(","));
    return [header, ...lines].join("\n");
  }

  // -----------------------------
  // State
  // -----------------------------
  const state = {
    raw: [],
    filtered: [],
    sourceLabel: "No data loaded",
    sourceUrl: "",
    liveMode: false,
    lastApiStatus: null,
    lastSuccessAt: null,
    lastLoadAt: null,
    errorCount: 0,

    refreshTimer: null,
    refreshEveryMs: 0,

    charts: {},
    table: null,
    qcMiniTable: null,

    // Map
    map: null,
    mapLayer: null,
    mapTiles: { light: null, dark: null },
    mapBoundaryLayer: null,
    coordsByPcode: new Map(),
    coordsMeta: { source: "sheet", pcodeKey: "PCode", latKey: null, lngKey: null, mapped: 0, total: 0 },
    mapColorBy: "Site Status",
    mapUseCluster: true,
  };

  const filters = {
    q: "",
    district: "All",
    cadaster: "All",
    siteStatus: "All",
    phoneStatus: "All",
    qc: "All",
  };

  // -----------------------------
  // Theme toggle
  // -----------------------------
  function applyTheme(theme) {
    document.documentElement.setAttribute("data-bs-theme", theme);
    localStorage.setItem("theme", theme);
    const icon = $("themeToggle")?.querySelector("i");
    if (icon) icon.className = theme === "dark" ? "bi bi-sun" : "bi bi-moon-stars";
  }

  function initTheme() {
    const saved = localStorage.getItem("theme");
    if (saved === "dark" || saved === "light") applyTheme(saved);
    $("themeToggle")?.addEventListener("click", () => {
      const current = document.documentElement.getAttribute("data-bs-theme") || "light";
      applyTheme(current === "dark" ? "light" : "dark");
      updateMapTheme();
      // Charts may need a redraw for better contrast; simplest is to re-render.
      if (state.raw.length) updateAll();
    });
  }

  // -----------------------------
  // Data loading
  // -----------------------------
  function showLoading(detail) {
    $("loadingDetail").textContent = detail || "Loading";
    $("loadingOverlay").classList.remove("d-none");
  }
  function hideLoading() {
    $("loadingOverlay").classList.add("d-none");
  }

  async function loadArrayBufferFromUrl(url) {
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error(`Failed to fetch (${resp.status})`);
    return await resp.arrayBuffer();
  }

  function pickSheet(workbook) {
    if (workbook.Sheets[PREFERRED_SHEET_NAME]) return PREFERRED_SHEET_NAME;
    // fallback: try close matches
    const names = workbook.SheetNames || [];
    const lower = names.map((n) => n.toLowerCase());
    const idx = lower.indexOf(PREFERRED_SHEET_NAME.toLowerCase());
    if (idx >= 0) return names[idx];
    return names[0];
  }

  function parseWorkbook(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = pickSheet(wb);
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // Normalize + derive fields
    const normalized = rows.map((r) => {
      const obj = { ...r };

      // numeric conversions
      for (const f of NUM_FIELDS) obj[f] = toNum(obj[f]);

      const phoneStatus = normText(obj["Phone call status"]);
      const siteStatus = normText(obj["Site Status"]);
      const district = normText(obj["District"]);
      const cadaster = normText(obj["Cadaster"]);

      const assessed = phoneStatus !== "";
      const siteStatusDisplay = siteStatus
        ? siteStatus
        : assessed
          ? "Not recorded"
          : "Not assessed";

      const qcAny = normText(obj["QC - Any issue"]) || "No";
      const qcIssueCount = toNum(obj["QC - Issue count"]);

      // QC flags list
      const qcFlags = [];
      for (const rule of QC_RULES) {
        if (truthy(obj[rule.col])) qcFlags.push(rule.label);
      }

      obj.__assessed = assessed;
      obj.__siteStatus = siteStatusDisplay;
      obj.__district = district || "—";
      obj.__cadaster = cadaster || "—";
      obj.__phoneStatus = assessed ? phoneStatus : "Not assessed";
      obj.__qcAny = qcAny === "Yes" ? "Yes" : "No";
      obj.__qcIssueCount = qcIssueCount;
      obj.__qcFlags = qcFlags;

      // short identifiers for searching
      obj.__search = [
        normText(obj["PCode"]),
        normText(obj["PCode Name"]),
        normText(obj["Local Name"]),
        district,
        cadaster,
      ].join(" ").toLowerCase();

      return obj;
    });

    return { rows: normalized, sheetName };
  }

  async function loadFromFile(file) {
    state.lastLoadAt = new Date();
    setHealth("Loading file…", false);

    showLoading("Reading file");
    const buf = await file.arrayBuffer();
    showLoading("Parsing spreadsheet");
    const { rows, sheetName } = parseWorkbook(buf);
    hideLoading();

    state.sourceLabel = file.name;
    $("fileNameHint").textContent = file.name;
    state.sourceUrl = "";

    // Loading a manual file should disable live mode + auto-refresh.
    if (state.liveMode) {
      setLiveUi(false);
      localStorage.setItem("liveMode", "0");
      setApiStatus("Live mode off (manual file loaded).", true);
    }
    applyRefresh();

    onDataLoaded(rows, `File: ${file.name}`, sheetName);
  }

  async function loadFromUrl(url, labelOverride = null) {
    if (!url) return;
    state.lastLoadAt = new Date();
    setHealth("Loading URL…", false);

    showLoading("Fetching xlsx from URL");
    const buf = await loadArrayBufferFromUrl(url);
    showLoading("Parsing spreadsheet");
    const { rows, sheetName } = parseWorkbook(buf);
    hideLoading();

    state.sourceLabel = url;
    state.sourceUrl = url;

    // If the user loads a manual URL, disable live mode so it doesn't surprise them later.
    if (url !== API_XLSX_ENDPOINT && state.liveMode) {
      setLiveUi(false);
      localStorage.setItem("liveMode", "0");
      setApiStatus("Live mode off (manual URL loaded).", true);
    }

    onDataLoaded(rows, labelOverride || ("URL: " + url), sheetName);
  }

  async function loadSample() {
    await loadFromUrl(DEFAULT_SAMPLE_URL);
  }

  // -----------------------------
  // Live mode (Vercel backend)
  // -----------------------------
  function setLiveUi(enabled) {
    state.liveMode = !!enabled;
    const badge = $("liveBadge");
    const toggle = $("liveToggle");
    if (badge) {
      badge.textContent = enabled ? "on" : "off";
      badge.classList.toggle("text-bg-success", enabled);
      badge.classList.toggle("text-bg-light", !enabled);
    }
    if (toggle) toggle.checked = enabled;
  }

  function setApiStatus(msg, ok = true) {
    const el = $("apiStatusText");
    if (!el) return;
    el.textContent = msg || "—";
    el.classList.toggle("text-danger", !ok);
    el.classList.toggle("text-body-secondary", ok);
  }

  async function testApiConnection() {
    setApiStatus("Testing connection…", true);
    const res = await fetch(API_STATUS_ENDPOINT, { cache: "no-store" });
    if (!res.ok) throw new Error(`API status failed (HTTP ${res.status})`);
    const data = await res.json();

    if (!data || data.ok === false) {
      const msg = data?.error || "API not configured.";
      throw new Error(msg);
    }

    state.lastApiStatus = data;
    const name = data.name ? `• ${data.name}` : "";
    const mod = data.lastModifiedDateTime ? `• modified ${new Date(data.lastModifiedDateTime).toLocaleString()}` : "";
    setApiStatus(`API connected ${name} ${mod}`.trim(), true);
    return data;
  }

  async function enableLiveMode(enabled) {
    setLiveUi(enabled);
    localStorage.setItem("liveMode", enabled ? "1" : "0");

    if (!enabled) {
      setApiStatus("Live mode off. Use Load File or Load URL.", true);
      return;
    }

    // Sanity check endpoint first so the error is clearer.
    try {
      await testApiConnection();
    } catch (err) {
      setApiStatus("Live mode failed: " + err.message, false);
      setLiveUi(false);
      localStorage.setItem("liveMode", "0");
      return;
    }

    // Load data through the same XLSX pipeline, but from the Vercel API.
    try {
      await loadFromUrl(API_XLSX_ENDPOINT, "LIVE: SharePoint (Vercel)");
      // Auto-refresh uses state.sourceUrl; keep it pointed to the API endpoint.
      state.sourceUrl = API_XLSX_ENDPOINT;
      applyRefresh();
    } catch (err) {
      setApiStatus("Live load failed: " + err.message, false);
    }
  }

  // -----------------------------
  // Filters
  // -----------------------------
  function uniqSorted(values) {
    return Array.from(new Set(values.filter((v) => v && v !== "—"))).sort((a, b) => a.localeCompare(b));
  }

  function setSelectOptions(selectEl, options, includeAll = true) {
    selectEl.innerHTML = "";
    if (includeAll) {
      const o = document.createElement("option");
      o.value = "All";
      o.textContent = "All";
      selectEl.appendChild(o);
    }
    for (const val of options) {
      const o = document.createElement("option");
      o.value = val;
      o.textContent = val;
      selectEl.appendChild(o);
    }
  }

  function populateFilterControls(rows) {
    const districts = uniqSorted(rows.map((r) => r.__district));
    const cadasters = uniqSorted(rows.map((r) => r.__cadaster));
    const siteStatuses = uniqSorted(rows.map((r) => r.__siteStatus));
    const phoneStatuses = uniqSorted(rows.map((r) => r.__phoneStatus));

    setSelectOptions($("districtFilter"), districts);
    setSelectOptions($("cadasterFilter"), cadasters);
    setSelectOptions($("siteStatusFilter"), siteStatuses);
    setSelectOptions($("phoneStatusFilter"), phoneStatuses);

    $("qcFilter").innerHTML = "";
    ["All", "Any issue", "No issues"].forEach((v) => {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      $("qcFilter").appendChild(o);
    });
  }

  function readFiltersFromUI() {
    filters.q = normText($("searchInput").value);
    filters.district = $("districtFilter").value || "All";
    filters.cadaster = $("cadasterFilter").value || "All";
    filters.siteStatus = $("siteStatusFilter").value || "All";
    filters.phoneStatus = $("phoneStatusFilter").value || "All";
    filters.qc = $("qcFilter").value || "All";
  }

  function applyFilters(rows) {
    const q = filters.q.trim().toLowerCase();
    return rows.filter((r) => {
      if (filters.district !== "All" && r.__district !== filters.district) return false;
      if (filters.cadaster !== "All" && r.__cadaster !== filters.cadaster) return false;
      if (filters.siteStatus !== "All" && r.__siteStatus !== filters.siteStatus) return false;
      if (filters.phoneStatus !== "All" && r.__phoneStatus !== filters.phoneStatus) return false;

      if (filters.qc === "Any issue" && r.__qcAny !== "Yes") return false;
      if (filters.qc === "No issues" && r.__qcAny !== "No") return false;

      if (q && !r.__search.includes(q)) return false;
      return true;
    });
  }

  function resetFilters() {
    $("searchInput").value = "";
    ["districtFilter", "cadasterFilter", "siteStatusFilter", "phoneStatusFilter"].forEach((id) => {
      const s = $(id);
      if (s) s.value = "All";
    });
    $("qcFilter").value = "All";
    readFiltersFromUI();
    updateAll();
  }

  // Share link: encode filters into URL query
  function encodeFiltersToQuery() {
    const params = new URLSearchParams();
    if (filters.q) params.set("q", filters.q);
    if (filters.district !== "All") params.set("district", filters.district);
    if (filters.cadaster !== "All") params.set("cadaster", filters.cadaster);
    if (filters.siteStatus !== "All") params.set("siteStatus", filters.siteStatus);
    if (filters.phoneStatus !== "All") params.set("phoneStatus", filters.phoneStatus);
    if (filters.qc !== "All") params.set("qc", filters.qc);
    return params.toString();
  }

  function applyQueryToFilters() {
    const params = new URLSearchParams(window.location.search);
    const q = params.get("q");
    const district = params.get("district");
    const cadaster = params.get("cadaster");
    const siteStatus = params.get("siteStatus");
    const phoneStatus = params.get("phoneStatus");
    const qc = params.get("qc");

    if (q) $("searchInput").value = q;
    if (district) $("districtFilter").value = district;
    if (cadaster) $("cadasterFilter").value = cadaster;
    if (siteStatus) $("siteStatusFilter").value = siteStatus;
    if (phoneStatus) $("phoneStatusFilter").value = phoneStatus;
    if (qc) $("qcFilter").value = qc;

    readFiltersFromUI();
  }

  // -----------------------------
  // KPIs
  // -----------------------------
  function updateKpis(rows, rawRows) {
    const total = rows.length;
    const totalRaw = rawRows.length;

    const assessed = rows.filter((r) => r.__assessed).length;
    const assessedRaw = rawRows.filter((r) => r.__assessed).length;

    const active = rows.filter((r) => r.__siteStatus === "Active").length;
    const activeRaw = rawRows.filter((r) => r.__siteStatus === "Active").length;

    const qcRecords = rows.filter((r) => r.__qcAny === "Yes").length;
    const qcRaw = rawRows.filter((r) => r.__qcAny === "Yes").length;

    const hh = rows.filter((r) => r.__siteStatus === "Active").reduce((s, r) => s + toNum(r["Total number of Households"]), 0);
    const ind = rows.filter((r) => r.__siteStatus === "Active").reduce((s, r) => s + toNum(r["Total number of Individuals"]), 0);
    const structures = rows.filter((r) => r.__siteStatus === "Active").reduce((s, r) => s + toNum(r["Total number of Structures"]), 0);
    const latrines = rows.filter((r) => r.__siteStatus === "Active").reduce((s, r) => s + toNum(r["Number of Latrines"]), 0);

    $("kpiTotal").textContent = fmtInt(total);
    $("kpiTotalSub").textContent = total === totalRaw ? "All records" : `Filtered from ${fmtInt(totalRaw)}`;

    $("kpiAssessed").textContent = fmtInt(assessed);
    $("kpiAssessedSub").textContent = `${fmtPct(total ? assessed / total : 0)} of filtered`;

    $("kpiActive").textContent = fmtInt(active);
    $("kpiActiveSub").textContent = `${fmtPct(total ? active / total : 0)} of filtered`;

    $("kpiQC").textContent = fmtInt(qcRecords);
    $("kpiQCSub").textContent = `${fmtPct(total ? qcRecords / total : 0)} of filtered`;

    $("kpiHH").textContent = fmtInt(hh);
    $("kpiHHSub").textContent = active ? `${(hh / active).toFixed(1)} avg/site` : "—";

    $("kpiIND").textContent = fmtInt(ind);
    $("kpiINDSub").textContent = active ? `${(ind / active).toFixed(1)} avg/site` : "—";

    $("kpiStruct").textContent = fmtInt(structures);
    $("kpiStructSub").textContent = active ? `${(structures / active).toFixed(1)} avg/site` : "—";

    $("kpiLat").textContent = fmtInt(latrines);
    $("kpiLatSub").textContent = active ? `${(latrines / active).toFixed(1)} avg/site` : "—";

    // Assessment note
    const pending = total - assessed;
    const pendingPct = total ? pending / total : 0;
    $("assessmentNote").textContent = `${fmtInt(assessed)} assessed • ${fmtInt(pending)} not assessed (${fmtPct(pendingPct)}).`;
  }

  // -----------------------------
  // Charts
  // -----------------------------
  function destroyChart(id) {
    if (state.charts[id]) {
      state.charts[id].destroy();
      delete state.charts[id];
    }
  }

  function applyChartTheme() {
    // Use Bootstrap CSS variables so charts automatically match light/dark themes.
    const theme = getTheme();
    const styles = getComputedStyle(document.documentElement);

    const bodyColor = styles.getPropertyValue("--bs-body-color").trim()
      || (theme === "dark" ? "rgba(255,255,255,.85)" : "rgba(0,0,0,.85)");
    const gridColor = theme === "dark" ? "rgba(255,255,255,.10)" : "rgba(0,0,0,.08)";

    Chart.defaults.color = bodyColor;
    Chart.defaults.borderColor = gridColor;
    Chart.defaults.font.family = 'system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, "Helvetica Neue", Arial, sans-serif';
    Chart.defaults.plugins.legend.labels.usePointStyle = true;
    Chart.defaults.plugins.legend.labels.boxWidth = 10;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.boxPadding = 4;

    // Smoother visuals
    Chart.defaults.elements.arc.borderWidth = 0;
    Chart.defaults.elements.bar.borderRadius = 8;
    Chart.defaults.elements.bar.borderSkipped = false;
    Chart.defaults.elements.line.tension = 0.35;

    // Tooltip polish
    Chart.defaults.plugins.tooltip.backgroundColor = theme === "dark" ? "rgba(17,24,39,.92)" : "rgba(255,255,255,.96)";
    Chart.defaults.plugins.tooltip.titleColor = theme === "dark" ? "rgba(255,255,255,.92)" : "rgba(0,0,0,.86)";
    Chart.defaults.plugins.tooltip.bodyColor = theme === "dark" ? "rgba(255,255,255,.86)" : "rgba(0,0,0,.78)";
    Chart.defaults.plugins.tooltip.borderColor = theme === "dark" ? "rgba(255,255,255,.14)" : "rgba(0,0,0,.10)";
    Chart.defaults.plugins.tooltip.borderWidth = 1;
  }

  function makeChart(ctx, config) {
    applyChartTheme();
    return new Chart(ctx, config);
  }

  function getPalette(n) {
    // Modern, high-contrast palette (works on both light + dark)
    const base = [
      "#2563eb", // blue
      "#10b981", // emerald
      "#8b5cf6", // violet
      "#f59e0b", // amber
      "#ef4444", // red
      "#06b6d4", // cyan
      "#f97316", // orange
      "#14b8a6", // teal
      "#a3e635", // lime
      "#e11d48", // rose
    ];
    const out = [];
    for (let i = 0; i < n; i++) out.push(base[i % base.length]);
    return out;
  }

  function updateCharts(rows, rawRows) {
    // Chart 1: Assessment progress donut
    destroyChart("chartAssessment");
    const assessed = rows.filter((r) => r.__assessed).length;
    const pending = rows.length - assessed;

    state.charts.chartAssessment = makeChart($("chartAssessment"), {
      type: "doughnut",
      data: {
        labels: ["Assessed", "Not assessed"],
        datasets: [{
          data: [assessed, pending],
          backgroundColor: getPalette(2),
          borderWidth: 0,
        }],
      },
      options: {
        responsive: true,
        plugins: { legend: { position: "bottom" } },
        cutout: "65%",
      },
    });

    // Chart 2: Site status mix
    destroyChart("chartSiteStatus");
    const statusCounts = countBy(rows, (r) => r.__siteStatus);
    const statusLabels = Object.keys(statusCounts);
    const statusValues = statusLabels.map((k) => statusCounts[k]);

    state.charts.chartSiteStatus = makeChart($("chartSiteStatus"), {
      type: "doughnut",
      data: {
        labels: statusLabels,
        datasets: [{
          data: statusValues,
          backgroundColor: getPalette(statusLabels.length),
          borderWidth: 0,
        }],
      },
      options: {
        responsive: true,
        plugins: { legend: { position: "bottom" } },
        cutout: "62%",
      },
    });

    // Chart 3: Phone outcomes (assessed)
    destroyChart("chartPhoneOutcomes");
    const assessedRows = rows.filter((r) => r.__assessed);
    const phoneCounts = countBy(assessedRows, (r) => r.__phoneStatus);
    const phoneEntries = Object.entries(phoneCounts).sort((a, b) => b[1] - a[1]);
    const phoneLabels = phoneEntries.map((x) => x[0]);
    const phoneValues = phoneEntries.map((x) => x[1]);

    state.charts.chartPhoneOutcomes = makeChart($("chartPhoneOutcomes"), {
      type: "bar",
      data: {
        labels: phoneLabels,
        datasets: [{
          label: "Records",
          data: phoneValues,
          backgroundColor: getPalette(1)[0],
        }],
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        indexAxis: "y",
        scales: {
          x: { ticks: { precision: 0 } },
          y: { ticks: { autoSkip: false } },
        },
      },
    });

    // Chart 4: Structure composition (active)
    destroyChart("chartStructures");
    const activeRows = rows.filter((r) => r.__siteStatus === "Active");
    const structureFields = [
      { key: "A- Number of Tents", label: "Tents" },
      { key: "B- Number of Self-built Structures with Non-Concrete Roof", label: "Self-built (non-concrete roof)" },
      { key: "C- Number of Prefab Structure", label: "Prefab" },
      { key: "D- Number of Self-built Structures with Concrete Roof", label: "Self-built (concrete roof)" },
    ];
    const structLabels = structureFields.map((x) => x.label);
    const structValues = structureFields.map((x) => activeRows.reduce((s, r) => s + toNum(r[x.key]), 0));

    state.charts.chartStructures = makeChart($("chartStructures"), {
      type: "bar",
      data: {
        labels: structLabels,
        datasets: [{
          label: "Structures",
          data: structValues,
          backgroundColor: getPalette(structLabels.length),
        }],
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
      },
    });

    // Chart 5: Top cadasters by active individuals
    destroyChart("chartTopCadaster");
    const byCad = groupSum(activeRows, (r) => r.__cadaster, (r) => toNum(r["Total number of Individuals"]));
    const top = Object.entries(byCad).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const cadLabels = top.map((x) => x[0]);
    const cadValues = top.map((x) => x[1]);

    state.charts.chartTopCadaster = makeChart($("chartTopCadaster"), {
      type: "bar",
      data: {
        labels: cadLabels,
        datasets: [{
          label: "Individuals",
          data: cadValues,
          backgroundColor: getPalette(1)[0],
        }],
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        indexAxis: "y",
        scales: {
          x: { ticks: { precision: 0 } },
          y: { ticks: { autoSkip: false } },
        },
      },
    });
  }

  function countBy(rows, fnKey) {
    const out = {};
    for (const r of rows) {
      const k = fnKey(r) || "—";
      out[k] = (out[k] || 0) + 1;
    }
    return out;
  }

  function groupSum(rows, fnKey, fnVal) {
    const out = {};
    for (const r of rows) {
      const k = fnKey(r) || "—";
      out[k] = (out[k] || 0) + (fnVal(r) || 0);
    }
    return out;
  }

  // QC chart + accordion counts
  function updateQualitySection(rows) {
    // Chart: QC by type
    destroyChart("chartQC");
    const qcCounts = QC_RULES.map((rule) => rows.reduce((s, r) => s + (truthy(r[rule.col]) ? 1 : 0), 0));
    const qcLabels = QC_RULES.map((r) => r.label);

    state.charts.chartQC = makeChart($("chartQC"), {
      type: "bar",
      data: {
        labels: qcLabels,
        datasets: [{
          label: "Records",
          data: qcCounts,
          backgroundColor: getPalette(1)[0],
        }],
      },
      options: {
        responsive: true,
        plugins: { legend: { display: false } },
        indexAxis: "y",
        scales: {
          x: { ticks: { precision: 0 } },
          y: { ticks: { autoSkip: false } },
        },
      },
    });

    // Accordion
    const acc = $("qcAccordion");
    if (!acc.dataset.built) {
      acc.innerHTML = "";
      QC_RULES.forEach((rule, idx) => {
        const itemId = `qcItem${idx}`;
        const headerId = `qcHeader${idx}`;
        const collapseId = `qcCollapse${idx}`;

        const item = el("div", { class: "accordion-item" }, [
          el("h2", { class: "accordion-header", id: headerId }, [
            el("button", {
              class: "accordion-button collapsed",
              type: "button",
              "data-bs-toggle": "collapse",
              "data-bs-target": "#" + collapseId,
              "aria-expanded": "false",
              "aria-controls": collapseId,
            }, []),
          ]),
          el("div", {
            id: collapseId,
            class: "accordion-collapse collapse",
            "aria-labelledby": headerId,
            "data-bs-parent": "#qcAccordion",
          }, [
            el("div", { class: "accordion-body" }, [
              el("div", { class: "small text-body-secondary mb-2", id: itemId + "-count" }, []),
              el("div", { class: "fw-semibold mb-1" }, [document.createTextNode("How to resolve")]),
              el("div", { class: "small" }, [document.createTextNode(rule.help)]),
              el("hr", { class: "my-2" }),
              el("div", { class: "small text-body-secondary" }, [document.createTextNode("QC column: " + rule.col)]),
            ]),
          ]),
        ]);

        // Label text is set later (so we can include counts)
        item.querySelector(".accordion-button").textContent = rule.label;
        acc.appendChild(item);
      });
      acc.dataset.built = "1";
    }

    // Update counts in accordion headers (without rebuilding)
    const headers = acc.querySelectorAll(".accordion-button");
    QC_RULES.forEach((rule, idx) => {
      const cnt = rows.reduce((s, r) => s + (truthy(r[rule.col]) ? 1 : 0), 0);
      if (headers[idx]) headers[idx].textContent = `${rule.label} (${fmtInt(cnt)})`;
    });

    // QC mini table
    const qcRows = rows.filter((r) => r.__qcAny === "Yes").slice(0, 250).map((r) => ({
      "PCode": r["PCode"],
      "PCode Name": r["PCode Name"],
      "District": r.__district,
      "Cadaster": r.__cadaster,
      "Site Status": r.__siteStatus,
      "Phone call status": r.__phoneStatus,
      "QC - Issue count": r.__qcIssueCount,
      "QC flags": r.__qcFlags.join("; "),
    }));
    if (!state.qcMiniTable) {
      state.qcMiniTable = new Tabulator("#qcMiniTable", {
        data: qcRows,
        layout: "fitColumns",
        height: 320,
        pagination: true,
        paginationSize: 25,
        columns: [
          { title: "PCode", field: "PCode", width: 120 },
          { title: "PCode Name", field: "PCode Name", widthGrow: 2 },
          { title: "District", field: "District" },
          { title: "Cadaster", field: "Cadaster" },
          { title: "Site Status", field: "Site Status" },
          { title: "Phone status", field: "Phone call status", widthGrow: 2 },
          { title: "Issue count", field: "QC - Issue count", hozAlign: "right" },
        ],
      });
    } else {
      state.qcMiniTable.replaceData(qcRows);
    }

    $("downloadQCBtn").onclick = () => {
      const allQCs = rows.filter((r) => r.__qcAny === "Yes").map((r) => sanitizeForExport(r));
      const cols = exportColumns();
      downloadText(`qc_records_${stampFile()}.csv`, toCSV(allQCs, cols), "text/csv");
    };
  }

  
  // -----------------------------
  // Map (Lebanon)
  // -----------------------------
  const LEBANON_BOUNDS = [
    [33.0, 35.0],
    [34.75, 36.7],
  ];
  const LEBANON_CENTER = [33.8547, 35.8623];

  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function hashCode(str) {
    const s = String(str ?? "");
    let h = 0;
    for (let i = 0; i < s.length; i++) h = ((h << 5) - h) + s.charCodeAt(i);
    return h | 0;
  }

  function colorForKey(key) {
    const k = String(key ?? "").trim();
    const fixed = {
      "Active": "#10b981",
      "Inactive": "#f59e0b",
      "Fully Demolished": "#ef4444",
      "Not assessed": "#94a3b8",
      "Not recorded": "#8b5cf6",
      "QC issue": "#ef4444",
      "No QC issue": "#10b981",
      "—": "#94a3b8",
    };
    if (fixed[k]) return fixed[k];

    // Fallback: hash-based hue so categories remain stable between reloads.
    const h = Math.abs(hashCode(k)) % 360;
    return `hsl(${h}, 72%, 46%)`;
  }

  function getTheme() {
    return document.documentElement.getAttribute("data-bs-theme") || "light";
  }

  function initMap() {
    const el = document.getElementById("lebanonMap");
    if (!el || !window.L) return;

    // Prevent double init
    if (state.map) return;

    state.map = L.map(el, {
      zoomSnap: 0.5,
      zoomControl: true,
    });

    // Tiles (switch based on theme)
    state.mapTiles.light = L.tileLayer(
      "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png",
      { maxZoom: 19, attribution: "&copy; OpenStreetMap &copy; CARTO" }
    );
    state.mapTiles.dark = L.tileLayer(
      "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png",
      { maxZoom: 19, attribution: "&copy; OpenStreetMap &copy; CARTO" }
    );

    updateMapTheme();

    // Initial view
    state.map.fitBounds(LEBANON_BOUNDS, { padding: [12, 12] });

    // Hint visibility
    setMapEmptyHintVisible(true);

    // Resize fix when switching to Map tab
    const mapTab = document.getElementById("tab-map");
    mapTab?.addEventListener("shown.bs.tab", () => {
      setTimeout(() => {
        try { state.map.invalidateSize(); } catch (_) {}
      }, 50);
    });
  }

  function updateMapTheme() {
    if (!state.map) return;
    const theme = getTheme();
    const targetLayer = theme === "dark" ? state.mapTiles.dark : state.mapTiles.light;

    // Remove both then add the correct one
    try {
      if (state.mapTiles.light) state.map.removeLayer(state.mapTiles.light);
      if (state.mapTiles.dark) state.map.removeLayer(state.mapTiles.dark);
    } catch (_) {}

    try { targetLayer.addTo(state.map); } catch (_) {}
  }

  function setMapEmptyHintVisible(visible) {
    const hint = $("mapEmptyHint");
    if (!hint) return;
    hint.classList.toggle("d-none", !visible);
  }

  function makeDotIcon(color) {
    return L.divIcon({
      className: "",
      html: `<span class="map-dot" style="background:${color}"></span>`,
      iconSize: [14, 14],
      iconAnchor: [7, 7],
      popupAnchor: [0, -8],
    });
  }

  function getCoordForRow(row) {
    const pcode = String(row["PCode"] ?? "").trim();
    if (pcode && state.coordsByPcode?.has(pcode)) return state.coordsByPcode.get(pcode);

    // Fallback: read from detected lat/lng keys (if available)
    const latKey = state.coordsMeta?.latKey;
    const lngKey = state.coordsMeta?.lngKey;
    if (latKey && lngKey) {
      const a = parseCoordinate(row[latKey]);
      const b = parseCoordinate(row[lngKey]);
      if (Number.isFinite(a) && Number.isFinite(b)) return { lat: a, lng: b };
    }

    return null;
  }

  function getCategory(row, mode) {
    if (mode === "QC") return row.__qcAny === "Yes" ? "QC issue" : "No QC issue";
    if (mode === "Phone call status") return row.__phoneStatus || "—";
    // default: Site Status
    return row.__siteStatus || "—";
  }

  function buildLegend(countsByCategory) {
    const legend = $("mapLegend");
    if (!legend) return;

    const entries = Object.entries(countsByCategory || {}).sort((a, b) => b[1] - a[1]);

    if (!entries.length) {
      legend.innerHTML = `<div class="text-body-secondary">No mapped points (after filters).</div>`;
      return;
    }

    legend.innerHTML = entries
      .slice(0, 18)
      .map(([cat, n]) => {
        const col = colorForKey(cat);
        return `
          <div class="map-legend-item">
            <span class="map-swatch" style="background:${col}"></span>
            <span class="flex-grow-1">${escapeHtml(cat)}</span>
            <span class="text-body-secondary">${n}</span>
          </div>
        `;
      })
      .join("");
  }

  function clearMapLayer() {
    if (!state.map || !state.mapLayer) return;
    try { state.map.removeLayer(state.mapLayer); } catch (_) {}
    state.mapLayer = null;
  }

  function updateMap(rows) {
    if (!state.map) return;

    // Read map controls
    const mode = $("mapColorBy")?.value || state.mapColorBy || "Site Status";
    const useCluster = $("toggleCluster") ? $("toggleCluster").checked : true;
    state.mapColorBy = mode;
    state.mapUseCluster = useCluster;

    const coordsLoaded = state.coordsByPcode && state.coordsByPcode.size > 0;

    // When no coordinates exist at all, show the hint overlay.
    if (!coordsLoaded) {
      clearMapLayer();
      setMapEmptyHintVisible(true);
      $("mapStatusText").textContent = "No usable Latitude/Longitude values detected in the spreadsheet.";
      buildLegend({});
      return;
    }

    setMapEmptyHintVisible(false);

    // Build markers
    const mapped = [];
    const missing = [];

    for (const r of rows) {
      const coord = getCoordForRow(r);
      if (!coord) { missing.push(r); continue; }
      mapped.push({ row: r, coord, category: getCategory(r, mode) });
    }

    // Layer selection
    const canCluster = !!(window.L && L.markerClusterGroup);
    const layer = (useCluster && canCluster)
      ? L.markerClusterGroup({ showCoverageOnHover: false, maxClusterRadius: 48 })
      : L.layerGroup();

    const counts = {};
    for (const item of mapped) {
      const cat = item.category || "—";
      counts[cat] = (counts[cat] || 0) + 1;

      const col = colorForKey(cat);
      const { lat, lng } = item.coord;

      const pcode = String(item.row["PCode"] ?? "").trim();
      const localName = item.row["Local Name"] || item.row["PCode Name"] || "";
      const district = item.row["District"] || "";
      const cadaster = item.row["Cadaster"] || "";
      const status = item.row["Site Status"] || "";
      const phone = item.row["Phone call status"] || "";
      const hh = item.row["Total number of Households"] ?? "";
      const ind = item.row["Total number of Individuals"] ?? "";

      const marker = L.marker([lat, lng], { icon: makeDotIcon(col) });
      marker.bindPopup(`
        <div class="fw-semibold">${escapeHtml(pcode || "Site")}</div>
        <div class="small text-body-secondary">${escapeHtml(localName)}</div>
        <hr class="my-2">
        <div class="small">
          <div><b>District:</b> ${escapeHtml(district)}</div>
          <div><b>Cadaster:</b> ${escapeHtml(cadaster)}</div>
          <div><b>Site Status:</b> ${escapeHtml(status)}</div>
          <div><b>Phone status:</b> ${escapeHtml(phone)}</div>
          <div class="mt-2 d-flex gap-3">
            <div><b>HH:</b> ${escapeHtml(hh)}</div>
            <div><b>IND:</b> ${escapeHtml(ind)}</div>
          </div>
          <div class="mt-2">
            <a class="link-primary" href="https://www.google.com/maps?q=${encodeURIComponent(lat + ',' + lng)}" target="_blank" rel="noopener">Open in Google Maps</a>
          </div>
        </div>
      `);

      layer.addLayer(marker);
    }

    clearMapLayer();
    state.mapLayer = layer;
    state.mapLayer.addTo(state.map);

    buildLegend(counts);

    // Status line
    const total = rows.length;
    const shown = mapped.length;
    const missingN = missing.length;
    const coordTotal = state.coordsByPcode.size;

    $("mapStatusText").textContent =
      `Showing ${shown.toLocaleString()} mapped site(s) out of ${total.toLocaleString()} filtered. ` +
      (missingN ? `${missingN.toLocaleString()} site(s) have no matching coordinates. ` : "") +
      `Coordinates mapped: ${coordTotal.toLocaleString()}.`;

    // Fit bounds if we have points and map hasn't been moved by the user recently.
    if (shown > 0) {
      try {
        const bounds = layer.getBounds ? layer.getBounds() : null;
        if (bounds && bounds.isValid && bounds.isValid()) state.map.fitBounds(bounds, { padding: [16, 16] });
      } catch (_) {}
    }
  }

  // -----------------------------
  // Coordinates (from spreadsheet)
  // -----------------------------
  function parseCoordinate(value) {
    if (value === null || value === undefined || value === "") return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;

    const raw = String(value).trim();
    if (!raw) return NaN;

    // DMS support: 33°54'12.3"N or 35 52 10 E
    const dms = raw.match(/(\d+(?:\.\d+)?)\D+(\d+(?:\.\d+)?)\D+(\d+(?:\.\d+)?)(?:\D*([NSEW]))?/i);
    if (dms) {
      const deg = Number(dms[1]);
      const min = Number(dms[2]);
      const sec = Number(dms[3]);
      if ([deg, min, sec].every(Number.isFinite)) {
        let out = deg + (min / 60) + (sec / 3600);
        const hemi = (dms[4] || "").toUpperCase();
        if (hemi === "S" || hemi === "W") out *= -1;
        return out;
      }
    }

    // Loose decimal parsing (handles 33,875 => 33.875)
    const cleaned = raw
      .replace(/,/g, ".")
      .replace(/[^0-9.+\-]/g, "")
      .replace(/(\+)(?=.)/g, "");
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function inferPcodeKey(keys) {
    const exact = keys.find((k) => String(k).trim().toLowerCase() === "pcode");
    if (exact) return exact;
    const fuzzy = keys.find((k) => /p\s*code/i.test(String(k)) || /pcode/i.test(String(k)));
    return fuzzy || "PCode";
  }

  function scoreKeyInRange(rows, key, min, max) {
    const sampleN = Math.min(1200, rows.length);
    let ok = 0;
    let valid = 0;
    for (let i = 0; i < sampleN; i++) {
      const v = parseCoordinate(rows[i]?.[key]);
      if (!Number.isFinite(v)) continue;
      valid += 1;
      if (v >= min && v <= max) ok += 1;
    }
    if (!valid) return 0;
    return ok / valid;
  }

  function inferLatLngKeys(rows) {
    if (!rows || !rows.length) return { pcodeKey: "PCode", latKey: null, lngKey: null };
    const keys = Object.keys(rows[0] || {});

    const pcodeKey = inferPcodeKey(keys);

    // Heuristic candidates by name
    const nameLat = keys.filter((k) => /lat|latitude/i.test(k) && !/latrine/i.test(k));
    const nameLng = keys.filter((k) => /lon|lng|longitude/i.test(k));

    // Numeric scoring across all keys (catches X/Y, GPS columns, etc.)
    const scored = keys.map((k) => {
      const latScore = scoreKeyInRange(rows, k, 32.0, 35.9); // Lebanon latitude band
      const lngScore = scoreKeyInRange(rows, k, 34.0, 37.9); // Lebanon longitude band
      return { key: k, latScore, lngScore };
    });

    const bestLat = [...scored]
      .sort((a, b) => b.latScore - a.latScore)
      .find((x) => x.latScore >= 0.65);
    const bestLng = [...scored]
      .sort((a, b) => b.lngScore - a.lngScore)
      .find((x) => x.lngScore >= 0.65 && (!bestLat || x.key !== bestLat.key));

    // Prefer name matches if they also score well.
    const pickBestNamed = (arr, scoreProp) => {
      if (!arr.length) return null;
      const ranked = arr
        .map((k) => ({ k, s: (scored.find((x) => x.key === k)?.[scoreProp] || 0) }))
        .sort((a, b) => b.s - a.s);
      return ranked[0]?.s >= 0.25 ? ranked[0].k : null;
    };

    const latKey = pickBestNamed(nameLat, "latScore") || bestLat?.key || null;
    const lngKey = pickBestNamed(nameLng, "lngScore") || bestLng?.key || null;

    return { pcodeKey, latKey, lngKey };
  }

  function buildCoordsMapFromRows(rows) {
    const meta = inferLatLngKeys(rows);
    const mp = new Map();
    if (!meta.latKey || !meta.lngKey) {
      return { mp, meta: { ...state.coordsMeta, ...meta, mapped: 0, total: rows?.length || 0 } };
    }

    let mapped = 0;
    for (const r of rows) {
      const p = String(r?.[meta.pcodeKey] ?? r?.["PCode"] ?? "").trim();
      if (!p) continue;
      const lat = parseCoordinate(r?.[meta.latKey]);
      const lng = parseCoordinate(r?.[meta.lngKey]);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) continue;
      mp.set(p, { lat, lng });
      mapped += 1;
    }

    return { mp, meta: { ...state.coordsMeta, ...meta, mapped, total: rows?.length || 0 } };
  }

  function setCoordsFromSpreadsheet(rows) {
    const { mp, meta } = buildCoordsMapFromRows(rows || []);
    state.coordsByPcode = mp;
    state.coordsMeta = meta;

    const badge = $("coordsBadge");
    const text = $("coordsDetectedText");

    if (badge) {
      badge.textContent = meta?.mapped ? `${meta.mapped.toLocaleString()} pts` : "missing";
      badge.classList.toggle("text-bg-warning", !meta?.mapped);
      badge.classList.toggle("text-bg-light", !!meta?.mapped);
    }

    if (text) {
      if (meta?.mapped) {
        text.innerHTML = `Detected: <code>${escapeHtml(meta.latKey)}</code> / <code>${escapeHtml(meta.lngKey)}</code><br>` +
          `<span class="text-body-secondary">Mapped ${meta.mapped.toLocaleString()} of ${meta.total.toLocaleString()} records.</span>`;
      } else {
        text.textContent = "No usable Latitude/Longitude columns detected (or values are empty).";
      }
    }
  }

  async function loadBoundaryFromFile(file) {
    if (!file || !state.map) return;
    const name = file.name || "boundary";

    const text = await file.text();
    const geojson = JSON.parse(text);

    if (!geojson || !geojson.type) throw new Error("Invalid GeoJSON file.");

    // Remove previous
    if (state.mapBoundaryLayer) {
      try { state.map.removeLayer(state.mapBoundaryLayer); } catch (_) {}
      state.mapBoundaryLayer = null;
    }

    state.mapBoundaryLayer = L.geoJSON(geojson, {
      style: () => ({
        color: getTheme() === "dark" ? "rgba(255,255,255,.55)" : "rgba(108,117,125,.85)",
        weight: 2,
        fillOpacity: 0.05,
      }),
    }).addTo(state.map);

    const badge = $("boundaryBadge");
    if (badge) badge.textContent = name;

    try {
      const b = state.mapBoundaryLayer.getBounds();
      if (b && b.isValid && b.isValid()) state.map.fitBounds(b, { padding: [12, 12] });
    } catch (_) {}
  }

  function hookMapControls() {
    $("fitLebanonBtn")?.addEventListener("click", () => {
      if (!state.map) return;
      state.map.fitBounds(LEBANON_BOUNDS, { padding: [12, 12] });
    });

    $("clearMapBtn")?.addEventListener("click", () => {
      if (!state.map) return;
      // Clear only the dynamic layer; keep optional boundaries
      clearMapLayer();
      buildLegend({});
      $("mapStatusText").textContent = "Cleared. Change filters or reload coordinates to draw markers again.";
      setMapEmptyHintVisible(!(state.coordsByPcode && state.coordsByPcode.size));
    });

    $("mapColorBy")?.addEventListener("change", () => updateMap(state.filtered || []));
    $("toggleCluster")?.addEventListener("change", () => updateMap(state.filtered || []));

    $("boundaryInput")?.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await loadBoundaryFromFile(file);
      } catch (err) {
        state.errorCount += 1;
        setHealth("Boundary load failed: " + err.message, false);
      }
    });
  }

// -----------------------------
  // Table
  // -----------------------------
  function exportColumns() {
    // A safe, mostly-non-PII set + QC flags (raw exports can still contain sensitive fields, so we keep this conservative)
    return [
      "PCode",
      "PCode Name",
      "Governorate",
      "District",
      "Cadaster",
      "Local Name",
      "Site Status",
      "Phone call status",
      "No response details",
      "Are you still living in this site",
      "Record status",
      "Total number of Structures",
      "Total number of Households",
      "Total number of Individuals",
      "Number of Latrines",
      "QC - Any issue",
      "QC - Issue count",
      "__qcFlagsText",
    ];
  }

  function sanitizeForExport(r) {
    const out = { ...r };
    // Add readable QC flags
    out.__qcFlagsText = (r.__qcFlags || []).join("; ");
    return out;
  }

  function buildTable(rows) {
    const data = rows.map((r) => sanitizeForExport(r));

    const baseCols = [
      { title: "PCode", field: "PCode", width: 120, headerFilter: true },
      { title: "PCode Name", field: "PCode Name", widthGrow: 2, headerFilter: true },
      { title: "District", field: "District", width: 120, headerFilter: true },
      { title: "Cadaster", field: "Cadaster", width: 160, headerFilter: true },
      { title: "Site Status", field: "Site Status", width: 130, headerFilter: true },
      { title: "Phone status", field: "Phone call status", widthGrow: 2, headerFilter: true },
      { title: "Record status", field: "Record status", width: 140, headerFilter: true },
      { title: "HH", field: "Total number of Households", hozAlign: "right", width: 90 },
      { title: "IND", field: "Total number of Individuals", hozAlign: "right", width: 90 },
      { title: "QC issues", field: "QC - Issue count", hozAlign: "right", width: 95 },
      { title: "QC flags", field: "__qcFlagsText", widthGrow: 3 },
    ];

    const piiCols = [
      { title: "Current Shawish Name", field: "Current Shawish Name", visible: false },
      { title: "Current Shawish Phone", field: "Current Shawish Phone", visible: false },
      { title: "New focal point name", field: "Name of the new focal point in the site", visible: false },
      { title: "New focal point phone", field: "Phone number of the new focal point in the site", visible: false },
    ];

    const columns = [...baseCols, ...piiCols];

    if (!state.table) {
      state.table = new Tabulator("#recordsTable", {
        data,
        layout: "fitColumns",
        height: 520,
        reactiveData: false,
        pagination: true,
        paginationSize: 50,
        selectable: true,
        columns,
      });
    } else {
      state.table.replaceData(data);
    }

    // Toggle PII
    $("togglePII").onchange = (e) => {
      const show = e.target.checked;
      for (const col of piiCols) {
        if (show) state.table.showColumn(col.field);
        else state.table.hideColumn(col.field);
      }
    };

    $("downloadTableBtn").onclick = async () => {
      const cols = exportColumns();
      const currentData = state.filtered.map((r) => sanitizeForExport(r));
      downloadText(`records_${stampFile()}.csv`, toCSV(currentData, cols), "text/csv");
    };
  }

  // -----------------------------
  // Export filtered CSV from sidebar
  // -----------------------------
  function stampFile() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  }

  function hookExports() {
    $("downloadFilteredBtn").addEventListener("click", () => {
      const cols = exportColumns();
      const data = state.filtered.map((r) => sanitizeForExport(r));
      downloadText(`filtered_${stampFile()}.csv`, toCSV(data, cols), "text/csv");
    });
  }

  // -----------------------------
  // Health UI
  // -----------------------------
  function setHealth(message, ok) {
    $("healthMessage").textContent = message;
    $("healthLastLoad").textContent = state.lastLoadAt ? state.lastLoadAt.toLocaleString() : "—";
    $("healthLastSuccess").textContent = state.lastSuccessAt ? state.lastSuccessAt.toLocaleString() : "—";
    $("healthErrors").textContent = String(state.errorCount);
  }

  function setDatasetBadge(label, sheetName) {
    $("datasetBadge").textContent = sheetName ? `${label} • ${sheetName}` : label;

    const refreshed = state.lastSuccessAt ? `Refreshed: ${state.lastSuccessAt.toLocaleString()}` : "—";
    const srcMod = (state.sourceUrl === API_XLSX_ENDPOINT && state.lastApiStatus?.lastModifiedDateTime)
      ? `Source modified: ${new Date(state.lastApiStatus.lastModifiedDateTime).toLocaleString()} • `
      : "";
    $("lastUpdatedText").textContent = refreshed === "—" ? "—" : (srcMod + refreshed);
  }

  // -----------------------------
  // Update pipeline
  // -----------------------------
  function onDataLoaded(rows, label, sheetName) {
    state.raw = rows;
    state.lastSuccessAt = new Date();
    state.errorCount = 0;
    setHealth("Loaded successfully.", true);
    setDatasetBadge(label, sheetName);

    // Coordinates are read directly from the main spreadsheet (Latitude/Longitude columns).
    setCoordsFromSpreadsheet(rows);

    populateFilterControls(rows);
    applyQueryToFilters(); // so shareable URLs work after reload
    updateAll();
  }

  function updateAll() {
    readFiltersFromUI();
    state.filtered = applyFilters(state.raw);

    // KPIs + charts + quality + table
    updateKpis(state.filtered, state.raw);
    updateCharts(state.filtered, state.raw);
    updateQualitySection(state.filtered);
    buildTable(state.filtered);
    updateMap(state.filtered);
  }

  // -----------------------------
  // Auto refresh
  // -----------------------------
  function applyRefresh() {
    const mins = toNum($("refreshMinutes").value);
    if (!state.sourceUrl) {
      state.refreshEveryMs = 0;
      if (state.refreshTimer) clearInterval(state.refreshTimer);
      state.refreshTimer = null;
      $("refreshStatus").textContent = "Auto-refresh: off (no URL loaded)";
      return;
    }

    state.refreshEveryMs = Math.max(1, mins) * 60 * 1000;
    if (state.refreshTimer) clearInterval(state.refreshTimer);
    state.refreshTimer = setInterval(async () => {
      try {
        await loadFromUrl(state.sourceUrl);
      } catch (e) {
        state.errorCount += 1;
        setHealth("Auto-refresh error: " + e.message, false);
      }
    }, state.refreshEveryMs);

    const modeLabel = state.sourceUrl === API_XLSX_ENDPOINT ? "API" : "URL";
    $("refreshStatus").textContent = `Auto-refresh: every ${mins} min (${modeLabel})`;
  }

  // -----------------------------
  // Share link
  // -----------------------------
  async function copyShareLink() {
    const q = encodeFiltersToQuery();
    const url = new URL(window.location.href);
    url.search = q ? "?" + q : "";
    await navigator.clipboard.writeText(url.toString());
    $("copyShareLinkBtn").innerHTML = '<i class="bi bi-check2"></i> Copied';
    setTimeout(() => {
      $("copyShareLinkBtn").innerHTML = '<i class="bi bi-link-45deg"></i> Share link';
    }, 1200);
  }

  // -----------------------------
  // Chart download buttons
  // -----------------------------
  function hookChartDownloads() {
    document.body.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-download-chart]");
      if (!btn) return;
      const id = btn.getAttribute("data-download-chart");
      const chart = state.charts[id];
      if (!chart) return;
      const dataUrl = chart.toBase64Image();
      const a = document.createElement("a");
      a.href = dataUrl;
      a.download = `${id}_${stampFile()}.png`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    });
  }

  // -----------------------------
  // Event wiring
  // -----------------------------
  function hookFilters() {
    ["searchInput", "districtFilter", "cadasterFilter", "siteStatusFilter", "phoneStatusFilter", "qcFilter"]
      .forEach((id) => $(id).addEventListener("input", () => updateAll()));
    $("resetFiltersBtn").addEventListener("click", resetFilters);
  }

  function hookDataPanel() {
    // Live mode wiring
    const liveToggle = $("liveToggle");
    if (liveToggle) {
      liveToggle.addEventListener("change", async (e) => {
        const enabled = !!e.target.checked;
        await enableLiveMode(enabled);
      });

      // Initialize UI state from localStorage (actual loading happens in init())
      const saved = localStorage.getItem("liveMode") === "1";
      setLiveUi(saved);
      setApiStatus(saved ? "Live mode is enabled. Loading will start on page load." : "Live mode off.", true);
    }

    $("testApiBtn")?.addEventListener("click", async () => {
      try {
        await testApiConnection();
      } catch (err) {
        setApiStatus("Test failed: " + err.message, false);
      }
    });

    $("reloadApiBtn")?.addEventListener("click", async () => {
      try {
        await loadFromUrl(API_XLSX_ENDPOINT, "LIVE: SharePoint (Vercel)");
        state.sourceUrl = API_XLSX_ENDPOINT;
        applyRefresh();
        setApiStatus("Reloaded from API.", true);
      } catch (err) {
        setApiStatus("Reload failed: " + err.message, false);
      }
    });

    $("fileInput").addEventListener("change", async (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      try {
        await loadFromFile(file);
      } catch (err) {
        hideLoading();
        state.errorCount += 1;
        setHealth("Load error: " + err.message, false);
      }
    });

    $("loadUrlBtn").addEventListener("click", async () => {
      const url = $("urlInput").value.trim();
      if (!url) return;
      try {
        await loadFromUrl(url);
        applyRefresh();
      } catch (err) {
        hideLoading();
        state.errorCount += 1;
        setHealth("Load URL error: " + err.message, false);
      }
    });

    $("applyRefreshBtn").addEventListener("click", applyRefresh);

    $("loadSampleBtn").addEventListener("click", async () => {
      try {
        await loadSample();
      } catch (err) {
        hideLoading();
        state.errorCount += 1;
        setHealth("Sample load error: " + err.message, false);
      }
    });
  }

  function hookMisc() {
    $("copyShareLinkBtn").addEventListener("click", () => copyShareLink());
    $("scrollTopLink").addEventListener("click", (e) => {
      e.preventDefault();
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  }

  // -----------------------------
  // Boot
  // -----------------------------
  async function init() {
    initTheme();
    hookFilters();
    hookExports();
    hookDataPanel();
    hookChartDownloads();
    hookMisc();
    initMap();
    hookMapControls();

    const params = new URLSearchParams(window.location.search);

    // Live mode (from localStorage, or forced with ?live=1)
    const forceLive = params.get("live") === "1";
    const forceNoLive = params.get("noLive") === "1";
    const savedLive = localStorage.getItem("liveMode") === "1";

    if (!forceNoLive && (forceLive || savedLive)) {
      try {
        await enableLiveMode(true);
        // If live mode successfully loaded data, stop here.
        if (state.raw.length) return;
      } catch (err) {
        // If something unexpected happens, fall back to sample/file.
        setApiStatus("Live mode error: " + err.message, false);
      }
    }

    // Load sample by default unless ?noSample=1
    const noSample = params.get("noSample") === "1";
    if (!noSample) {
      try {
        await loadSample();
      } catch (err) {
        hideLoading();
        state.errorCount += 1;
        setHealth("Auto-load sample failed: " + err.message, false);
      }
    } else {
      setHealth("Ready. Load a file or URL from Data panel.", true);
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();