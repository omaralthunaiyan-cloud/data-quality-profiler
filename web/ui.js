/* ---------------------------------------------------------------------
   Data Quality Profiler — UI wiring
------------------------------------------------------------------------ */

const TYPE_CLASS = {
  email: "type-contact", url: "type-contact", phone: "type-contact",
  saudi_id: "type-id",
  date: "type-temporal",
  integer: "type-numeric", float: "type-numeric", currency: "type-numeric", percent: "type-numeric",
  categorical: "type-category", boolean: "type-category",
  free_text: "type-text",
};

const DIMENSION_COLOR_VAR = {
  Completeness: "--type-contact",
  Validity: "--type-id",
  Uniqueness: "--type-category",
  Outliers: "--type-temporal",
};

const TYPE_LABEL = {
  email: "email", url: "url", phone: "phone", saudi_id: "national ID",
  date: "date", integer: "integer", float: "decimal", currency: "currency",
  percent: "percent", categorical: "categorical", boolean: "boolean", free_text: "free text",
};

const els = {
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  sampleBtns: document.querySelectorAll("[data-sample]"),
  emptyState: document.getElementById("empty-state"),
  report: document.getElementById("report"),
  fileLabel: document.getElementById("file-label"),
  statScore: document.getElementById("stat-score"),
  scoreRingFill: document.getElementById("score-ring-fill"),
  scoreHeroSummary: document.getElementById("score-hero-summary"),
  statRows: document.getElementById("stat-rows"),
  statCols: document.getElementById("stat-cols"),
  statDupes: document.getElementById("stat-dupes"),
  statAttention: document.getElementById("stat-attention"),
  statAvgCompleteness: document.getElementById("stat-avg-completeness"),
  statAvgValidity: document.getElementById("stat-avg-validity"),
  statAvgUniqueness: document.getElementById("stat-avg-uniqueness"),
  statMissing: document.getElementById("stat-missing"),
  worstColumn: document.getElementById("worst-column"),
  bestColumn: document.getElementById("best-column"),
  dimBars: document.getElementById("dim-bars"),
  donutWrap: document.getElementById("donut-wrap"),
  donut: document.getElementById("donut"),
  donutLegend: document.getElementById("donut-legend"),
  printMeta: document.getElementById("print-meta"),
  exportPdfBtn: document.getElementById("export-pdf-btn"),
  sheetPicker: document.getElementById("sheet-picker"),
  sheetSelect: document.getElementById("sheet-select"),
  sheetNameBadge: document.getElementById("sheet-name-badge"),
  barChart: document.getElementById("bar-chart"),
  columnList: document.getElementById("column-list"),
  downloadBtn: document.getElementById("download-btn"),
  errorBox: document.getElementById("error-box"),
};

let currentReport = null;
let currentFileLabel = "";
let currentWorkbook = null;   // parsed SheetJS workbook, kept around so switching sheets doesn't need a re-read
let currentBaseFileName = ""; // original file name, without the " — SheetName" suffix
let currentSheetName = "";

function scoreLevel(score) {
  if (score >= 90) return "good";
  if (score >= 75) return "warn";
  return "bad";
}

// Builds a filesystem-safe slug for downloaded filenames, e.g.
// "employees.xlsx" + "Q1 Sales" -> "employees-Q1_Sales". Built from the
// underlying file/sheet names directly (not by parsing currentFileLabel's
// display text) so the sheet name can't collide with a real "." in the
// original filename.
function safeFilenameSlug(fileName, sheetName) {
  const base = (fileName || "report").replace(/\.[^./\\]+$/, "");
  const parts = [base];
  if (sheetName) parts.push(sheetName);
  return parts.join("-")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "") || "report";
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function workbookToRows(workbook, sheetName) {
  sheetName = sheetName || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const headers = rows.length
    ? Object.keys(rows[0])
    : (XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || []);
  return { headers, rows };
}

// Sheets that have no header row / no data at all still show up in the
// picker (so the count matches what's really in the workbook), but we skip
// straight past them when picking the sheet to profile by default.
function sheetLooksUsable(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  const preview = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  return preview.length > 0 && preview[0].length > 0;
}

function populateSheetPicker(workbook, selectedName) {
  const names = workbook.SheetNames;
  if (names.length <= 1) {
    els.sheetPicker.hidden = true;
    els.sheetSelect.innerHTML = "";
    return;
  }
  els.sheetPicker.hidden = false;
  els.sheetSelect.innerHTML = names.map((name) =>
    `<option value="${escapeHtml(name)}"${name === selectedName ? " selected" : ""}>${escapeHtml(name)}</option>`
  ).join("");
}

function profileWorkbookSheet(sheetName) {
  if (!currentWorkbook) return;
  try {
    const { headers, rows } = workbookToRows(currentWorkbook, sheetName);
    const isMultiSheet = currentWorkbook.SheetNames.length > 1;
    currentSheetName = isMultiSheet ? sheetName : "";
    // Fold the sheet name into the label so it automatically flows through
    // to the CSV/PDF export filenames and report titles, not just the
    // on-screen badge.
    const label = isMultiSheet ? `${currentBaseFileName} — ${sheetName}` : currentBaseFileName;
    runProfilerOn(headers, rows, label, isMultiSheet ? sheetName : "");
  } catch (err) {
    showError("Couldn't read that sheet. It may be empty, hidden, or in an unsupported format.");
  }
}

function showError(msg) {
  els.errorBox.textContent = msg;
  els.errorBox.hidden = !msg;
}

function runProfilerOn(headers, rows, label, sheetName) {
  showError("");
  if (!headers.length || !rows.length) {
    showError(sheetName
      ? `The "${sheetName}" sheet has no rows. Make sure its first row has column headers, or pick a different sheet.`
      : "Couldn't find any rows in this file. Make sure the first row has column headers.");
    return;
  }
  currentReport = profileDataset(headers, rows);
  currentFileLabel = label;
  render();
}

function loadSample(name) {
  currentWorkbook = null;
  currentBaseFileName = "";
  currentSheetName = "";
  els.sheetPicker.hidden = true;
  const text = name === "employees" ? SAMPLE_EMPLOYEES_CSV : SAMPLE_PROJECTS_CSV;
  const { headers, rows } = parseCsv(text);
  runProfilerOn(headers, rows, name === "employees" ? "employees.csv (sample)" : "projects.csv (sample)");
}

function loadFile(file) {
  const reader = new FileReader();
  const isCsv = /\.csv$/i.test(file.name);
  reader.onload = (e) => {
    try {
      if (isCsv) {
        // Raw text parser — preserves exactly what's in the file, no
        // implicit date/number reformatting.
        currentWorkbook = null;
        currentBaseFileName = "";
        currentSheetName = "";
        els.sheetPicker.hidden = true;
        const { headers, rows } = parseCsv(e.target.result);
        runProfilerOn(headers, rows, file.name);
      } else {
        // Real .xlsx/.xls: SheetJS is the only practical way to read the
        // binary format, so we use it here (and only here).
        const wb = XLSX.read(e.target.result, { type: "array" });
        if (!wb.SheetNames.length) {
          showError("Couldn't find any sheets in that workbook.");
          return;
        }
        currentWorkbook = wb;
        currentBaseFileName = file.name;
        // default to the first sheet that actually has a header row + data,
        // falling back to the very first sheet if none look usable
        const defaultSheet = wb.SheetNames.find((n) => sheetLooksUsable(wb, n)) || wb.SheetNames[0];
        populateSheetPicker(wb, defaultSheet);
        profileWorkbookSheet(defaultSheet);
      }
    } catch (err) {
      showError("Couldn't read that file. Make sure it's a valid .xlsx, .xls, or .csv file.");
    }
  };
  if (isCsv) reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
}

els.sheetSelect.addEventListener("change", () => {
  profileWorkbookSheet(els.sheetSelect.value);
});

const RING_CIRCUMFERENCE = 2 * Math.PI * 52; // matches the SVG circle's r="52"

function issueChipHtml(primaryIssue, level) {
  if (!primaryIssue || level === "good") return "";
  return `<span class="issue-chip ${level}">${escapeHtml(primaryIssue.dimension)} gap</span>`;
}

function heroSummaryHtml(r) {
  const parts = [];
  if (r.needsAttention === 0) {
    parts.push(`All <strong>${r.columnCount}</strong> columns look healthy — no major issues detected.`);
  } else {
    const level = r.needsAttention / r.columnCount > 0.4 ? "bad" : "warn";
    parts.push(`<span class="chip ${level}">${r.needsAttention} of ${r.columnCount}</span> columns need attention.`);
  }
  if (r.topIssue) {
    parts.push(`Most common driver: <span class="chip warn">${escapeHtml(r.topIssue.dimension)}</span> (${r.topIssue.count} column${r.topIssue.count > 1 ? "s" : ""}).`);
  }
  return parts.join(" ");
}

function renderDonut(counts) {
  const entries = Object.entries(counts || {}).filter(([, n]) => n > 0);
  const total = entries.reduce((sum, [, n]) => sum + n, 0);

  if (total === 0) {
    els.donut.style.background = "var(--surface-2)";
    els.donut.innerHTML = "";
    els.donutLegend.innerHTML = `<p class="donut-empty">No dimension issues detected — every column is within healthy range.</p>`;
    return;
  }

  let offset = 0;
  const stops = [];
  entries.forEach(([dim, n]) => {
    const colorVar = DIMENSION_COLOR_VAR[dim] || "--text-muted";
    const pct = (n / total) * 100;
    stops.push(`var(${colorVar}) ${offset}% ${offset + pct}%`);
    offset += pct;
  });
  els.donut.style.background = `conic-gradient(${stops.join(", ")})`;
  els.donut.innerHTML = "";

  els.donutLegend.innerHTML = entries.map(([dim, n]) => {
    const colorVar = DIMENSION_COLOR_VAR[dim] || "--text-muted";
    return `
      <div class="donut-legend-item">
        <span class="donut-dot" style="background:var(${colorVar})"></span>
        <span>${escapeHtml(dim)}</span>
        <span class="donut-legend-count tabular">${n}</span>
      </div>`;
  }).join("");
}

function render() {
  const r = currentReport;
  els.emptyState.hidden = true;
  els.report.hidden = false;
  els.fileLabel.textContent = currentFileLabel;
  if (currentWorkbook && currentWorkbook.SheetNames.length > 1) {
    const n = currentWorkbook.SheetNames.length;
    els.sheetNameBadge.textContent = `${n} sheets`;
    els.sheetNameBadge.hidden = false;
  } else {
    els.sheetNameBadge.hidden = true;
  }

  const generatedAt = new Date().toLocaleString(undefined, {
    year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit",
  });
  els.printMeta.innerHTML = `
    <div class="print-masthead-brand">
      <span class="print-masthead-mark">DQ</span>
      <strong>Data Quality Profiler — Diagnostic Report</strong>
    </div>
    <div class="print-masthead-meta">
      <span>File: ${escapeHtml(currentFileLabel)}</span>
      <span>Generated: ${escapeHtml(generatedAt)}</span>
    </div>`;

  const overallLevel = scoreLevel(r.overallQualityScore);
  els.statScore.textContent = r.overallQualityScore.toFixed(1);
  els.statScore.className = overallLevel;
  els.scoreRingFill.setAttribute("class", "score-ring-fill " + overallLevel);
  els.scoreRingFill.style.strokeDasharray = String(RING_CIRCUMFERENCE);
  els.scoreRingFill.style.strokeDashoffset = String(RING_CIRCUMFERENCE * (1 - r.overallQualityScore / 100));
  els.scoreHeroSummary.innerHTML = heroSummaryHtml(r);

  els.statRows.textContent = r.rowCount.toLocaleString();
  els.statCols.textContent = r.columnCount;
  els.statDupes.textContent = r.duplicateRows;
  els.statDupes.className = "stat-value " + (r.duplicateRows ? "bad" : "good");
  els.statAttention.textContent = r.needsAttention;
  els.statAttention.className = "stat-value " + (r.needsAttention === 0 ? "good" : r.needsAttention / r.columnCount > 0.4 ? "bad" : "warn");

  els.statAvgCompleteness.textContent = r.avgCompleteness.toFixed(1) + "%";
  els.statAvgCompleteness.className = "stat-value tabular " + scoreLevel(r.avgCompleteness);
  els.statAvgValidity.textContent = r.avgValidity.toFixed(1) + "%";
  els.statAvgValidity.className = "stat-value tabular " + scoreLevel(r.avgValidity);
  els.statAvgUniqueness.textContent = r.avgUniqueness.toFixed(1) + "%";
  els.statAvgUniqueness.className = "stat-value tabular " + scoreLevel(r.avgUniqueness);
  els.statMissing.textContent = r.totalMissingValues.toLocaleString();
  els.statMissing.className = "stat-value tabular " + (r.totalMissingValues === 0 ? "good" : "bad");

  if (r.worstColumn) {
    const wLevel = scoreLevel(r.worstColumn.qualityScore);
    els.worstColumn.innerHTML = `${escapeHtml(r.worstColumn.column)} <span class="tabular">${r.worstColumn.qualityScore.toFixed(0)}/100</span>`;
    els.worstColumn.className = "highlight-value " + wLevel;
  } else {
    els.worstColumn.textContent = "–";
    els.worstColumn.className = "highlight-value";
  }
  if (r.bestColumn) {
    const bLevel = scoreLevel(r.bestColumn.qualityScore);
    els.bestColumn.innerHTML = `${escapeHtml(r.bestColumn.column)} <span class="tabular">${r.bestColumn.qualityScore.toFixed(0)}/100</span>`;
    els.bestColumn.className = "highlight-value " + bLevel;
  } else {
    els.bestColumn.textContent = "–";
    els.bestColumn.className = "highlight-value";
  }

  const dimAverages = [
    { label: "Completeness", value: r.avgCompleteness },
    { label: "Validity", value: r.avgValidity },
    { label: "Uniqueness", value: r.avgUniqueness },
  ];
  els.dimBars.innerHTML = dimAverages.map((d) => {
    const colorVar = DIMENSION_COLOR_VAR[d.label];
    return `
      <div class="dim-bar-row">
        <span class="dim-bar-label">${d.label}</span>
        <div class="dim-bar-track">
          <div class="dim-bar-fill" style="width:${d.value}%;background:var(${colorVar})"></div>
        </div>
        <span class="dim-bar-value tabular">${d.value.toFixed(1)}%</span>
      </div>`;
  }).join("");

  renderDonut(r.dimensionIssueCounts);

  els.barChart.innerHTML = r.columns.map((c) => {
    const level = scoreLevel(c.qualityScore);
    const issueText = c.primaryIssue && level !== "good" ? `${c.primaryIssue.dimension} gap` : "";
    return `
      <div class="bar-row">
        <span class="bar-label" title="${escapeHtml(c.column)}">${escapeHtml(c.column)}</span>
        <div class="bar-track">
          <div class="bar-fill ${level}" style="width:${c.qualityScore}%"></div>
        </div>
        <span class="bar-issue ${level}">${escapeHtml(issueText)}</span>
        <span class="bar-value tabular">${c.qualityScore.toFixed(0)}</span>
      </div>`;
  }).join("");

  els.columnList.innerHTML = r.columns.map((c, i) => {
    const level = scoreLevel(c.qualityScore);
    const typeClass = TYPE_CLASS[c.inferredType] || "type-text";
    const typeLabel = TYPE_LABEL[c.inferredType] || c.inferredType;
    const openAttr = c.qualityScore < 85 ? "open" : "";
    return `
      <details class="col-card" ${openAttr}>
        <summary>
          <span class="severity-dot ${level}"></span>
          <span class="col-name">${escapeHtml(c.column)}</span>
          <span class="type-badge ${typeClass}">${escapeHtml(typeLabel)}</span>
          ${issueChipHtml(c.primaryIssue, level)}
          <span class="col-score tabular ${level}">${c.qualityScore.toFixed(0)}<span class="col-score-max">/100</span></span>
        </summary>
        <div class="col-detail">
          <div class="metric-row">
            <div class="metric"><span class="metric-label">Completeness</span><span class="metric-value tabular">${c.completenessPct.toFixed(1)}%</span></div>
            <div class="metric"><span class="metric-label">Validity</span><span class="metric-value tabular">${c.validityPct.toFixed(1)}%</span></div>
            <div class="metric"><span class="metric-label">Uniqueness</span><span class="metric-value tabular">${c.uniquenessPct.toFixed(1)}%</span></div>
            <div class="metric"><span class="metric-label">Confidence</span><span class="metric-value tabular">${(c.confidence * 100).toFixed(0)}%</span></div>
          </div>
          ${c.primaryIssue ? `<p class="note">Primary driver: <strong>${escapeHtml(c.primaryIssue.dimension)}</strong> — this dimension accounts for most of this column's score deduction.</p>` : ""}
          ${c.notes ? `<p class="note">${escapeHtml(c.notes)}</p>` : ""}
          ${c.sampleInvalid.length ? `
            <div class="samples">
              <span class="samples-label">Sample problem values</span>
              <div class="samples-list">${c.sampleInvalid.map((s) => `<code>${escapeHtml(s || "(empty)")}</code>`).join("")}</div>
            </div>` : ""}
          ${(!c.notes && c.issueCount === 0) ? `<p class="note ok">No issues detected in this column.</p>` : ""}
        </div>
      </details>`;
  }).join("");

  els.downloadBtn.onclick = () => downloadCsv(r);
}

function reportToCsv(r) {
  const cols = ["column", "inferred_type", "confidence", "completeness_pct", "validity_pct",
    "uniqueness_pct", "quality_score", "issue_count", "notes"];
  const lines = [cols.join(",")];
  for (const c of r.columns) {
    const row = [c.column, c.inferredType, c.confidence, c.completenessPct, c.validityPct,
      c.uniquenessPct, c.qualityScore, c.issueCount, c.notes];
    lines.push(row.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","));
  }
  return lines.join("\n");
}

async function downloadCsv(r) {
  const csv = reportToCsv(r);
  const filename = `quality_report_${safeFilenameSlug(currentBaseFileName || currentFileLabel, currentSheetName)}.csv`;
  document.getElementById("download-fallback").hidden = true;

  const downloads = window.__claudeDownloads;
  if (downloads) {
    try {
      await downloads.save({ filename, data: csv });
      return;
    } catch (err) {
      if (err && err.code === "extension_not_enabled") {
        try {
          await downloads.save({ filename: filename.replace(/\.csv$/, ".txt"), data: csv });
          return;
        } catch (err2) { /* fall through to manual copy */ }
      }
      // declined / rate_limited / unavailable / anything else -> fall back below
    }
  }
  showCsvFallback(csv);
}

function showCsvFallback(csv) {
  const box = document.getElementById("download-fallback");
  const textarea = document.getElementById("download-fallback-text");
  textarea.value = csv;
  box.hidden = false;
  textarea.focus();
  textarea.select();
}

// --- wiring ---
els.fileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) loadFile(e.target.files[0]);
});

els.dropzone.addEventListener("click", () => els.fileInput.click());
els.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); els.fileInput.click(); }
});
els.dropzone.addEventListener("dragover", (e) => { e.preventDefault(); els.dropzone.classList.add("dragging"); });
els.dropzone.addEventListener("dragleave", () => els.dropzone.classList.remove("dragging"));
els.dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  els.dropzone.classList.remove("dragging");
  if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
});

els.sampleBtns.forEach((btn) => {
  btn.addEventListener("click", () => loadSample(btn.dataset.sample));
});

document.getElementById("fallback-close").addEventListener("click", () => {
  document.getElementById("download-fallback").hidden = true;
});

// The Artifact viewer runs this page inside a sandboxed frame that silently
// swallows window.print() (same restriction that blocks direct <a download>
// clicks — see downloadCsv() above). So instead of printing in-place, we
// package a standalone, print-ready DOCUMENT and hand it to the viewer via
// the downloads capability. That file is a plain, unsandboxed HTML document,
// so when the person opens it, its own onload handler can call
// window.print() for real and the browser's native "Save as PDF" dialog
// appears.
//
// This export is deliberately its own design, not a screenshot of the
// dashboard: a dashboard is scanned and clicked through; a PDF is read
// top to bottom and often skimmed once, printed, or forwarded. So it reads
// like an analyst's report — a serif editorial voice, a real findings
// table instead of a stack of repeated cards, and one page-length verdict
// up top — rather than a dump of the app's UI chrome.

const REPORT_DIM_COLOR = {
  Completeness: "#2568c9",
  Validity: "#b2318f",
  Uniqueness: "#b3690b",
  Outliers: "#7a3fd6",
};

function verdictLabel(score) {
  if (score >= 90) return "Excellent";
  if (score >= 75) return "Needs attention";
  return "Poor";
}

function reportExecutiveTakeaway(r) {
  if (r.needsAttention === 0) {
    return `All ${r.columnCount} columns fall within a healthy range — this dataset is ready to use as-is.`;
  }
  const share = r.needsAttention / r.columnCount > 0.4 ? "a substantial share of" : "a handful of";
  return `${share[0].toUpperCase()}${share.slice(1)} the columns need attention before this dataset is production-ready — ${r.needsAttention} of ${r.columnCount} in total${r.topIssue ? `, most often flagged for ${r.topIssue.dimension.toLowerCase()} gaps` : ""}.`;
}

function reportExecutiveBody(r, fileLabel) {
  const sentences = [];
  sentences.push(`This report profiles <strong>${escapeHtml(fileLabel)}</strong> — ${r.rowCount.toLocaleString()} rows across ${r.columnCount} columns — inferring each column's data type automatically and scoring it across four dimensions: completeness, validity, uniqueness, and outliers.`);
  if (r.worstColumn && r.bestColumn && r.needsAttention > 0) {
    sentences.push(`<strong>${escapeHtml(r.worstColumn.column)}</strong> is the weakest column at ${r.worstColumn.qualityScore.toFixed(0)}/100; <strong>${escapeHtml(r.bestColumn.column)}</strong> is the strongest at ${r.bestColumn.qualityScore.toFixed(0)}/100.`);
  }
  sentences.push(`${r.totalMissingValues.toLocaleString()} missing value${r.totalMissingValues === 1 ? "" : "s"} and ${r.duplicateRows.toLocaleString()} duplicate row${r.duplicateRows === 1 ? "" : "s"} were found across the dataset. Full detector methodology is in the appendix below.`);
  return sentences.join(" ");
}

function reportMeterHtml(score, level) {
  return `
    <div class="meter">
      <div class="meter-track">
        <div class="meter-zone bad"></div><div class="meter-zone warn"></div><div class="meter-zone good"></div>
        <div class="meter-marker ${level}" style="left:${score}%"></div>
      </div>
      <div class="meter-scale"><span>0</span><span>50</span><span>75</span><span>90</span><span>100</span></div>
    </div>`;
}

function reportDimBarsHtml(r) {
  const dims = [
    { label: "Completeness", value: r.avgCompleteness },
    { label: "Validity", value: r.avgValidity },
    { label: "Uniqueness", value: r.avgUniqueness },
  ];
  return dims.map((d) => `
      <div class="rbar-row">
        <span class="rbar-label">${d.label}</span>
        <div class="rbar-track"><div class="rbar-fill" style="width:${d.value}%;background:${REPORT_DIM_COLOR[d.label]}"></div></div>
        <span class="rbar-value">${d.value.toFixed(1)}%</span>
      </div>`).join("");
}

function reportColumnBarsHtml(r) {
  const sorted = [...r.columns].sort((a, b) => a.qualityScore - b.qualityScore);
  return sorted.map((c) => {
    const level = scoreLevel(c.qualityScore);
    return `
      <div class="rbar-row">
        <span class="rbar-label" title="${escapeHtml(c.column)}">${escapeHtml(c.column)}</span>
        <div class="rbar-track"><div class="rbar-fill ${level}" style="width:${c.qualityScore}%"></div></div>
        <span class="rbar-value">${c.qualityScore.toFixed(0)}</span>
      </div>`;
  }).join("");
}

function reportDonutHtml(r) {
  const entries = Object.entries(r.dimensionIssueCounts || {}).filter(([, n]) => n > 0);
  const total = entries.reduce((sum, [, n]) => sum + n, 0);
  if (total === 0) {
    return `<div class="rdonut" style="background:#eee"></div><p class="rdonut-empty">No dimension issues detected.</p>`;
  }
  let offset = 0;
  const stops = [];
  entries.forEach(([dim, n]) => {
    const color = REPORT_DIM_COLOR[dim] || "#999";
    const pct = (n / total) * 100;
    stops.push(`${color} ${offset}% ${offset + pct}%`);
    offset += pct;
  });
  const legend = entries.map(([dim, n]) => {
    const color = REPORT_DIM_COLOR[dim] || "#999";
    return `<div class="rdonut-legend-item"><span class="rdonut-dot" style="background:${color}"></span><span>${escapeHtml(dim)}</span><span class="rdonut-count">${n}</span></div>`;
  }).join("");
  return `<div class="rdonut" style="background:conic-gradient(${stops.join(", ")})"></div><div class="rdonut-legend">${legend}</div>`;
}

function reportFindingsTableHtml(r) {
  const sorted = [...r.columns].sort((a, b) => a.qualityScore - b.qualityScore);
  const rows = sorted.map((c) => {
    const level = scoreLevel(c.qualityScore);
    const typeLabel = TYPE_LABEL[c.inferredType] || c.inferredType;
    return `
      <tr>
        <td><span class="rdot ${level}"></span>${escapeHtml(c.column)}</td>
        <td class="rmuted">${escapeHtml(typeLabel)}</td>
        <td class="rnum ${level}">${c.qualityScore.toFixed(0)}</td>
        <td class="rmuted">${c.primaryIssue ? escapeHtml(c.primaryIssue.dimension) : "—"}</td>
        <td class="rnum rmuted">${c.completenessPct.toFixed(0)}%</td>
        <td class="rnum rmuted">${c.validityPct.toFixed(0)}%</td>
        <td class="rnum rmuted">${c.uniquenessPct.toFixed(0)}%</td>
      </tr>`;
  }).join("");
  return `
    <table class="rtable">
      <thead><tr>
        <th>Column</th><th>Type</th><th>Score</th><th>Primary issue</th>
        <th>Completeness</th><th>Validity</th><th>Uniqueness</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
}

function reportNotesHtml(r) {
  const flagged = r.columns.filter((c) => c.notes || c.sampleInvalid.length);
  if (!flagged.length) return `<p class="rnote-empty">No columns produced additional notes — nothing further to flag.</p>`;
  return `<ul class="rnotes">${flagged.map((c) => `
      <li>
        <strong>${escapeHtml(c.column)}</strong> — ${escapeHtml(c.notes || "flagged sample values below.")}
        ${c.sampleInvalid.length ? `<span class="rnote-samples">${c.sampleInvalid.slice(0, 4).map((s) => `<code>${escapeHtml(s || "(empty)")}</code>`).join(" ")}</span>` : ""}
      </li>`).join("")}</ul>`;
}

function buildStandaloneReportHtml() {
  const r = currentReport;
  const overallLevel = scoreLevel(r.overallQualityScore);
  const generatedAt = new Date().toLocaleString(undefined, {
    year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit",
  });

  const css = `
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; }
    body {
      background: #fdfcfb;
      color: #1b1730;
      font-family: "Source Sans 3", "Manrope", system-ui, sans-serif;
      font-size: 14.5px;
      line-height: 1.55;
      -webkit-print-color-adjust: exact; print-color-adjust: exact;
    }
    .rdoc { max-width: 860px; margin: 0 auto; padding: 46px 30px 70px; }
    h1, h2, h3 { font-family: "Source Serif 4", Georgia, serif; margin: 0; text-wrap: balance; }
    .tabular, .rnum, .rbar-value { font-variant-numeric: tabular-nums; font-feature-settings: "tnum"; }

    .rmast { display: flex; align-items: center; justify-content: space-between; gap: 16px; padding-bottom: 18px; border-bottom: 3px solid #1b1730; margin-bottom: 30px; flex-wrap: wrap; }
    .rmast-brand { display: flex; align-items: center; gap: 10px; }
    .rmast-mark { width: 32px; height: 32px; border-radius: 8px; background: linear-gradient(135deg, #6a35f2, #00967e); color: #fff; display: flex; align-items: center; justify-content: center; font-family: "Source Serif 4", serif; font-weight: 700; font-size: 13px; }
    .rmast-brand span { font-size: 12.5px; letter-spacing: 0.08em; text-transform: uppercase; color: #6a6280; font-weight: 600; }
    .rmast-meta { text-align: right; font-size: 12px; color: #6a6280; }
    .rmast-meta strong { color: #1b1730; }

    .rtitle { font-size: 30px; font-weight: 700; margin-bottom: 6px; }
    .rsubtitle { font-size: 14.5px; color: #6a6280; margin-bottom: 34px; }
    .rsubtitle code { font-family: "JetBrains Mono", monospace; background: #efeaf9; border-radius: 4px; padding: 1px 6px; font-size: 0.92em; }

    .rverdict { display: flex; gap: 30px; align-items: center; padding: 26px 28px; border: 1.5px solid #1b1730; border-radius: 4px; margin-bottom: 8px; break-inside: avoid; }
    .rverdict-score { flex-shrink: 0; text-align: center; min-width: 108px; }
    .rverdict-score .num { font-family: "Source Serif 4", serif; font-size: 52px; font-weight: 700; line-height: 1; }
    .rverdict-score .num.good { color: #1f9d5c; } .rverdict-score .num.warn { color: #b3690b; } .rverdict-score .num.bad { color: #d33a5c; }
    .rverdict-score .max { font-size: 12px; color: #6a6280; }
    .rverdict-score .tag { display: block; margin-top: 4px; font-size: 11px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; }
    .rverdict-score .tag.good { color: #1f9d5c; } .rverdict-score .tag.warn { color: #b3690b; } .rverdict-score .tag.bad { color: #d33a5c; }
    .rverdict-body { flex: 1; }
    .rverdict-body p { margin: 0 0 12px; font-size: 15px; }
    .rverdict-body p:last-child { margin-bottom: 0; }

    .meter-track { position: relative; height: 8px; border-radius: 5px; overflow: hidden; display: flex; }
    .meter-zone { height: 100%; }
    .meter-zone.bad { width: 75%; background: #f6d4dc; }
    .meter-zone.warn { width: 15%; background: #f7e3c4; }
    .meter-zone.good { width: 10%; background: #cdeddb; }
    .meter-marker { position: absolute; top: -3px; width: 3px; height: 14px; border-radius: 2px; background: #1b1730; transform: translateX(-1.5px); }
    .meter-marker.good { background: #1f9d5c; } .meter-marker.warn { background: #b3690b; } .meter-marker.bad { background: #d33a5c; }
    .meter-scale { display: flex; justify-content: space-between; font-size: 10px; color: #a49dbd; margin-top: 4px; font-family: "JetBrains Mono", monospace; }

    .rsection { margin-top: 40px; break-inside: avoid-page; }
    .rsection-head { display: flex; align-items: baseline; gap: 10px; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 1.5px solid #1b1730; }
    .rsection-head .rn { font-family: "JetBrains Mono", monospace; font-size: 12px; color: #6a35f2; font-weight: 700; }
    .rsection-head h2 { font-size: 18px; }
    .rsection-note { font-size: 11.5px; color: #6a6280; font-weight: 400; font-family: "Source Sans 3", sans-serif; margin-left: 4px; }

    .rmetric-strip { display: flex; }
    .rmetric { flex: 1; padding: 0 18px; }
    .rmetric:first-child { padding-left: 0; }
    .rmetric:not(:first-child) { border-left: 1px solid #e3ddf2; }
    .rmetric .v { font-family: "Source Serif 4", serif; font-size: 26px; font-weight: 700; }
    .rmetric .v.good { color: #1f9d5c; } .rmetric .v.warn { color: #b3690b; } .rmetric .v.bad { color: #d33a5c; }
    .rmetric .l { display: block; font-size: 11.5px; color: #6a6280; margin-top: 2px; }

    .rtwocol { display: flex; gap: 40px; }
    .rtwocol > div { flex: 1; }
    .rhighlight { font-size: 12.5px; margin-top: 14px; color: #4a4460; }
    .rhighlight b { color: #1b1730; }

    .rbar-row { display: grid; grid-template-columns: 128px 1fr 34px; align-items: center; gap: 10px; margin-bottom: 9px; font-size: 12px; }
    .rbar-label { color: #4a4460; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .rbar-track { background: #eeeaf8; border-radius: 5px; height: 8px; overflow: hidden; }
    .rbar-fill { height: 100%; border-radius: 5px; background: #6a35f2; }
    .rbar-fill.good { background: #1f9d5c; } .rbar-fill.warn { background: #b3690b; } .rbar-fill.bad { background: #d33a5c; }
    .rbar-value { text-align: right; color: #6a6280; font-family: "JetBrains Mono", monospace; }

    .rroot { display: flex; gap: 30px; align-items: center; }
    .rroot-chart { display: flex; align-items: center; gap: 20px; }
    .rdonut { width: 108px; height: 108px; border-radius: 50%; position: relative; flex-shrink: 0; }
    .rdonut::after { content: ""; position: absolute; inset: 20px; border-radius: 50%; background: #fdfcfb; }
    .rdonut-empty { font-size: 12px; color: #6a6280; }
    .rdonut-legend { display: flex; flex-direction: column; gap: 6px; font-size: 12px; }
    .rdonut-legend-item { display: flex; align-items: center; gap: 7px; }
    .rdonut-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
    .rdonut-count { color: #6a6280; font-family: "JetBrains Mono", monospace; margin-left: 4px; }
    .rroot-takeaway { flex: 1; font-size: 13px; color: #4a4460; }

    .rtable { width: 100%; border-collapse: collapse; font-size: 12.5px; }
    .rtable th { text-align: left; font-size: 10.5px; letter-spacing: 0.05em; text-transform: uppercase; color: #6a6280; font-weight: 600; padding: 0 10px 8px; border-bottom: 1.5px solid #1b1730; }
    .rtable td { padding: 8px 10px; border-bottom: 1px solid #e9e4f5; }
    .rtable th:not(:first-child), .rtable td:not(:first-child) { text-align: right; }
    .rtable tbody tr:nth-child(even) { background: #f7f4fc; }
    .rtable .rnum { font-family: "JetBrains Mono", monospace; font-weight: 600; }
    .rtable .rnum.good { color: #1f9d5c; } .rtable .rnum.warn { color: #b3690b; } .rtable .rnum.bad { color: #d33a5c; }
    .rtable .rmuted { color: #6a6280; }
    .rdot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; margin-right: 8px; }
    .rdot.good { background: #1f9d5c; } .rdot.warn { background: #b3690b; } .rdot.bad { background: #d33a5c; }

    .rnotes { list-style: none; margin: 0; padding: 0; font-size: 12.5px; color: #4a4460; }
    .rnotes li { padding: 10px 0; border-bottom: 1px solid #e9e4f5; }
    .rnotes li:last-child { border-bottom: none; }
    .rnote-samples { display: block; margin-top: 5px; }
    .rnote-samples code { font-family: "JetBrains Mono", monospace; font-size: 11px; background: #efeaf9; border-radius: 4px; padding: 1px 6px; margin-right: 5px; display: inline-block; margin-top: 3px; }
    .rnote-empty { font-size: 12.5px; color: #6a6280; font-style: italic; }

    .rfooter { margin-top: 46px; padding-top: 16px; border-top: 1px solid #e3ddf2; font-size: 11px; color: #a49dbd; display: flex; justify-content: space-between; }

    @media print {
      @page { margin: 15mm; }
      .rsection { break-inside: auto; }
      .rsection-head, .rverdict, .rtable tr { break-inside: avoid; }
    }
  `;

  const bodyHtml = `
    <div class="rdoc">
      <div class="rmast">
        <div class="rmast-brand"><span class="rmast-mark">DQ</span><span>Data Quality Profiler</span></div>
        <div class="rmast-meta"><strong>${escapeHtml(currentFileLabel)}</strong><br>Generated ${escapeHtml(generatedAt)}</div>
      </div>

      <h1 class="rtitle">Data Quality Assessment</h1>
      <p class="rsubtitle">${r.rowCount.toLocaleString()} rows · ${r.columnCount} columns · profiled from <code>${escapeHtml(currentFileLabel)}</code></p>

      <div class="rverdict">
        <div class="rverdict-score">
          <span class="num ${overallLevel}">${r.overallQualityScore.toFixed(1)}</span><span class="max">/100</span>
          <span class="tag ${overallLevel}">${verdictLabel(r.overallQualityScore)}</span>
        </div>
        <div class="rverdict-body">
          <p>${reportExecutiveTakeaway(r)}</p>
          ${reportMeterHtml(r.overallQualityScore, overallLevel)}
        </div>
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">01</span><h2>Summary</h2></div>
        <p>${reportExecutiveBody(r, currentFileLabel)}</p>
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">02</span><h2>Key metrics</h2></div>
        <div class="rmetric-strip">
          <div class="rmetric"><span class="v">${r.rowCount.toLocaleString()}</span><span class="l">Rows</span></div>
          <div class="rmetric"><span class="v">${r.columnCount}</span><span class="l">Columns</span></div>
          <div class="rmetric"><span class="v ${r.duplicateRows ? "bad" : "good"}">${r.duplicateRows}</span><span class="l">Duplicate rows</span></div>
          <div class="rmetric"><span class="v ${r.needsAttention === 0 ? "good" : "warn"}">${r.needsAttention}</span><span class="l">Need attention</span></div>
          <div class="rmetric"><span class="v ${r.totalMissingValues === 0 ? "good" : "bad"}">${r.totalMissingValues.toLocaleString()}</span><span class="l">Missing values</span></div>
        </div>
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">03</span><h2>Quality by dimension</h2></div>
        <div class="rtwocol">
          <div>${reportDimBarsHtml(r)}</div>
          <div>
            ${r.worstColumn ? `<p class="rhighlight">▾ Weakest column: <b>${escapeHtml(r.worstColumn.column)}</b> — ${r.worstColumn.qualityScore.toFixed(0)}/100</p>` : ""}
            ${r.bestColumn ? `<p class="rhighlight">▴ Strongest column: <b>${escapeHtml(r.bestColumn.column)}</b> — ${r.bestColumn.qualityScore.toFixed(0)}/100</p>` : ""}
          </div>
        </div>
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">04</span><h2>Root cause analysis</h2></div>
        <div class="rroot">
          <div class="rroot-chart">${reportDonutHtml(r)}</div>
          <p class="rroot-takeaway">${r.topIssue ? `<strong>${escapeHtml(r.topIssue.dimension)}</strong> is the leading driver of quality loss, affecting ${r.topIssue.count} column${r.topIssue.count === 1 ? "" : "s"}.` : "No single dimension dominates — issues are spread thinly, if present at all."}</p>
        </div>
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">05</span><h2>Column quality scores</h2><span class="rsection-note">lowest first</span></div>
        ${reportColumnBarsHtml(r)}
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">06</span><h2>Findings by column</h2><span class="rsection-note">lowest score first</span></div>
        ${reportFindingsTableHtml(r)}
      </div>

      <div class="rsection">
        <div class="rsection-head"><span class="rn">07</span><h2>Notes &amp; sample problem values</h2></div>
        ${reportNotesHtml(r)}
      </div>

      <div class="rfooter">
        <span>Data Quality Profiler — client-side engine, no data leaves the browser.</span>
        <span>${escapeHtml(currentFileLabel)} · ${escapeHtml(generatedAt)}</span>
      </div>
    </div>`;

  const title = `Data Quality Report — ${currentFileLabel}`;

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@600;700&family=Source+Sans+3:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>${css}</style>
</head>
<body>
${bodyHtml}
<script>
  window.addEventListener("load", function () {
    setTimeout(function () { try { window.print(); } catch (e) {} }, 350);
  });
<\/script>
</body>
</html>`;
}

async function exportPdfReport() {
  if (!currentReport) return;
  const html = buildStandaloneReportHtml();
  const filename = `quality_report_${safeFilenameSlug(currentBaseFileName || currentFileLabel, currentSheetName)}.html`;
  document.getElementById("download-fallback").hidden = true;

  const downloads = window.__claudeDownloads;
  if (downloads) {
    try {
      await downloads.save({ filename, data: html });
      return;
    } catch (err) {
      if (err && err.code === "extension_not_enabled") {
        try {
          await downloads.save({ filename: filename.replace(/\.html$/, ".txt"), data: html });
          return;
        } catch (err2) { /* fall through to manual copy */ }
      }
      // declined / rate_limited / unavailable / anything else -> fall back below
    }
  }
  showCsvFallback(html);
}

els.exportPdfBtn.addEventListener("click", exportPdfReport);

// downloads capability init: use Claude's native downloads capability when
// embedded in claude.ai; otherwise fall back to a plain browser download,
// which works fine here because this page (unlike the claude.ai artifact
// viewer) isn't running inside a sandboxed iframe.
(async () => {
  try {
    if (window.claude && typeof window.claude.use === "function") {
      window.__claudeDownloads = await window.claude.use("downloads");
      return;
    }
  } catch (e) { /* capability unavailable, fall through to polyfill below */ }
  window.__claudeDownloads = {
    async save({ filename, data }) {
      const isHtml = /\.html?$/i.test(filename);
      const blob = new Blob([data], { type: isHtml ? "text/html" : "text/csv" });
      const url = URL.createObjectURL(blob);
      if (isHtml) {
        // The generated report is a standalone document with its own
        // onload -> window.print(), so opening it directly triggers the
        // browser's native "Save as PDF" dialog.
        window.open(url, "_blank");
      } else {
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
      }
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    },
  };
})();
