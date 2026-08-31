/* ---------------------------------------------------------------------
   Data Quality Profiler — client-side engine
   A JS port of the Python type-inference + profiling engine. No column
   name or schema is hardcoded anywhere here — every column's semantic
   type is inferred purely from its values.
------------------------------------------------------------------------ */

/**
 * A minimal, dependency-free CSV parser that preserves cell text EXACTLY
 * as written (RFC 4180 quoting only). This matters: libraries like SheetJS
 * try to be "helpful" by auto-detecting dates/numbers in CSV cells and
 * silently reformatting them (e.g. "2024-08-01" -> "8/1/24") before we
 * ever get to profile the column — which corrupts the very inconsistency
 * we're trying to detect. Raw text in, raw text out.
 */
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  const s = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inQuotes) {
      if (c === '"') {
        if (s[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field); field = "";
    } else if (c === "\n") {
      row.push(field); field = "";
      rows.push(row); row = [];
    } else {
      field += c;
    }
  }
  if (field !== "" || row.length) { row.push(field); rows.push(row); }

  const nonEmptyRows = rows.filter((r) => !(r.length === 1 && r[0] === ""));
  if (nonEmptyRows.length === 0) return { headers: [], rows: [] };

  const headers = nonEmptyRows[0].map((h) => h.trim());
  const dataRows = nonEmptyRows.slice(1).map((r) => {
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = r[idx] ?? ""; });
    return obj;
  });
  return { headers, rows: dataRows };
}

const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$/;
const URL_RE = /^(https?:\/\/|www\.)[^\s]+\.[a-zA-Z]{2,}[^\s]*$/i;
const PHONE_RE = /^\+?\d[\d\s\-()]{6,15}\d$/;
const SAUDI_ID_RE = /^[12]\d{9}$/;
const INTEGER_RE = /^-?\d+$/;
const FLOAT_RE = /^-?\d+\.\d+$/;
const CURRENCY_RE = /^-?\d[\d,]*(\.\d+)?\s?(SAR|USD|\$|ر\.س)?$/i;
const PERCENT_RE = /^-?\d+(\.\d+)?\s?%$/;
const BOOLEAN_VALUES = new Set(["true", "false", "yes", "no", "y", "n", "1", "0", "نعم", "لا"]);

// [regex-testable format] -> parser. Order mirrors the Python DATE_FORMATS list.
const DATE_PARSERS = [
  { re: /^(\d{4})-(\d{2})-(\d{2})$/, make: (m) => new Date(+m[1], +m[2] - 1, +m[3]) },        // YYYY-MM-DD
  { re: /^(\d{2})\/(\d{2})\/(\d{4})$/, make: (m) => new Date(+m[3], +m[2] - 1, +m[1]) },       // DD/MM/YYYY
  { re: /^(\d{4})\/(\d{2})\/(\d{2})$/, make: (m) => new Date(+m[1], +m[2] - 1, +m[3]) },       // YYYY/MM/DD
  { re: /^(\d{2})-(\d{2})-(\d{4})$/, make: (m) => new Date(+m[3], +m[2] - 1, +m[1]) },         // DD-MM-YYYY
  { re: /^(\d{2})-(\d{2})-(\d{2})$/, make: (m) => new Date(2000 + +m[3], +m[1] - 1, +m[2]) },  // MM-DD-YY
  { re: /^(\d{2})\.(\d{2})\.(\d{4})$/, make: (m) => new Date(+m[3], +m[2] - 1, +m[1]) },       // DD.MM.YYYY
];

function isValidDate(v) {
  for (const { re, make } of DATE_PARSERS) {
    const m = re.exec(v);
    if (!m) continue;
    const d = make(m);
    if (!isNaN(d.getTime())) return true;
  }
  return false;
}

function matchRate(values, predicate) {
  const nonEmpty = values.filter((v) => v.trim() !== "");
  if (nonEmpty.length === 0) return 0;
  let hits = 0;
  for (const v of nonEmpty) if (predicate(v.trim())) hits++;
  return hits / nonEmpty.length;
}

function categoricalScore(values) {
  const nonEmpty = values.filter((v) => v.trim() !== "");
  if (nonEmpty.length === 0) return 0;
  const uniqueSet = new Set(nonEmpty.map((v) => v.trim()));
  const nUnique = uniqueSet.size;
  const ratioUnique = nUnique / nonEmpty.length;
  const avgLen = nonEmpty.reduce((s, v) => s + v.trim().length, 0) / nonEmpty.length;
  if (ratioUnique <= 0.5 && nUnique <= 50 && avgLen <= 40) return 1 - ratioUnique;
  return 0;
}

// Tie-break order: date before phone (an ISO date is digits+separators and
// would otherwise satisfy a loose phone regex too).
const TYPE_ORDER = ["email", "url", "saudi_id", "date", "phone", "boolean", "percent",
  "integer", "float", "currency", "categorical", "free_text"];

const DETECTORS = {
  email: (vals) => matchRate(vals, (v) => EMAIL_RE.test(v)),
  url: (vals) => matchRate(vals, (v) => URL_RE.test(v)),
  saudi_id: (vals) => matchRate(vals, (v) => SAUDI_ID_RE.test(v)),
  date: (vals) => matchRate(vals, isValidDate),
  phone: (vals) => matchRate(vals, (v) => PHONE_RE.test(v) && !isValidDate(v)),
  boolean: (vals) => matchRate(vals, (v) => BOOLEAN_VALUES.has(v.toLowerCase())),
  percent: (vals) => matchRate(vals, (v) => PERCENT_RE.test(v)),
  integer: (vals) => matchRate(vals, (v) => INTEGER_RE.test(v)),
  float: (vals) => matchRate(vals, (v) => FLOAT_RE.test(v)),
  currency: (vals) => matchRate(vals, (v) => CURRENCY_RE.test(v) && /\d/.test(v)),
  categorical: categoricalScore,
  free_text: () => 0.15,
};

const MIN_CONFIDENCE = 0.6;
const NUMERIC_TYPES = new Set(["integer", "float", "currency", "percent"]);

function inferColumnType(values) {
  const scores = {};
  for (const t of TYPE_ORDER) {
    try {
      scores[t] = Math.round(DETECTORS[t](values) * 1000) / 1000;
    } catch (e) {
      scores[t] = 0;
    }
  }
  let bestType = TYPE_ORDER[0];
  let bestScore = -1;
  for (const t of TYPE_ORDER) {
    if (scores[t] > bestScore) { bestScore = scores[t]; bestType = t; }
  }
  if (bestScore < MIN_CONFIDENCE && bestType !== "categorical") {
    bestType = "free_text";
    bestScore = scores.free_text;
  }
  return { inferredType: bestType, confidence: bestScore, scores };
}

function cleanNumeric(values) {
  const out = [];
  for (const v of values) {
    const stripped = v.replace(/[^\d.\-]/g, "");
    const n = parseFloat(stripped);
    if (!isNaN(n) && stripped !== "" && stripped !== "-" && stripped !== ".") out.push(n);
  }
  return out;
}

function outliersIQR(nums) {
  if (nums.length < 4) return 0;
  const sorted = [...nums].sort((a, b) => a - b);
  const q = (p) => {
    const idx = p * (sorted.length - 1);
    const lo = Math.floor(idx), hi = Math.ceil(idx);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
  };
  const q1 = q(0.25), q3 = q(0.75);
  const iqr = q3 - q1;
  if (iqr === 0) return 0;
  const lower = q1 - 1.5 * iqr, upper = q3 + 1.5 * iqr;
  return nums.filter((n) => n < lower || n > upper).length;
}

function profileColumn(name, rawValues) {
  const total = rawValues.length;
  const missingMask = rawValues.map((v) => (v ?? "").toString().trim() === "");
  const missing = missingMask.filter(Boolean).length;
  const completenessPct = total ? round1((100 * (total - missing)) / total) : 0;

  const strValues = rawValues.map((v) => (v ?? "").toString());
  const guess = inferColumnType(strValues);
  const nonEmpty = strValues.filter((v) => v.trim() !== "").map((v) => v.trim());

  let validityPct = round1(100 * (guess.scores[guess.inferredType] || 0));
  if (guess.inferredType === "free_text" || guess.inferredType === "categorical") validityPct = 100;

  let invalidSamples = [];
  let outlierCount = 0;
  let notes = "";

  if (guess.inferredType === "email") {
    invalidSamples = nonEmpty.filter((v) => !EMAIL_RE.test(v)).slice(0, 5);
  } else if (guess.inferredType === "date") {
    invalidSamples = nonEmpty.filter((v) => !isValidDate(v)).slice(0, 5);
  } else if (NUMERIC_TYPES.has(guess.inferredType)) {
    const nums = cleanNumeric(nonEmpty);
    outlierCount = outliersIQR(nums);
    let negOrZero = 0;
    if (["currency", "integer", "float"].includes(guess.inferredType)) {
      negOrZero = nums.filter((n) => n <= 0).length;
    }
    const parts = [];
    if (negOrZero) parts.push(`${negOrZero} non-positive value(s)`);
    if (outlierCount) parts.push(`${outlierCount} statistical outlier(s) (IQR method)`);
    if (parts.length) notes = parts.join("; ") + " — worth a manual check.";
    const numericSet = new Set(nums.map(String));
    invalidSamples = nonEmpty.filter((v) => {
      const stripped = v.replace(/[^\d.\-]/g, "");
      const n = parseFloat(stripped);
      return isNaN(n) || stripped === "" || stripped === "-";
    }).slice(0, 5);
  } else if (guess.inferredType === "categorical") {
    const counts = {};
    for (const v of nonEmpty) counts[v] = (counts[v] || 0) + 1;
    const values = Object.values(counts);
    const rare = values.filter((c) => c === 1).length;
    if (values.length > 0 && rare / values.length > 0.3) {
      notes = `${rare} value(s) appear only once — possible free-text contamination or typos.`;
    }
  }

  const uniqueCount = new Set(nonEmpty).size;
  const uniquenessPct = nonEmpty.length ? round1((100 * uniqueCount) / nonEmpty.length) : 0;

  const issueCount = missing + invalidSamples.length + outlierCount;
  const outlierPenalty = Math.min(15, outlierCount * 1.5);
  let qualityScore =
    0.5 * completenessPct + 0.4 * validityPct + 0.1 * Math.min(uniquenessPct, 100) -
    outlierPenalty - (notes ? 3 : 0);
  qualityScore = Math.min(100, Math.max(0, round1(qualityScore)));

  const primaryIssue = diagnosePrimaryIssue({
    completenessPct, validityPct, uniquenessPct, outlierCount, outlierPenalty, inferredType: guess.inferredType,
  });

  return {
    column: name,
    inferredType: guess.inferredType,
    confidence: Math.round(guess.confidence * 100) / 100,
    completenessPct,
    validityPct,
    uniquenessPct,
    issueCount,
    missingCount: missing,
    qualityScore,
    sampleInvalid: invalidSamples,
    outlierCount,
    notes,
    primaryIssue, // { dimension, pct, gap } | null — the single biggest driver of a low score
  };
}

// Uniqueness is only a fair quality dimension for columns that are SUPPOSED
// to be unique per row (identifiers, contact info) — a low-uniqueness
// categorical or date column is the expected shape, not a defect.
const UNIQUENESS_RELEVANT_TYPES = new Set(["email", "saudi_id", "phone", "url"]);
const WEAKNESS_THRESHOLD = 95; // a dimension below this % counts as "weak"

/**
 * Which single dimension is actually responsible for a column's score being
 * low: completeness (missing data), validity (values that don't match the
 * inferred type), uniqueness (for identifier-like columns only), or
 * statistical outliers?
 */
function diagnosePrimaryIssue({ completenessPct, validityPct, uniquenessPct, outlierCount, outlierPenalty, inferredType }) {
  const dims = { Completeness: completenessPct, Validity: validityPct };
  if (UNIQUENESS_RELEVANT_TYPES.has(inferredType)) dims.Uniqueness = uniquenessPct;

  const candidates = [];
  let weakestDim = null, weakestPct = 100;
  for (const [dim, pct] of Object.entries(dims)) {
    if (pct < weakestPct) { weakestPct = pct; weakestDim = dim; }
  }
  if (weakestDim !== null && weakestPct < WEAKNESS_THRESHOLD) {
    candidates.push({ dimension: weakestDim, pct: round1(weakestPct), gap: round1(100 - weakestPct) });
  }
  if (outlierCount > 0) {
    candidates.push({ dimension: "Outliers", pct: null, gap: round1(outlierPenalty) });
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.gap - a.gap);
  return candidates[0];
}

function profileDataset(headers, rows) {
  const rowStrings = rows.map((r) => headers.map((h) => (r[h] ?? "").toString().trim()).join(""));
  const seen = new Set();
  let duplicateRows = 0;
  for (const s of rowStrings) {
    if (seen.has(s)) duplicateRows++;
    else seen.add(s);
  }

  const columns = headers.map((h) => profileColumn(h, rows.map((r) => r[h])));
  const overallRaw = columns.length ? columns.reduce((s, c) => s + c.qualityScore, 0) / columns.length : 0;
  const dupPenalty = rows.length ? Math.min(15, round1((100 * duplicateRows) / rows.length)) : 0;
  const overallQualityScore = round1(Math.max(0, overallRaw - dupPenalty * 0.3));

  const needsAttention = columns.filter((c) => c.qualityScore < 85).length;

  // diagnosePrimaryIssue already applies a 95%-weakness threshold per
  // dimension, so tallying across ALL columns (not just low scorers) is
  // safe here — a column only shows up if it genuinely has a weak spot.
  const dimensionIssueCounts = {};
  for (const c of columns) {
    if (c.primaryIssue) dimensionIssueCounts[c.primaryIssue.dimension] = (dimensionIssueCounts[c.primaryIssue.dimension] || 0) + 1;
  }
  let topIssue = null;
  for (const [dim, count] of Object.entries(dimensionIssueCounts)) {
    if (!topIssue || count > topIssue.count) topIssue = { dimension: dim, count };
  }

  const n = columns.length;
  const avgCompleteness = n ? round1(columns.reduce((s, c) => s + c.completenessPct, 0) / n) : 0;
  const avgValidity = n ? round1(columns.reduce((s, c) => s + c.validityPct, 0) / n) : 0;
  const avgUniqueness = n ? round1(columns.reduce((s, c) => s + c.uniquenessPct, 0) / n) : 0;
  const totalMissingValues = columns.reduce((s, c) => s + c.missingCount, 0);
  const worstColumn = n ? columns.reduce((a, b) => (a.qualityScore <= b.qualityScore ? a : b)) : null;
  const bestColumn = n ? columns.reduce((a, b) => (a.qualityScore >= b.qualityScore ? a : b)) : null;

  return {
    rowCount: rows.length,
    columnCount: headers.length,
    duplicateRows,
    overallQualityScore,
    needsAttention,
    topIssue, // { dimension, count } | null
    dimensionIssueCounts,
    avgCompleteness,
    avgValidity,
    avgUniqueness,
    totalMissingValues,
    worstColumn: worstColumn ? { column: worstColumn.column, qualityScore: worstColumn.qualityScore } : null,
    bestColumn: bestColumn ? { column: bestColumn.column, qualityScore: bestColumn.qualityScore } : null,
    columns,
  };
}

function round1(n) {
  return Math.round(n * 10) / 10;
}
