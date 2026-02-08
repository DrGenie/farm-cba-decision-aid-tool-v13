
// Farming CBA Decision Aid - Newcastle Business School
// Trial cost-benefit decision aid.
// - Can load an embedded example dataset on start (if present).
// - Uses all rows and columns; no sampling or removal.
// - Detects key columns, including replicate_id, treatment_name, is_control,
//   yield_t_ha, total_cost_per_ha (variable cost proxy), and
//   cost_amendment_input_per_ha (capital cost proxy).
// - Treats the measured control as the baseline for all comparisons.
// - Presents a control-centric comparison table with plain-language labels.
// - Provides leaderboard, charts, exports (TSV/CSV/XLSX), and AI summary prompt.

(() => {
  "use strict";

  // =========================
  // 0) STATE
  // =========================

  const state = {
    rawText: "",
    headers: [],
    rows: [],
    cleanedHeaders: [],
    cleanedRows: [],
    pendingFile: null,
    lastPreview: null,
    audit: {
      fileName: "",
      uploadedAt: null,
      rowCount: 0,
      treatmentCount: 0,
      controlCandidates: [],
      columnMapping: {},
      version: "1.6.0"
    },
    templateHeaders: [],
    columnMap: null,
    treatments: [], // aggregated per treatment
    controlName: null,
    params: {
      pricePerTonne: 500,
      years: 10,
      persistenceYears: 10,
      discountRate: 5
    },
    results: {
      treatments: [],
      control: null
    },
    charts: {
      netProfitDelta: null,
      costsBenefits: null
    }
  };

  const VERSION = "1.6.0";

  const PROJECT = {
    name: "Trial Cost-Benefit Decision Aid",
    partnerPlaceholder: "Project partner logos (placeholder)"
  };

  // =========================
  // 1) UTILITIES
  // =========================

  function showToast(message, type = "info", timeoutMs = 3200) {
    const container = document.getElementById("toastContainer");
    if (!container) return;
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.innerHTML = `
      <div>${message}</div>
      <button class="toast-dismiss" aria-label="Dismiss">&times;</button>
    `;
    container.appendChild(toast);

    const remove = () => {
      if (!toast.parentElement) return;
      toast.style.opacity = "0";
      toast.style.transform = "translateY(4px)";
      setTimeout(() => {
        if (toast.parentElement) {
          toast.parentElement.removeChild(toast);
        }
      }, 150);
    };

    toast.querySelector(".toast-dismiss").addEventListener("click", remove);
    setTimeout(remove, timeoutMs);
  }
  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function parseNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    const trimmed = String(value).trim();
    if (!trimmed) return NaN;
    const cleaned = trimmed.replace(/,/g, "");
    const num = Number(cleaned);
    return Number.isFinite(num) ? num : NaN;
  }

  function parseBoolean(value) {
    if (value === null || value === undefined) return false;
    const s = String(value).trim().toLowerCase();
    return s === "true" || s === "1" || s === "yes" || s === "y";
  }

  function discountFactorSum(discountRatePct, years) {
    const r = discountRatePct / 100;
    if (years <= 0) return 0;
    if (r === 0) return years;
    let sum = 0;
    for (let t = 1; t <= years; t++) {
      sum += 1 / Math.pow(1 + r, t);
    }
    return sum;
  }

  function meanIgnoringNaN(values) {
    if (!values || values.length === 0) return NaN;
    let sum = 0;
    let n = 0;
    for (const v of values) {
      const x = parseNumber(v);
      if (!Number.isNaN(x)) {
        sum += x;
        n++;
      }
    }
    return n === 0 ? NaN : sum / n;
  }

  function formatCurrency(value) {
    if (value === null || value === undefined || Number.isNaN(value)) return "-";
    try {
      return new Intl.NumberFormat("en-AU", {
        style: "currency",
        currency: "AUD",
        maximumFractionDigits: Math.abs(value) < 1000 ? 2 : 0
      }).format(value);
    } catch (_) {
      return value.toFixed(0);
    }
  }

  function formatNumber(value, decimals = 2) {
    if (value === null || value === undefined || Number.isNaN(value)) return "-";
    return value.toFixed(decimals);
  }

  function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  // =========================
  // 2) DATA PARSING
  // =========================

  function detectDelimiter(headerLine) {
    if (!headerLine) return "\t";
    const tabCount = (headerLine.match(/\t/g) || []).length;
    const commaCount = (headerLine.match(/,/g) || []).length;
    return tabCount >= commaCount ? "\t" : ",";
  }

  function parseDelimitedText(text) {
    const trimmed = text.replace(/\r\n/g, "\n").trim();
    if (!trimmed) {
      throw new Error("No data found in the provided text.");
    }

    const lines = trimmed.split("\n").filter((l) => l.trim().length > 0);
    const delimiter = detectDelimiter(lines[0]);
    const headers = lines[0].split(delimiter).map((h) => h.trim());
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split(delimiter);
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = j < parts.length ? parts[j] : "";
      }
      rows.push(row);
    }
    return { headers, rows };
  }

  // Map required semantic columns to actual dataset headers
  function normaliseHeaderForMatch(value) {
    const s = String(value || "").trim().toLowerCase();
    if (!s) return "";
    return s
      .replace(/\uFEFF/g, "")
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .replace(/_+/g, "_");
  }

  function normaliseHeaderForLooseMatch(value) {
    return normaliseHeaderForMatch(value).replace(/_/g, "");
  }

  // Map required semantic columns to actual dataset headers (tolerant to common variants)
  function detectColumnMap(headers) {
    const lcHeaders = headers.map((h) => String(h || "").toLowerCase());
    const normHeaders = headers.map((h) => normaliseHeaderForMatch(h));
    const looseHeaders = headers.map((h) => normaliseHeaderForLooseMatch(h));

    const findCol = (aliases) => {
      for (const alias of aliases) {
        const a = String(alias || "");
        const aLc = a.toLowerCase();
        const aNorm = normaliseHeaderForMatch(a);
        const aLoose = normaliseHeaderForLooseMatch(a);

        // Exact match
        let idx = lcHeaders.indexOf(aLc);
        if (idx !== -1) return headers[idx];

        // Normalised match
        idx = normHeaders.indexOf(aNorm);
        if (idx !== -1) return headers[idx];

        // Loose match (ignores underscores and punctuation)
        idx = looseHeaders.indexOf(aLoose);
        if (idx !== -1) return headers[idx];

        // Partial match as a last resort (safe only for unique hits)
        const partialHits = [];
        for (let i = 0; i < looseHeaders.length; i++) {
          if (!looseHeaders[i]) continue;
          if (looseHeaders[i] === aLoose) {
            partialHits.push(i);
          } else if (looseHeaders[i].includes(aLoose) && aLoose.length >= 6) {
            partialHits.push(i);
          }
        }
        if (partialHits.length === 1) return headers[partialHits[0]];
      }
      return null;
    };

    const map = {
      plotId: findCol(["plot_id", "plotid", "plot", "plot_no", "plotnumber"]),
      replicate: findCol(["replicate_id", "replicate", "rep", "rep_no", "repno", "replicateid"]),
      treatmentName: findCol(["treatment_name", "treatment", "amendment_name", "amendment", "treatmentname"]),
      isControl: findCol(["is_control", "control_flag", "control", "iscontrol", "controlindicator"]),
      yield: findCol(["yield_t_ha", "yield", "yieldtha", "yield_t/ha", "yield_tonnes_ha", "yield_tonnes_per_ha"]),
      variableCost: findCol([
        "total_cost_per_ha",
        "variable_cost_per_ha",
        "totalcostperha",
        "variablecostperha",
        "total_cost_ha",
        "total_cost_perha"
      ]),
      capitalCost: findCol([
        "cost_amendment_input_per_ha",
        "capital_cost_per_ha",
        "capitalcostperha",
        "amendment_cost_per_ha"
      ])
    };

    return map;
  }

  function setUploadStatus(message = "", type = "") {
    const el = document.getElementById("uploadStatus");
    if (!el) return;
    el.textContent = message || "";
    el.className = "upload-status" + (type ? ` ${type}` : "");
  }

  
  function buildNearMatchSuggestions(headers, missingCanonical) {
    const suggestions = {};
    const loose = headers.map((h) => normaliseHeaderForLooseMatch(h));
    for (const canon of missingCanonical) {
      const cLoose = normaliseHeaderForLooseMatch(canon);
      const hits = [];
      for (let i = 0; i < headers.length; i++) {
        if (!loose[i]) continue;
        if (loose[i] === cLoose) hits.push(headers[i]);
        else if (loose[i].includes(cLoose) || cLoose.includes(loose[i])) hits.push(headers[i]);
      }
      suggestions[canon] = hits.slice(0, 5);
    }
    return suggestions;
  }

  function validateParsedData(headers, rows) {
    const map = detectColumnMap(headers);

    const required = {
      treatment_name: map.treatmentName,
      is_control: map.isControl,
      yield_t_ha: map.yield,
      total_cost_per_ha: map.variableCost
    };

    const missing = Object.entries(required)
      .filter(([, v]) => !v)
      .map(([k]) => k);

    const errors = [];
    const warnings = [];

    if (missing.length) {
      const near = buildNearMatchSuggestions(headers, missing);
      const parts = missing.map((c) => {
        const hits = near[c] && near[c].length ? ` Near matches found: ${near[c].join(", ")}.` : "";
        return `${c}.${hits}`;
      });
      errors.push(
        `Missing required column${missing.length === 1 ? "" : "s"}: ${missing.join(", ")}. ` +
          `Download the template and copy these column names exactly, or rename your headers to match. ` +
          parts.join(" ")
      );
    }

    let controlRows = 0;
    let treatmentRows = 0;
    let controlNames = new Set();
    let controlHasCore = false;
    let treatmentHasCore = false;

    if (map.treatmentName && map.isControl) {
      for (const row of rows) {
        const nameRaw = row[map.treatmentName];
        const name = nameRaw ? String(nameRaw).trim() : "";
        if (!name) continue;

        const isCtrl = parseBoolean(row[map.isControl]);
        const y = map.yield ? parseNumber(row[map.yield]) : NaN;
        const c = map.variableCost ? parseNumber(row[map.variableCost]) : NaN;
        const hasCore = !Number.isNaN(y) && !Number.isNaN(c);

        if (isCtrl) {
          controlRows += 1;
          controlNames.add(name);
          if (hasCore) controlHasCore = true;
        } else {
          treatmentRows += 1;
          if (hasCore) treatmentHasCore = true;
        }
      }
    }

    if (!missing.length) {
      if (controlRows < 1) {
        errors.push(
          `No control identified. Look for the is_control column and set at least one row to TRUE (or 1 or yes). Example: for your control treatment rows, set is_control to TRUE.`
        );
      }
      if (treatmentRows < 1) {
        errors.push(
          `No treatment rows found. Look for the is_control column and ensure at least one row is marked FALSE (or 0 or no). Example: for non-control treatments, set is_control to FALSE.`
        );
      }
      if (!controlHasCore || !treatmentHasCore) {
        errors.push(
          `Core fields are missing or not numeric. Look for yield_t_ha and total_cost_per_ha and ensure there are numeric values for some control rows and some treatment rows. Example: yield_t_ha = 2.6 and total_cost_per_ha = 150.`
        );
      }
    }

    // Non-blocking warnings for optional fields and common unit issues.
    if (!map.capitalCost) {
      warnings.push(
        `Optional column not detected: cost_amendment_input_per_ha. If you have one-off establishment costs, include them here. Otherwise the tool assumes zero.`
      );
    }

    // Unit and outlier checks (non-blocking)
    if (!missing.length && map.yield) {
      let maxYield = -Infinity;
      let minYield = Infinity;
      for (const row of rows) {
        const y = parseNumber(row[map.yield]);
        if (!Number.isNaN(y)) {
          if (y > maxYield) maxYield = y;
          if (y < minYield) minYield = y;
        }
      }
      if (minYield < 0) {
        warnings.push(
          `Some yields are negative. Check yield_t_ha values and units.`
        );
      }
      if (maxYield > 60) {
        warnings.push(
          `Some yields are very large (maximum ${formatNumber(maxYield, 2)}). Check that yield is in tonnes per hectare, not kilograms per hectare. Example: 3000 kg/ha should be entered as 3.0 t/ha.`
        );
      }
    }

    if (!missing.length && map.variableCost) {
      let maxCost = -Infinity;
      let minCost = Infinity;
      for (const row of rows) {
        const c = parseNumber(row[map.variableCost]);
        if (!Number.isNaN(c)) {
          if (c > maxCost) maxCost = c;
          if (c < minCost) minCost = c;
        }
      }
      if (minCost < 0) {
        warnings.push(
          `Some costs are negative. Check total_cost_per_ha values.`
        );
      }
      if (maxCost > 200000) {
        warnings.push(
          `Some costs are very large (maximum ${formatCurrency(maxCost)}). Check whether costs are in cents rather than dollars. Example: 25000 cents should be entered as 250 dollars.`
        );
      }
    }

    const stats = {
      rowCount: rows.length,
      columnCount: headers.length,
      treatmentCount: 0,
      controlRowCount: controlRows,
      treatmentRowCount: treatmentRows
    };

    // Treatment count based on detected treatment_name
    if (map.treatmentName) {
      const names = new Set();
      for (const row of rows) {
        const v = row[map.treatmentName];
        const n = v ? String(v).trim() : "";
        if (n) names.add(n);
      }
      stats.treatmentCount = names.size;
    }

    const controlCandidates = Array.from(controlNames.values());

    // Mapping report for transparency
    const mapping = {
      treatment_name: map.treatmentName || "",
      is_control: map.isControl || "",
      yield_t_ha: map.yield || "",
      total_cost_per_ha: map.variableCost || "",
      cost_amendment_input_per_ha: map.capitalCost || ""
    };

    const canRun = errors.length === 0;
    return { canRun, errors, warnings, map, stats, controlCandidates, mapping };
  }
function normaliseTemplateHeaders(headers) {
    // Ensure we preserve the exact column names users will upload (case-sensitive).
    return headers.filter((h) => String(h || "").trim().length > 0);
  }

  // =========================
  // 3) AGGREGATION & CBA
  // =========================

  function aggregateTreatments() {
    const { headers, rows } = state;
    if (!headers.length || !rows.length) {
      state.treatments = [];
      state.controlName = null;
      state.columnMap = null;
      return;
    }

    const columnMap = detectColumnMap(headers);
    state.columnMap = columnMap;

    const summaryByName = new Map();
    const flaggedControls = new Set();

    let yieldMissingCount = 0;
    let costMissingCount = 0;

    for (const row of rows) {
      const tNameRaw = columnMap.treatmentName ? row[columnMap.treatmentName] : "";
      const treatmentName = tNameRaw ? String(tNameRaw).trim() : "";
      if (!treatmentName) continue;

      if (!summaryByName.has(treatmentName)) {
        summaryByName.set(treatmentName, {
          name: treatmentName,
          replicates: [],
          yields: [],
          variableCosts: [],
          capitalCosts: [],
          isControlFlagged: false
        });
      }
      const t = summaryByName.get(treatmentName);

      const repVal = columnMap.replicate ? row[columnMap.replicate] : null;
      if (repVal !== null && repVal !== undefined && String(repVal).trim()) {
        t.replicates.push(repVal);
      }

      if (columnMap.yield) {
        const y = parseNumber(row[columnMap.yield]);
        if (!Number.isNaN(y)) t.yields.push(y);
        else yieldMissingCount += 1;
      }
      if (columnMap.variableCost) {
        const vc = parseNumber(row[columnMap.variableCost]);
        if (!Number.isNaN(vc)) t.variableCosts.push(vc);
        else costMissingCount += 1;
      }
      if (columnMap.capitalCost) {
        const cc = parseNumber(row[columnMap.capitalCost]);
        if (!Number.isNaN(cc)) t.capitalCosts.push(cc);
      }
      if (columnMap.isControl) {
        const isCtrl = parseBoolean(row[columnMap.isControl]);
        if (isCtrl) {
          t.isControlFlagged = true;
          flaggedControls.add(treatmentName);
        }
      }
    }

    // Decide control treatment.
    let controlName = null;

    if (flaggedControls.size === 1) {
      controlName = Array.from(flaggedControls)[0];
    } else if (flaggedControls.size > 1) {
      // Keep current selection if still valid, otherwise choose the first flagged option.
      if (state.controlName && flaggedControls.has(state.controlName)) controlName = state.controlName;
      else controlName = Array.from(flaggedControls)[0];
    } else {
      // No explicit control flag, fall back to name heuristic
      for (const name of summaryByName.keys()) {
        if (name.toLowerCase().includes("control")) {
          controlName = name;
          break;
        }
      }
      if (!controlName && summaryByName.size > 0) controlName = summaryByName.keys().next().value;
    }

    state.controlName = controlName || null;

    const treatments = [];
    for (const [name, t] of summaryByName.entries()) {
      const avgYield = meanIgnoringNaN(t.yields);
      const avgVarCost = meanIgnoringNaN(t.variableCosts);

      const capMean = meanIgnoringNaN(t.capitalCosts);
      const avgCapCost = Number.isNaN(capMean) ? 0 : capMean;

      treatments.push({
        name,
        isControl: name === state.controlName,
        avgYield,
        avgVarCost,
        avgCapCost,
        replicateCount: Math.max(t.yields.length, t.variableCosts.length, 0),
        isControlFlagged: t.isControlFlagged
      });
    }

    state.treatments = treatments;

    // Assumptions and defaults applied
    state.assumptions = {
      capitalCostAssumedZero: !columnMap.capitalCost,
      missingYieldCells: yieldMissingCount,
      missingCostCells: costMissingCount
    };
  }

  function computeCBA() {
    const { treatments, controlName, params } = state;
    if (!treatments.length || !controlName) return;

    const control = treatments.find((t) => t.name === controlName) || treatments[0];

    const price = parseNumber(params.pricePerTonne) || 0;
    const years = Math.max(1, parseInt(params.years, 10) || 1);
    const persistenceYears = Math.max(
      1,
      Math.min(years, parseInt(params.persistenceYears, 10) || years)
    );
    const discountRate = parseNumber(params.discountRate);
    const factorBenefits = discountFactorSum(discountRate, persistenceYears);
    const factorCosts = discountFactorSum(discountRate, years);

    const results = [];
    for (const t of treatments) {
      const avgYield = parseNumber(t.avgYield);
      const avgVarCost = parseNumber(t.avgVarCost);
      const avgCapCost = parseNumber(t.avgCapCost);

      const pvBenefits = Number.isNaN(avgYield)
        ? NaN
        : avgYield * price * factorBenefits;
      const pvVarCosts = Number.isNaN(avgVarCost) ? NaN : avgVarCost * factorCosts;
      const pvCapCosts = Number.isNaN(avgCapCost) ? NaN : avgCapCost;
      let pvTotalCosts = NaN;
      if (!Number.isNaN(pvVarCosts)) {
        pvTotalCosts = pvVarCosts + (Number.isNaN(pvCapCosts) ? 0 : pvCapCosts);
      } else if (!Number.isNaN(pvCapCosts) && pvCapCosts > 0) {
        pvTotalCosts = pvCapCosts;
      }
      const npv =
        Number.isNaN(pvBenefits) || Number.isNaN(pvTotalCosts)
          ? NaN
          : pvBenefits - pvTotalCosts;
      const bcr =
        !Number.isNaN(pvBenefits) && pvTotalCosts > 0
          ? pvBenefits / pvTotalCosts
          : NaN;
      const roi =
        !Number.isNaN(npv) && pvTotalCosts > 0
          ? (npv / pvTotalCosts) * 100
          : NaN;

      results.push({
        name: t.name,
        isControl: t.name === controlName,
        avgYield,
        avgVarCost,
        avgCapCost,
        pvBenefits,
        pvTotalCosts,
        npv,
        bcr,
        roi,
        meanOutcome: avgYield,
        meanCost: (Number.isNaN(avgVarCost) ? NaN : avgVarCost) + (Number.isNaN(avgCapCost) ? 0 : avgCapCost),
        netAnnualValue:
          Number.isNaN(avgYield) || Number.isNaN(avgVarCost)
            ? NaN
            : avgYield * price - avgVarCost - (Number.isNaN(avgCapCost) ? 0 : avgCapCost)
      });
    }

    // Control metrics
    const controlRes = results.find((r) => r.isControl) || results[0];
    const cPvBenefits = controlRes.pvBenefits;
    const cPvCosts = controlRes.pvTotalCosts;
    const cNpv = controlRes.npv;
    const cNetAnnual = controlRes.netAnnualValue;

    // Differences vs control (full and indicative)
    for (const r of results) {
      r.deltaPvBenefits = (Number.isNaN(r.pvBenefits) || Number.isNaN(cPvBenefits)) ? NaN : r.pvBenefits - cPvBenefits;
      r.deltaPvCosts = (Number.isNaN(r.pvTotalCosts) || Number.isNaN(cPvCosts)) ? NaN : r.pvTotalCosts - cPvCosts;
      r.deltaNpv = (Number.isNaN(r.npv) || Number.isNaN(cNpv)) ? NaN : r.npv - cNpv;
      r.deltaNetAnnual = (Number.isNaN(r.netAnnualValue) || Number.isNaN(cNetAnnual)) ? NaN : r.netAnnualValue - cNetAnnual;
    }

    const basicOnly =
      results.every((r) => Number.isNaN(r.npv)) &&
      results.some((r) => !Number.isNaN(r.netAnnualValue));

    if (basicOnly) {
      // Rank on net annual value as an indicative comparison.
      results.sort((a, b) => {
        const aVal = Number.isNaN(a.netAnnualValue) ? -Infinity : a.netAnnualValue;
        const bVal = Number.isNaN(b.netAnnualValue) ? -Infinity : b.netAnnualValue;
        return bVal - aVal;
      });
      results.forEach((r, idx) => {
        r.rank = idx + 1;
        // Reuse deltaNpv slot for indicative display and exports when full NPV is not available.
        r.deltaNpv = r.deltaNetAnnual;
      });
    } else {
      results.sort((a, b) => {
        const aVal = Number.isNaN(a.npv) ? -Infinity : a.npv;
        const bVal = Number.isNaN(b.npv) ? -Infinity : b.npv;
        return bVal - aVal;
      });
      results.forEach((r, idx) => {
        r.rank = idx + 1;
      });
    }

state.results = {
      treatments: results,
      control: controlRes,
      price,
      years,
      persistenceYears,
      discountRate,
      basicOnly
    };
  }

  // =========================
  // 4) RENDERING
  // =========================

  function renderOverview() {
    const container = document.getElementById("overviewSummary");
    if (!container) return;
    container.innerHTML = "";

    if (!state.rows.length || !state.treatments.length) {
      container.innerHTML = `<p class="small muted">No dataset loaded yet.</p>`;
      return;
    }

    const nPlots = state.rows.length;
    const nTreatments = state.treatments.length;
    const nReplicates = new Set(
      state.rows
        .map((r) =>
          state.columnMap && state.columnMap.replicate
            ? r[state.columnMap.replicate]
            : null
        )
        .filter((x) => x !== null && x !== undefined && String(x).trim() !== "")
    ).size;

    const controlName = state.controlName || "Not identified";

    const cards = [
      {
        label: "Plots in dataset",
        value: nPlots
      },
      {
        label: "Distinct treatments",
        value: nTreatments
      },
      {
        label: "Replicates",
        value: nReplicates || "-"
      },
      {
        label: "Control treatment",
        value: controlName
      }
    ];

    for (const c of cards) {
      const div = document.createElement("div");
      div.className = "summary-card";
      div.innerHTML = `
        <div class="summary-card-label">${c.label}</div>
        <div class="summary-card-value">${c.value}</div>
      `;
      container.appendChild(div);
    }
  }

  function renderDataSummaryAndChecks() {
    const summaryEl = document.getElementById("dataSummary");
    const checksEl = document.getElementById("dataChecks");
    if (!summaryEl || !checksEl) return;

    if (!state.rows.length || !state.headers.length) {
      summaryEl.textContent = "No dataset loaded.";
      checksEl.innerHTML = "";
      return;
    }

    const nRows = state.rows.length;
    const nCols = state.headers.length;
    summaryEl.textContent = `${nRows} plot rows, ${nCols} columns. All rows and columns are used in calculations.`;

    const cm = state.columnMap;
    const checks = [];

    const addCheck = (ok, label, detail, severity) => {
      checks.push({ ok, label, detail, severity });
    };

    const addColCheck = (purpose, colName) => {
      if (colName) {
        addCheck(
          true,
          `${purpose} column found`,
          `"${colName}" is used for ${purpose.toLowerCase()}.`,
          "ok"
        );
      } else {
        addCheck(
          false,
          `${purpose} column missing`,
          `No column was found for ${purpose.toLowerCase()}. The tool will still run but some metrics may be incomplete.`,
          "warn"
        );
      }
    };

    addColCheck("Treatment name", cm.treatmentName);
    addColCheck("Control flag", cm.isControl);
    addColCheck("Replicate", cm.replicate);
    addColCheck("Yield", cm.yield);
    addColCheck("Variable cost", cm.variableCost);
    addColCheck("Capital cost", cm.capitalCost);

    const missingYieldRows = state.rows.filter((r) => {
      if (!cm.yield) return false;
      const v = r[cm.yield];
      return v === null || v === undefined || String(v).trim() === "";
    }).length;

    if (missingYieldRows > 0) {
      addCheck(
        false,
        "Missing yield values",
        `${missingYieldRows} rows have missing yield; these rows still count for costs but are excluded from yield averages.`,
        "warn"
      );
    } else {
      addCheck(
        true,
        "Yield values present",
        "All rows have yield values in the detected yield column.",
        "ok"
      );
    }

    checksEl.innerHTML = "";
    for (const c of checks) {
      const li = document.createElement("li");
      li.className = "check-item";
      const pillClass =
        c.severity === "err" ? "err" : c.severity === "warn" ? "warn" : "ok";
      li.innerHTML = `
        <span class="check-pill ${pillClass}">${c.ok ? "OK" : "Check"}</span>
        <span>${c.label}. ${c.detail}</span>
      `;
      checksEl.appendChild(li);
    }
  }

  function getColumnHelpText(columnName) {
    const key = String(columnName || "").trim();
    const map = {
      plot_id: "Unique identifier for each plot or observation (optional).",
      replicate_id: "Replicate number or label for the same treatment (optional).",
      treatment_name: "Name of the treatment for this row. Use the same name across replicates.",
      is_control: "Set TRUE/1/yes for control rows, FALSE/0/no for treatment rows.",
      yield_t_ha: "Outcome for the row (for example yield) expressed per hectare.",
      total_cost_per_ha: "Total variable cost per hectare for this row.",
      total_cost_per_ha_raw: "Raw total variable cost per hectare (optional).",
      cost_amendment_input_per_ha: "One-off capital / upfront cost per hectare for this row (optional).",
      amendment_name: "Alternative label for treatment name (optional).",
      practice_change_label: "Short label describing the treatment (optional)."
    };
    return map[key] || "Optional column. Leave blank if not applicable.";
  }

  function renderTemplateColumns() {
    const container = document.getElementById("templateColumns");
    if (!container) return;
    const headers = (state.templateHeaders && state.templateHeaders.length)
      ? state.templateHeaders
      : (state.headers && state.headers.length ? state.headers : []);

    if (!headers.length) {
      container.innerHTML = '<div class="small muted">Template columns will appear once a dataset is loaded.</div>';
      return;
    }

    const required = new Set(["treatment_name", "is_control", "yield_t_ha", "total_cost_per_ha"]);

    const rows = headers.map((h) => {
      const help = getColumnHelpText(h);
      const req = required.has(h) ? "Required" : "Optional";
      return `<tr><td class="col-name"><span class="tip" data-tooltip="${escapeHtml(help)}">${escapeHtml(h)}</span></td><td>${escapeHtml(help)}</td><td>${req}</td></tr>`;
    });

    container.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>Column</th>
            <th>What to enter</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${rows.join("")}
        </tbody>
      </table>
    `;
  }

  function renderControlChoice() {
    const select = document.getElementById("controlChoice");
    const help = document.getElementById("controlChoiceHelp");
    if (!select) return;

    select.innerHTML = "";

    if (!state.treatments.length) {
      if (help) help.textContent = "";
      return;
    }

    const flagged = state.treatments.filter((t) => t.isControlFlagged).map((t) => t.name);
    const flaggedSet = new Set(flagged);

    const ordered = [];
    // Put flagged control candidates first
    for (const t of state.treatments) {
      if (flaggedSet.has(t.name)) ordered.push(t);
    }
    for (const t of state.treatments) {
      if (!flaggedSet.has(t.name)) ordered.push(t);
    }

    for (const t of ordered) {
      const opt = document.createElement("option");
      opt.value = t.name;
      opt.textContent = flaggedSet.has(t.name) ? `${t.name} (marked as control)` : t.name;
      if (t.name === state.controlName) opt.selected = true;
      select.appendChild(opt);
    }

    if (help) {
      if (flagged.length > 1) {
        help.textContent = "Multiple treatments are marked as control. Select the reference control to use in the comparison.";
      } else if (flagged.length === 1) {
        help.textContent = "Reference control detected from the is_control column. You can change it if needed.";
      } else {
        help.textContent = "No control flag was detected. Select the reference control used for comparisons.";
      }
    }
  }

  function renderLeaderboard() {
    const card = document.getElementById("leaderboardCard");
    const table = document.getElementById("leaderboardTable");

    if (!card || !table) return;

    const hasResults = state.results && state.results.rows && state.results.rows.length;
    if (!hasResults) {
      card.style.display = "none";
      return;
    }

    card.style.display = "block";

    const rows = state.results.rows.slice();
    const basicOnly = !!state.results.basicOnly;

    const controlName = state.controlName;
    const nonControl = rows.filter((r) => r.name !== controlName);

    const top = (() => {
      if (!nonControl.length) return null;
      if (basicOnly) {
        return nonControl
          .slice()
          .sort((a, b) => {
            const aVal = Number.isNaN(a.netAnnualValue) ? -Infinity : a.netAnnualValue;
            const bVal = Number.isNaN(b.netAnnualValue) ? -Infinity : b.netAnnualValue;
            return bVal - aVal;
          })[0];
      }
      return nonControl
        .slice()
        .sort((a, b) => {
          const aVal = Number.isNaN(a.npv) ? -Infinity : a.npv;
          const bVal = Number.isNaN(b.npv) ? -Infinity : b.npv;
          return bVal - aVal;
        })[0];
    })();

    const header = basicOnly
      ? `<h2>Indicative ranking</h2>
         <p class="tab-intro">Based on average yield and average costs. Full discounted results require complete inputs.</p>`
      : `<h2>Top ranked treatment</h2>
         <p class="tab-intro">Based on discounted net profit per hectare relative to the selected control.</p>`;

    const topHtml = (() => {
      if (!top) return `<p class="muted">No treatment results are available.</p>`;

      if (basicOnly) {
        const net = Number.isNaN(top.netAnnualValue) ? "NA" : formatCurrency(top.netAnnualValue);
        const delta = Number.isNaN(top.deltaNetAnnual) ? "NA" : formatCurrency(top.deltaNetAnnual);
        return `
          <div class="leaderboard-highlight">
            <p><strong>${escapeHtml(top.name)}</strong> has the highest indicative net value (${net}), which is ${delta} relative to the control.</p>
          </div>
        `;
      }

      const npv = Number.isNaN(top.npv) ? "NA" : formatCurrency(top.npv);
      const delta = Number.isNaN(top.deltaNpv) ? "NA" : formatCurrency(top.deltaNpv);
      return `
        <div class="leaderboard-highlight">
          <p><strong>${escapeHtml(top.name)}</strong> has the highest net profit (${npv}), which is ${delta} relative to the control.</p>
        </div>
      `;
    })();

    const cols = basicOnly
      ? ["Rank", "Treatment", "Yield mean (t/ha)", "Cost mean ($/ha)", "Net value (1 year)", "Change vs control"]
      : ["Rank", "Treatment", "NPV ($/ha)", "Change vs control", "Change in costs", "BCR"];

    const headerRow = `<tr>${cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr>`;

    const bodyRows = rows
      .filter((r) => r.name !== controlName)
      .slice(0, 10)
      .map((r) => {
        if (basicOnly) {
          const y = Number.isNaN(r.meanOutcome) ? "NA" : formatNumber(r.meanOutcome, 2);
          const c = Number.isNaN(r.meanCost) ? "NA" : formatCurrency(r.meanCost);
          const n = Number.isNaN(r.netAnnualValue) ? "NA" : formatCurrency(r.netAnnualValue);
          const d = Number.isNaN(r.deltaNetAnnual) ? "NA" : formatCurrency(r.deltaNetAnnual);
          return `<tr>
            <td>${escapeHtml(String(r.rank))}</td>
            <td>${escapeHtml(r.name)}</td>
            <td>${escapeHtml(y)}</td>
            <td>${escapeHtml(c)}</td>
            <td>${escapeHtml(n)}</td>
            <td>${escapeHtml(d)}</td>
          </tr>`;
        }

        const npv = Number.isNaN(r.npv) ? "NA" : formatCurrency(r.npv);
        const delta = Number.isNaN(r.deltaNpv) ? "NA" : formatCurrency(r.deltaNpv);
        const costDelta = Number.isNaN(r.deltaPvCosts) ? "NA" : formatCurrency(r.deltaPvCosts);
        const bcr = r.bcr == null || Number.isNaN(r.bcr) ? "NA" : r.bcr.toFixed(2);

        return `<tr>
          <td>${escapeHtml(String(r.rank))}</td>
          <td>${escapeHtml(r.name)}</td>
          <td>${escapeHtml(npv)}</td>
          <td>${escapeHtml(delta)}</td>
          <td>${escapeHtml(costDelta)}</td>
          <td>${escapeHtml(bcr)}</td>
        </tr>`;
      })
      .join("");

    table.innerHTML = `
      ${header}
      ${topHtml}
      <div class="table-wrap">
        <table class="result-table">
          <thead>${headerRow}</thead>
          <tbody>${bodyRows || ""}</tbody>
        </table>
      </div>
    `;
  }

  function renderComparisonTable() {
    const table = document.getElementById("comparisonTable");
    if (!table) return;

    const basicNote = document.getElementById("basicAnalysisNote");
    if (basicNote) {
      if (state.results && state.results.basicOnly) {
        basicNote.style.display = "inline";
        basicNote.textContent = " Indicative summary shown because full discounted results could not be calculated from the available inputs.";
      } else {
        basicNote.style.display = "none";
        basicNote.textContent = "";
      }
    }

    table.innerHTML = "";

    const { treatments } = state.results;
    if (!treatments || !treatments.length) {
      table.innerHTML =
        '<tbody><tr><td class="small muted">No results yet. Load data and apply scenario settings.</td></tr></tbody>';
      return;
    }

    const control = treatments.find((t) => t.isControl) || treatments[0];
    const nonControl = treatments.filter((t) => !t.isControl);
    const ordered = [control, ...nonControl];

    const indicators = state.results.basicOnly ? [
      { key: "meanOutcome", label: "Average yield (t/ha)" },
      { key: "meanCost", label: "Average total cost ($/ha)" },
      { key: "netAnnualValue", label: "Net value (one year, $/ha)" },
      { key: "deltaNetAnnual", label: "Change vs control (one year, $/ha)" }
    ] : [
      { key: "pvBenefits", label: "Present value benefits ($/ha)" },
      { key: "pvTotalCosts", label: "Present value costs ($/ha)" },
      { key: "npv", label: "Net profit (NPV, $/ha)" },
      { key: "deltaNpv", label: "Change in NPV vs control ($/ha)" },
      { key: "deltaPvCosts", label: "Change in costs vs control ($/ha)" },
      { key: "bcr", label: "Benefit cost ratio (BCR)" }
    ];

    const lines = [];
    lines.push(headers.join(","));

    for (const t of treatments) {
      const row = [
        `"${t.name.replace(/"/g, '""')}"`,
        t.isControl ? "TRUE" : "FALSE",
        Number.isNaN(t.avgYield) ? "" : t.avgYield.toFixed(4),
        Number.isNaN(t.avgVarCost) ? "" : t.avgVarCost.toFixed(4),
        Number.isNaN(t.avgCapCost) ? "" : t.avgCapCost.toFixed(4),
        Number.isNaN(t.pvBenefits) ? "" : t.pvBenefits.toFixed(2),
        Number.isNaN(t.pvTotalCosts) ? "" : t.pvTotalCosts.toFixed(2),
        Number.isNaN(t.npv) ? "" : t.npv.toFixed(2),
        Number.isNaN(t.bcr) ? "" : t.bcr.toFixed(4),
        Number.isNaN(t.roi) ? "" : t.roi.toFixed(2),
        t.rank,
        Number.isNaN(t.deltaNpv) ? "" : t.deltaNpv.toFixed(2),
        Number.isNaN(t.deltaPvCosts) ? "" : t.deltaPvCosts.toFixed(2)
      ];
      lines.push(row.join(","));
    }

    lines.push("#");
    lines.push(`# Generated by ${PROJECT.name}`);
    lines.push(`# ${PROJECT.partnerPlaceholder}`);

    const blob = new Blob([lines.join("\n")], {
      type: "text/csv;charset=utf-8;"
    });
    downloadBlob(blob, "treatment_summary.csv");
    showToast("Treatment summary (CSV) downloaded.", "success");
  }

  function exportComparisonCSV() {
    const { treatments } = state.results;
    if (!treatments || !treatments.length) {
      showToast("No results table to export.", "error");
      return;
    }

    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (discounted)" },
      { key: "pvTotalCosts", label: "Total costs over time (discounted)" },
      { key: "npv", label: "Net profit over time" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment (percent)" },
      { key: "rank", label: "Overall ranking" },
      {
        key: "deltaNpv",
        label: "Difference in net profit compared with control"
      },
      {
        key: "deltaPvCosts",
        label: "Difference in total cost compared with control"
      }
    ];

    const control = treatments.find((t) => t.isControl) || treatments[0];
    const nonControl = treatments.filter((t) => !t.isControl);
    const ordered = [control, ...nonControl];

    const header = ["What is measured"];
    for (const t of ordered) {
      header.push(t.isControl ? "Control (baseline)" : t.name);
    }

    const lines = [];
    lines.push(header.map(csvEscape).join(","));

    for (const ind of indicators) {
      const row = [ind.label];
      for (const t of ordered) {
        const r = treatments.find((x) => x.name === t.name);
        let value = "";
        if (ind.key === "pvBenefits") {
          value = Number.isNaN(r.pvBenefits) ? "" : r.pvBenefits.toFixed(2);
        } else if (ind.key === "pvTotalCosts") {
          value = Number.isNaN(r.pvTotalCosts) ? "" : r.pvTotalCosts.toFixed(2);
        } else if (ind.key === "npv") {
          value = Number.isNaN(r.npv) ? "" : r.npv.toFixed(2);
        } else if (ind.key === "bcr") {
          value = Number.isNaN(r.bcr) ? "" : r.bcr.toFixed(4);
        } else if (ind.key === "roi") {
          value = Number.isNaN(r.roi) ? "" : r.roi.toFixed(2);
        } else if (ind.key === "rank") {
          value = r.rank;
        } else if (ind.key === "deltaNpv") {
          value = Number.isNaN(r.deltaNpv) ? "" : r.deltaNpv.toFixed(2);
        } else if (ind.key === "deltaPvCosts") {
          value = Number.isNaN(r.deltaPvCosts) ? "" : r.deltaPvCosts.toFixed(2);
        }
        row.push(value);
      }
      lines.push(row.map(csvEscape).join(","));
    }

    lines.push("#");
    lines.push(`# Generated by ${PROJECT.name}`);
    lines.push(`# ${PROJECT.partnerPlaceholder}`);

    lines.push("#");
    lines.push(`# Generated by ${PROJECT.name}`);
    lines.push(`# ${PROJECT.partnerPlaceholder}`);

    const blob = new Blob([lines.join("\n")], {
      type: "text/csv;charset=utf-8;"
    });
    downloadBlob(blob, "comparison_to_control.csv");
    showToast("Comparison-to-control results (CSV) downloaded.", "success");
  }

  function csvEscape(value) {
    if (value === null || value === undefined) return "";
    const s = String(value);
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  function exportWorkbook() {
    if (!window.XLSX) {
      showToast("Excel library not available.", "error");
      return;
    }
    if (!state.headers.length || !state.rows.length) {
      showToast("No dataset loaded for workbook export.", "error");
      return;
    }
    const wb = XLSX.utils.book_new();

    // Attribution sheet (kept minimal so it appears clearly in exports)
    const attributionAoA = [
      [PROJECT.name],
      [PROJECT.partnerPlaceholder],
      [""],
      ["Note"],
      [
        "Treatments may appear in multiple rows (replicates). The tool aggregates rows by treatment and compares the average treatment with the average control."
      ]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(attributionAoA), "Attribution");

    // Sheet 1: Cleaned dataset
    const dataSheetAoA = [state.headers.slice()];
    for (const row of state.rows) {
      dataSheetAoA.push(
        state.headers.map((h) =>
          row[h] === undefined || row[h] === null ? "" : row[h]
        )
      );
    }
    const wsData = XLSX.utils.aoa_to_sheet(dataSheetAoA);
    XLSX.utils.book_append_sheet(wb, wsData, "Dataset");

    // Sheet 2: Treatment summary
    const { treatments } = state.results;
    if (treatments && treatments.length) {
      const summaryAoA = [
        [
          "Treatment name",
          "Is control",
          "Average yield (t per ha)",
          "Average variable cost (per ha)",
          "Average capital cost (per ha)",
          "Total benefits over time (discounted)",
          "Total costs over time (discounted)",
          "Net profit over time",
          "Benefit per dollar spent",
          "Return on investment (percent)",
          "Rank",
          "Difference in net profit vs control",
          "Difference in total cost vs control"
        ]
      ];

      for (const t of treatments) {
        summaryAoA.push([
          t.name,
          t.isControl ? "TRUE" : "FALSE",
          Number.isNaN(t.avgYield) ? "" : t.avgYield,
          Number.isNaN(t.avgVarCost) ? "" : t.avgVarCost,
          Number.isNaN(t.avgCapCost) ? "" : t.avgCapCost,
          Number.isNaN(t.pvBenefits) ? "" : t.pvBenefits,
          Number.isNaN(t.pvTotalCosts) ? "" : t.pvTotalCosts,
          Number.isNaN(t.npv) ? "" : t.npv,
          Number.isNaN(t.bcr) ? "" : t.bcr,
          Number.isNaN(t.roi) ? "" : t.roi,
          t.rank,
          Number.isNaN(t.deltaNpv) ? "" : t.deltaNpv,
          Number.isNaN(t.deltaPvCosts) ? "" : t.deltaPvCosts
        ]);
      }

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoA);
      XLSX.utils.book_append_sheet(wb, wsSummary, "Treatment summary");
    }

    // Sheet 3: Comparison to control
    const compAoA = [];
    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (discounted)" },
      { key: "pvTotalCosts", label: "Total costs over time (discounted)" },
      { key: "npv", label: "Net profit over time" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment (percent)" },
      { key: "rank", label: "Overall ranking" },
      {
        key: "deltaNpv",
        label: "Difference in net profit compared with control"
      },
      {
        key: "deltaPvCosts",
        label: "Difference in total cost compared with control"
      }
    ];

    if (treatments && treatments.length) {
      const control = treatments.find((t) => t.isControl) || treatments[0];
      const nonControl = treatments.filter((t) => !t.isControl);
      const ordered = [control, ...nonControl];

      const headerRow = ["What is measured"];
      for (const t of ordered) {
        headerRow.push(t.isControl ? "Control (baseline)" : t.name);
      }
      compAoA.push(headerRow);

      for (const ind of indicators) {
        const row = [ind.label];
        for (const t of ordered) {
          const r = treatments.find((x) => x.name === t.name);
          let v = "";
          if (ind.key === "pvBenefits") v = r.pvBenefits;
          else if (ind.key === "pvTotalCosts") v = r.pvTotalCosts;
          else if (ind.key === "npv") v = r.npv;
          else if (ind.key === "bcr") v = r.bcr;
          else if (ind.key === "roi") v = r.roi;
          else if (ind.key === "rank") v = r.rank;
          else if (ind.key === "deltaNpv") v = r.deltaNpv;
          else if (ind.key === "deltaPvCosts") v = r.deltaPvCosts;
          row.push(v);
        }
        compAoA.push(row);
      }

      const wsComp = XLSX.utils.aoa_to_sheet(compAoA);
      XLSX.utils.book_append_sheet(wb, wsComp, "Comparison to control");
    }

    // Sheet 4: Attribution
    const now = new Date();
    const attrAoA = [
      ["Project", PROJECT.name],
      ["Partners", PROJECT.partnerPlaceholder],
      ["Generated", now.toISOString()],
      [
        "Note",
        "Treatments with multiple rows (replicates) are aggregated and compared as average treatment vs average control."
      ]
    ];
    const wsAttr = XLSX.utils.aoa_to_sheet(attrAoA);
    XLSX.utils.book_append_sheet(wb, wsAttr, "Attribution");

    XLSX.writeFile(wb, "cba_results.xlsx");
    showToast("Excel workbook downloaded.", "success");
  }


  function exportCleanDatasetXLSX() {
    if (!window.XLSX) {
      showToast("Excel library not available.", "error");
      return;
    }
    const headers = (state.cleanedHeaders && state.cleanedHeaders.length) ? state.cleanedHeaders : state.headers;
    const rows = (state.cleanedRows && state.cleanedRows.length) ? state.cleanedRows : state.rows;

    if (!headers.length || !rows.length) {
      showToast("No dataset to export.", "error");
      return;
    }

    const wb = XLSX.utils.book_new();
    const aoa = [headers.slice()];
    for (const r of rows) {
      aoa.push(headers.map((h) => (r[h] == null ? "" : r[h])));
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, "Cleaned data");

    const meta = [
      ["Upload timestamp", state.audit.uploadedAt || ""],
      ["File name", state.audit.fileName || ""],
      ["Rows", String(state.audit.rowCount || rows.length)],
      ["Treatments", String(state.audit.treatmentCount || "")],
      ["Tool version", state.audit.version || VERSION]
    ];
    const wsMeta = XLSX.utils.aoa_to_sheet(meta);
    XLSX.utils.book_append_sheet(wb, wsMeta, "Audit");

    XLSX.writeFile(wb, "cleaned_dataset.xlsx");
    showToast("Cleaned dataset downloaded.", "success");
  }

  function exportWordReport() {
    if (!state.treatments || !state.treatments.length) {
      showToast("Run an analysis first.", "error");
      return;
    }

    const title = escapeHtml(PROJECT.name || "Cost benefit analysis report");
    const now = new Date();
    const uploadedAt = state.audit.uploadedAt || now.toISOString();
    const fileName = state.audit.fileName || "Uploaded dataset";

    const controlName = state.controlName || "";
    const treatments = state.treatments.slice().sort((a, b) => (b.npv || 0) - (a.npv || 0));
    const nonControl = treatments.filter((t) => t.name !== controlName);
    const headline = nonControl.length ? nonControl[0] : null;

    const assumptionsHtml = buildAssumptionsHtml(true);
    const auditHtml = `
      <table class="mini">
        <tr><th>Upload timestamp</th><td>${escapeHtml(uploadedAt)}</td></tr>
        <tr><th>File name</th><td>${escapeHtml(fileName)}</td></tr>
        <tr><th>Rows</th><td>${escapeHtml(String(state.audit.rowCount || 0))}</td></tr>
        <tr><th>Treatments</th><td>${escapeHtml(String(state.audit.treatmentCount || 0))}</td></tr>
        <tr><th>Reference control</th><td>${escapeHtml(controlName || "Not set")}</td></tr>
        <tr><th>Tool version</th><td>${escapeHtml(state.audit.version || VERSION)}</td></tr>
      </table>
    `;

    const execSummary = (() => {
      if (!headline) return `<p>No treatment results are available.</p>`;
      const delta = formatCurrency(headline.deltaNpv);
      const npv = formatCurrency(headline.npv);
      const costDelta = formatCurrency(headline.deltaPvCosts);
      const bcr = (headline.bcr == null || Number.isNaN(headline.bcr)) ? "NA" : headline.bcr.toFixed(2);
      return `
        <p><strong>Headline result:</strong> ${escapeHtml(headline.name)} ranks first on discounted net profit per hectare (${npv}), which is ${delta} relative to the control.</p>
        <p><strong>Cost change relative to control:</strong> ${costDelta}. <strong>Benefit cost ratio:</strong> ${escapeHtml(bcr)}.</p>
      `;
    })();

    const rankingRows = nonControl.slice(0, 10).map((t, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>${escapeHtml(t.name)}</td>
        <td>${escapeHtml(formatCurrency(t.npv))}</td>
        <td>${escapeHtml(formatCurrency(t.deltaNpv))}</td>
        <td>${escapeHtml(formatCurrency(t.deltaPvCosts))}</td>
        <td>${escapeHtml((t.bcr == null || Number.isNaN(t.bcr)) ? "NA" : t.bcr.toFixed(2))}</td>
      </tr>
    `).join("");

    const rankingTable = `
      <table class="report-table">
        <thead>
          <tr>
            <th>Rank</th>
            <th>Treatment</th>
            <th>Net profit (NPV)</th>
            <th>Change vs control</th>
            <th>Change in costs vs control</th>
            <th>BCR</th>
          </tr>
        </thead>
        <tbody>${rankingRows}</tbody>
      </table>
    `;

    const compTableHtml = document.getElementById("comparisonTable") ? document.getElementById("comparisonTable").outerHTML : "";

    const chartNet = document.getElementById("chartNetProfitVsControl");
    const chartCb = document.getElementById("chartCostsBenefits");
    const netImg = chartNet && chartNet.toDataURL ? chartNet.toDataURL("image/png") : "";
    const cbImg = chartCb && chartCb.toDataURL ? chartCb.toDataURL("image/png") : "";

    const chartsHtml = `
      <h2>Charts</h2>
      ${netImg ? `<h3>Extra net profit compared with control</h3><img class="chart-img" src="${netImg}" alt="Extra net profit compared with control">` : ""}
      ${cbImg ? `<h3>Total discounted costs and benefits</h3><img class="chart-img" src="${cbImg}" alt="Total discounted costs and benefits">` : ""}
    `;

    const footerHtml = `
      <div class="report-footer">
        <p><strong>${escapeHtml(PROJECT.name)}</strong></p>
        <p>${escapeHtml(PROJECT.partnerPlaceholder)}</p>
      </div>
    `;

    const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${title}</title>
<style>
  body { font-family: Arial, Helvetica, sans-serif; color: #313131; line-height: 1.45; margin: 28px; }
  h1 { font-size: 22px; margin: 0 0 8px; }
  h2 { font-size: 16px; margin: 22px 0 8px; }
  h3 { font-size: 13px; margin: 14px 0 6px; }
  p { margin: 0 0 10px; }
  .muted { color: #626262; }
  .mini { border-collapse: collapse; width: 100%; margin: 8px 0 16px; }
  .mini th, .mini td { border: 1px solid #d3d3d3; padding: 6px 8px; text-align: left; vertical-align: top; }
  .mini th { background: #f5f5f5; width: 32%; }
  .report-table { border-collapse: collapse; width: 100%; margin: 10px 0 16px; }
  .report-table th, .report-table td { border: 1px solid #d3d3d3; padding: 6px 8px; text-align: left; vertical-align: top; }
  .report-table th { background: #f5f5f5; }
  .chart-img { max-width: 100%; height: auto; border: 1px solid #d3d3d3; padding: 6px; margin: 6px 0 14px; }
  .report-footer { margin-top: 26px; border-top: 2px solid #d3d3d3; padding-top: 10px; }
</style>
</head>
<body>
  <h1>${title}</h1>
  <p class="muted">Reference control: ${escapeHtml(controlName || "Not set")}</p>

  <h2>Executive summary</h2>
  ${execSummary}

  <h2>Top ranked treatments</h2>
  ${rankingTable}

  <h2>Assumptions used</h2>
  ${assumptionsHtml}

  <h2>Results audit trail</h2>
  ${auditHtml}

  <h2>Comparison to control</h2>
  ${compTableHtml}

  ${chartsHtml}

  ${footerHtml}
</body>
</html>`;

    const blob = new Blob([html], { type: "application/msword;charset=utf-8;" });
    downloadBlob(blob, "cba_report.doc");
    showToast("Word report downloaded.", "success");
  }

  function buildTemplateRows(headers) {
  // Blank rows only. Users can add as many rows as needed offline.
  const blank = {};
  for (const h of headers) blank[h] = "";
  // Provide a few blank rows to make the template immediately editable in spreadsheet software.
  return [ { ...blank }, { ...blank }, { ...blank } ];
}


  async function ensureTemplateHeaders() {
    if (state.templateHeaders && state.templateHeaders.length) {
      return state.templateHeaders;
    }
    // Use example dataset headers as the authoritative template if available
    try {
      const response = await fetch("faba_beans_trial_clean_named.tsv", {
        cache: "no-cache"
      });
      if (!response.ok) throw new Error("Template headers could not be fetched.");
      const text = await response.text();
      const parsed = parseDelimitedText(text);
      state.templateHeaders = normaliseTemplateHeaders(parsed.headers);
      return state.templateHeaders;
    } catch (e) {
      // Fallback to minimum required fields
      state.templateHeaders = [
        "plot_id",
        "replicate_id",
        "treatment_name",
        "is_control",
        "yield_t_ha",
        "total_cost_per_ha",
        "cost_amendment_input_per_ha"
      ];
      return state.templateHeaders;
    }
  }

  async function downloadTemplateTSV() {
    const headers = await ensureTemplateHeaders();
    const headerLine = headers.join("\t");
    const rows = buildTemplateRows(headers);
    const lines = [headerLine];
    for (const row of rows) {
      lines.push(headers.map((h) => (row[h] == null ? "" : String(row[h]))).join("\t"));
    }
    const blob = new Blob([lines.join("\n")], {
      type: "text/tab-separated-values;charset=utf-8;"
    });
    downloadBlob(blob, "trial_data_template.tsv");
    showToast("Template downloaded.", "success");
  }

  async function downloadTemplateXLSX() {
    if (!window.XLSX) {
      showToast("Excel library not available.", "error");
      return;
    }
    const headers = await ensureTemplateHeaders();
    const wb = XLSX.utils.book_new();
    const rows = buildTemplateRows(headers);
    const aoa = [headers.slice()];
    for (const r of rows) aoa.push(headers.map((h) => (r[h] == null ? "" : r[h])));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    const wsInfo = XLSX.utils.aoa_to_sheet([
      ["How to use"],
      ["Fill one row per plot/replicate."],
      ["Mark control rows with is_control = TRUE/1/yes."],
      ["Provide numeric yield_t_ha and total_cost_per_ha for at least some rows."],
      ["Optional columns may be left blank."]
    ]);
    XLSX.utils.book_append_sheet(wb, wsInfo, "Instructions");
    XLSX.writeFile(wb, "trial_data_template.xlsx");
    showToast("Template downloaded.", "success");
  }

  // =========================
  // 6) DATA LOAD & COMMIT
  // =========================

  async function loadDefaultDataset() {
    try {
      const response = await fetch("faba_beans_trial_clean_named.tsv", {
        cache: "no-cache"
      });
      if (!response.ok) {
        throw new Error("Default dataset could not be fetched.");
      }
      const text = await response.text();
      commitParsedData(text, "Example dataset loaded. You can upload your own data at any time.");
      setUploadStatus("Example dataset loaded. You can upload your own data at any time.", "info");
    } catch (err) {
      console.error(err);
      showToast(
        "Could not load default dataset. You can still upload or paste data.",
        "error"
      );
    }
  }

  function buildCleanedDataset(parsed, templateHeaders, mapping) {
    const originalHeaders = parsed.headers.slice();
    const template = (templateHeaders && templateHeaders.length) ? templateHeaders.slice() : originalHeaders.slice();

    const templateSet = new Set(template);
    const extraHeaders = originalHeaders.filter((h) => !templateSet.has(h));

    // Canonical-first headers: keep template order, then any extra columns from the file.
    const cleanedHeaders = template.concat(extraHeaders);

    const cleanedRows = parsed.rows.map((row) => {
      const out = {};
      // Template columns: pull from mapped columns or direct name match.
      for (const h of template) {
        // If this template header is one of the canonical columns, pull from mapping.
        let sourceHeader = null;
        if (h === "treatment_name") sourceHeader = mapping.treatment_name || null;
        else if (h === "is_control") sourceHeader = mapping.is_control || null;
        else if (h === "yield_t_ha") sourceHeader = mapping.yield_t_ha || null;
        else if (h === "total_cost_per_ha") sourceHeader = mapping.total_cost_per_ha || null;
        else if (h === "cost_amendment_input_per_ha") sourceHeader = mapping.cost_amendment_input_per_ha || null;
        else sourceHeader = originalHeaders.includes(h) ? h : null;

        out[h] = sourceHeader ? (row[sourceHeader] == null ? "" : row[sourceHeader]) : "";
      }
      // Extra columns: keep verbatim.
      for (const h of extraHeaders) {
        out[h] = row[h] == null ? "" : row[h];
      }
      return out;
    });

    return { cleanedHeaders, cleanedRows };
  }

  function renderDataPreview(preview) {
    const panel = document.getElementById("dataPreviewPanel");
    if (!panel) return;

    if (!preview) {
      panel.innerHTML = "";
      return;
    }

    const { stats, errors, warnings, mapping, controlCandidates } = preview;

    const issuesHtml = (() => {
      const items = [];
      for (const e of errors || []) {
        items.push(`<li class="status error">${escapeHtml(e)}</li>`);
      }
      for (const w of warnings || []) {
        items.push(`<li class="status warning">${escapeHtml(w)}</li>`);
      }
      if (!items.length) {
        return `<p class="status success">Validation passed. You can upload and run the analysis.</p>`;
      }
      return `<ul class="issues-list">${items.join("")}</ul>`;
    })();

    const mappingRows = Object.keys(mapping || {}).map((k) => {
      const v = mapping[k] || "(not found)";
      return `<tr><th>${escapeHtml(k)}</th><td>${escapeHtml(v)}</td></tr>`;
    }).join("");

    const controlList = (controlCandidates && controlCandidates.length)
      ? controlCandidates.map((c) => `<li>${escapeHtml(c)}</li>`).join("")
      : "<li>None detected</li>";

    panel.innerHTML = `
      <h3>Preview your data</h3>
      <div class="preview-grid">
        <div class="preview-item"><div class="preview-label">Rows</div><div class="preview-value">${stats.rowCount}</div></div>
        <div class="preview-item"><div class="preview-label">Columns</div><div class="preview-value">${stats.columnCount}</div></div>
        <div class="preview-item"><div class="preview-label">Treatments</div><div class="preview-value">${stats.treatmentCount || "-"}</div></div>
        <div class="preview-item"><div class="preview-label">Control rows</div><div class="preview-value">${stats.controlRowCount}</div></div>
      </div>

      <div class="field-group">
        <p class="small muted">Column mapping used by the tool</p>
        <table class="mini-table">
          <tbody>${mappingRows}</tbody>
        </table>
      </div>

      <div class="field-group">
        <p class="small muted">Rows marked as control (by treatment name)</p>
        <ul class="compact-list">${controlList}</ul>
      </div>

      <div class="field-group">
        <p class="small muted">Issues and warnings</p>
        ${issuesHtml}
      </div>

      <div class="field-group preview-actions" id="previewActions" style="display:${(errors && errors.length) ? "block" : "none"};">
        <button id="btnDownloadHeaderTemplate" class="btn secondary" type="button">
          Download template with your detected headers
        </button>
        <button id="btnDownloadCleanedNow" class="btn ghost" type="button">
          Download cleaned data (TSV)
        </button>
      </div>
    `;

    const btnHdrTpl = document.getElementById("btnDownloadHeaderTemplate");
    if (btnHdrTpl) {
      btnHdrTpl.addEventListener("click", () => {
        if (!state.lastPreview || !state.lastPreview.parsed) return;
        downloadTemplateWithDetectedHeaders(state.lastPreview.parsed.headers);
      });
    }

    const btnCleanNow = document.getElementById("btnDownloadCleanedNow");
    if (btnCleanNow) {
      btnCleanNow.addEventListener("click", () => {
        exportCleanDatasetTSV();
      });
    }
  }

  function downloadTemplateWithDetectedHeaders(detectedHeaders) {
    const baseHeaders = (detectedHeaders && detectedHeaders.length) ? detectedHeaders.slice() : [];
    // Add canonical required headers if missing
    const required = ["treatment_name", "is_control", "yield_t_ha", "total_cost_per_ha"];
    const out = baseHeaders.slice();
    for (const r of required) {
      if (!out.includes(r)) out.push(r);
    }
    const line1 = out.join("\t");
    const blankRow = out.map(() => "").join("\t");
    const blob = new Blob([line1 + "\n" + blankRow + "\n"], { type: "text/tab-separated-values;charset=utf-8;" });
    downloadBlob(blob, "template_with_detected_headers.tsv");
    showToast("Template downloaded. Rename or add columns as needed.", "success");
  }

  function commitParsedData(text, meta) {
    const parsed = parseDelimitedText(text);

    // Keep template headers aligned with the authoritative dataset structure.
    if (!state.templateHeaders || !state.templateHeaders.length) {
      state.templateHeaders = normaliseTemplateHeaders(parsed.headers);
    }

    const preview = validateParsedData(parsed.headers, parsed.rows);
    preview.parsed = parsed;
    state.lastPreview = preview;
    renderDataPreview(preview);

    if (!preview.canRun) {
      const msg = (preview.errors && preview.errors.length)
        ? preview.errors[0]
        : "File could not be loaded. Check required columns and minimum data requirements.";
      showToast(msg, "error");
      return { ok: false, message: msg, preview };
    }

    // Save audit trail
    state.audit.fileName = meta && meta.fileName ? meta.fileName : "";
    state.audit.uploadedAt = meta && meta.uploadedAt ? meta.uploadedAt : new Date().toISOString();
    state.audit.rowCount = preview.stats.rowCount;
    state.audit.treatmentCount = preview.stats.treatmentCount;
    state.audit.controlCandidates = preview.controlCandidates || [];
    state.audit.columnMapping = preview.mapping || {};
    state.audit.version = VERSION;

    // Store raw and parsed
    state.rawText = text;
    state.headers = parsed.headers;
    state.rows = parsed.rows;

    // Build cleaned dataset for export and transparency
    const cleaned = buildCleanedDataset(parsed, state.templateHeaders, preview.mapping || {});
    state.cleanedHeaders = cleaned.cleanedHeaders;
    state.cleanedRows = cleaned.cleanedRows;

    aggregateTreatments();
    computeCBA();
    renderAll();

    
    const btnRun = document.getElementById("btnRunAnalysis");
    if (btnRun) btnRun.disabled = false;
const detected = (() => {
      const nRows = (state.rows && state.rows.length) ? state.rows.length : 0;
      const nTreat = (state.treatments && state.treatments.length) ? state.treatments.length : 0;
      const control = state.controlName ? `Control: ${state.controlName}` : "Control: not set";
      return `${nRows} rows loaded; ${nTreat} treatments detected; ${control}.`;
    })();

    const successMessage = meta && meta.successMessage ? meta.successMessage : "Data uploaded successfully. Next: review results.";
    const finalMsg = `${successMessage} ${detected}`;
    showToast(finalMsg, "success");
    return { ok: true, message: finalMsg, preview };
  }

  // =========================
  // 7) UI WIRING
  // =========================

  function onTabClick(event) {
    const button = event.currentTarget;
    const tabId = button.getAttribute("data-tab");
    if (!tabId) return;

    document.querySelectorAll(".tab-button").forEach((btn) => {
      btn.classList.toggle("active", btn === button);
    });
    document.querySelectorAll(".tab-panel").forEach((panel) => {
      panel.classList.toggle("active", panel.id === tabId);
    });
  }

  function onApplyScenario() {
    const priceInput = document.getElementById("pricePerTonne");
    const yearsInput = document.getElementById("years");
    const persInput = document.getElementById("persistenceYears");
    const discInput = document.getElementById("discountRate");
    const controlSelect = document.getElementById("controlChoice");

    if (priceInput) {
      state.params.pricePerTonne = parseNumber(priceInput.value) || 0;
    }
    if (yearsInput) {
      state.params.years = parseInt(yearsInput.value, 10) || 1;
    }
    if (persInput) {
      state.params.persistenceYears =
        parseInt(persInput.value, 10) || state.params.years;
    }
    if (discInput) {
      state.params.discountRate = parseNumber(discInput.value) || 0;
    }
    if (controlSelect && controlSelect.value) {
      state.controlName = controlSelect.value;
    }

    computeCBA();
    renderAll();
    showToast("Scenario settings updated.", "success");
  }

  function onFileInputChange(event) {
    const file = event.target.files && event.target.files[0];
    state.pendingFile = file || null;

    if (!file) {
      setUploadStatus("No file selected.", "warning");
      renderDataPreview(null);
      return;
    }

    setUploadStatus(`File selected: ${file.name}. Next: validate or upload.`, "info");
    showToast("File selected. Validate or upload.", "info");

    // Quick preview (non-blocking) using the first read
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseDelimitedText(String(e.target.result || ""));
        const preview = validateParsedData(parsed.headers, parsed.rows);
        preview.parsed = parsed;
        state.lastPreview = preview;
        renderDataPreview(preview);
      } catch (err) {
        console.error(err);
        const msg = err && err.message ? err.message : "Could not read the selected file.";
        setUploadStatus(msg, "error");
        showToast(msg, "error");
      }
    };
    reader.readAsText(file);
  }

  function onValidateSelectedFileClick() {
    const file = state.pendingFile || (document.getElementById("fileInput") && document.getElementById("fileInput").files && document.getElementById("fileInput").files[0]);
    if (!file) {
      showToast("Select a file first.", "warning");
      return;
    }
    setUploadStatus("Validating selected file locally.", "info");

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseDelimitedText(String(e.target.result || ""));
        const preview = validateParsedData(parsed.headers, parsed.rows);
        preview.parsed = parsed;
        state.lastPreview = preview;
        renderDataPreview(preview);

        if (preview.canRun) {
          setUploadStatus("Validation passed. You can upload and run the analysis.", "success");
          showToast("Validation passed. Next: upload the file.", "success");
        } else {
          setUploadStatus("Validation found issues. See the preview panel for details.", "warning");
          showToast("Validation found issues. See the preview panel.", "warning");
        }
      } catch (err) {
        console.error(err);
        const msg = err && err.message ? err.message : "Could not validate the selected file.";
        setUploadStatus(msg, "error");
        showToast(msg, "error");
      }
    };
    reader.readAsText(file);
  }

  function onUploadSelectedFileClick() {
    const file = state.pendingFile || (document.getElementById("fileInput") && document.getElementById("fileInput").files && document.getElementById("fileInput").files[0]);
    if (!file) {
      showToast("Select a file first.", "warning");
      return;
    }

    setUploadStatus("Uploading and processing data.", "info");

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = String(e.target.result || "");
        const meta = {
          fileName: file.name,
          uploadedAt: new Date().toISOString(),
          successMessage: "Data uploaded successfully. Next: review results."
        };
        const result = commitParsedData(text, meta);
        if (result && result.ok) {
          setUploadStatus(result.message, "success");
        } else {
          setUploadStatus(result.message || "Upload failed. See the preview panel for details.", "error");
        }
      } catch (err) {
        console.error(err);
        const msg = err && err.message ? err.message : "Upload failed.";
        setUploadStatus(msg, "error");
        showToast(msg, "error");
      }
    };
    reader.readAsText(file);
  }

  function onResetToolClick() {
    // Clear state while keeping parameters and template headers
    state.rawText = "";
    state.headers = [];
    state.rows = [];
    state.cleanedHeaders = [];
    state.cleanedRows = [];
    state.treatments = [];
    state.controlName = null;
    state.columnMap = null;
    state.pendingFile = null;
    state.lastPreview = null;
    state.audit.fileName = "";
    state.audit.uploadedAt = null;
    state.audit.rowCount = 0;
    state.audit.treatmentCount = 0;
    state.audit.controlCandidates = [];
    state.audit.columnMapping = {};
    state.assumptions = {};

    const fileInput = document.getElementById("fileInput");
    if (fileInput) fileInput.value = "";

    renderDataPreview(null);
    renderAll();

    // Return user to the first tab
    const btn = document.querySelector('.tab-button[data-tab="overviewTab"]');
    if (btn) btn.click();
    setUploadStatus("Reset complete. Next: download the template or load your data.", "info");
    showToast("Reset complete.", "success");
  }

  function onLoadPastedClick() {
    const ta = document.getElementById("pasteInput");
    if (!ta) return;
    const text = ta.value;
    if (!text || !text.trim()) {
      showToast("Paste data before loading.", "error");
      return;
    }
    try {
      setUploadStatus("", "");
      commitParsedData(text, "Data uploaded successfully. Next: review results.");
      setUploadStatus("Data uploaded successfully. Next: run analysis and review results.", "success");
    } catch (err) {
      console.error(err);
      const msg = err && err.message ? err.message : "Could not parse pasted data. Check its format.";
      showToast(msg, "error");
      setUploadStatus(msg, "error");
    }
  }

  function onReloadDefaultClick() {
    loadDefaultDataset();
  }

  function onLeaderboardFilterChange() {
    renderLeaderboard();
  }

  function onCopyAiBriefing() {
    const ta = document.getElementById("aiBriefing");
    if (!ta) return;

    const text = (ta.value || "").trim();
    if (!text) {
      showToast("No summary prompt is available yet. Run an analysis first.", "warning");
      return;
    }

    // Prefer modern clipboard API; fall back to execCommand for older browsers.
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard
        .writeText(text)
        .then(() => showToast("AI summary prompt copied. Paste it into your AI assistant.", "success"))
        .catch(() => fallbackCopy(ta));
      return;
    }

    fallbackCopy(ta);

    function fallbackCopy(textarea) {
      textarea.focus();
      textarea.select();
      textarea.setSelectionRange(0, textarea.value.length);
      try {
        document.execCommand("copy");
        showToast("AI summary prompt copied. Paste it into your AI assistant.", "success");
      } catch (err) {
        showToast("Copy failed. Select the text and copy it manually.", "warning", 4200);
      }
    }
  }

  function openAiAssistant(url, label) {
    const promptEl = document.getElementById("aiBriefing");
    const promptText = (promptEl && promptEl.value ? promptEl.value.trim() : "");

    const openTarget = () => window.open(url, "_blank", "noopener");

    if (!promptText) {
      showToast("No summary prompt is available yet. Run an analysis first.", "warning");
      openTarget();
      return;
    }

    // Copy the prompt to clipboard when possible, then open the assistant in a new tab.
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard
        .writeText(promptText)
        .then(() => {
          showToast(`Prompt copied. Paste it into ${label}.`, "success");
          openTarget();
        })
        .catch(() => {
          showToast("Copy did not work. Copy the prompt manually, then paste it into the assistant.", "warning", 4200);
          openTarget();
        });
      return;
    }

    showToast("Clipboard is not available in this browser. Copy the prompt manually, then paste it into the assistant.", "warning", 5200);
    openTarget();
  }

  function attachEventListeners() {
    document.querySelectorAll(".tab-button").forEach((btn) => {
      btn.addEventListener("click", onTabClick);
    });

    const applyBtn = document.getElementById("applyScenario");
    if (applyBtn) applyBtn.addEventListener("click", onApplyScenario);

    const fileInput = document.getElementById("fileInput");
    if (fileInput) fileInput.addEventListener("change", onFileInputChange);

    const btnValidateSelected = document.getElementById("btnValidateSelectedFile");
    if (btnValidateSelected) btnValidateSelected.addEventListener("click", onValidateSelectedFileClick);

    const btnUploadSelected = document.getElementById("btnUploadSelectedFile");
    if (btnUploadSelected) btnUploadSelected.addEventListener("click", onUploadSelectedFileClick);

    const btnResetTool = document.getElementById("btnResetTool");
    if (btnResetTool) btnResetTool.addEventListener("click", onResetToolClick);

    const btnLoadPasted = document.getElementById("btnLoadPasted");
    if (btnLoadPasted)
      btnLoadPasted.addEventListener("click", onLoadPastedClick);

    const btnReloadDefault = document.getElementById("btnReloadDefault");
    if (btnReloadDefault)
      btnReloadDefault.addEventListener("click", onReloadDefaultClick);

    const btnTplTsv = document.getElementById("btnDownloadTemplateTSV");
    if (btnTplTsv) btnTplTsv.addEventListener("click", downloadTemplateTSV);

    const btnTplXlsx = document.getElementById("btnDownloadTemplateXLSX");
    if (btnTplXlsx) btnTplXlsx.addEventListener("click", downloadTemplateXLSX);

    const filterSelect = document.getElementById("leaderboardFilter");
    if (filterSelect)
      filterSelect.addEventListener("change", onLeaderboardFilterChange);

    const btnCleanData = document.getElementById("btnExportCleanData");
    if (btnCleanData)
      btnCleanData.addEventListener("click", exportCleanDatasetTSV);

    const btnCleanDataXlsx = document.getElementById("btnExportCleanDataXLSX");
    if (btnCleanDataXlsx)
      btnCleanDataXlsx.addEventListener("click", exportCleanDatasetXLSX);

    const btnSummary = document.getElementById("btnExportSummary");
    if (btnSummary)
      btnSummary.addEventListener("click", exportTreatmentSummaryCSV);

    const btnResults = document.getElementById("btnExportResults");
    if (btnResults)
      btnResults.addEventListener("click", exportComparisonCSV);

    const btnWorkbook = document.getElementById("btnExportWorkbook");
    if (btnWorkbook)
      btnWorkbook.addEventListener("click", exportWorkbook);

    const btnWord = document.getElementById("btnExportWordReport");
    if (btnWord) btnWord.addEventListener("click", exportWordReport);

    const btnCopyAi = document.getElementById("btnCopyAiBriefing");
    if (btnCopyAi)
      btnCopyAi.addEventListener("click", onCopyAiBriefing);

    

    const btnOpenChatGPT = document.getElementById("btnOpenChatGPT");
    if (btnOpenChatGPT)
      btnOpenChatGPT.addEventListener("click", () =>
        openAiAssistant("https://chat.openai.com/", "ChatGPT")
      );

    const btnOpenCopilot = document.getElementById("btnOpenCopilot");
    if (btnOpenCopilot)
      btnOpenCopilot.addEventListener("click", () =>
        openAiAssistant("https://copilot.microsoft.com/", "Microsoft Copilot")
      );

// Live scenario updates
    const priceInput = document.getElementById("pricePerTonne");
    const yearsInput = document.getElementById("years");
    const persInput = document.getElementById("persistenceYears");
    const discInput = document.getElementById("discountRate");
    const controlSelect = document.getElementById("controlChoice");

    if (priceInput)
      priceInput.addEventListener("change", onApplyScenario);
    if (yearsInput)
      yearsInput.addEventListener("change", onApplyScenario);
    if (persInput)
      persInput.addEventListener("change", onApplyScenario);
    if (discInput)
      discInput.addEventListener("change", onApplyScenario);
    if (controlSelect)
      controlSelect.addEventListener("change", onApplyScenario);
  }

  
  function buildAssumptionsHtml(asHtml) {
    const p = state.params || {};
    const price = parseNumber(p.pricePerTonne);
    const years = parseInt(p.years, 10);
    const dr = parseNumber(p.discountRate);
    const persistence = parseInt(p.persistenceYears, 10);

    const parts = [];
    parts.push(`Reference control: ${state.controlName || "Not set"}`);
    parts.push(`Grain price: ${formatCurrency(price)} per tonne`);
    parts.push(`Time horizon: ${Number.isFinite(years) ? years : ""} years`);
    parts.push(`Discount rate: ${Number.isFinite(dr) ? (dr * 100).toFixed(1) : ""} percent`);
    parts.push(`Benefit persistence: ${Number.isFinite(persistence) ? persistence : ""} years`);

    if (state.assumptions && state.assumptions.capitalCostAssumedZero) {
      parts.push("One-off establishment cost assumed to be zero because cost_amendment_input_per_ha was not provided");
    }

    const text = parts.join("; ") + ".";
    if (!asHtml) return text;
    return `<p class="small">${escapeHtml(text)}</p>`;
  }

  function renderAssumptionsUsed() {
    const el = document.getElementById("assumptionsUsed");
    if (!el) return;
    el.innerHTML = buildAssumptionsHtml(true);
  }

  function renderIndicativeBadge() {
    const badge = document.getElementById("indicativeBadge");
    if (!badge) return;

    const columnMap = state.columnMap || {};
    const assumptions = state.assumptions || {};
    const indicative = !!(assumptions.capitalCostAssumedZero || assumptions.missingYieldCells > 0 || assumptions.missingCostCells > 0 || !columnMap.replicate);

    badge.style.display = indicative ? "inline-flex" : "none";
  }

  function renderReplicationTable() {
    const table = document.getElementById("replicationTable");
    if (!table) return;

    if (!state.rows || !state.rows.length || !state.columnMap) {
      table.innerHTML = "<tr><td>No data loaded.</td></tr>";
      return;
    }

    const m = state.columnMap;
    const headers = [
      "treatment_name",
      "is_control",
      "plot_id",
      "replicate_id",
      "yield_t_ha",
      "total_cost_per_ha",
      "cost_amendment_input_per_ha"
    ];

    const getVal = (row, canon) => {
      const map = {
        treatment_name: m.treatmentName,
        is_control: m.isControl,
        plot_id: m.plotId,
        replicate_id: m.replicate,
        yield_t_ha: m.yield,
        total_cost_per_ha: m.variableCost,
        cost_amendment_input_per_ha: m.capitalCost
      };
      const h = map[canon];
      return h ? row[h] : "";
    };

    const rowsHtml = state.rows.slice(0, 1000).map((r) => {
      const t = String(getVal(r, "treatment_name") || "").trim();
      const isC = parseBoolean(getVal(r, "is_control")) ? "TRUE" : "FALSE";
      const plot = String(getVal(r, "plot_id") || "");
      const rep = String(getVal(r, "replicate_id") || "");
      const y = parseNumber(getVal(r, "yield_t_ha"));
      const c = parseNumber(getVal(r, "total_cost_per_ha"));
      const cc = parseNumber(getVal(r, "cost_amendment_input_per_ha"));

      return `<tr>
        <td>${escapeHtml(t)}</td>
        <td>${escapeHtml(isC)}</td>
        <td>${escapeHtml(plot)}</td>
        <td>${escapeHtml(rep)}</td>
        <td>${escapeHtml(Number.isNaN(y) ? "" : String(y))}</td>
        <td>${escapeHtml(Number.isNaN(c) ? "" : String(c))}</td>
        <td>${escapeHtml(Number.isNaN(cc) ? "" : String(cc))}</td>
      </tr>`;
    }).join("");

    table.innerHTML = `
      <thead>
        <tr>
          <th>Treatment</th>
          <th>Control</th>
          <th>Plot</th>
          <th>Replicate</th>
          <th>Yield (t/ha)</th>
          <th>Total cost ($/ha)</th>
          <th>One-off cost ($/ha)</th>
        </tr>
      </thead>
      <tbody>${rowsHtml || ""}</tbody>
    `;
  }

function renderAll() {
    renderOverview();
    renderDataSummaryAndChecks();
    renderTemplateColumns();
    renderControlChoice();
    computeCBA();
    renderLeaderboard();
    renderComparisonTable();
    renderCharts();
    renderAssumptionsUsed();
    renderIndicativeBadge();
    renderReplicationTable();
    buildAiBriefingPrompt();
  }

  document.addEventListener("DOMContentLoaded", () => {
    attachEventListeners();
    loadDefaultDataset();
  });
})();
