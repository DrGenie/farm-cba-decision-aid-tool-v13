/* global Chart, XLSX */

const PROJECT = {
  name: "Trial Cost-Benefit Decision Aid",
  partnerPlaceholder: "Project partners (letter logos)"
};

const state = {
  rawRows: [],
  rows: [],
  treatments: [],
  controlName: null,
  scenario: {
    pricePerTonne: 400,
    timeHorizon: 10,
    discountRate: 4
  },
  aggregates: [],
  cbaResults: [],
  basicAnalysisOnly: false,
  charts: {
    npvChart: null,
    bcrChart: null,
    paybackChart: null
  },
  uploadPreview: null,
  pendingFile: null,
  templateHeaders: null
};

const CORE_COLUMNS = [
  "trial_id",
  "treatment_name",
  "is_control",
  "yield_t_ha",
  "total_cost_per_ha"
];

const OPTIONAL_COLUMNS = [
  "variable_cost_per_ha",
  "fixed_cost_per_ha",
  "capital_cost_per_ha",
  "other_benefit_per_ha"
];

const UPLOAD_ERRORS = {
  MISSING_COLUMNS: "Missing required columns.",
  NO_CONTROL: "No control rows found in the dataset.",
  NO_TREATMENT: "No treatment rows found in the dataset.",
  NO_ROWS: "The file does not contain any data rows.",
  PARSE: "The file could not be parsed as a table.",
  INVALID_NUMERIC: "Some numeric fields are not in a usable format."
};

const TOAST = {
  element: null,
  timeout: null
};

function initToast() {
  TOAST.element = document.getElementById("toast");
}

function showToast(message, type = "info") {
  if (!TOAST.element) return;

  TOAST.element.textContent = message;
  TOAST.element.className = "toast show";

  if (type === "success") {
    TOAST.element.classList.add("toast-success");
  } else if (type === "warning") {
    TOAST.element.classList.add("toast-warn");
  } else if (type === "error") {
    TOAST.element.classList.add("toast-error");
  }

  if (TOAST.timeout) {
    clearTimeout(TOAST.timeout);
  }

  TOAST.timeout = setTimeout(() => {
    TOAST.element.className = "toast";
  }, 3500);
}

function activateTab(tabName) {
  const tabButtons = document.querySelectorAll(".tab");
  const tabPanels = document.querySelectorAll(".tab-panel");

  tabButtons.forEach((btn) => {
    const isActive = btn.getAttribute("data-tab") === tabName;
    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  tabPanels.forEach((panel) => {
    const id = panel.id.replace("Tab", "");
    const isActive = id === tabName;
    panel.hidden = !isActive;
  });
}

function setupTabs() {
  const tabButtons = document.querySelectorAll(".tab");
  tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const tabName = btn.getAttribute("data-tab");
      activateTab(tabName);
    });
  });
}

function describeScenario(scenario) {
  return `Price ${formatNumber(scenario.pricePerTonne, 0)} $/t; time horizon ${scenario.timeHorizon} years; discount rate ${formatNumber(
    scenario.discountRate,
    1
  )}%`;
}

function setAssumptionsSummary() {
  const el = document.getElementById("assumptionsSummary");
  if (!el) {
    return;
  }
  const s = state.scenario;
  const lines = [
    `Grain price: $${formatNumber(s.pricePerTonne, 0)} per tonne.`,
    `Time horizon: ${s.timeHorizon} years.`,
    `Discount rate: ${formatNumber(s.discountRate, 1)}% per year.`,
    state.controlName
      ? `Control treatment: ${state.controlName}.`
      : "Control treatment: not yet selected."
  ];
  el.innerHTML = `<p>${lines.join(" ")}</p>`;
}

function setScenarioStatus(message) {
  const el = document.getElementById("scenarioStatus");
  if (!el) return;
  el.textContent = message;
}

function setUploadStatus(message, isError = false) {
  const el = document.getElementById("uploadStatus");
  if (!el) return;
  el.textContent = message;
  el.classList.remove("error", "success");
  if (message) {
    el.classList.add(isError ? "error" : "success");
  }
}

function setUploadPreview(html) {
  const el = document.getElementById("uploadPreview");
  if (!el) return;
  el.innerHTML = html || "";
}

function formatNumber(value, decimals = 1) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  const factor = Math.pow(10, decimals);
  const rounded = Math.round(value * factor) / factor;
  return rounded.toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
}

function formatCurrency(value, decimals = 0) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  const factor = Math.pow(10, decimals);
  const rounded = Math.round(value * factor) / factor;
  return `$${rounded.toLocaleString(undefined, {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  })}`;
}

function parseNumber(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  const trimmed = String(value).trim();
  if (!trimmed) return null;
  const normalised = trimmed.replace(/,/g, "");
  const parsed = Number(normalised);
  return Number.isFinite(parsed) ? parsed : null;
}

function normaliseHeader(header) {
  if (!header) return "";
  return header
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_")
    .replace(/[^a-z0-9_]/g, "");
}

function normaliseRow(rawRow, headerMap) {
  const normalised = {};
  Object.keys(headerMap).forEach((key) => {
    const idx = headerMap[key];
    normalised[key] = rawRow[idx];
  });
  return normalised;
}

function parseDelimitedText(text, delimiter) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0);
  if (lines.length < 2) {
    return { header: [], rows: [] };
  }

  const headerRaw = lines[0].split(delimiter);
  const header = headerRaw.map((h) => h.trim());
  const rows = [];

  for (let i = 1; i < lines.length; i += 1) {
    const parts = lines[i].split(delimiter);
    if (parts.every((p) => p.trim() === "")) {
      continue;
    }
    const row = {};
    header.forEach((h, idx) => {
      row[h] = parts[idx] !== undefined ? parts[idx].trim() : "";
    });
    rows.push(row);
  }

  return { header, rows };
}

function parseSpreadsheetFile(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (!json.length) {
    return { header: [], rows: [] };
  }
  const headerRow = json[0];
  const header = headerRow.map((h) => String(h || "").trim());
  const rows = [];
  for (let i = 1; i < json.length; i += 1) {
    const rowArr = json[i];
    if (!rowArr || rowArr.every((c) => c === null || c === undefined || String(c).trim() === "")) {
      continue;
    }
    const row = {};
    header.forEach((h, idx) => {
      row[h] = rowArr[idx] !== undefined && rowArr[idx] !== null ? String(rowArr[idx]).trim() : "";
    });
    rows.push(row);
  }
  return { header, rows };
}

function buildHeaderMap(header) {
  const headerMap = {};
  header.forEach((h, idx) => {
    const key = normaliseHeader(h);
    if (key) {
      if (headerMap[key] === undefined) {
        headerMap[key] = idx;
      }
    }
  });
  return headerMap;
}

function detectDelimiterFromName(fileName) {
  if (!fileName) return "\t";
  const lower = fileName.toLowerCase();
  if (lower.endsWith(".csv")) return ",";
  return "\t";
}

function classifyRows(header, rows) {
  const headerMap = buildHeaderMap(header);
  const normalisedRows = [];
  const missingCore = [];

  CORE_COLUMNS.forEach((col) => {
    const found = Object.keys(headerMap).includes(col);
    if (!found) {
      missingCore.push(col);
    }
  });

  rows.forEach((row) => {
    const norm = {};
    Object.keys(headerMap).forEach((key) => {
      norm[key] = row[header[headerMap[key]]];
    });
    normalisedRows.push(norm);
  });

  return {
    headerMap,
    normalisedRows,
    missingCore
  };
}

function validateParsedData(header, rows) {
  const preview = {
    ok: false,
    canRun: false,
    errors: [],
    warnings: [],
    info: [],
    detectedTreatments: [],
    detectedControlName: null
  };

  if (!rows.length) {
    preview.errors.push(UPLOAD_ERRORS.NO_ROWS);
    return preview;
  }

  const headerMap = buildHeaderMap(header);
  const missing = [];

  CORE_COLUMNS.forEach((col) => {
    if (!Object.prototype.hasOwnProperty.call(headerMap, col)) {
      missing.push(col);
    }
  });

  if (missing.length) {
    preview.errors.push(
      `Missing required column(s): ${missing.join(
        ", "
      )}. Please download the template and ensure these columns are present.`
    );
  }

  if (Object.prototype.hasOwnProperty.call(headerMap, "is_control")) {
    const idx = headerMap.is_control;
    const controlRows = rows.filter((row) => {
      const raw = row[header[idx]];
      if (raw === null || raw === undefined) return false;
      const text = String(raw).trim().toLowerCase();
      return text === "1" || text === "true" || text === "yes" || text === "control";
    });

    const nonControlRows = rows.filter((row) => {
      const raw = row[header[idx]];
      const text = raw === null || raw === undefined ? "" : String(raw).trim().toLowerCase();
      return !(text === "1" || text === "true" || text === "yes" || text === "control");
    });

    if (!controlRows.length) {
      preview.errors.push(
        "No control identified. At least one row must be flagged as control in the is_control column."
      );
    }

    if (!nonControlRows.length) {
      preview.errors.push(
        "No treatment rows detected. At least one non-control row is required for comparison."
      );
    }

    const treatmentNames = new Set();
    const controlNames = new Set();
    const treatmentIndex = headerMap.treatment_name;

    if (treatmentIndex !== undefined) {
      controlRows.forEach((row) => {
        const raw = row[header[treatmentIndex]];
        const name = raw === null || raw === undefined ? "" : String(raw).trim();
        if (name) {
          controlNames.add(name);
        }
      });

      nonControlRows.forEach((row) => {
        const raw = row[header[treatmentIndex]];
        const name = raw === null || raw === undefined ? "" : String(raw).trim();
        if (name) {
          treatmentNames.add(name);
        }
      });
    }

    preview.detectedTreatments = Array.from(treatmentNames).sort();
    if (controlNames.size === 1) {
      preview.detectedControlName = Array.from(controlNames)[0];
    } else if (controlNames.size > 1) {
      preview.warnings.push(
        "Multiple control treatment names were found. Please check the treatment_name column for control rows."
      );
    }
  } else {
    preview.errors.push(
      "The is_control column is missing. Please download the template and ensure there is a column that identifies control rows."
    );
  }

  const numericProblems = [];
  const headerMap2 = buildHeaderMap(header);

  ["yield_t_ha", "total_cost_per_ha"].forEach((col) => {
    const idx = headerMap2[col];
    if (idx === undefined) return;
    let numericCount = 0;
    rows.forEach((row) => {
      const value = row[header[idx]];
      const parsed = parseNumber(value);
      if (parsed !== null) {
        numericCount += 1;
      }
    });
    if (!numericCount) {
      numericProblems.push(col);
    }
  });

  if (numericProblems.length) {
    preview.errors.push(
      `Core fields ${numericProblems.join(
        ", "
      )} do not contain usable numeric values. Please check formatting (no text in numeric cells).`
    );
  }

  if (!preview.errors.length) {
    preview.ok = true;
    preview.canRun = true;
  }

  return preview;
}

function renderDataPreview(preview) {
  const parts = [];

  if (preview.errors.length) {
    parts.push("<div class=\"upload-preview-errors\"><strong>Problems detected:</strong><ul>");
    preview.errors.forEach((err) => {
      parts.push(`<li>${err}</li>`);
    });
    parts.push("</ul></div>");
  }

  if (preview.warnings.length) {
    parts.push("<div class=\"upload-preview-warnings\"><strong>Warnings:</strong><ul>");
    preview.warnings.forEach((w) => {
      parts.push(`<li>${w}</li>`);
    });
    parts.push("</ul></div>");
  }

  if (preview.info.length) {
    parts.push("<div class=\"upload-preview-info\"><ul>");
    preview.info.forEach((i) => {
      parts.push(`<li>${i}</li>`);
    });
    parts.push("</ul></div>");
  }

  if (preview.detectedTreatments && preview.detectedTreatments.length) {
    parts.push(
      `<div class="upload-preview-info"><strong>Detected treatments:</strong> ${preview.detectedTreatments.join(
        ", "
      )}</div>`
    );
  }

  if (preview.detectedControlName) {
    parts.push(
      `<div class="upload-preview-info"><strong>Detected control:</strong> ${preview.detectedControlName}</div>`
    );
  }

  if (!preview.errors.length && !preview.warnings.length && !preview.info.length) {
    parts.push(
      "<div class=\"upload-preview-info\">No structural problems detected in the file. You can upload it for analysis.</div>"
    );
  }

  setUploadPreview(parts.join(""));
}

function buildTemplateColumnsPanel(headers) {
  const panel = document.getElementById("templateColumnsPanel");
  if (!panel) return;

  const makeRole = (name) => {
    if (CORE_COLUMNS.includes(name)) {
      if (name === "trial_id") return "Identifier for the trial, site, or block.";
      if (name === "treatment_name") return "Name of the treatment or management option.";
      if (name === "is_control") return "Flag for control rows (1, true, yes, or control).";
      if (name === "yield_t_ha") return "Main outcome per hectare (for example yield in tonnes per hectare).";
      if (name === "total_cost_per_ha") return "Total cost per hectare, including all variable and fixed costs.";
    }
    if (OPTIONAL_COLUMNS.includes(name)) {
      if (name === "variable_cost_per_ha") return "Variable cost per hectare (for example fertiliser, chemicals, seed).";
      if (name === "fixed_cost_per_ha") return "Fixed cost per hectare (for example machinery and overheads attributed to the treatment).";
      if (name === "capital_cost_per_ha") return "Capital cost per hectare spread over the time horizon.";
      if (name === "other_benefit_per_ha") return "Other benefit per hectare not captured in yield (for example quality premiums).";
    }
    return "Optional column that can be used for additional information.";
  };

  const allNames = [...CORE_COLUMNS, ...OPTIONAL_COLUMNS].filter((name, idx, arr) => arr.indexOf(name) === idx);
  panel.innerHTML = "";

  allNames.forEach((name) => {
    const div = document.createElement("div");
    div.className = "template-column";
    const labelSpan = document.createElement("span");
    labelSpan.className = "template-column-name tip";
    labelSpan.textContent = name;
    labelSpan.setAttribute(
      "data-tooltip",
      makeRole(name)
    );

    const roleDiv = document.createElement("div");
    roleDiv.className = "template-column-role";
    roleDiv.textContent = makeRole(name);

    div.appendChild(labelSpan);
    div.appendChild(roleDiv);
    panel.appendChild(div);
  });
}

function coerceRow(row, headerMap) {
  const out = {};
  Object.keys(headerMap).forEach((normKey) => {
    const idx = headerMap[normKey];
    out[normKey] = row[idx];
  });
  return out;
}

function normaliseRows(header, rows) {
  const headerMap = buildHeaderMap(header);
  const normRows = rows.map((row) => {
    const normalised = {};
    Object.keys(headerMap).forEach((key) => {
      normalised[key] = row[header[headerMap[key]]];
    });
    return normalised;
  });
  return { headerMap, normRows };
}

function interpretControlFlag(value) {
  if (value === null || value === undefined) return false;
  const text = String(value).trim().toLowerCase();
  return text === "1" || text === "true" || text === "yes" || text === "control";
}

function prepareRowsForAnalysis(header, rows) {
  const { headerMap, normRows } = normaliseRows(header, rows);
  const prepared = [];
  normRows.forEach((row) => {
    const treatmentName =
      headerMap.treatment_name !== undefined ? String(row[Object.keys(header)[headerMap.treatment_name]]).trim() : "";
    const isControl =
      headerMap.is_control !== undefined
        ? interpretControlFlag(row[Object.keys(header)[headerMap.is_control]])
        : false;

    const yieldValue =
      headerMap.yield_t_ha !== undefined
        ? parseNumber(row[Object.keys(header)[headerMap.yield_t_ha]])
        : null;

    const totalCost =
      headerMap.total_cost_per_ha !== undefined
        ? parseNumber(row[Object.keys(header)[headerMap.total_cost_per_ha]])
        : null;

    const variableCost =
      headerMap.variable_cost_per_ha !== undefined
        ? parseNumber(row[Object.keys(header)[headerMap.variable_cost_per_ha]])
        : null;

    const fixedCost =
      headerMap.fixed_cost_per_ha !== undefined
        ? parseNumber(row[Object.keys(header)[headerMap.fixed_cost_per_ha]])
        : null;

    const capitalCost =
      headerMap.capital_cost_per_ha !== undefined
        ? parseNumber(row[Object.keys(header)[headerMap.capital_cost_per_ha]])
        : null;

    const otherBenefit =
      headerMap.other_benefit_per_ha !== undefined
        ? parseNumber(row[Object.keys(header)[headerMap.other_benefit_per_ha]])
        : null;

    prepared.push({
      trial_id:
        headerMap.trial_id !== undefined
          ? String(row[Object.keys(header)[headerMap.trial_id]] || "").trim()
          : "",
      treatment_name: treatmentName,
      is_control: isControl,
      yield_t_ha: yieldValue,
      total_cost_per_ha: totalCost,
      variable_cost_per_ha: variableCost,
      fixed_cost_per_ha: fixedCost,
      capital_cost_per_ha: capitalCost,
      other_benefit_per_ha: otherBenefit
    });
  });

  const controlRows = prepared.filter((r) => r.is_control);
  const treatmentRows = prepared.filter((r) => !r.is_control);

  return {
    prepared,
    controlRows,
    treatmentRows
  };
}

function aggregateTreatments() {
  const rows = state.rawRows;
  if (!rows.length) {
    state.aggregates = [];
    return;
  }

  const byTreatment = new Map();
  rows.forEach((row) => {
    const name = row.treatment_name || "";
    if (!name) return;
    if (!byTreatment.has(name)) {
      byTreatment.set(name, []);
    }
    byTreatment.get(name).push(row);
  });

  const aggregates = [];
  byTreatment.forEach((group, name) => {
    const yields = group.map((r) => r.yield_t_ha).filter((v) => v !== null && !Number.isNaN(v));
    const totalCosts = group.map((r) => r.total_cost_per_ha).filter((v) => v !== null && !Number.isNaN(v));
    const variableCosts = group.map((r) => r.variable_cost_per_ha).filter((v) => v !== null && !Number.isNaN(v));
    const fixedCosts = group.map((r) => r.fixed_cost_per_ha).filter((v) => v !== null && !Number.isNaN(v));
    const capitalCosts = group.map((r) => r.capital_cost_per_ha).filter((v) => v !== null && !Number.isNaN(v));
    const otherBenefits = group.map((r) => r.other_benefit_per_ha).filter((v) => v !== null && !Number.isNaN(v));

    const avg = (arr) => {
      if (!arr.length) return null;
      const sum = arr.reduce((a, b) => a + b, 0);
      return sum / arr.length;
    };

    aggregates.push({
      treatment_name: name,
      is_control: group.some((r) => r.is_control),
      n_plots: group.length,
      avg_yield_t_ha: avg(yields),
      avg_total_cost_per_ha: avg(totalCosts),
      avg_variable_cost_per_ha: avg(variableCosts),
      avg_fixed_cost_per_ha: avg(fixedCosts),
      avg_capital_cost_per_ha: avg(capitalCosts),
      avg_other_benefit_per_ha: avg(otherBenefits)
    });
  });

  state.aggregates = aggregates;
}

function computeCBA() {
  const s = state.scenario;
  const price = s.pricePerTonne;
  const T = s.timeHorizon;
  const r = s.discountRate / 100;

  const agg = state.aggregates;
  if (!agg.length || !state.controlName) {
    state.cbaResults = [];
    state.basicAnalysisOnly = false;
    return;
  }

  const control = agg.find((a) => a.treatment_name === state.controlName && a.is_control);
  if (!control) {
    state.cbaResults = [];
    state.basicAnalysisOnly = false;
    return;
  }

  const includeCapital = agg.some((a) => a.avg_capital_cost_per_ha !== null);
  const includeOtherBenefit = agg.some((a) => a.avg_other_benefit_per_ha !== null);

  const rows = [];
  let basicOnly = false;

  agg.forEach((t) => {
    const yieldVal = t.avg_yield_t_ha;
    const totalCost = t.avg_total_cost_per_ha;
    if (yieldVal === null || totalCost === null) {
      basicOnly = true;
    }

    const revenuePerHa = yieldVal !== null ? yieldVal * price : null;
    const netAnnualPerHa =
      revenuePerHa !== null && totalCost !== null ? revenuePerHa - totalCost : null;

    const otherBenefit = t.avg_other_benefit_per_ha;
    const capitalCost = t.avg_capital_cost_per_ha;

    let npv = null;
    let bcr = null;
    let payback = null;

    if (!basicOnly && netAnnualPerHa !== null) {
      const annuityFactor =
        r === 0 ? T : (1 - Math.pow(1 + r, -T)) / r;
      const totalNet = netAnnualPerHa * annuityFactor;

      const capitalOutlay = capitalCost !== null ? capitalCost : 0;
      const capitalNPV = capitalOutlay;

      const otherNPV =
        otherBenefit !== null ? otherBenefit * annuityFactor : 0;

      npv = totalNet + otherNPV - capitalNPV;

      const totalCostPresent =
        capitalNPV +
        totalCost * annuityFactor -
        (otherBenefit !== null ? otherBenefit * annuityFactor : 0);

      if (totalCostPresent > 0) {
        bcr = (totalNet + otherNPV) / totalCostPresent;
      }

      if (netAnnualPerHa > 0) {
        payback = capitalOutlay / netAnnualPerHa;
      }
    }

    rows.push({
      treatment_name: t.treatment_name,
      is_control: t.is_control,
      n_plots: t.n_plots,
      avg_yield_t_ha: yieldVal,
      avg_total_cost_per_ha: totalCost,
      avg_variable_cost_per_ha: t.avg_variable_cost_per_ha,
      avg_fixed_cost_per_ha: t.avg_fixed_cost_per_ha,
      avg_capital_cost_per_ha: capitalCost,
      avg_other_benefit_per_ha: otherBenefit,
      revenue_per_ha: revenuePerHa,
      net_annual_per_ha: netAnnualPerHa,
      npv_per_ha: npv,
      bcr,
      payback_years: payback,
      includeCapital,
      includeOtherBenefit
    });
  });

  state.cbaResults = rows;
  state.basicAnalysisOnly = basicOnly;
}

function renderResultsSummary() {
  const container = document.getElementById("resultsSummary");
  if (!container) return;

  if (!state.cbaResults.length) {
    container.innerHTML =
      "<p class=\"small muted\">Results will appear here once data have been uploaded and a control has been selected.</p>";
    document.getElementById("basicAnalysisNote").style.display = "none";
    document.getElementById("indicativeBadge").style.display = "none";
    return;
  }

  const cards = [];
  const controlRow = state.cbaResults.find((r) => r.is_control);
  const treatments = state.cbaResults.filter((r) => !r.is_control);

  const bestByNPV = [...treatments].filter((t) => t.npv_per_ha !== null);
  bestByNPV.sort((a, b) => (b.npv_per_ha || 0) - (a.npv_per_ha || 0));
  const leading = bestByNPV[0] || null;

  const bodyParts = [];

  if (leading && controlRow && leading.npv_per_ha !== null && controlRow.npv_per_ha !== null) {
    const diff = leading.npv_per_ha - controlRow.npv_per_ha;
    bodyParts.push(
      `<div class="results-card">
        <div class="results-card-title">Top treatment by NPV</div>
        <div class="results-card-value">${leading.treatment_name}</div>
        <div class="results-card-sub">Improves NPV by ${formatCurrency(
          diff,
          0
        )} per hectare relative to the control.</div>
      </div>`
    );
  }

  const positiveCount = treatments.filter((t) => t.npv_per_ha && t.npv_per_ha > 0).length;
  const nonPositiveCount = treatments.length - positiveCount;
  bodyParts.push(
    `<div class="results-card">
      <div class="results-card-title">How many look attractive?</div>
      <div class="results-card-value">${positiveCount} of ${treatments.length}</div>
      <div class="results-card-sub">Treatments with NPV above zero per hectare (given current assumptions).</div>
    </div>`
  );

  const betterBCR = treatments.filter(
    (t) => t.bcr !== null && controlRow && controlRow.bcr !== null && t.bcr > controlRow.bcr
  ).length;
  bodyParts.push(
    `<div class="results-card">
      <div class="results-card-title">Better benefit-cost ratio</div>
      <div class="results-card-value">${betterBCR}</div>
      <div class="results-card-sub">Treatments with higher BCR than the control.</div>
    </div>`
  );

  if (nonPositiveCount > 0) {
    bodyParts.push(
      `<div class="results-card">
        <div class="results-card-title">Treatments that do not pay back</div>
        <div class="results-card-value">${nonPositiveCount}</div>
        <div class="results-card-sub">Treatments where net present value is at or below zero per hectare.</div>
      </div>`
    );
  }

  container.innerHTML = `<div class="results-summary-grid">${bodyParts.join("")}</div>`;

  const basicNote = document.getElementById("basicAnalysisNote");
  const indicativeBadge = document.getElementById("indicativeBadge");

  if (state.basicAnalysisOnly) {
    basicNote.style.display = "block";
    indicativeBadge.style.display = "inline-flex";
  } else {
    basicNote.style.display = "none";
    indicativeBadge.style.display = "none";
  }
}

function renderComparisonTable() {
  const table = document.getElementById("comparisonTable").querySelector("tbody");
  if (!table) return;

  table.innerHTML = "";

  if (!state.cbaResults.length || !state.controlName) {
    return;
  }

  const control = state.cbaResults.find((r) => r.is_control);
  if (!control) return;

  const treatments = state.cbaResults.filter((r) => !r.is_control);

  const rows = [
    {
      label: "Net present value (NPV) per hectare",
      controlVal: control.npv_per_ha,
      getTreatmentVal: (t) => t.npv_per_ha,
      formatter: (v) => formatCurrency(v, 0)
    },
    {
      label: "Benefit-cost ratio (BCR)",
      controlVal: control.bcr,
      getTreatmentVal: (t) => t.bcr,
      formatter: (v) =>
        v === null || v === undefined || Number.isNaN(v) ? "" : v.toFixed(2)
    },
    {
      label: "Payback period (years)",
      controlVal: control.payback_years,
      getTreatmentVal: (t) => t.payback_years,
      formatter: (v) =>
        v === null || v === undefined || Number.isNaN(v) ? "" : v.toFixed(1)
    },
    {
      label: "Net return per hectare (annual)",
      controlVal: control.net_annual_per_ha,
      getTreatmentVal: (t) => t.net_annual_per_ha,
      formatter: (v) => formatCurrency(v, 0)
    },
    {
      label: "Average yield (t/ha)",
      controlVal: control.avg_yield_t_ha,
      getTreatmentVal: (t) => t.avg_yield_t_ha,
      formatter: (v) => formatNumber(v, 2)
    },
    {
      label: "Average total cost ($/ha)",
      controlVal: control.avg_total_cost_per_ha,
      getTreatmentVal: (t) => t.avg_total_cost_per_ha,
      formatter: (v) => formatCurrency(v, 0)
    }
  ];

  treatments.forEach((treatment) => {
    rows.forEach((rowDef) => {
      const tr = document.createElement("tr");
      const labelCell = document.createElement("td");
      const controlCell = document.createElement("td");
      const treatmentCell = document.createElement("td");
      const diffCell = document.createElement("td");

      labelCell.textContent = rowDef.label;

      controlCell.textContent = rowDef.formatter(rowDef.controlVal);
      const treatmentVal = rowDef.getTreatmentVal(treatment);
      treatmentCell.textContent = rowDef.formatter(treatmentVal);

      let diff = null;
      if (rowDef.controlVal !== null && treatmentVal !== null) {
        diff = treatmentVal - rowDef.controlVal;
      }

      if (rowDef.label.includes("BCR")) {
        diffCell.textContent =
          diff === null || Number.isNaN(diff)
            ? ""
            : diff >= 0
            ? `+${diff.toFixed(2)}`
            : diff.toFixed(2);
      } else if (rowDef.label.includes("Payback")) {
        diffCell.textContent =
          diff === null || Number.isNaN(diff)
            ? ""
            : diff >= 0
            ? `+${diff.toFixed(1)}`
            : diff.toFixed(1);
      } else if (rowDef.label.includes("yield")) {
        diffCell.textContent =
          diff === null || Number.isNaN(diff)
            ? ""
            : diff >= 0
            ? `+${formatNumber(diff, 2)}`
            : formatNumber(diff, 2);
      } else {
        diffCell.textContent =
          diff === null || Number.isNaN(diff)
            ? ""
            : diff >= 0
            ? `+${formatCurrency(diff, 0)}`
            : formatCurrency(diff, 0);
      }

      tr.appendChild(labelCell);
      tr.appendChild(controlCell);
      tr.appendChild(treatmentCell);
      tr.appendChild(diffCell);

      table.appendChild(tr);
    });
  });
}

function renderReplicateSummary() {
  const tbody = document.getElementById("replicateSummaryTable").querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!state.aggregates.length) return;

  state.aggregates.forEach((a) => {
    const tr = document.createElement("tr");
    const tCell = document.createElement("td");
    tCell.textContent = a.treatment_name;
    const yCell = document.createElement("td");
    yCell.textContent = formatNumber(a.avg_yield_t_ha, 2);
    const cCell = document.createElement("td");
    cCell.textContent = formatCurrency(a.avg_total_cost_per_ha, 0);
    const nCell = document.createElement("td");
    nCell.textContent = a.n_plots;

    tr.appendChild(tCell);
    tr.appendChild(yCell);
    tr.appendChild(cCell);
    tr.appendChild(nCell);

    tbody.appendChild(tr);
  });
}

function renderLeaderboard() {
  const tbody = document.getElementById("leaderboardTable").querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!state.cbaResults.length) {
    return;
  }

  const metricSelect = document.getElementById("leaderboardMetric");
  const metric = metricSelect ? metricSelect.value : "npv";

  const records = state.cbaResults.filter((r) => !r.is_control);
  if (!records.length) return;

  let comparator;

  if (metric === "bcr") {
    comparator = (a, b) => (b.bcr || 0) - (a.bcr || 0);
  } else if (metric === "payback") {
    comparator = (a, b) => {
      const pa = a.payback_years || Number.POSITIVE_INFINITY;
      const pb = b.payback_years || Number.POSITIVE_INFINITY;
      return pa - pb;
    };
  } else if (metric === "netReturnPerHa") {
    comparator = (a, b) => (b.net_annual_per_ha || 0) - (a.net_annual_per_ha || 0);
  } else {
    comparator = (a, b) => (b.npv_per_ha || 0) - (a.npv_per_ha || 0);
  }

  records.sort(comparator);

  records.forEach((r) => {
    const tr = document.createElement("tr");
    const nameCell = document.createElement("td");
    nameCell.textContent = r.treatment_name;

    const npvCell = document.createElement("td");
    npvCell.textContent = formatCurrency(r.npv_per_ha, 0);

    const bcrCell = document.createElement("td");
    bcrCell.textContent =
      r.bcr === null || r.bcr === undefined || Number.isNaN(r.bcr)
        ? ""
        : r.bcr.toFixed(2);

    const paybackCell = document.createElement("td");
    paybackCell.textContent =
      r.payback_years === null || r.payback_years === undefined || Number.isNaN(r.payback_years)
        ? ""
        : r.payback_years.toFixed(1);

    const netCell = document.createElement("td");
    netCell.textContent = formatCurrency(r.net_annual_per_ha, 0);

    tr.appendChild(nameCell);
    tr.appendChild(npvCell);
    tr.appendChild(bcrCell);
    tr.appendChild(paybackCell);
    tr.appendChild(netCell);

    tbody.appendChild(tr);
  });
}

function buildChartOptions(title) {
  return {
    responsive: true,
    scales: {
      x: {
        ticks: {
          autoSkip: false,
          maxRotation: 0,
          minRotation: 0
        }
      },
      y: {
        beginAtZero: true
      }
    },
    plugins: {
      title: {
        display: false
      },
      legend: {
        display: false
      },
      tooltip: {
        callbacks: {
          label(context) {
            if (context.parsed.y === null || context.parsed.y === undefined) {
              return "";
            }
            return context.parsed.y.toLocaleString();
          }
        }
      }
    }
  };
}

function ensureCharts() {
  const labels = state.cbaResults
    .filter((r) => !r.is_control)
    .map((r) => r.treatment_name);

  const npvData = state.cbaResults
    .filter((r) => !r.is_control)
    .map((r) => (r.npv_per_ha !== null ? Math.round(r.npv_per_ha) : null));

  const bcrData = state.cbaResults
    .filter((r) => !r.is_control)
    .map((r) => (r.bcr !== null ? Number(r.bcr.toFixed(2)) : null));

  const paybackData = state.cbaResults
    .filter((r) => !r.is_control)
    .map((r) =>
      r.payback_years !== null && Number.isFinite(r.payback_years)
        ? Number(r.payback_years.toFixed(1))
        : null
    );

  function buildDataset(data) {
    return {
      label: "Treatment",
      data,
      borderWidth: 1
    };
  }

  const npvCtx = document.getElementById("npvChart");
  const bcrCtx = document.getElementById("bcrChart");
  const paybackCtx = document.getElementById("paybackChart");

  if (state.charts.npvChart) {
    state.charts.npvChart.destroy();
  }
  if (state.charts.bcrChart) {
    state.charts.bcrChart.destroy();
  }
  if (state.charts.paybackChart) {
    state.charts.paybackChart.destroy();
  }

  if (npvCtx && labels.length) {
    state.charts.npvChart = new Chart(npvCtx, {
      type: "bar",
      data: {
        labels,
        datasets: [buildDataset(npvData)]
      },
      options: buildChartOptions("Net present value")
    });
  }

  if (bcrCtx && labels.length) {
    state.charts.bcrChart = new Chart(bcrCtx, {
      type: "bar",
      data: {
        labels,
        datasets: [buildDataset(bcrData)]
      },
      options: buildChartOptions("Benefit-cost ratio")
    });
  }

  if (paybackCtx && labels.length) {
    state.charts.paybackChart = new Chart(paybackCtx, {
      type: "bar",
      data: {
        labels,
        datasets: [buildDataset(paybackData)]
      },
      options: buildChartOptions("Payback period")
    });
  }
}

function renderAll() {
  setAssumptionsSummary();
  renderResultsSummary();
  renderComparisonTable();
  renderReplicateSummary();
  renderLeaderboard();
  ensureCharts();
  updateAIPrompt();
}

function updateControlTreatmentOptions() {
  const select = document.getElementById("controlTreatment");
  if (!select) return;
  const current = state.controlName;

  const names = Array.from(
    new Set(state.aggregates.filter((a) => a.is_control).map((a) => a.treatment_name))
  );

  select.innerHTML = "";

  if (!names.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "No control detected yet";
    select.appendChild(opt);
    return;
  }

  names.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    select.appendChild(opt);
  });

  if (current && names.includes(current)) {
    select.value = current;
  } else {
    state.controlName = names[0];
    select.value = names[0];
  }
}

function onLeaderboardMetricChange() {
  renderLeaderboard();
}

function onApplyScenario() {
  const priceEl = document.getElementById("pricePerTonne");
  const horizonEl = document.getElementById("timeHorizon");
  const discountEl = document.getElementById("discountRate");
  const controlEl = document.getElementById("controlTreatment");

  const price = parseNumber(priceEl.value);
  const horizon = parseInt(horizonEl.value, 10);
  const discount = parseNumber(discountEl.value);

  if (price === null || horizon <= 0 || discount === null) {
    setScenarioStatus(
      "Please provide a valid grain price, time horizon, and discount rate before applying."
    );
    return;
  }

  state.scenario.pricePerTonne = price;
  state.scenario.timeHorizon = horizon;
  state.scenario.discountRate = discount;

  if (controlEl && controlEl.value) {
    state.controlName = controlEl.value;
  }

  aggregateTreatments();
  computeCBA();
  renderAll();
  setScenarioStatus("Scenario updated.");
  showToast("Scenario settings applied.", "success");
}

function onControlTreatmentChange() {
  const controlEl = document.getElementById("controlTreatment");
  if (!controlEl || !controlEl.value) return;
  state.controlName = controlEl.value;
  aggregateTreatments();
  computeCBA();
  renderAll();
  setScenarioStatus(`Control changed to ${state.controlName}.`);
}

function describeDataQuality(rows) {
  const total = rows.length;
  const numericYield = rows.filter((r) => r.yield_t_ha !== null).length;
  const numericCost = rows.filter((r) => r.total_cost_per_ha !== null).length;

  const parts = [];
  parts.push(`${total} data rows loaded.`);

  if (numericYield === total && numericCost === total) {
    parts.push("All rows have numeric yield and total cost values.");
  } else {
    parts.push(
      `${numericYield} rows have usable yield values and ${numericCost} rows have usable total cost values.`
    );
    parts.push(
      "The analysis uses all rows with valid numbers and ignores missing or non-numeric entries in optional columns."
    );
  }

  return parts.join(" ");
}

function commitParsedData(header, rows, preview) {
  const { headerMap, normRows } = normaliseRows(header, rows);
  const prepared = [];
  normRows.forEach((row) => {
    const treatmentName =
      headerMap.treatment_name !== undefined
        ? String(row[header[headerMap.treatment_name]] || "").trim()
        : "";
    const isControl =
      headerMap.is_control !== undefined
        ? interpretControlFlag(row[header[headerMap.is_control]])
        : false;

    const yieldValue =
      headerMap.yield_t_ha !== undefined
        ? parseNumber(row[header[headerMap.yield_t_ha]])
        : null;

    const totalCost =
      headerMap.total_cost_per_ha !== undefined
        ? parseNumber(row[header[headerMap.total_cost_per_ha]])
        : null;

    const variableCost =
      headerMap.variable_cost_per_ha !== undefined
        ? parseNumber(row[header[headerMap.variable_cost_per_ha]])
        : null;

    const fixedCost =
      headerMap.fixed_cost_per_ha !== undefined
        ? parseNumber(row[header[headerMap.fixed_cost_per_ha]])
        : null;

    const capitalCost =
      headerMap.capital_cost_per_ha !== undefined
        ? parseNumber(row[header[headerMap.capital_cost_per_ha]])
        : null;

    const otherBenefit =
      headerMap.other_benefit_per_ha !== undefined
        ? parseNumber(row[header[headerMap.other_benefit_per_ha]])
        : null;

    prepared.push({
      trial_id:
        headerMap.trial_id !== undefined
          ? String(row[header[headerMap.trial_id]] || "").trim()
          : "",
      treatment_name: treatmentName,
      is_control: isControl,
      yield_t_ha: yieldValue,
      total_cost_per_ha: totalCost,
      variable_cost_per_ha: variableCost,
      fixed_cost_per_ha: fixedCost,
      capital_cost_per_ha: capitalCost,
      other_benefit_per_ha: otherBenefit
    });
  });

  state.rawRows = prepared;
  state.rows = prepared;

  const treatments = Array.from(
    new Set(prepared.filter((r) => r.treatment_name).map((r) => r.treatment_name))
  ).sort();
  state.treatments = treatments;

  const controls = Array.from(
    new Set(prepared.filter((r) => r.is_control).map((r) => r.treatment_name))
  );
  if (controls.length === 1) {
    state.controlName = controls[0];
  } else if (controls.length > 1) {
    state.controlName = controls[0];
  } else {
    state.controlName = null;
  }

  aggregateTreatments();
  computeCBA();
  updateControlTreatmentOptions();
  renderAll();

  const qualityText = describeDataQuality(prepared);
  const previewText = [];
  previewText.push("<div class=\"upload-preview-info\">");
  previewText.push(
    preview.detectedControlName
      ? `<strong>Control detected:</strong> ${preview.detectedControlName}.`
      : "Control detected from is_control column."
  );
  if (preview.detectedTreatments && preview.detectedTreatments.length) {
    previewText.push(
      `<strong>Treatments detected:</strong> ${preview.detectedTreatments.join(", ")}.`
    );
  }
  previewText.push(`<span>${qualityText}</span>`);
  previewText.push("</div>");

  setUploadPreview(previewText.join(""));
  setUploadStatus("Data uploaded successfully. Next: review results.", false);
  showToast("Data uploaded successfully. Next: review results.", "success");
}

function onFileSelected(file) {
  if (!file) {
    state.pendingFile = null;
    setUploadStatus("No file selected.", true);
    setUploadPreview("");
    return;
  }

  state.pendingFile = file;
  setUploadStatus(`Selected file: ${file.name}. Click “Check file” to validate.`, false);
  setUploadPreview("");
}

function handleParsedFile(header, rows) {
  const preview = validateParsedData(header, rows);
  state.uploadPreview = preview;
  renderDataPreview(preview);

  if (!preview.errors.length && preview.canRun) {
    setUploadStatus(
      "File structure looks valid. You can now upload and use this data in the analysis.",
      false
    );
  } else if (preview.errors.length) {
    setUploadStatus(
      "File has issues that must be fixed before the analysis can run. See details below.",
      true
    );
  } else {
    setUploadStatus(
      "The file was read but some issues were detected. You may still be able to use it for a basic analysis.",
      false
    );
  }
}

function onValidateSelectedFileClick() {
  const file =
    state.pendingFile ||
    (document.getElementById("fileInput") &&
      document.getElementById("fileInput").files &&
      document.getElementById("fileInput").files[0]);

  if (!file) {
    setUploadStatus("Please choose a TSV, CSV, or Excel file before checking.", true);
    showToast("No file selected.", "warning");
    return;
  }

  const reader = new FileReader();
  const extension = file.name.toLowerCase();

  reader.onload = (event) => {
    try {
      let header;
      let rows;
      if (extension.endsWith(".xlsx")) {
        const result = parseSpreadsheetFile(event.target.result);
        header = result.header;
        rows = result.rows;
      } else {
        const text = event.target.result;
        const delimiter = detectDelimiterFromName(file.name);
        const parsed = parseDelimitedText(text, delimiter);
        header = parsed.header;
        rows = parsed.rows;
      }
      handleParsedFile(header, rows);
      showToast("File checked. Review the messages below.", "success");
    } catch (err) {
      console.error(err);
      setUploadStatus(
        "The file could not be read. Please ensure it is a valid TSV, CSV, or Excel file.",
        true
      );
      showToast("File could not be read.", "error");
    }
  };

  if (extension.endsWith(".xlsx")) {
    reader.readAsArrayBuffer(file);
  } else {
    reader.readAsText(file);
  }
}

function onUploadSelectedFileClick() {
  const file =
    state.pendingFile ||
    (document.getElementById("fileInput") &&
      document.getElementById("fileInput").files &&
      document.getElementById("fileInput").files[0]);

  if (!file) {
    setUploadStatus("Please choose a file and run “Check file” before uploading.", true);
    showToast("No file selected.", "warning");
    return;
  }

  if (!state.uploadPreview) {
    setUploadStatus(
      "Please use “Check file” first so the tool can validate the structure before uploading.",
      true
    );
    showToast("Please check the file before uploading.", "warning");
    return;
  }

  if (state.uploadPreview.errors && state.uploadPreview.errors.length) {
    setUploadStatus(
      "The file still has issues that must be fixed before the analysis can run. See the list below.",
      true
    );
    showToast("Upload blocked due to structural issues.", "error");
    return;
  }

  const reader = new FileReader();
  const extension = file.name.toLowerCase();

  reader.onload = (event) => {
    try {
      let header;
      let rows;
      if (extension.endsWith(".xlsx")) {
        const result = parseSpreadsheetFile(event.target.result);
        header = result.header;
        rows = result.rows;
      } else {
        const text = event.target.result;
        const delimiter = detectDelimiterFromName(file.name);
        const parsed = parseDelimitedText(text, delimiter);
        header = parsed.header;
        rows = parsed.rows;
      }
      commitParsedData(header, rows, state.uploadPreview);
    } catch (err) {
      console.error(err);
      setUploadStatus(
        "An error occurred while processing the file for analysis. Please try again or check the format.",
        true
      );
      showToast("An error occurred while uploading the data.", "error");
    }
  };

  if (extension.endsWith(".xlsx")) {
    reader.readAsArrayBuffer(file);
  } else {
    reader.readAsText(file);
  }
}

function onResetToolClick() {
  loadDefaultDataset()
    .then(() => {
      showToast("Tool reset to built-in example dataset.", "success");
      setUploadStatus("Reset to built-in example dataset.", false);
      setUploadPreview("");
      state.pendingFile = null;
      state.uploadPreview = null;
      const fileInput = document.getElementById("fileInput");
      if (fileInput) {
        fileInput.value = "";
      }
    })
    .catch((err) => {
      console.error(err);
      showToast("Unable to reset to the built-in dataset.", "error");
    });
}

async function loadDefaultDataset() {
  const response = await fetch("faba_beans_trial_clean_named.tsv");
  const text = await response.text();
  const parsed = parseDelimitedText(text, "\t");
  const header = parsed.header;
  const rows = parsed.rows;

  const preview = validateParsedData(header, rows);
  state.uploadPreview = preview;
  renderDataPreview(preview);

  commitParsedData(header, rows, preview);

  const priceEl = document.getElementById("pricePerTonne");
  const horizonEl = document.getElementById("timeHorizon");
  const discountEl = document.getElementById("discountRate");

  if (priceEl) priceEl.value = state.scenario.pricePerTonne;
  if (horizonEl) horizonEl.value = state.scenario.timeHorizon;
  if (discountEl) discountEl.value = state.scenario.discountRate;

  buildTemplateColumnsPanel(header.map((h) => normaliseHeader(h)));
  await ensureTemplateHeaders();
}

async function ensureTemplateHeaders() {
  if (state.templateHeaders && Array.isArray(state.templateHeaders)) {
    return state.templateHeaders;
  }
  try {
    const response = await fetch("faba_beans_trial_clean_named.tsv");
    const text = await response.text();
    const { header } = parseDelimitedText(text, "\t");
    const headerMap = buildHeaderMap(header);

    const desired = [];
    CORE_COLUMNS.forEach((col) => {
      if (Object.prototype.hasOwnProperty.call(headerMap, col)) {
        desired.push(col);
      }
    });
    OPTIONAL_COLUMNS.forEach((col) => {
      if (Object.prototype.hasOwnProperty.call(headerMap, col)) {
        desired.push(col);
      }
    });

    const extras = [];
    header.forEach((h) => {
      const norm = normaliseHeader(h);
      if (norm && !desired.includes(norm)) {
        extras.push(norm);
      }
    });

    state.templateHeaders = [...desired, ...extras];
  } catch (err) {
    console.error(err);
    state.templateHeaders = [...CORE_COLUMNS, ...OPTIONAL_COLUMNS];
  }
  return state.templateHeaders;
}

function buildTemplateRows(headers) {
  const coreCols = headers.filter((h) => CORE_COLUMNS.includes(h));
  const otherCols = headers.filter((h) => !CORE_COLUMNS.includes(h));

  const ordered = [...coreCols, ...otherCols];

  const rows = [];
  for (let i = 0; i < 5; i += 1) {
    const row = {};
    ordered.forEach((h) => {
      row[h] = "";
    });
    rows.push(row);
  }

  if (rows.length) {
    rows[0].is_control = "1";
    rows[0].treatment_name = "Control";
  }

  return { header: ordered, rows };
}

async function downloadTemplateTSV() {
  const headers = await ensureTemplateHeaders();
  const { header, rows } = buildTemplateRows(headers);

  const lines = [];
  lines.push(header.join("\t"));
  rows.forEach((row) => {
    const line = header.map((h) => (row[h] !== undefined ? row[h] : "")).join("\t");
    lines.push(line);
  });

  const blob = new Blob([lines.join("\n")], { type: "text/tab-separated-values" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "trial_cost_benefit_template.tsv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function downloadTemplateXLSX() {
  const headers = await ensureTemplateHeaders();
  const { header, rows } = buildTemplateRows(headers);

  const aoa = [];
  aoa.push(header);
  rows.forEach((row) => {
    const line = header.map((h) => (row[h] !== undefined ? row[h] : ""));
    aoa.push(line);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, "trial_cost_benefit_template.xlsx");
}

function buildWordParagraph(text) {
  return `<p>${text}</p>`;
}

function buildWordHeading(text, level) {
  const tag = level === 1 ? "h1" : level === 2 ? "h2" : "h3";
  return `<${tag}>${text}</${tag}>`;
}

function buildWordTable(headers, rows) {
  const thead = `<tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr>`;
  const tbody = rows
    .map((row) => `<tr>${row.map((cell) => `<td>${cell}</td>`).join("")}</tr>`)
    .join("");
  return `<table>${thead}${tbody}</table>`;
}

function buildResultsSummaryForWord() {
  if (!state.cbaResults.length || !state.controlName) {
    return buildWordParagraph(
      "Results are not available because no dataset is loaded or no control treatment has been selected."
    );
  }

  const control = state.cbaResults.find((r) => r.is_control);
  const treatments = state.cbaResults.filter((r) => !r.is_control);
  const rows = [];

  treatments.forEach((t) => {
    const npvDiff =
      t.npv_per_ha !== null && control.npv_per_ha !== null
        ? t.npv_per_ha - control.npv_per_ha
        : null;
    const netDiff =
      t.net_annual_per_ha !== null && control.net_annual_per_ha !== null
        ? t.net_annual_per_ha - control.net_annual_per_ha
        : null;
    rows.push([
      t.treatment_name,
      formatCurrency(t.npv_per_ha, 0),
      formatCurrency(control.npv_per_ha, 0),
      npvDiff !== null ? formatCurrency(npvDiff, 0) : "",
      t.bcr !== null && control.bcr !== null
        ? (t.bcr - control.bcr >= 0 ? `+${(t.bcr - control.bcr).toFixed(2)}` : (t.bcr - control.bcr).toFixed(2))
        : "",
      formatCurrency(t.net_annual_per_ha, 0),
      netDiff !== null ? formatCurrency(netDiff, 0) : ""
    ]);
  });

  const headers = [
    "Treatment",
    "NPV ($/ha)",
    "Control NPV ($/ha)",
    "Difference in NPV",
    "Difference in BCR",
    "Net return ($/ha/year)",
    "Difference in net return"
  ];

  return buildWordTable(headers, rows);
}

function buildReplicateSummaryForWord() {
  if (!state.aggregates.length) {
    return buildWordParagraph("No replicate-level summary is available yet.");
  }

  const rows = state.aggregates
    .map((a) => [
      a.treatment_name,
      a.n_plots,
      formatNumber(a.avg_yield_t_ha, 2),
      formatCurrency(a.avg_total_cost_per_ha, 0)
    ])
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  const headers = ["Treatment", "Number of plots", "Average yield (t/ha)", "Average total cost ($/ha)"];

  return buildWordTable(headers, rows);
}

function buildTechnicalSummaryForWord() {
  const s = state.scenario;
  const lines = [];

  lines.push(
    `The analysis compares a control treatment with alternative treatments using a trial dataset with one row per plot or replicate.`
  );
  lines.push(
    `The current scenario assumes a grain price of $${formatNumber(
      s.pricePerTonne,
      0
    )} per tonne, a time horizon of ${s.timeHorizon} years, and a discount rate of ${formatNumber(
      s.discountRate,
      1
    )}% per year.`
  );
  lines.push(
    `For each treatment, the tool calculates an average yield per hectare and an average total cost per hectare across all plots. These averages are used to compute net annual returns, net present value (NPV), benefit-cost ratio (BCR), and payback period where the data allow.`
  );
  if (state.basicAnalysisOnly) {
    lines.push(
      "In this scenario, some information required for full discounted analysis is missing. The tool therefore reports treatment means and simple differences relative to the control as indicative results."
    );
  }

  return lines.map((l) => buildWordParagraph(l)).join("");
}

function buildExportHeader() {
  const s = state.scenario;
  const controlName = state.controlName || "not yet selected";

  const parts = [];
  parts.push(buildWordHeading(PROJECT.name, 1));
  parts.push(
    buildWordParagraph(
      `Project partners: ${PROJECT.partnerPlaceholder}.`
    )
  );
  parts.push(
    buildWordParagraph(
      `Scenario: price $${formatNumber(s.pricePerTonne, 0)} per tonne; time horizon ${s.timeHorizon} years; discount rate ${formatNumber(
        s.discountRate,
        1
      )}% per year; control treatment: ${controlName}.`
    )
  );

  return parts.join("");
}

function buildExportFooterNote() {
  return buildWordParagraph(
    "These figures are generated automatically from the uploaded trial dataset using the Trial Cost-Benefit Decision Aid. If the dataset or assumptions change, the results in this document will also change when the export is refreshed."
  );
}

function buildWordDocumentHtml(bodyContent) {
  const html = `
<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${PROJECT.name} - Export</title>
<style>
body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; color: #111827; }
h1 { font-size: 18pt; margin-bottom: 4pt; }
h2 { font-size: 14pt; margin-top: 14pt; margin-bottom: 4pt; }
h3 { font-size: 12pt; margin-top: 12pt; margin-bottom: 4pt; }
p { margin: 4pt 0; }
table { border-collapse: collapse; width: 100%; margin: 6pt 0; }
th, td { border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 10pt; }
th { background: #e5eff9; text-align: left; }
.small { font-size: 9pt; color: #6b7280; }
.footer { margin-top: 12pt; font-size: 9pt; color: #6b7280; border-top: 1px solid #e5e7eb; padding-top: 6pt; }
</style>
</head>
<body>
${bodyContent}
</body>
</html>
`;
  return html;
}

function triggerWordDownload(filename, htmlContent) {
  const blob = new Blob([htmlContent], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function onExportWordReportClick() {
  if (!state.cbaResults.length || !state.controlName) {
    showToast("No results to export. Please upload data and run the analysis first.", "warning");
    return;
  }

  const header = buildExportHeader();
  const mainSummaryHeading = buildWordHeading("Summary of results", 2);
  const mainSummaryTable = buildResultsSummaryForWord();
  const replicateHeading = buildWordHeading("Replicate-level summary", 2);
  const replicateTable = buildReplicateSummaryForWord();
  const technicalHeading = buildWordHeading("How these numbers were calculated", 2);
  const technicalBody = buildTechnicalSummaryForWord();
  const footerNote = `<div class="footer">${buildExportFooterNote()}</div>`;

  const body = [
    header,
    mainSummaryHeading,
    mainSummaryTable,
    replicateHeading,
    replicateTable,
    technicalHeading,
    technicalBody,
    footerNote
  ].join("");

  const html = buildWordDocumentHtml(body);
  triggerWordDownload("trial_cost_benefit_results.doc", html);
  const statusEl = document.getElementById("exportStatus");
  if (statusEl) {
    statusEl.textContent = "Results summary exported as Word document.";
  }
  showToast("Results summary exported.", "success");
}

function onExportWordTechnicalClick() {
  const header = buildExportHeader();
  const technicalHeading = buildWordHeading("Technical appendix", 2);
  const technicalBody = buildTechnicalSummaryForWord();
  const footerNote = `<div class="footer">${buildExportFooterNote()}</div>`;

  const body = [header, technicalHeading, technicalBody, footerNote].join("");

  const html = buildWordDocumentHtml(body);
  triggerWordDownload("trial_cost_benefit_technical_appendix.doc", html);
  const statusEl = document.getElementById("exportStatus");
  if (statusEl) {
    statusEl.textContent = "Technical appendix exported as Word document.";
  }
  showToast("Technical appendix exported.", "success");
}

function updateAIPrompt() {
  const textarea = document.getElementById("aiPrompt");
  if (!textarea) return;

  if (!state.cbaResults.length || !state.controlName) {
    textarea.value =
      "No scenario is currently loaded. Upload a dataset, set the scenario parameters, and then refresh this prompt.";
    return;
  }

  const s = state.scenario;
  const control = state.cbaResults.find((r) => r.is_control);
  const treatments = state.cbaResults.filter((r) => !r.is_control);

  const lines = [];

  lines.push(
    "You are an agricultural economist preparing a short, plain-language summary of a replicated trial comparing a control treatment with alternative treatments."
  );
  lines.push("");
  lines.push(`Project: ${PROJECT.name}.`);
  lines.push(
    `Scenario assumptions: grain price $${formatNumber(
      s.pricePerTonne,
      0
    )} per tonne; time horizon ${s.timeHorizon} years; discount rate ${formatNumber(
      s.discountRate,
      1
    )}% per year.`
  );
  lines.push(`Control treatment: ${state.controlName}.`);
  lines.push("");

  lines.push("For each treatment, you are given net present value (NPV), benefit-cost ratio (BCR), payback period, and net annual return per hectare, all calculated relative to the control.");
  if (state.basicAnalysisOnly) {
    lines.push(
      "In this scenario, some information required for full discounted analysis is missing. Treat the figures as indicative rather than definitive."
    );
  }
  lines.push("");

  lines.push("Here are the results for each treatment (including the control):");
  lines.push("");

  const header = [
    "treatment_name",
    "is_control",
    "npv_per_ha",
    "bcr",
    "payback_years",
    "net_annual_per_ha",
    "avg_yield_t_ha",
    "avg_total_cost_per_ha",
    "n_plots"
  ];
  lines.push(header.join("\t"));
  state.cbaResults.forEach((r) => {
    lines.push(
      [
        r.treatment_name,
        r.is_control ? "control" : "treatment",
        r.npv_per_ha !== null ? Math.round(r.npv_per_ha) : "",
        r.bcr !== null ? r.bcr.toFixed(2) : "",
        r.payback_years !== null && Number.isFinite(r.payback_years)
          ? r.payback_years.toFixed(1)
          : "",
        r.net_annual_per_ha !== null ? Math.round(r.net_annual_per_ha) : "",
        r.avg_yield_t_ha !== null ? r.avg_yield_t_ha.toFixed(2) : "",
        r.avg_total_cost_per_ha !== null ? Math.round(r.avg_total_cost_per_ha) : "",
        r.n_plots
      ].join("\t")
    );
  });

  lines.push("");
  lines.push(
    "Please write a concise, non-technical summary (2–4 short paragraphs) that explains which treatments look most economically attractive, how large the gains or losses are per hectare, and how sensitive these conclusions might be to the price and discount rate assumptions. Avoid equations and avoid giving specific methodological details unless needed for clarity."
  );

  textarea.value = lines.join("\n");
}

function copyAIPromptToClipboard() {
  const textarea = document.getElementById("aiPrompt");
  if (!textarea) return;
  textarea.select();
  textarea.setSelectionRange(0, textarea.value.length);
  const ok = document.execCommand("copy");
  if (ok) {
    showToast("AI prompt copied to clipboard.", "success");
  } else {
    showToast("Unable to copy the prompt. Please copy it manually.", "warning");
  }
}

function openAiAssistant(kind) {
  const textarea = document.getElementById("aiPrompt");
  if (!textarea) return;

  textarea.select();
  textarea.setSelectionRange(0, textarea.value.length);
  document.execCommand("copy");

  let url = "";
  if (kind === "chatgpt") {
    url = "https://chat.openai.com/";
  } else if (kind === "copilot") {
    url = "https://copilot.microsoft.com/";
  }

  if (url) {
    window.open(url, "_blank", "noopener");
    showToast("Prompt copied. Paste it into the new chat window.", "success");
  } else {
    showToast("Unsupported assistant. Copy the prompt manually.", "warning");
  }
}

function onRunAnalysisClick() {
  if (!state.rows || !state.rows.length) {
    showToast("Load or upload data first, then run the analysis.", "warning");
    return;
  }
  try {
    aggregateTreatments();
    computeCBA();
    renderAll();
    showToast("Analysis run successfully. Results updated.", "success");
  } catch (err) {
    console.error(err);
    showToast("Something went wrong while running the analysis.", "error");
  }
}

function attachEventListeners() {
  const tabButtons = document.querySelectorAll(".tab");
  tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const tabName = btn.getAttribute("data-tab");
      activateTab(tabName);
    });
  });

  const metricSelect = document.getElementById("leaderboardMetric");
  if (metricSelect) {
    metricSelect.addEventListener("change", onLeaderboardMetricChange);
  }

  const applyBtn = document.getElementById("applyScenario");
  if (applyBtn) applyBtn.addEventListener("click", onApplyScenario);

  const runBtn = document.getElementById("btnRunAnalysis");
  if (runBtn) runBtn.addEventListener("click", onRunAnalysisClick);

  const fileInput = document.getElementById("fileInput");
  if (fileInput) {
    fileInput.addEventListener("change", (event) => {
      const file = event.target.files[0];
      onFileSelected(file);
    });
  }

  const validateBtn = document.getElementById("btnValidateSelectedFile");
  if (validateBtn) {
    validateBtn.addEventListener("click", onValidateSelectedFileClick);
  }

  const uploadBtn = document.getElementById("btnUploadSelectedFile");
  if (uploadBtn) {
    uploadBtn.addEventListener("click", onUploadSelectedFileClick);
  }

  const resetBtn = document.getElementById("btnResetTool");
  if (resetBtn) {
    resetBtn.addEventListener("click", onResetToolClick);
  }

  const controlSelect = document.getElementById("controlTreatment");
  if (controlSelect) {
    controlSelect.addEventListener("change", onControlTreatmentChange);
  }

  const templateTsvBtn = document.getElementById("btnDownloadTemplateTSV");
  if (templateTsvBtn) {
    templateTsvBtn.addEventListener("click", downloadTemplateTSV);
  }

  const templateXlsxBtn = document.getElementById("btnDownloadTemplateXLSX");
  if (templateXlsxBtn) {
    templateXlsxBtn.addEventListener("click", downloadTemplateXLSX);
  }

  const exportReportBtn = document.getElementById("btnExportWordReport");
  if (exportReportBtn) {
    exportReportBtn.addEventListener("click", onExportWordReportClick);
  }

  const exportTechBtn = document.getElementById("btnExportWordTechnical");
  if (exportTechBtn) {
    exportTechBtn.addEventListener("click", onExportWordTechnicalClick);
  }

  const copyPromptBtn = document.getElementById("btnCopyAIPrompt");
  if (copyPromptBtn) {
    copyPromptBtn.addEventListener("click", copyAIPromptToClipboard);
  }

  const chatgptBtn = document.getElementById("btnOpenChatGPT");
  if (chatgptBtn) {
    chatgptBtn.addEventListener("click", () => openAiAssistant("chatgpt"));
  }

  const copilotBtn = document.getElementById("btnOpenCopilot");
  if (copilotBtn) {
    copilotBtn.addEventListener("click", () => openAiAssistant("copilot"));
  }
}

document.addEventListener("DOMContentLoaded", () => {
  initToast();
  setupTabs();
  attachEventListeners();
  loadDefaultDataset().catch((err) => {
    console.error(err);
    showToast("Unable to load built-in example dataset.", "error");
  });
});
