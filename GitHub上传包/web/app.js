const schemas = {
  resources: {
    mount: "resources",
    columns: [["period", "旬度"], ["total", "总资源"], ["foreign", "外贸资源"], ["domestic", "内贸资源", true], ["hrb500e", "HRB500E保供"], ["six_hundred", "600兆帕保供"], ["four_hundred", "400兆帕资源", true]],
    blank: { period: "上旬", total: "", foreign: "", domestic: "", hrb500e: "", six_hundred: "", four_hundred: "" }
  },
  calendar: {
    mount: "calendar",
    columns: [["period", "旬度"], ["line", "产线"], ["days", "生产天数"]],
    blank: { period: "上旬", line: "", days: "" }
  },
  foreign: {
    mount: "foreignTable",
    columns: [["period", "旬度"], ["line", "产线"], ["grade", "牌号"], ["spec", "规格"], ["tons", "吨位"]],
    blank: { period: "上旬", line: "", grade: "", spec: "", tons: "" }
  },
  efficiency: {
    mount: "efficiency",
    columns: [["line", "产线"], ["grade", "牌号"], ["spec", "规格"], ["daily", "日产量"]],
    blank: { line: "", grade: "", spec: "", daily: "" }
  },
  cashflow: {
    mount: "cashflow",
    columns: [["grade", "牌号"], ["spec", "规格"], ["cashflow", "现金流"]],
    blank: { grade: "", spec: "", cashflow: "" }
  },
  demands: {
    mount: "demands",
    columns: [["period", "旬度"], ["grade", "牌号"], ["spec", "规格"], ["tons", "需求吨位"]],
    blank: { period: "上旬", grade: "", spec: "", tons: "" }
  },
  ratios: {
    mount: "ratios",
    columns: [["spec", "\u89c4\u683c"], ["lower", "\u6bd4\u4f8b\u4e0b\u9650(%)"], ["upper", "\u6bd4\u4f8b\u4e0a\u9650(%)"]],
    blank: { spec: "", lower: "", upper: "" }
  },
  grade_ratios: {
    mount: "gradeRatios",
    columns: [["grade", "\u724c\u53f7"], ["lower", "\u6bd4\u4f8b\u4e0b\u9650(%)"], ["upper", "\u6bd4\u4f8b\u4e0a\u9650(%)"]],
    blank: { grade: "HRB400E", lower: "", upper: "" }
  },
  forecast: {
    mount: "forecast",
    columns: [["period", "旬度"], ["grade", "牌号"], ["spec", "规格"], ["tons", "客户预报吨位"]],
    blank: { period: "上旬", grade: "", spec: "", tons: "" }
  }
};

const state = Object.fromEntries(Object.keys(schemas).map((key) => [key, []]));
const defaultGradeRatios = [
  { grade: "HRB400E", lower: 60, upper: 100 },
  { grade: "HRB400", lower: 30, upper: 40 }
];
const selectOptions = {
  calendar: {
    period: ["上旬", "中旬", "下旬"],
    line: ["棒三", "棒五A线", "棒五B线", "线三", "线五"]
  }
};
let pendingImportTable = null;
let lastResult = { plan: [], cashflow_summary: [] };
let adjustmentPlan = [];
let selectedAdjustmentKey = "";

function setStatus(text) {
  document.getElementById("status").textContent = text;
}

function syncStickyOffsets() {
  const topbar = document.querySelector(".topbar");
  const dashboard = document.getElementById("dashboardShell");
  const tabs = document.querySelector(".tabs");
  const root = document.documentElement;
  const topbarH = topbar ? Math.ceil(topbar.getBoundingClientRect().height) : 0;
  const dashboardH = dashboard ? Math.ceil(dashboard.getBoundingClientRect().height) : 0;
  const tabsH = tabs ? Math.ceil(tabs.getBoundingClientRect().height) : 0;
  root.style.setProperty("--topbar-h", `${topbarH}px`);
  root.style.setProperty("--dashboard-shell-h", `${dashboardH}px`);
  root.style.setProperty("--tabs-h", `${tabsH}px`);
  root.style.setProperty("--fixed-header-h", `${topbarH + dashboardH + tabsH + 18}px`);
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "";
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(4)));
}

function percentForDisplay(value) {
  if (value === "" || value === null || value === undefined) return null;
  const number = Number(value);
  if (!Number.isFinite(number)) return null;
  return number > 0 && number <= 1 ? number * 100 : number;
}

function percentInputValue(key, field, value) {
  if (!["ratios", "grade_ratios"].includes(key) || !["lower", "upper"].includes(field)) return value;
  const display = percentForDisplay(value);
  return display === null ? value : formatNumber(display);
}
function escapeAttr(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
}

function syncGradeRatios() {
}

function recalcResources() {
  state.resources.forEach((row) => {
    const total = toNumber(row.total);
    const foreign = toNumber(row.foreign);
    const domestic = total - foreign;
    const fourHundred = domestic - toNumber(row.hrb500e) - toNumber(row.six_hundred);
    row.domestic = total || foreign ? formatNumber(domestic) : "";
    row.four_hundred = total || foreign || row.hrb500e || row.six_hundred ? formatNumber(fourHundred) : "";
  });
}

function normalizeDefaults(data) {
  Object.keys(schemas).forEach((key) => {
    state[key] = data[key] || [];
  });
  recalcResources();
  syncGradeRatios();
}

function controlHtml(key, row, rowIndex, field, readonly) {
  const options = selectOptions[key]?.[field];
  if (options) {
    const current = row[field] ?? "";
    const optionHtml = options.map((option) => `<option value="${escapeAttr(option)}" ${String(current) === option ? "selected" : ""}>${option}</option>`).join("");
    return `<select data-table="${key}" data-row="${rowIndex}" data-field="${field}"><option value=""></option>${optionHtml}</select>`;
  }
  const attr = readonly ? "readonly aria-readonly=\"true\" class=\"readonly\"" : "";
  return `<input ${attr} data-table="${key}" data-row="${rowIndex}" data-field="${field}" value="${escapeAttr(percentInputValue(key, field, row[field]))}">`;
}

function tableSummaryTfoot(key) {
  if (key === "foreign") {
    const total = state.foreign.reduce((sum, row) => sum + toNumber(row.tons), 0);
    return `<tfoot><tr class="total-row"><td colspan="4">\u5408\u8ba1</td><td data-foreign-total>${formatNumber(total)}</td><td></td></tr></tfoot>`;
  }
  if (key === "demands") {
    const hrb500e = state.demands.reduce((sum, row) => String(row.grade || "").trim() === "HRB500E" ? sum + toNumber(row.tons) : sum, 0);
    const sixHundred = state.demands.reduce((sum, row) => ["T63E/E/G", "T63/E/G"].includes(String(row.grade || "").trim()) ? sum + toNumber(row.tons) : sum, 0);
    return `<tfoot><tr class="total-row"><td colspan="3">HRB500E\u603b\u91cf</td><td data-demand-hrb500e>${formatNumber(hrb500e)}</td><td></td></tr><tr class="total-row"><td colspan="3">600\u5146\u5e15\u603b\u91cf</td><td data-demand-six>${formatNumber(sixHundred)}</td><td></td></tr></tfoot>`;
  }
  return "";
}

function updateTableSummaryFooter(key) {
  if (key === "foreign") {
    const cell = document.querySelector('[data-foreign-total]');
    if (!cell) return;
    const total = state.foreign.reduce((sum, row) => sum + toNumber(row.tons), 0);
    cell.textContent = formatNumber(total);
  }
  if (key === "demands") {
    const hrb500eCell = document.querySelector('[data-demand-hrb500e]');
    const sixCell = document.querySelector('[data-demand-six]');
    const hrb500e = state.demands.reduce((sum, row) => String(row.grade || "").trim() === "HRB500E" ? sum + toNumber(row.tons) : sum, 0);
    const sixHundred = state.demands.reduce((sum, row) => ["T63E/E/G", "T63/E/G"].includes(String(row.grade || "").trim()) ? sum + toNumber(row.tons) : sum, 0);
    if (hrb500eCell) hrb500eCell.textContent = formatNumber(hrb500e);
    if (sixCell) sixCell.textContent = formatNumber(sixHundred);
  }
}

function renderEditableTable(key) {
  if (key === "grade_ratios") syncGradeRatios();
  const schema = schemas[key];
  const container = document.getElementById(schema.mount);
  const rows = state[key];
  const headers = schema.columns.map(([, label]) => `<th>${label}</th>`).join("");
  const body = rows.map((row, index) => {
    const cells = schema.columns.map(([field, , readonly]) => `<td>${controlHtml(key, row, index, field, readonly)}</td>`).join("");
    return `<tr>${cells}<td><button class="row-remove" data-remove="${key}" data-row="${index}" type="button" aria-label="删除">&times;</button></td></tr>`;
  }).join("");
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${headers}<th></th></tr></thead><tbody>${body}</tbody>${tableSummaryTfoot(key)}</table></div>`;
}

function renderAllInputs() {
  recalcResources();
  syncGradeRatios();
  Object.keys(schemas).forEach(renderEditableTable);
  renderDashboard();
  renderAdjustmentInsight();
}

function updateResourceReadonlyCells(rowIndex) {
  ["domestic", "four_hundred"].forEach((field) => {
    const input = document.querySelector(`input[data-table="resources"][data-row="${rowIndex}"][data-field="${field}"]`);
    if (input) input.value = state.resources[rowIndex]?.[field] ?? "";
  });
}

function updateGradeRatioReadonlyCells() {
  state.grade_ratios.forEach((row, rowIndex) => {
    ["grade", "lower", "upper"].forEach((field) => {
      const input = document.querySelector(`input[data-table="grade_ratios"][data-row="${rowIndex}"][data-field="${field}"]`);
      if (input && input !== document.activeElement) input.value = row[field] ?? "";
    });
  });
}

const PLAN_HEADERS = ["旬度", "牌号", "规格", "吨位", "生产条线"];
const CASHFLOW_HEADERS = ["分类", "吨位", "现金流", "说明"];

function rowCashflow(row) {
  const spec = Number(row["规格"]);
  const grade = String(row["牌号"] ?? "").trim();
  const tons = toNumber(row["吨位"]);
  const item = state.cashflow.find((cash) => String(cash.grade ?? "").trim() === grade && Number(cash.spec) === spec);
  return tons * toNumber(item?.cashflow);
}

function planTotals(rows) {
  return {
    tons: rows.reduce((sum, row) => sum + toNumber(row["吨位"]), 0),
    cash: rows.reduce((sum, row) => sum + rowCashflow(row), 0)
  };
}

function planBreakdown(rows) {
  const list = rows || [];
  const four = list.reduce((sum, row) => ["HRB400", "HRB400E"].includes(String(row["\u724c\u53f7"] ?? "").trim()) ? sum + toNumber(row["\u5428\u4f4d"]) : sum, 0);
  const hrb500e = list.reduce((sum, row) => String(row["\u724c\u53f7"] ?? "").trim() === "HRB500E" ? sum + toNumber(row["\u5428\u4f4d"]) : sum, 0);
  const sixHundred = list.reduce((sum, row) => ["T63E/E/G", "T63/E/G"].includes(String(row["\u724c\u53f7"] ?? "").trim()) ? sum + toNumber(row["\u5428\u4f4d"]) : sum, 0);
  return { ...planTotals(list), four, hrb500e, sixHundred };
}


function adjustmentKey(row) {
  return [row["\u65ec\u5ea6"] ?? "", row["\u724c\u53f7"] ?? "", row["\u89c4\u683c"] ?? ""].join("|");
}

function sumBy(rows, getter) {
  return rows.reduce((sum, row) => sum + toNumber(getter(row)), 0);
}

function activeTabId() {
  return document.querySelector(".tabs button.active")?.dataset.tab || "conditions";
}

function dashboardCardsForActiveTab() {
  const resources = state.resources || [];
  const totalDomestic = sumBy(resources, (row) => row.domestic);
  const fourHundred = sumBy(resources, (row) => row.four_hundred);
  const foreignTons = sumBy(state.foreign || [], (row) => row.tons);
  const hrb500e = (state.demands || []).reduce((sum, row) => String(row.grade || "").trim() === "HRB500E" ? sum + toNumber(row.tons) : sum, 0);
  const sixHundred = (state.demands || []).reduce((sum, row) => ["T63E/E/G", "T63/E/G"].includes(String(row.grade || "").trim()) ? sum + toNumber(row.tons) : sum, 0);
  const active = activeTabId();
  if (active === "initialPlan" || active === "adjustedPlan") {
    const rows = active === "adjustedPlan" ? adjustmentPlan : (lastResult.plan || []);
    const summary = planBreakdown(rows);
    const planName = active === "adjustedPlan" ? "\u8c03\u6574\u65b9\u6848" : "\u521d\u59cb\u65b9\u6848";
    const meta = rows.length ? "\u5df2\u6709\u65b9\u6848" : "\u5f85\u751f\u6210";
    return [
      [`${planName}\u603b\u5428\u4f4d`, formatNumber(summary.tons), "\u5428"],
      ["400\u5146\u5e15\u5428\u4f4d", formatNumber(summary.four), "HRB400 + HRB400E"],
      ["HRB500E\u5428\u4f4d", formatNumber(summary.hrb500e), "\u5428"],
      ["600\u5146\u5e15\u5428\u4f4d", formatNumber(summary.sixHundred), "T63E/E/G + T63/E/G"],
      [`${planName}\u73b0\u91d1\u6d41`, formatNumber(summary.cash), meta]
    ];
  }
  return [
    ["\u5185\u8d38\u8d44\u6e90", formatNumber(totalDomestic), "\u5428"],
    ["400\u5146\u5e15\u8d44\u6e90", formatNumber(fourHundred), "\u5428"],
    ["\u5916\u8d38\u6392\u4ea7", formatNumber(foreignTons), "\u5428"],
    ["HRB500E\u4fdd\u4f9b", formatNumber(hrb500e), "\u5428"],
    ["600\u5146\u5e15\u4fdd\u4f9b", formatNumber(sixHundred), "\u5428"],
    ["\u6761\u4ef6\u72b6\u6001", (state.resources || []).length ? "\u5df2\u5f55\u5165" : "\u5f85\u5f55\u5165", "\u70b9\u51fb\u751f\u6210\u65b9\u6848"]
  ];
}

function renderDashboard() {
  const container = document.getElementById("dashboard");
  if (!container) return;
  const cards = dashboardCardsForActiveTab();
  container.innerHTML = cards.map(([label, value, meta], index) => `
    <article class="kpi-card ${index === cards.length - 1 ? "kpi-card-strong" : ""}">
      <span>${label}</span>
      <strong>${value || "0"}</strong>
      <em>${meta}</em>
    </article>
  `).join("");
  syncStickyOffsets();
}


function adjustmentGroups() {
  const groups = new Map();
  adjustmentPlan.forEach((row) => {
    const key = adjustmentKey(row);
    if (!groups.has(key)) {
      groups.set(key, { key, period: row["\u65ec\u5ea6"] ?? "", grade: row["\u724c\u53f7"] ?? "", spec: row["\u89c4\u683c"] ?? "", tons: 0, lines: new Map(), cash: 0 });
    }
    const group = groups.get(key);
    const tons = toNumber(row["\u5428\u4f4d"]);
    const line = row["\u751f\u4ea7\u6761\u7ebf"] ?? "";
    group.tons += tons;
    group.cash += rowCashflow(row);
    if (line) group.lines.set(line, (group.lines.get(line) || 0) + tons);
  });
  return groups;
}

function renderAdjustmentInsight() {
  const container = document.getElementById("adjustmentInsight");
  if (!container) return;
  const groups = adjustmentGroups();
  if (!groups.size) {
    container.innerHTML = `<div class="empty-state">\u751f\u6210\u6216\u590d\u5236\u65b9\u6848\u540e\uff0c\u8fd9\u91cc\u4f1a\u663e\u793a\u5ba2\u6237\u9700\u6c42\u3001\u6298\u6263\u548c\u4ea7\u7ebf\u5206\u5e03\u3002</div>`;
    return;
  }
  if (!selectedAdjustmentKey || !groups.has(selectedAdjustmentKey)) selectedAdjustmentKey = Array.from(groups.keys())[0];
  const group = groups.get(selectedAdjustmentKey);
  const forecast = forecastTotal(group.period, group.grade, group.spec);
  const initial = (lastResult.plan || []).reduce((sum, row) => adjustmentKey(row) === selectedAdjustmentKey ? sum + toNumber(row["\u5428\u4f4d"]) : sum, 0);
  const discount = forecast ? group.tons / forecast : null;
  const gap = forecast ? group.tons - forecast : null;
  const lineHtml = Array.from(group.lines.entries()).map(([line, tons]) => `<span class="line-chip"><b>${line}</b>${formatNumber(tons)}\u5428</span>`).join("") || `<span class="line-chip muted-chip">\u672a\u5206\u914d</span>`;
  const tone = gap === null ? "neutral" : gap >= 0 ? "good" : "warn";
  container.innerHTML = `
    <div class="insight-title">
      <span>${group.period || "-"}</span>
      <strong>${group.grade || "-"} ${group.spec || "-"}\u89c4\u683c</strong>
    </div>
    <div class="insight-metrics">
      <div><span>\u5ba2\u6237\u9884\u62a5</span><strong>${forecast ? formatNumber(forecast) : "\u672a\u62a5"}</strong></div>
      <div><span>\u521d\u59cb\u65b9\u6848</span><strong>${formatNumber(initial)}</strong></div>
      <div><span>\u8c03\u6574\u65b9\u6848</span><strong>${formatNumber(group.tons)}</strong></div>
      <div class="${tone}"><span>\u5dee\u989d</span><strong>${gap === null ? "-" : formatNumber(gap)}</strong></div>
      <div><span>\u6298\u6263</span><strong>${discount === null ? "-" : formatNumber(discount)}</strong></div>
      <div><span>\u73b0\u91d1\u6d41</span><strong>${formatNumber(group.cash)}</strong></div>
    </div>
    <div class="line-list">${lineHtml}</div>
  `;
}


function renderPlanWithSummary(containerId, rows) {
  const container = document.getElementById(containerId);
  const body = (rows || []).map((row) => `<tr>${PLAN_HEADERS.map((header) => `<td>${row[header] ?? ""}</td>`).join("")}</tr>`).join("");
  const totals = planTotals(rows || []);
  const foot = `<tfoot><tr class="total-row"><td colspan="3">合计</td><td>${formatNumber(totals.tons)}</td><td>总现金流：${formatNumber(totals.cash)}</td></tr></tfoot>`;
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${PLAN_HEADERS.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>${body}</tbody>${foot}</table></div>`;
}

function availableLines() {
  const lines = [
    ...state.calendar.map((row) => row.line),
    ...state.efficiency.map((row) => row.line),
    ...state.foreign.map((row) => row.line),
    ...(lastResult.plan || []).map((row) => row["\u751f\u4ea7\u6761\u7ebf"]),
    ...adjustmentPlan.map((row) => row["\u751f\u4ea7\u6761\u7ebf"]),
    "\u68d2\u4e09", "\u68d2\u4e94A\u7ebf", "\u68d2\u4e94B\u7ebf", "\u7ebf\u4e09", "\u7ebf\u4e94"
  ];
  return Array.from(new Set(lines.map((line) => String(line || "").trim()).filter(Boolean)));
}

function dailyFor(line, grade, spec) {
  const item = state.efficiency.find((row) => String(row.line ?? "").trim() === String(line ?? "").trim() && String(row.grade ?? "").trim() === String(grade ?? "").trim() && Number(row.spec) === Number(spec));
  return toNumber(item?.daily);
}

function foreignDaysByPeriodLine() {
  const days = new Map();
  state.foreign.forEach((row) => {
    const daily = dailyFor(row.line, row.grade, row.spec);
    const key = `${row.period}|${row.line}`;
    days.set(key, (days.get(key) || 0) + (daily ? toNumber(row.tons) / daily : 0));
  });
  return days;
}

function calendarDays(period, line) {
  const item = state.calendar.find((row) => String(row.period ?? "") === String(period ?? "") && String(row.line ?? "") === String(line ?? ""));
  return item ? toNumber(item.days) : 0;
}

function evaluateAdjustmentPlan() {
  const foreignDays = foreignDaysByPeriodLine();
  const used = new Map();
  const problems = [];
  adjustmentPlan.forEach((row, index) => {
    const daily = dailyFor(row["\u751f\u4ea7\u6761\u7ebf"], row["\u724c\u53f7"], row["\u89c4\u683c"]);
    if (!daily && toNumber(row["\u5428\u4f4d"]) > 0) {
      problems.push(`\u7b2c${index + 1}\u884c\u7f3a\u5c11\u65e5\u4ea7\u91cf`);
      return;
    }
    const key = `${row["\u65ec\u5ea6"]}|${row["\u751f\u4ea7\u6761\u7ebf"]}`;
    used.set(key, (used.get(key) || 0) + toNumber(row["\u5428\u4f4d"]) / daily);
  });
  used.forEach((domesticDays, key) => {
    const [period, line] = key.split("|");
    const total = (foreignDays.get(key) || 0) + domesticDays;
    const capacity = calendarDays(period, line);
    if (!capacity) problems.push(`${period}-${line}\u672a\u586b\u751f\u4ea7\u5929\u6570`);
    if (capacity && total > capacity + 1e-8) problems.push(`${period}-${line}\u8d85\u51fa${formatNumber(total - capacity)}\u5929`);
  });
  return { ok: problems.length === 0, problems };
}

function forecastTotal(period, grade, spec) {
  return state.forecast.reduce((sum, row) => {
    const samePeriod = String(row.period ?? "") === String(period ?? "");
    const sameGrade = String(row.grade ?? "").trim() === String(grade ?? "").trim();
    const sameSpec = Number(row.spec) === Number(spec);
    return samePeriod && sameGrade && sameSpec ? sum + toNumber(row.tons) : sum;
  }, 0);
}

function renderAdjustmentDemandSummary() {
  const groups = new Map();
  adjustmentPlan.forEach((row) => {
    const key = [row["\u65ec\u5ea6"], row["\u724c\u53f7"], row["\u89c4\u683c"]].join("|");
    if (!groups.has(key)) {
      groups.set(key, { "\u65ec\u5ea6": row["\u65ec\u5ea6"] ?? "", "\u724c\u53f7": row["\u724c\u53f7"] ?? "", "\u89c4\u683c": row["\u89c4\u683c"] ?? "", tons: 0, lines: new Map() });
    }
    const group = groups.get(key);
    const tons = toNumber(row["\u5428\u4f4d"]);
    const line = row["\u751f\u4ea7\u6761\u7ebf"] ?? "";
    group.tons += tons;
    if (line) group.lines.set(line, (group.lines.get(line) || 0) + tons);
  });
  const rows = Array.from(groups.values()).map((group) => {
    const forecast = forecastTotal(group["\u65ec\u5ea6"], group["\u724c\u53f7"], group["\u89c4\u683c"]);
    const discount = forecast ? group.tons / forecast : null;
    const lineText = Array.from(group.lines.entries()).map(([line, tons]) => `${line} ${formatNumber(tons)}`).join("\u3001");
    return {
      "\u65ec\u5ea6": group["\u65ec\u5ea6"],
      "\u724c\u53f7": group["\u724c\u53f7"],
      "\u89c4\u683c": group["\u89c4\u683c"],
      "\u5ba2\u6237\u9884\u62a5\u5428\u4f4d": forecast ? formatNumber(forecast) : "\u672a\u62a5",
      "\u8c03\u6574\u65b9\u6848\u5408\u8ba1\u5428\u4f4d": formatNumber(group.tons),
      "\u6298\u6263": discount === null ? "-" : formatNumber(discount),
      "\u5dee\u989d": forecast ? formatNumber(group.tons - forecast) : "-",
      "\u751f\u4ea7\u6761\u7ebf": lineText
    };
  }).sort((a, b) => String(a["\u65ec\u5ea6"]).localeCompare(String(b["\u65ec\u5ea6"])) || String(a["\u724c\u53f7"]).localeCompare(String(b["\u724c\u53f7"])) || Number(a["\u89c4\u683c"]) - Number(b["\u89c4\u683c"]));
  renderReadOnly("adjustmentDemandSummary", rows, ["\u65ec\u5ea6", "\u724c\u53f7", "\u89c4\u683c", "\u5ba2\u6237\u9884\u62a5\u5428\u4f4d", "\u8c03\u6574\u65b9\u6848\u5408\u8ba1\u5428\u4f4d", "\u6298\u6263", "\u5dee\u989d", "\u751f\u4ea7\u6761\u7ebf"]);
}

function renderAdjustmentSpecRatios() {
  const specTotals = new Map();
  const allTons = adjustmentPlan.reduce((sum, row) => sum + toNumber(row["吨位"]), 0);
  adjustmentPlan.forEach((row) => {
    const spec = row["规格"];
    specTotals.set(spec, (specTotals.get(spec) || 0) + toNumber(row["吨位"]));
  });
  const rows = Array.from(specTotals.entries()).sort((a, b) => Number(a[0]) - Number(b[0])).map(([spec, tons]) => ({
    "规格": spec,
    "吨位": formatNumber(tons),
    "占比": allTons ? formatNumber((tons / allTons) * 100) + "%" : ""
  }));
  renderReadOnly("adjustmentSpecRatios", rows, ["规格", "吨位", "占比"]);
}

function renderAdjustmentPlan() {
  const container = document.getElementById("adjustmentPlan");
  if (!container) return;
  const lines = availableLines();
  const body = adjustmentPlan.map((row, index) => {
    const key = adjustmentKey(row);
    const selectedClass = selectedAdjustmentKey === key ? " class=\"selected-row\"" : "";
    const lineOptions = lines.map((line) => `<option value="${escapeAttr(line)}" ${String(row["\u751f\u4ea7\u6761\u7ebf"]) === line ? "selected" : ""}>${line}</option>`).join("");
    const forecast = forecastTotal(row["\u65ec\u5ea6"], row["\u724c\u53f7"], row["\u89c4\u683c"]);
    const forecastText = forecast ? `\u603b\u9884\u62a5 ${formatNumber(forecast)}` : "\u672a\u62a5";
    return `<tr${selectedClass} data-adjust-key="${escapeAttr(key)}"><td>${row["\u65ec\u5ea6"] ?? ""}</td><td>${row["\u724c\u53f7"] ?? ""}</td><td>${row["\u89c4\u683c"] ?? ""}</td><td><input data-adjust-row="${index}" data-adjust-field="\u5428\u4f4d" value="${escapeAttr(row["\u5428\u4f4d"])}"></td><td><select data-adjust-row="${index}" data-adjust-field="\u751f\u4ea7\u6761\u7ebf">${lineOptions}</select></td><td><span class="forecast-pill">${forecastText}</span></td></tr>`;
  }).join("");
  const totals = planTotals(adjustmentPlan);
  const evaluation = evaluateAdjustmentPlan();
  const status = evaluation.ok ? "\u53ef\u4ee5\u6392\u5165\u4ea7\u7ebf" : `\u4e0d\u53ef\u6392\uff1a${evaluation.problems.join("\uff1b")}`;
  const foot = `<tfoot><tr class="total-row"><td colspan="3">\u5408\u8ba1</td><td>${formatNumber(totals.tons)}</td><td colspan="2">\u603b\u73b0\u91d1\u6d41\uff1a${formatNumber(totals.cash)}\uff1b${status}</td></tr></tfoot>`;
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${[...PLAN_HEADERS, "\u5ba2\u6237\u9884\u62a5\u53c2\u8003"].map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>${body}</tbody>${foot}</table></div>`;
  renderAdjustmentDemandSummary();
  renderAdjustmentSpecRatios();
  renderAdjustmentInsight();
  renderDashboard();
}
function copyInitialToAdjustment() {
  selectedAdjustmentKey = "";
  adjustmentPlan = (lastResult.plan || []).map((row) => ({
    "\u65ec\u5ea6": row["\u65ec\u5ea6"] ?? "",
    "\u724c\u53f7": row["\u724c\u53f7"] ?? "",
    "\u89c4\u683c": row["\u89c4\u683c"] ?? "",
    "\u5428\u4f4d": row["\u5428\u4f4d"] ?? "",
    "\u751f\u4ea7\u6761\u7ebf": row["\u751f\u4ea7\u6761\u7ebf"] ?? ""
  }));
  renderAdjustmentPlan();
  renderDashboard();
}
function renderReadOnly(containerId, rows, headers) {
  const container = document.getElementById(containerId);
  const body = rows.map((row) => `<tr>${headers.map((header) => `<td>${row[header] ?? ""}</td>`).join("")}</tr>`).join("");
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function clearResults() {
  ["plan", "checks", "suggestions", "diagnostics", "cashflowSummary", "adjustmentPlan", "adjustmentDemandSummary", "adjustmentSpecRatios", "adjustmentInsight"].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.innerHTML = "";
  });
}

function renderDiagnostics(rows) {
  const panel = document.getElementById("diagnosticsPanel");
  const list = rows || [];
  if (!list.length) {
    if (panel) panel.classList.add("hidden");
    const container = document.getElementById("diagnostics");
    if (container) container.innerHTML = "";
    return;
  }
  if (panel) panel.classList.remove("hidden");
  renderReadOnly("diagnostics", list, ["\u5206\u6790\u9879", "\u5bf9\u8c61", "\u5224\u65ad", "\u8bf4\u660e"]);
}

async function loadDefaults() {
  setStatus("正在重置空模板");
  const response = await fetch("/api/defaults");
  const data = await response.json();
  if (!response.ok) {
    setStatus(data.error || "空模板读取失败");
    return;
  }
  normalizeDefaults(data);
  renderAllInputs();
  setStatus(data.message || "已打开空模板");
}

async function solve() {
  recalcResources();
  const errorBox = document.getElementById("errorBox");
  errorBox.classList.add("hidden");
  clearResults();
  renderDiagnostics([]);
  setStatus("\u6b63\u5728\u751f\u6210\u65b9\u6848");
  const response = await fetch("/api/solve", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(state)
  });
  const data = await response.json();
  document.querySelector('[data-tab="initialPlan"]').click();
  if (!response.ok) {
    errorBox.textContent = data.error || "\u751f\u6210\u5931\u8d25";
    errorBox.classList.remove("hidden");
    renderDiagnostics(data.diagnostics || []);
    lastResult = { plan: [], cashflow_summary: [] };
    adjustmentPlan = [];
    selectedAdjustmentKey = "";
    renderDashboard();
    renderAdjustmentInsight();
    setStatus("\u751f\u6210\u5931\u8d25");
    return;
  }
  lastResult = data;
  renderPlanWithSummary("plan", data.plan || []);
  copyInitialToAdjustment();
  renderReadOnly("cashflowSummary", data.cashflow_summary || [], ["\u5206\u7c7b", "\u5428\u4f4d", "\u73b0\u91d1\u6d41", "\u8bf4\u660e"]);
  renderReadOnly("checks", data.checks, ["\u6821\u9a8c\u9879", "\u65ec\u5ea6", "\u4ea7\u7ebf", "\u7ed3\u679c", "\u6570\u503c", "\u8bf4\u660e"]);
  renderReadOnly("suggestions", data.suggestions, ["\u65ec\u5ea6", "\u724c\u53f7", "\u89c4\u683c", "\u5ba2\u6237\u9884\u62a5\u5428\u4f4d", "\u6392\u4ea7\u65b9\u6848\u5428\u4f4d", "\u6298\u6263", "\u7b80\u8981\u539f\u56e0"]);
  renderDiagnostics([]);
  renderDashboard();
  setStatus("\u65b9\u6848\u5df2\u751f\u6210");
}

async function importExcel(table, file) {
  const form = new FormData();
  form.append("file", file);
  setStatus("正在导入 Excel");
  const response = await fetch(`/api/import-excel?table=${encodeURIComponent(table)}`, { method: "POST", body: form });
  const data = await response.json();
  if (!response.ok) {
    setStatus(data.error || "导入失败");
    alert(data.error || "导入失败");
    return;
  }
  state[table] = data.rows || [];
  if (table === "demands") state[table] = state[table].filter((row) => toNumber(row.tons) !== 0);
  if (table === "ratios") {
    state[table] = state[table].map((row) => ({ ...row, lower: percentInputValue("ratios", "lower", row.lower), upper: percentInputValue("ratios", "upper", row.upper) }));
  }
  renderEditableTable(table);
  setStatus(`已导入 ${state[table].length} 行`);
}

function rowsFromRenderedTable(containerId) {
  const table = document.querySelector(`#${containerId} table`);
  if (!table) return [];
  const headers = Array.from(table.querySelectorAll("thead th")).map((th) => th.textContent.trim()).filter(Boolean);
  return Array.from(table.querySelectorAll("tbody tr")).map((tr) => {
    const cells = Array.from(tr.querySelectorAll("td"));
    const row = {};
    headers.forEach((header, index) => {
      row[header] = cells[index]?.textContent.trim() ?? "";
    });
    return row;
  }).filter((row) => Object.values(row).some((value) => value !== ""));
}

async function exportExcel(rows, headers, filename, sheetName) {
  if (!rows || !rows.length) {
    alert("\u6ca1\u6709\u53ef\u5bfc\u51fa\u7684\u6570\u636e\uff0c\u8bf7\u5148\u751f\u6210\u65b9\u6848\u3002");
    return;
  }
  const response = await fetch("/api/export-excel", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows, headers, filename, sheetName })
  });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    alert(data.error || "\u5bfc\u51fa\u5931\u8d25");
    return;
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function updateField(target) {
  if (!target.matches("input[data-table], select[data-table]") || target.readOnly) return;
  const table = target.dataset.table;
  const row = Number(target.dataset.row);
  const field = target.dataset.field;
  state[table][row][field] = target.value;
  if (table === "resources") {
    recalcResources();
    updateResourceReadonlyCells(row);
    renderDashboard();
  }
  if (table === "foreign") {
    updateTableSummaryFooter("foreign");
    renderDashboard();
  }
  if (table === "demands") {
    updateTableSummaryFooter("demands");
    renderDashboard();
  }
}

document.addEventListener("input", (event) => updateField(event.target));
document.addEventListener("change", (event) => updateField(event.target));

document.addEventListener("input", (event) => {
  const target = event.target;
  if (!target.matches("input[data-adjust-row], select[data-adjust-row]")) return;
  const row = Number(target.dataset.adjustRow);
  const field = target.dataset.adjustField;
  adjustmentPlan[row][field] = target.value;
  selectedAdjustmentKey = adjustmentKey(adjustmentPlan[row]);
  renderAdjustmentDemandSummary();
  renderAdjustmentSpecRatios();
  renderAdjustmentInsight();
  renderDashboard();
});

document.addEventListener("focusin", (event) => {
  const target = event.target;
  if (!target.matches("input[data-adjust-row], select[data-adjust-row]")) return;
  const row = Number(target.dataset.adjustRow);
  selectedAdjustmentKey = adjustmentKey(adjustmentPlan[row]);
  renderAdjustmentInsight();
});

document.addEventListener("change", (event) => {
  const target = event.target;
  if (!target.matches("input[data-adjust-row], select[data-adjust-row]")) return;
  const row = Number(target.dataset.adjustRow);
  const field = target.dataset.adjustField;
  adjustmentPlan[row][field] = target.value;
  renderAdjustmentPlan();
  renderDashboard();
});

document.addEventListener("click", (event) => {
  const target = event.target;
  if (target.matches(".tabs button")) {
    document.querySelectorAll(".tabs button").forEach((button) => button.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
    target.classList.add("active");
    document.getElementById(target.dataset.tab).classList.add("active");
    renderDashboard();
    if (target.dataset.tab === "adjustedPlan") renderAdjustmentInsight();
  }
  if (target.dataset.add) {
    const key = target.dataset.add;
    state[key].push({ ...schemas[key].blank });
    renderEditableTable(key);
  }
  if (target.dataset.clear) {
    const key = target.dataset.clear;
    state[key] = key === "grade_ratios" ? defaultGradeRatios.map((row) => ({ ...row })) : ["resources", "calendar"].includes(key) ? [{ ...schemas[key].blank }] : [];
    if (key === "resources") recalcResources();
    if (key === "grade_ratios") syncGradeRatios();
    renderEditableTable(key);
    renderDashboard();
    setStatus("\u5df2\u6e05\u9664\u5f53\u524d\u6a21\u5757");
  }
  if (target.dataset.remove) {
    const key = target.dataset.remove;
    state[key].splice(Number(target.dataset.row), 1);
    renderEditableTable(key);
    renderDashboard();
  }
  const adjustmentRow = target.closest("tr[data-adjust-key]");
  if (adjustmentRow && !target.matches("input, select, button")) {
    selectedAdjustmentKey = adjustmentRow.dataset.adjustKey || "";
    renderAdjustmentPlan();
  }
  if (target.dataset.togglePanel !== undefined) {
    target.closest(".panel").classList.toggle("collapsed");
  }
  if (target.dataset.import) {
    pendingImportTable = target.dataset.import;
    const input = document.getElementById("excelImport");
    input.value = "";
    input.click();
  }
});

document.getElementById("excelImport").addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file && pendingImportTable) importExcel(pendingImportTable, file);
});

document.getElementById("loadDefaults").addEventListener("click", loadDefaults);
document.getElementById("solve").addEventListener("click", solve);
document.getElementById("toggleDashboard").addEventListener("click", () => {
  const shell = document.getElementById("dashboardShell");
  const button = document.getElementById("toggleDashboard");
  const collapsed = shell.classList.toggle("collapsed");
  button.setAttribute("aria-expanded", String(!collapsed));
  button.querySelector(".toggle-text").textContent = collapsed ? "\u5c55\u5f00\u603b\u6570" : "\u6536\u8d77\u603b\u6570";
  syncStickyOffsets();
});
document.getElementById("exportPlan").addEventListener("click", () => exportExcel(lastResult.plan || [], PLAN_HEADERS, "内贸排产初始方案.xlsx", "内贸排产初始方案"));
document.getElementById("copyAdjustment").addEventListener("click", copyInitialToAdjustment);
document.getElementById("exportAdjustment").addEventListener("click", () => exportExcel(adjustmentPlan || [], PLAN_HEADERS, "\u5185\u8d38\u6392\u4ea7\u8c03\u6574\u65b9\u6848.xlsx", "\u5185\u8d38\u6392\u4ea7\u8c03\u6574\u65b9\u6848"));
window.addEventListener("resize", syncStickyOffsets);

document.getElementById("exportCashflow").addEventListener("click", () => exportExcel(lastResult.cashflow_summary || [], CASHFLOW_HEADERS, "现金流汇总.xlsx", "现金流汇总"));

renderAllInputs();
loadDefaults();






