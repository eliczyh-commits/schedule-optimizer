const schemas = {
  resources: {
    mount: "resources",
    columns: [
      ["period", "旬度"],
      ["total", "总资源"],
      ["foreign", "外贸资源"],
      ["domestic", "内贸资源", true],
      ["hrb500e", "HRB500E保供"],
      ["six_hundred", "600兆帕保供"],
      ["four_hundred", "400兆帕资源", true]
    ],
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
    columns: [["spec", "规格"], ["lower", "比例下限"], ["upper", "比例上限"]],
    blank: { spec: "", lower: "", upper: "" }
  },
  forecast: {
    mount: "forecast",
    columns: [["period", "旬度"], ["grade", "牌号"], ["spec", "规格"], ["tons", "客户预报吨位"]],
    blank: { period: "上旬", grade: "", spec: "", tons: "" }
  }
};

const state = Object.fromEntries(Object.keys(schemas).map((key) => [key, []]));
let pendingImportTable = null;

function setStatus(text) {
  document.getElementById("status").textContent = text;
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "";
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(4)));
}

function escapeAttr(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;");
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
}

function renderEditableTable(key) {
  const schema = schemas[key];
  const container = document.getElementById(schema.mount);
  const rows = state[key];
  const headers = schema.columns.map(([, label]) => `<th>${label}</th>`).join("");
  const body = rows.map((row, index) => {
    const cells = schema.columns.map(([field, , readonly]) => {
      const attr = readonly ? "readonly aria-readonly=\"true\" class=\"readonly\"" : "";
      return `<td><input ${attr} data-table="${key}" data-row="${index}" data-field="${field}" value="${escapeAttr(row[field])}"></td>`;
    }).join("");
    return `<tr>${cells}<td><button class="row-remove" data-remove="${key}" data-row="${index}" type="button">×</button></td></tr>`;
  }).join("");
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${headers}<th></th></tr></thead><tbody>${body}</tbody></table></div>`;
}

function renderAllInputs() {
  recalcResources();
  Object.keys(schemas).forEach(renderEditableTable);
}

function renderReadOnly(containerId, rows, headers) {
  const container = document.getElementById(containerId);
  const body = rows.map((row) => `<tr>${headers.map((header) => `<td>${row[header] ?? ""}</td>`).join("")}</tr>`).join("");
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table></div>`;
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
  setStatus("正在生成方案");
  const response = await fetch("/api/solve", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(state)
  });
  const data = await response.json();
  document.querySelector('[data-tab="result"]').click();
  if (!response.ok) {
    errorBox.textContent = data.error || "生成失败";
    errorBox.classList.remove("hidden");
    setStatus("生成失败");
    return;
  }
  renderReadOnly("plan", data.plan, ["旬度", "牌号", "规格", "吨位", "生产条线"]);
  renderReadOnly("checks", data.checks, ["校验项", "旬度", "产线", "结果", "数值", "说明"]);
  renderReadOnly("suggestions", data.suggestions, ["旬度", "牌号", "规格", "客户预报吨位", "排产方案吨位", "折扣", "简要原因"]);
  setStatus("方案已生成");
}

async function importExcel(table, file) {
  const form = new FormData();
  form.append("file", file);
  setStatus("正在导入 Excel");
  const response = await fetch(`/api/import-excel?table=${encodeURIComponent(table)}`, {
    method: "POST",
    body: form
  });
  const data = await response.json();
  if (!response.ok) {
    setStatus(data.error || "导入失败");
    alert(data.error || "导入失败");
    return;
  }
  state[table] = data.rows || [];
  renderEditableTable(table);
  setStatus(`已导入 ${state[table].length} 行`);
}

document.addEventListener("input", (event) => {
  const target = event.target;
  if (!target.matches("input[data-table]") || target.readOnly) return;
  const table = target.dataset.table;
  const row = Number(target.dataset.row);
  const field = target.dataset.field;
  state[table][row][field] = target.value;
  if (table === "resources") {
    recalcResources();
    renderEditableTable("resources");
  }
});

document.addEventListener("click", (event) => {
  const target = event.target;
  if (target.matches(".tabs button")) {
    document.querySelectorAll(".tabs button").forEach((button) => button.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
    target.classList.add("active");
    document.getElementById(target.dataset.tab).classList.add("active");
  }
  if (target.dataset.add) {
    const key = target.dataset.add;
    state[key].push({ ...schemas[key].blank });
    renderEditableTable(key);
  }
  if (target.dataset.remove) {
    const key = target.dataset.remove;
    state[key].splice(Number(target.dataset.row), 1);
    renderEditableTable(key);
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
  if (file && pendingImportTable) {
    importExcel(pendingImportTable, file);
  }
});

document.getElementById("loadDefaults").addEventListener("click", loadDefaults);
document.getElementById("solve").addEventListener("click", solve);

renderAllInputs();
loadDefaults();
