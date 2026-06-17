/* global Office, Excel */

const API_BASE_URL = "http://localhost:8000/api";

const COLOR_OPTIONS = [
  { value: "yellow",  label: "Amarelo" },
  { value: "green",   label: "Verde" },
  { value: "red",     label: "Vermelho" },
  { value: "blue",    label: "Azul" },
  { value: "orange",  label: "Laranja" },
  { value: "gray",    label: "Cinza" },
  { value: "pink",    label: "Rosa" },
];
const OPERATOR_OPTIONS = [
  { value: "equals",       label: "Igual a" },
  { value: "not_equals",   label: "Diferente de" },
  { value: "starts_with",  label: "Começa com" },
  { value: "ends_with",    label: "Termina com" },
  { value: "contains",     label: "Contém" },
  { value: "greater_than", label: "Maior que" },
  { value: "less_than",    label: "Menor que" },
  { value: "is_empty",     label: "Está vazio" },
  { value: "is_not_empty", label: "Não está vazio" },
];

let columns = [];

// ─── Office init ────────────────────────────────────────────────────────────

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    columns = await getSheetColumns();
    populateColumnSelects();
    await loadProfiles();
    bindEvents();
  }
});

// ─── Excel helpers ───────────────────────────────────────────────────────────

async function getSheetColumns() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();
    if (!range.values || range.values.length === 0) return [];
    return range.values[0].map(String);
  });
}

async function getSheetDataAsBlob() {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await ctx.sync();

    const rows = range.values;
    const header = rows[0];
    const csvLines = rows.map((r) =>
      r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")
    );
    const csv = csvLines.join("\n");
    return new Blob([csv], { type: "text/csv" });
  });
}

// ─── Populate selects ────────────────────────────────────────────────────────

function populateColumnSelects() {
  document.querySelectorAll(".sort-col-select, #dup-cols, #split-col").forEach((sel) => {
    const isMultiple = sel.id === "dup-cols";
    if (!isMultiple) sel.innerHTML = "";
    if (sel.id === "split-col") sel.innerHTML = '<option value="">-- Não dividir --</option>';
    columns.forEach((col) => {
      const opt = document.createElement("option");
      opt.value = col;
      opt.textContent = col;
      sel.appendChild(opt);
    });
  });

  const textRules = document.getElementById("text-rules");
  textRules.innerHTML = "";
  columns.forEach((col) => {
    const row = document.createElement("div");
    row.style.display = "flex";
    row.style.gap = "6px";
    row.style.alignItems = "center";
    row.innerHTML = `<span style="flex:1;font-size:12px;">${col}</span>
      <select class="text-op" data-col="${col}">
        <option value="">não alterar</option>
        <option value="upper">MAIÚSCULAS</option>
        <option value="lower">minúsculas</option>
        <option value="capitalize">Capitalizar</option>
        <option value="strip">Remover espaços</option>
      </select>`;
    textRules.appendChild(row);
  });
}

// ─── Color rules ─────────────────────────────────────────────────────────────

function addColorRule() {
  const container = document.getElementById("color-rules");
  const row = document.createElement("div");
  row.className = "color-rule-row";
  row.innerHTML = `
    <select class="cr-col">${columns.map((c) => `<option value="${c}">${c}</option>`).join("")}</select>
    <select class="cr-op">${OPERATOR_OPTIONS.map((o) => `<option value="${o.value}">${o.label}</option>`).join("")}</select>
    <input type="text" class="cr-val" placeholder="valor" />
    <select class="cr-color">${COLOR_OPTIONS.map((c) => `<option value="${c.value}">${c.label}</option>`).join("")}</select>
    <button class="btn-secondary cr-remove" style="margin:0;padding:2px 6px;">✕</button>`;
  row.querySelector(".cr-remove").addEventListener("click", () => row.remove());
  container.appendChild(row);
}

// ─── Sort rules ──────────────────────────────────────────────────────────────

function addSortRow() {
  const container = document.getElementById("sort-rules");
  const row = document.createElement("div");
  row.className = "sort-row";
  row.innerHTML = `
    <select class="sort-col-select">${columns.map((c) => `<option value="${c}">${c}</option>`).join("")}</select>
    <select class="sort-dir-select">
      <option value="asc">Crescente</option>
      <option value="desc">Decrescente</option>
    </select>
    <button class="btn-secondary" style="margin:0;padding:2px 6px;">✕</button>`;
  row.querySelector("button").addEventListener("click", () => row.remove());
  container.appendChild(row);
}

// ─── Collect params ──────────────────────────────────────────────────────────

function collectParams() {
  const params = {};

  const sortRows = document.querySelectorAll(".sort-row");
  const sortBy = [];
  sortRows.forEach((row) => {
    const col = row.querySelector(".sort-col-select")?.value;
    const dir = row.querySelector(".sort-dir-select")?.value;
    if (col) sortBy.push({ column: col, direction: dir });
  });
  if (sortBy.length) params.sort_by = sortBy;

  if (document.getElementById("chk-duplicates").checked) {
    params.remove_duplicates = true;
    const dupSel = document.getElementById("dup-cols");
    const selected = Array.from(dupSel.selectedOptions).map((o) => o.value);
    if (selected.length) params.duplicate_columns = selected;
    params.keep_duplicate = document.querySelector('input[name="keep"]:checked')?.value || "first";
  }

  const textRules = {};
  document.querySelectorAll(".text-op").forEach((sel) => {
    if (sel.value) textRules[sel.dataset.col] = sel.value;
  });
  if (Object.keys(textRules).length) params.standardize_text = textRules;

  const colorRules = [];
  document.querySelectorAll(".color-rule-row").forEach((row) => {
    colorRules.push({
      column: row.querySelector(".cr-col").value,
      operator: row.querySelector(".cr-op").value,
      value: row.querySelector(".cr-val").value || null,
      color: row.querySelector(".cr-color").value,
    });
  });
  if (colorRules.length) params.color_rules = colorRules;

  const splitCol = document.getElementById("split-col").value;
  if (splitCol) {
    params.split_by_category = splitCol;
    params.keep_original_sheet = document.getElementById("chk-keep-original").checked;
    params.create_summary_sheet = document.getElementById("chk-summary").checked;
  }

  return params;
}

// ─── API calls ───────────────────────────────────────────────────────────────

async function organizeWithAI(instruction, fileBlob) {
  const form = new FormData();
  form.append("file", fileBlob, "planilha.xlsx");
  form.append("parameters", JSON.stringify({ natural_language_instruction: instruction }));
  const res = await fetch(`${API_BASE_URL}/organize`, { method: "POST", body: form });
  if (!res.ok) throw new Error((await res.json()).detail || res.statusText);
  return res.json();
}

async function organizeWithParams(params, fileBlob) {
  const form = new FormData();
  form.append("file", fileBlob, "planilha.xlsx");
  form.append("parameters", JSON.stringify(params));
  const res = await fetch(`${API_BASE_URL}/organize`, { method: "POST", body: form });
  if (!res.ok) throw new Error((await res.json()).detail || res.statusText);
  return res.json();
}

async function downloadResult(token) {
  const res = await fetch(`${API_BASE_URL}/download/${token}`);
  if (!res.ok) throw new Error("Erro ao baixar resultado.");
  return res.blob();
}

async function loadProfiles() {
  const res = await fetch(`${API_BASE_URL}/profiles`);
  if (!res.ok) return;
  const profiles = await res.json();
  renderProfileList(profiles);
  const sel = document.getElementById("profile-select");
  sel.innerHTML = '<option value="">-- Selecionar perfil --</option>';
  profiles.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = p.name;
    sel.appendChild(opt);
  });
}

async function saveProfile(name, params) {
  const res = await fetch(`${API_BASE_URL}/profiles`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, parameters: params }),
  });
  if (!res.ok) throw new Error((await res.json()).detail || "Erro ao salvar perfil.");
  return res.json();
}

async function deleteProfile(id) {
  await fetch(`${API_BASE_URL}/profiles/${id}`, { method: "DELETE" });
}

// ─── Profile list render ─────────────────────────────────────────────────────

function renderProfileList(profiles) {
  const ul = document.getElementById("profile-list");
  ul.innerHTML = "";
  profiles.forEach((p) => {
    const li = document.createElement("li");
    li.innerHTML = `<span>${p.name}</span>
      <div class="actions">
        <button class="btn-secondary" data-id="${p.id}" data-action="delete">Deletar</button>
      </div>`;
    li.querySelector("[data-action='delete']").addEventListener("click", async () => {
      await deleteProfile(p.id);
      await loadProfiles();
      showToast("Perfil deletado.", "success");
    });
    ul.appendChild(li);
  });
}

// ─── Main flow ───────────────────────────────────────────────────────────────

async function runOrganize(getParamsFn) {
  showSpinner(true);
  try {
    const fileBlob = await getSheetDataAsBlob();
    const result = await getParamsFn(fileBlob);
    showResult(result);
    showToast("Planilha organizada!", "success");
  } catch (e) {
    showToast(`Erro: ${e.message}`, "error");
  } finally {
    showSpinner(false);
  }
}

function showResult(result) {
  const section = document.getElementById("section-result");
  const logArea = document.getElementById("log-area");
  section.style.display = "block";
  logArea.innerHTML = [
    `<div class="log-line"><b>Status:</b> ${result.status === "success" ? "Sucesso" : result.status}</div>`,
    `<div class="log-line"><b>Linhas antes:</b> ${result.rows_before} → <b>depois:</b> ${result.rows_after}</div>`,
    ...result.transformations_applied.map(
      (t) => `<div class="log-line">✓ ${t}</div>`
    ),
    result.sheets_created.length
      ? `<div class="log-line"><b>Abas criadas:</b> ${result.sheets_created.join(", ")}</div>`
      : "",
    `<div class="log-line"><b>Token:</b> <a href="${API_BASE_URL}/download/${result.download_token}" target="_blank">Baixar arquivo</a></div>`,
  ].join("");
}

// ─── UI helpers ──────────────────────────────────────────────────────────────

function showSpinner(show) {
  document.getElementById("spinner").style.display = show ? "flex" : "none";
}

function showToast(msg, type = "") {
  const toast = document.getElementById("toast");
  toast.textContent = msg;
  toast.className = `toast ${type}`;
  toast.style.display = "block";
  setTimeout(() => { toast.style.display = "none"; }, 3000);
}

// ─── Event bindings ──────────────────────────────────────────────────────────

function bindEvents() {
  document.getElementById("btn-ai").addEventListener("click", () => {
    const instruction = document.getElementById("ai-instruction").value.trim();
    if (!instruction) { showToast("Digite uma instrução.", "error"); return; }
    runOrganize((blob) => organizeWithAI(instruction, blob));
  });

  document.getElementById("btn-apply").addEventListener("click", () => {
    const params = collectParams();
    runOrganize((blob) => organizeWithParams(params, blob));
  });

  document.getElementById("btn-add-sort").addEventListener("click", addSortRow);
  document.getElementById("btn-add-color").addEventListener("click", addColorRule);

  document.getElementById("chk-duplicates").addEventListener("change", (e) => {
    document.getElementById("dup-cols-wrapper").style.display = e.target.checked ? "block" : "none";
  });

  document.getElementById("btn-save-profile").addEventListener("click", async () => {
    const name = document.getElementById("profile-name").value.trim();
    if (!name) { showToast("Digite um nome para o perfil.", "error"); return; }
    try {
      await saveProfile(name, collectParams());
      await loadProfiles();
      document.getElementById("profile-name").value = "";
      showToast("Perfil salvo!", "success");
    } catch (e) {
      showToast(`Erro: ${e.message}`, "error");
    }
  });

  document.getElementById("btn-apply-profile").addEventListener("click", async () => {
    const id = document.getElementById("profile-select").value;
    if (!id) { showToast("Selecione um perfil.", "error"); return; }
    const res = await fetch(`${API_BASE_URL}/profiles/${id}`);
    const profile = await res.json();
    runOrganize((blob) => organizeWithParams(profile.parameters, blob));
  });
}
