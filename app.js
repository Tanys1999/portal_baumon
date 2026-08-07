// ---- CONFIGURAÇÃO ----
// Caminho do Excel dentro do próprio repositório (relativo ao index.html).
// Troque o nome do arquivo aqui se o seu Excel tiver outro nome.
const CONFIG = {
  dataUrl: "data/MRP - Baumon.xlsx",
  pollIntervalMs: 15000 // verifica se o arquivo mudou a cada 15s
};

const SHEETS = ["Base de Dados", "Precificação"];

const state = {
  data: {},      // { sheetName: { headers: [...], rows: [...] } }
  active: null,
  filters: {},   // { sheetName: searchTerm }
  lastHash: null
};

const fileInput = document.getElementById("fileInput");
const uploadStatus = document.getElementById("uploadStatus");
const syncStatus = document.getElementById("syncStatus");
const refreshBtn = document.getElementById("refreshBtn");
const tabsEl = document.getElementById("tabs");
const toolbarEl = document.getElementById("toolbar");
const tableWrapEl = document.getElementById("tableWrap");

// Upload manual continua funcionando (útil pra testar antes de subir no GitHub,
// ou para abrir um arquivo pontual sem esperar o deploy).
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  file.arrayBuffer().then((buf) => {
    processWorkbook(new Uint8Array(buf), "Arquivo local: " + file.name, true);
  });
  fileInput.value = ""; // permite reselecionar o mesmo arquivo depois
});

refreshBtn.addEventListener("click", () => fetchFromServer(true));

async function fetchFromServer(forced) {
  try {
    const res = await fetch(CONFIG.dataUrl + "?t=" + Date.now(), { cache: "no-store" });
    if (!res.ok) {
      setSyncStatus(
        'Aguardando arquivo em <code>' + CONFIG.dataUrl + "</code> (status " + res.status + ")",
        true
      );
      return;
    }
    const buf = new Uint8Array(await res.arrayBuffer());
    const hash = await hashBytes(buf);

    if (!forced && hash === state.lastHash) {
      setSyncStatus("Sincronizado \u2014 sem alterações desde a última leitura (" + nowTime() + ")");
      return;
    }

    state.lastHash = hash;
    processWorkbook(buf, "data/" + CONFIG.dataUrl.split("/").pop(), false);
  } catch (err) {
    setSyncStatus("Não foi possível buscar o arquivo: " + escapeHtml(err.message), true);
  }
}

function processWorkbook(buf, sourceLabel, isManual) {
  try {
    const workbook = XLSX.read(buf, { type: "array", cellDates: true });
    const found = SHEETS.filter((name) => workbook.SheetNames.includes(name));

    if (found.length === 0) {
      setSyncStatus(
        "Nenhuma das abas esperadas (" + SHEETS.join(", ") + ") foi encontrada neste arquivo.",
        true
      );
      return;
    }

    const previousActive = state.active;
    state.data = {};
    found.forEach((name) => {
      const sheet = workbook.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
      const headers = (rows[0] || []).map((h) => String(h || "").trim());
      const body = rows.slice(1).filter((r) => r.some((c) => c !== ""));
      state.data[name] = { headers, rows: body };
      if (!(name in state.filters)) state.filters[name] = "";
    });

    state.active = found.includes(previousActive) ? previousActive : found[0];

    uploadStatus.innerHTML = 'Fonte atual: <span class="filename">' + escapeHtml(sourceLabel) + "</span>";
    setSyncStatus(
      (isManual ? "Carregado manualmente" : "Sincronizado automaticamente") + " \u2014 " + nowTime()
    );

    renderTabs();
    renderToolbar();
    renderTable();
  } catch (err) {
    setSyncStatus("Não foi possível ler o arquivo: " + escapeHtml(err.message), true);
  }
}

function setSyncStatus(html, isError) {
  syncStatus.className = "upload-status" + (isError ? " error" : "");
  syncStatus.innerHTML = html;
}

function nowTime() {
  return new Date().toLocaleTimeString("pt-BR");
}

async function hashBytes(buf) {
  const digest = await crypto.subtle.digest("SHA-256", buf);
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

// Início: tenta buscar o arquivo hospedado e passa a verificar periodicamente.
fetchFromServer(true);
setInterval(() => fetchFromServer(false), CONFIG.pollIntervalMs);

function renderTabs() {
  const names = Object.keys(state.data);
  tabsEl.innerHTML = names
    .map((name) => {
      const count = state.data[name].rows.length;
      const active = name === state.active ? "active" : "";
      return (
        '<button class="tab-btn ' +
        active +
        '" data-sheet="' +
        escapeHtml(name) +
        '">' +
        escapeHtml(name) +
        '<span class="count">' +
        count +
        "</span></button>"
      );
    })
    .join("");

  tabsEl.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.active = btn.dataset.sheet;
      renderTabs();
      renderToolbar();
      renderTable();
    });
  });
}

function renderToolbar() {
  if (!state.active) {
    toolbarEl.innerHTML = "";
    return;
  }
  const term = state.filters[state.active] || "";
  toolbarEl.innerHTML =
    '<input type="text" class="search-input" id="searchBox" placeholder="Filtrar linhas em ' +
    escapeHtml(state.active) +
    '..." value="' +
    escapeHtml(term) +
    '" />' +
    '<span class="row-count" id="rowCount"></span>';

  const box = document.getElementById("searchBox");
  box.addEventListener("input", () => {
    state.filters[state.active] = box.value;
    renderTable();
  });
}

function renderTable() {
  if (!state.active) {
    tableWrapEl.innerHTML =
      '<div class="empty-state"><i class="ti ti-table-off" aria-hidden="true"></i>Selecione um arquivo Excel para carregar os dados das abas "' +
      SHEETS.join('" e "') +
      '".</div>';
    return;
  }

  const { headers, rows } = state.data[state.active];
  const term = (state.filters[state.active] || "").trim().toLowerCase();

  const filtered = term
    ? rows.filter((row) => row.some((cell) => String(cell).toLowerCase().includes(term)))
    : rows;

  const rowCountEl = document.getElementById("rowCount");
  if (rowCountEl) {
    rowCountEl.textContent = filtered.length + " de " + rows.length + " linhas";
  }

  if (filtered.length === 0) {
    tableWrapEl.innerHTML = '<div class="empty-state"><i class="ti ti-search-off" aria-hidden="true"></i>Nenhuma linha corresponde ao filtro.</div>';
    return;
  }

  let html = "<table><thead><tr>";
  headers.forEach((h) => {
    html += "<th>" + escapeHtml(h || "\u2014") + "</th>";
  });
  html += "</tr></thead><tbody>";

  filtered.forEach((row) => {
    html += "<tr>";
    headers.forEach((_, i) => {
      html += "<td>" + escapeHtml(formatCell(row[i])) + "</td>";
    });
    html += "</tr>";
  });
  html += "</tbody></table>";

  tableWrapEl.innerHTML = html;
}

function formatCell(value) {
  if (value === undefined || value === null || value === "") return "";
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toLocaleString("pt-BR", { maximumFractionDigits: 4 });
  }
  return String(value);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

renderTable();
