"use strict";

const SHEET_NAME = "base call center";

const COLUMNS = Object.freeze({
  openedAt: "DATA ABERTURA CALL CENTER",
  personName: "NOME PRESTADOR/BENEFICIÁRIO",
  subject: "ASSUNTO CALL CENTER",
  callNumber: "NÚMERO CALL CENTER",
  demandType: "TIPO DE DEMANDA",
  subgroup: "SUBGRUPO",
  receivedAtGrr: "CALL CENTER RECEBIDO CAIXA GRR",
  sentToOwnerAt: "ENVIADO PARA O RESPONSÁVEL DA CARTEIRA NA DATA",
  owner: "RESPONSÁVEL",
  deadline: "PRAZO (DATA MÁXIMA) DE RESPOSTA",
  receivedTiming: "RECEBIDO NO PRAZO OU ATRASADO",
  unreliableDeadline: "PRAZO HOJE",
  status: "STATUS",
  answeredOnTime: "RESPONDIDO NO PRAZO?"
});

const EXPECTED_COLUMNS = Object.values(COLUMNS);
const MAX_SOURCE_COLUMN = 15; // Coluna P; evita a formatação residual até XFC.
const STATUS_VALUES = Object.freeze(["PENDENTE", "EM TRATATIVA", "ENCERRADO"]);
const EMPTY_FILTER_VALUE = "__EMPTY__";

const TABLE_COLUMNS = Object.freeze([
  { key: COLUMNS.callNumber, label: "Número Call Center", cellClass: "cell-id" },
  { key: COLUMNS.personName, label: "Prestador / Beneficiário", cellClass: "cell-text" },
  { key: COLUMNS.subject, label: "Assunto", cellClass: "cell-text cell-subject" },
  { key: COLUMNS.demandType, label: "Tipo de demanda", cellClass: "cell-text" },
  { key: COLUMNS.subgroup, label: "Subgrupo", cellClass: "cell-text" },
  { key: COLUMNS.owner, label: "Responsável", cellClass: "cell-text" },
  { key: COLUMNS.status, label: "Status", cellClass: "cell-status" },
  { key: COLUMNS.openedAt, label: "Data de abertura", type: "date", cellClass: "cell-date" },
  { key: COLUMNS.sentToOwnerAt, label: "Enviado ao responsável", type: "date", cellClass: "cell-date" },
  { key: COLUMNS.deadline, label: "Prazo máximo", type: "date", cellClass: "cell-date" }
]);

const CATEGORY_CONFIG = Object.freeze([
  {
    key: "receivedToday",
    tone: "blue",
    countElementId: "count-received-today",
    titleToday: "Chamados recebidos hoje",
    titlePeriod: "Chamados recebidos no período"
  },
  {
    key: "receivedLate",
    title: "Chamados recebidos em atraso",
    tone: "red",
    countElementId: "count-received-late"
  },
  {
    key: "open",
    title: "Chamados em aberto",
    tone: "yellow",
    countElementId: "count-open"
  },
  {
    key: "dueToday",
    title: "Chamados vencendo hoje",
    tone: "orange",
    countElementId: "count-due-today"
  },
  {
    key: "overdue",
    title: "Chamados fora do prazo",
    tone: "red",
    countElementId: "count-overdue"
  }
]);

const appState = {
  records: [],
  fileName: "",
  loadedAt: null
};

const uploadView = document.querySelector("#upload-view");
const dashboardView = document.querySelector("#dashboard-view");
const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const resetButton = document.querySelector("#reset-button");
const errorMessage = document.querySelector("#error-message");
const loadingMessage = document.querySelector("#loading-message");
const loadedAt = document.querySelector("#loaded-at");
const detailsContainer = document.querySelector("#details-container");
const invalidDatesMessage = document.querySelector("#invalid-dates-message");
const startDateInput = document.querySelector("#filter-start-date");
const endDateInput = document.querySelector("#filter-end-date");
const todayButton = document.querySelector("#filter-today-button");
const subgroupSelect = document.querySelector("#filter-subgroup");
const ownerSelect = document.querySelector("#filter-owner");
const statusFilters = document.querySelectorAll(".status-filter");
const filterEmptyMessage = document.querySelector("#filter-empty-message");
const receivedLabel = document.querySelector("#received-label");
const receivedCaption = document.querySelector("#received-caption");

fileInput.addEventListener("change", () => {
  const [file] = fileInput.files;
  if (file) {
    processFile(file);
  }
});

dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

for (const eventName of ["dragenter", "dragover"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("is-dragging");
  });
}

dropZone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (file) {
    processFile(file);
  }
});

document.addEventListener("dragover", (event) => event.preventDefault());
document.addEventListener("drop", (event) => event.preventDefault());
resetButton.addEventListener("click", resetApplication);
document.querySelector("#filter-bar").addEventListener("submit", (event) => {
  event.preventDefault();
});
todayButton.addEventListener("click", () => {
  setTodayRange();
  applyFilters();
});
startDateInput.addEventListener("change", applyFilters);
endDateInput.addEventListener("change", applyFilters);
subgroupSelect.addEventListener("change", () => {
  populateOwnerOptions(ownerSelect.value);
  applyFilters();
});
ownerSelect.addEventListener("change", applyFilters);

for (const checkbox of statusFilters) {
  checkbox.addEventListener("change", applyFilters);
}

async function processFile(file) {
  clearError();

  if (!/\.(xlsx|xls|ods)$/i.test(file.name)) {
    showError("Formato não suportado. Selecione um arquivo .xlsx, .xls ou .ods.");
    return;
  }

  if (typeof XLSX === "undefined") {
    showError(
      "Não foi possível carregar a biblioteca de leitura. Verifique a conexão e recarregue a página."
    );
    return;
  }

  setLoading(true);

  try {
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const fileBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(fileBuffer, {
      type: "array",
      cellDates: true,
      cellNF: true
    });
    const worksheet = workbook.Sheets[SHEET_NAME];

    if (!worksheet) {
      throw new Error(`A aba “${SHEET_NAME}” não foi encontrada na planilha.`);
    }

    const headerMap = readAndValidateHeaders(worksheet);
    const date1904 = Boolean(workbook.Workbook?.WBProps?.date1904);

    appState.records = enrichRecords(readRecords(worksheet, headerMap), date1904);
    appState.fileName = file.name;
    appState.loadedAt = new Date();

    populateSubgroupOptions();
    populateOwnerOptions();
    setTodayRange();
    renderDashboard(calculateIndicators(appState.records, getActiveFilters()), {
      isInitial: true
    });
  } catch (error) {
    console.error("Erro ao processar a planilha:", error);
    showError(error instanceof Error ? error.message : "Não foi possível processar o arquivo.");
  } finally {
    setLoading(false);
    fileInput.value = "";
  }
}

function readAndValidateHeaders(worksheet) {
  if (!worksheet["!ref"]) {
    throw new Error(`A aba “${SHEET_NAME}” está vazia.`);
  }

  const sheetRange = XLSX.utils.decode_range(worksheet["!ref"]);
  const lastColumn = Math.min(sheetRange.e.c, MAX_SOURCE_COLUMN);
  const headerMap = new Map();

  for (let columnIndex = 0; columnIndex <= lastColumn; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: 0, c: columnIndex });
    const value = worksheet[address]?.v;
    const normalizedHeader = value == null ? "" : String(value).trim();

    if (normalizedHeader) {
      headerMap.set(normalizedHeader, columnIndex);
    }
  }

  const missingColumns = EXPECTED_COLUMNS.filter((column) => !headerMap.has(column));

  if (missingColumns.length > 0) {
    const formattedList = missingColumns.map((column) => `“${column}”`).join(", ");
    throw new Error(
      `A planilha não contém ${missingColumns.length === 1 ? "a coluna esperada" : "as colunas esperadas"}: ${formattedList}.`
    );
  }

  return headerMap;
}

function readRecords(worksheet, headerMap) {
  const sheetRange = XLSX.utils.decode_range(worksheet["!ref"]);
  const records = [];

  for (let rowIndex = 1; rowIndex <= sheetRange.e.r; rowIndex += 1) {
    const record = { __rowNumber: rowIndex + 1 };
    let hasData = false;

    for (const columnName of EXPECTED_COLUMNS) {
      const columnIndex = headerMap.get(columnName);
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = worksheet[address];
      let value = cell?.v ?? null;

      if (
        columnName === COLUMNS.callNumber &&
        cell?.w != null &&
        typeof value === "number"
      ) {
        value = cell.w;
      }

      record[columnName] = value;

      if (value !== null && String(value).trim() !== "") {
        hasData = true;
      }
    }

    if (hasData) {
      records.push(record);
    }
  }

  return records;
}

function enrichRecords(records, date1904 = false) {
  return records.map((record) => ({
    ...record,
    __parsedDates: {
      [COLUMNS.openedAt]: parseSpreadsheetDate(record[COLUMNS.openedAt], date1904),
      [COLUMNS.sentToOwnerAt]: parseSpreadsheetDate(record[COLUMNS.sentToOwnerAt], date1904),
      [COLUMNS.deadline]: parseSpreadsheetDate(record[COLUMNS.deadline], date1904)
    }
  }));
}

function getActiveFilters(now = new Date()) {
  const today = toLocalDateOnly(now);
  let startDate = parseDateInput(startDateInput.value) ?? today;
  let endDate = parseDateInput(endDateInput.value) ?? today;

  if (startDate.getTime() > endDate.getTime()) {
    [startDate, endDate] = [endDate, startDate];
    startDateInput.value = toInputDateValue(startDate);
    endDateInput.value = toInputDateValue(endDate);
  }

  const statuses = new Set(
    [...statusFilters]
      .filter((checkbox) => checkbox.checked)
      .map((checkbox) => checkbox.value)
  );

  return {
    startDate,
    endDate,
    statuses,
    today,
    subgroup: subgroupSelect.value,
    owner: ownerSelect.value
  };
}

function applyFilters() {
  if (appState.records.length === 0) {
    return;
  }

  renderDashboard(calculateIndicators(appState.records, getActiveFilters()));
}

function calculateIndicators(records, filters) {
  const { startDate, endDate, statuses, today, subgroup, owner } = filters;
  const periodIncludesToday =
    today.getTime() >= startDate.getTime() && today.getTime() <= endDate.getTime();
  const overdueReference = periodIncludesToday ? today : endDate;
  const isTodayRange =
    startDate.getTime() === today.getTime() && endDate.getTime() === today.getTime();
  const invalidRows = new Set();
  const categories = {
    receivedToday: [],
    receivedLate: [],
    open: [],
    dueToday: [],
    overdue: []
  };

  for (const record of records) {
    const sentAt = record.__parsedDates[COLUMNS.sentToOwnerAt];
    const deadline = record.__parsedDates[COLUMNS.deadline];

    if (!sentAt) {
      invalidRows.add(record.__rowNumber);
      continue;
    }

    if (sentAt.getTime() < startDate.getTime() || sentAt.getTime() > endDate.getTime()) {
      continue;
    }

    if (!deadline) {
      invalidRows.add(record.__rowNumber);
    }

    if (statuses.size === 0) {
      continue;
    }

    const status = normalizeText(record[COLUMNS.status]);

    if (!statuses.has(status)) {
      continue;
    }

    if (!matchesSelectFilter(record[COLUMNS.subgroup], subgroup)) {
      continue;
    }

    if (!matchesSelectFilter(record[COLUMNS.owner], owner)) {
      continue;
    }

    categories.receivedToday.push(record);

    if (normalizeText(record[COLUMNS.receivedTiming]) === "EM ATRASO") {
      categories.receivedLate.push(record);
    }

    if (status === "PENDENTE" || status === "EM TRATATIVA") {
      categories.open.push(record);
    }

    if (normalizeText(record[COLUMNS.unreliableDeadline]) === "VENCE HOJE") {
      categories.dueToday.push(record);
    }

    if (
      deadline &&
      deadline.getTime() < overdueReference.getTime() &&
      status !== "ENCERRADO"
    ) {
      categories.overdue.push(record);
    }
  }

  return {
    categories,
    invalidDateRows: invalidRows.size,
    isTodayRange,
    noStatusSelected: statuses.size === 0,
    startDate,
    endDate,
    subgroup,
    owner
  };
}

function parseSpreadsheetDate(value, date1904 = false) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return null;
    }

    return createValidatedDate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate()
    );
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value, { date1904 });
    return parsed ? createValidatedDate(parsed.y, parsed.m, parsed.d) : null;
  }

  if (typeof value !== "string") {
    return null;
  }

  const text = value.trim();

  if (!text) {
    return null;
  }

  const brazilianDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s|$)/);
  if (brazilianDate) {
    return createValidatedDate(
      Number(brazilianDate[3]),
      Number(brazilianDate[2]),
      Number(brazilianDate[1])
    );
  }

  const isoDate = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T|\s|$)/);
  if (isoDate) {
    return createValidatedDate(
      Number(isoDate[1]),
      Number(isoDate[2]),
      Number(isoDate[3])
    );
  }

  if (/^\d+(?:[.,]\d+)?$/.test(text)) {
    const serial = Number(text.replace(",", "."));
    if (Number.isFinite(serial)) {
      const parsed = XLSX.SSF.parse_date_code(serial, { date1904 });
      return parsed ? createValidatedDate(parsed.y, parsed.m, parsed.d) : null;
    }
  }

  return null;
}

function createValidatedDate(year, month, day) {
  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}

function toLocalDateOnly(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function toInputDateValue(date) {
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDateInput(value) {
  if (!value) {
    return null;
  }

  const [year, month, day] = value.split("-").map(Number);
  return createValidatedDate(year, month, day);
}

function setTodayRange() {
  const todayValue = toInputDateValue(toLocalDateOnly(new Date()));
  startDateInput.value = todayValue;
  endDateInput.value = todayValue;
}

function normalizeText(value) {
  return value == null ? "" : String(value).trim().toLocaleUpperCase("pt-BR");
}

function displayText(value) {
  return value == null ? "" : String(value).trim();
}

function matchesSelectFilter(rawValue, selectedValue) {
  if (!selectedValue) {
    return true;
  }

  const text = displayText(rawValue);

  if (selectedValue === EMPTY_FILTER_VALUE) {
    return text === "";
  }

  return text === selectedValue;
}

function uniqueColumnValues(records, columnName) {
  const unique = new Map();

  for (const record of records) {
    const text = displayText(record[columnName]);
    const key = text === "" ? EMPTY_FILTER_VALUE : text;

    if (!unique.has(key)) {
      unique.set(key, text);
    }
  }

  return [...unique.entries()].sort((left, right) => {
    if (left[0] === EMPTY_FILTER_VALUE) {
      return 1;
    }

    if (right[0] === EMPTY_FILTER_VALUE) {
      return -1;
    }

    return left[1].localeCompare(right[1], "pt-BR", { sensitivity: "base" });
  });
}

function replaceSelectOptions(select, values, allLabel, emptyLabel) {
  const previousValue = select.value;
  const fragment = document.createDocumentFragment();
  const allOption = document.createElement("option");
  allOption.value = "";
  allOption.textContent = allLabel;
  fragment.append(allOption);

  for (const [value, label] of values) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value === EMPTY_FILTER_VALUE ? emptyLabel : label;
    fragment.append(option);
  }

  select.replaceChildren(fragment);
  select.value = [...select.options].some((option) => option.value === previousValue)
    ? previousValue
    : "";
}

function populateSubgroupOptions() {
  replaceSelectOptions(
    subgroupSelect,
    uniqueColumnValues(appState.records, COLUMNS.subgroup),
    "Todos os setores",
    "Sem setor"
  );
}

function populateOwnerOptions(preferredOwner = "") {
  const selectedSubgroup = subgroupSelect.value;
  const scopedRecords = selectedSubgroup
    ? appState.records.filter((record) =>
        matchesSelectFilter(record[COLUMNS.subgroup], selectedSubgroup)
      )
    : appState.records;

  replaceSelectOptions(
    ownerSelect,
    uniqueColumnValues(scopedRecords, COLUMNS.owner),
    "Todos os responsáveis",
    "Sem responsável"
  );

  if (
    preferredOwner &&
    [...ownerSelect.options].some((option) => option.value === preferredOwner)
  ) {
    ownerSelect.value = preferredOwner;
  }
}

function getCategoryTitle(category, isTodayRange) {
  if (category.key === "receivedToday") {
    return isTodayRange ? category.titleToday : category.titlePeriod;
  }

  return category.title;
}

function renderDashboard(result, options = {}) {
  const { isInitial = false } = options;

  receivedLabel.textContent = result.isTodayRange ? "Recebidos hoje" : "Recebidos no período";
  receivedCaption.textContent = result.isTodayRange
    ? "Enviados ao responsável na data atual"
    : "Enviados ao responsável no período selecionado";
  filterEmptyMessage.hidden = !result.noStatusSelected;

  for (const category of CATEGORY_CONFIG) {
    const records = result.categories[category.key];
    document.querySelector(`#${category.countElementId}`).textContent =
      records.length.toLocaleString("pt-BR");
  }

  loadedAt.textContent = `Planilha carregada em ${appState.loadedAt.toLocaleString("pt-BR")}`;
  invalidDatesMessage.textContent =
    result.invalidDateRows === 0
      ? "Nenhuma linha com data inválida foi ignorada."
      : `${result.invalidDateRows.toLocaleString("pt-BR")} ${
          result.invalidDateRows === 1
            ? "linha com data inválida foi ignorada"
            : "linhas com datas inválidas foram ignoradas"
        } nos indicadores de data.`;

  detailsContainer.replaceChildren(
    ...CATEGORY_CONFIG.map((category) =>
      createDetailSection(
        {
          ...category,
          title: getCategoryTitle(category, result.isTodayRange)
        },
        result.categories[category.key],
        result.noStatusSelected
      )
    )
  );

  console.log("Dashboard de chamados processado:", {
    arquivo: appState.fileName,
    aba: SHEET_NAME,
    linhasProcessadas: appState.records.length,
    dataInicial: toInputDateValue(result.startDate),
    dataFinal: toInputDateValue(result.endDate),
    statusSelecionados: [...statusFilters]
      .filter((checkbox) => checkbox.checked)
      .map((checkbox) => checkbox.value),
    setor: result.subgroup || "Todos os setores",
    responsavel: result.owner || "Todos os responsáveis",
    recebidosNoPeriodo: result.categories.receivedToday.length,
    recebidosEmAtraso: result.categories.receivedLate.length,
    emAberto: result.categories.open.length,
    vencendoHoje: result.categories.dueToday.length,
    foraDoPrazo: result.categories.overdue.length,
    linhasComDatasInvalidas: result.invalidDateRows
  });

  uploadView.hidden = true;
  dashboardView.hidden = false;
  resetButton.hidden = false;

  if (isInitial) {
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
}

function createDetailSection(category, records, noStatusSelected = false) {
  const details = document.createElement("details");
  details.className = "detail-section";
  details.dataset.tone = category.tone;

  const summary = document.createElement("summary");
  const title = document.createElement("span");
  title.className = "summary-title";

  const dot = document.createElement("span");
  dot.className = "summary-dot";
  dot.setAttribute("aria-hidden", "true");

  const titleText = document.createElement("span");
  titleText.textContent = category.title;
  title.append(dot, titleText);

  const meta = document.createElement("span");
  meta.className = "summary-meta";

  const recordCount = document.createElement("span");
  recordCount.textContent = `${records.length.toLocaleString("pt-BR")} ${
    records.length === 1 ? "chamado" : "chamados"
  }`;

  const chevron = document.createElement("span");
  chevron.className = "chevron";
  chevron.setAttribute("aria-hidden", "true");
  meta.append(recordCount, chevron);
  summary.append(title, meta);
  details.append(summary);

  if (records.length === 0) {
    const emptyMessage = document.createElement("p");
    emptyMessage.className = "empty-table";
    emptyMessage.textContent = noStatusSelected
      ? "Nenhum status selecionado. Marque ao menos um status para ver os chamados."
      : "Nenhum chamado nesta categoria com os filtros atuais.";
    details.append(emptyMessage);
    return details;
  }

  const wrapper = document.createElement("div");
  wrapper.className = "table-wrapper";
  wrapper.setAttribute("tabindex", "0");
  wrapper.setAttribute("aria-label", `Tabela: ${category.title}`);

  const table = document.createElement("table");
  table.className = "data-table";

  const tableHead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  for (const column of TABLE_COLUMNS) {
    const header = document.createElement("th");
    header.scope = "col";
    header.className = column.cellClass;
    header.textContent = column.label;
    headerRow.append(header);
  }
  tableHead.append(headerRow);

  const tableBody = document.createElement("tbody");
  const rowsFragment = document.createDocumentFragment();

  for (const record of records) {
    const row = document.createElement("tr");

    for (const column of TABLE_COLUMNS) {
      const cell = document.createElement("td");
      const rawValue = record[column.key];
      cell.className = column.cellClass;

      if (column.type === "date") {
        const parsedDate = record.__parsedDates[column.key];
        cell.textContent = parsedDate ? parsedDate.toLocaleDateString("pt-BR") : "—";
      } else {
        cell.textContent = rawValue == null || String(rawValue).trim() === ""
          ? "—"
          : String(rawValue).trim();
      }

      row.append(cell);
    }

    rowsFragment.append(row);
  }

  tableBody.append(rowsFragment);
  table.append(tableHead, tableBody);
  wrapper.append(table);
  details.append(wrapper);

  return details;
}

function resetApplication() {
  appState.records = [];
  appState.fileName = "";
  appState.loadedAt = null;
  fileInput.value = "";
  detailsContainer.replaceChildren();
  dashboardView.hidden = true;
  resetButton.hidden = true;
  uploadView.hidden = false;
  filterEmptyMessage.hidden = true;

  for (const checkbox of statusFilters) {
    checkbox.checked = STATUS_VALUES.includes(checkbox.value);
  }

  replaceSelectOptions(subgroupSelect, [], "Todos os setores", "Sem setor");
  replaceSelectOptions(ownerSelect, [], "Todos os responsáveis", "Sem responsável");
  startDateInput.value = "";
  endDateInput.value = "";
  clearError();
  setLoading(false);
  dropZone.focus();
}

function setLoading(isLoading) {
  loadingMessage.hidden = !isLoading;
  fileInput.disabled = isLoading;
  dropZone.setAttribute("aria-busy", String(isLoading));
}

function showError(message) {
  errorMessage.textContent = message;
  errorMessage.hidden = false;
}

function clearError() {
  errorMessage.textContent = "";
  errorMessage.hidden = true;
}
