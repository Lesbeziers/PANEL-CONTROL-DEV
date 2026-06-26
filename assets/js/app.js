
const columns = [
  { key: "listo", label: "LISTO", type: "checkbox" },
  { key: "title", label: "TÍTULO", type: "text" },
  { key: "startDate", label: "INICIO VIG", type: "text" },
  { key: "endDate", label: "FIN VIG", type: "text" },
  {
    key: "genre",
    label: "GÉNERO",
    type: "text",
    cellType: "select",
    options: ["Caza y Pesca", "Cine Compra", "Cine Original", "Deportes", "Documentales", "Entretenimiento", "M+", "No Ficción", "Series Compra", "Series Originales"],
  },
  { key: "id", label: "ID", type: "text" },
];

const headers = columns.map((column) => column.label);
const MONTH_NAMES_ES = [
  "enero",
  "febrero",
  "marzo",
  "abril",
  "mayo",
  "junio",
  "julio",
  "agosto",
  "septiembre",
  "octubre",
  "noviembre",
  "diciembre",
];
const WEEKDAY_INITIALS_ES = ["L", "M", "X", "J", "V", "S", "D"];
const now = new Date();
const DEFAULT_CALENDAR_CONTEXT = {
  month: now.getMonth() + 1,
  year: now.getFullYear(),
  daysInMonth: new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate(),
};
const MONTH_LABEL_TO_NUMBER = {
  enero: 1,
  febrero: 2,
  marzo: 3,
  abril: 4,
  mayo: 5,
  junio: 6,
  julio: 7,
  agosto: 8,
  septiembre: 9,
  setiembre: 9,
  octubre: 10,
  noviembre: 11,
  diciembre: 12,
};
const DATE_COLUMNS = new Set(["startDate", "endDate"]);
const DEFAULT_MAX_SIMULTANEOUS = 5;
const GLOBAL_COLLAPSE_BUTTON_ID = "global-collapse-toggle";
const ENABLE_EXCEL_IMPORT = window.PANEL_FEATURES?.excelImportV2 !== false;
const IS_VIEWER_MODE = window.PANEL_FEATURES?.viewerMode === true;
const EXCEL_BLOCK_HEADER_CANDIDATES = ["BLOQUE", "TIPO BLOQUE", "TIPO", "FORMATO"];

// ── Red de seguridad contra pérdida de datos ─────────────────────────────────
// El token de Google caduca a la hora; si el guardado en Drive se interrumpe al
// re-loguear, el trabajo en memoria podría perderse. Para evitarlo:
//  1) Autoguardamos el estado en localStorage tras cada cambio (debounce).
//  2) Al cargar, si hay un borrador local sin sincronizar, ofrecemos restaurarlo.
//  3) La carga desde Drive nunca machaca cambios locales sin guardar.
const DRAFT_STORAGE_KEY = `panelControlDraft:${window.PANEL_CONFIG?.GOOGLE_DRIVE_FILE_ID || "default"}`;
const DRAFT_AUTOSAVE_DELAY_MS = 1500;
let hasUnsavedChanges = false;
let draftAutosaveTimer = null;
let initialDriveLoadDone = false;

function hasLocalStorage() {
  try {
    return typeof window !== "undefined" && !!window.localStorage;
  } catch (_) {
    return false;
  }
}

function writeDraftNow() {
  if (IS_VIEWER_MODE || !hasLocalStorage()) {
    return;
  }
  // Skip autosave while the user is mid-edit in a cell — the JSON.stringify
  // of the full blocks structure can stall the main thread for tens of ms on
  // older machines, which manifests as input lag. The autosave will re-fire
  // shortly after editing ends via the next markDirty().
  if (editingCell) {
    scheduleDraftAutosave();
    return;
  }
  // Defer the heavy serialise+write off the input-handling stack so any
  // pending paint/scroll finishes first.
  const runWrite = () => {
    try {
      window.localStorage.setItem(
        DRAFT_STORAGE_KEY,
        JSON.stringify({ savedAt: Date.now(), blocks })
      );
    } catch (err) {
      console.error("[draft] no se pudo autoguardar en local:", err);
    }
  };
  if (typeof window.requestIdleCallback === "function") {
    window.requestIdleCallback(runWrite, { timeout: 2000 });
  } else {
    setTimeout(runWrite, 0);
  }
}

function scheduleDraftAutosave() {
  if (IS_VIEWER_MODE || !hasLocalStorage()) {
    return;
  }
  if (draftAutosaveTimer) {
    clearTimeout(draftAutosaveTimer);
  }
  draftAutosaveTimer = setTimeout(writeDraftNow, DRAFT_AUTOSAVE_DELAY_MS);
}

// Marca que hay cambios sin guardar en Drive y programa el autoguardado local.
function markDirty() {
  if (IS_VIEWER_MODE) {
    return;
  }
  hasUnsavedChanges = true;
  scheduleDraftAutosave();
}

function clearDraft() {
  hasUnsavedChanges = false;
  if (draftAutosaveTimer) {
    clearTimeout(draftAutosaveTimer);
    draftAutosaveTimer = null;
  }
  if (!hasLocalStorage()) {
    return;
  }
  try {
    window.localStorage.removeItem(DRAFT_STORAGE_KEY);
  } catch (_) {
    // ignore
  }
}

function readDraft() {
  if (!hasLocalStorage()) {
    return null;
  }
  try {
    const raw = window.localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!raw) {
      return null;
    }
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.blocks) || !parsed.blocks.length) {
      return null;
    }
    return parsed;
  } catch (_) {
    return null;
  }
}

const EXCEL_COLUMN_ALIASES = {
  listo: ["LISTO", "READY", "OK"],
  title: ["TÍTULO", "TITULO", "TITLE", "NOMBRE"],
  startDate: ["INICIO VIG", "INICIO VIG.", "INICIO VIGENCIA", "INICIO", "FECHA INICIO", "START", "START DATE"],
  endDate: ["FIN VIG", "FIN VIG.", "FIN VIGENCIA", "FIN", "FECHA FIN", "END", "END DATE"],
  genre: ["GÉNERO", "GENERO", "GÉNERO/PROGRAMA", "GENRE"],
  id: ["ID", "CÓDIGO", "CODIGO", "IDENTIFICADOR"],
  actualizado: ["ACTUALIZADO", "UPDATED"],
  rowKey: ["ROW_KEY", "ROWKEY"],
};

// Random session prefix so that rowKeys generated by two different browser
// sessions cannot collide. Critical for the merge-on-save flow: each row's
// rowKey is its persistent identity across saves and reloads.
const ROW_KEY_SESSION_PREFIX = `s${Math.random().toString(36).slice(2, 10)}`;
let rowId = 0;

function newRow() {
  rowId += 1;
  return {
    rowKey: `${ROW_KEY_SESSION_PREFIX}-${Date.now()}-${rowId}`,
    _autoPlaceholder: false,
    id: "",
    blockType: "",
    listoByMonth: {},
    actualizado: false,
    title: "",
    genre: "",
    startDateText: "",
    startDateISO: null,
    startDateError: null,
    endDateText: "",
    endDateISO: null,
    endDateError: null,
    dateRangeError: null,
    homeMonth: DEFAULT_CALENDAR_CONTEXT.month,
    homeYear: DEFAULT_CALENDAR_CONTEXT.year,
  };
}

function newRowForBlock(blockType, homeContext = DEFAULT_CALENDAR_CONTEXT) {
  const row = newRow();
  row.blockType = blockType;
  if (homeContext && Number.isInteger(homeContext.month) && Number.isInteger(homeContext.year)) {
    row.homeMonth = homeContext.month;
    row.homeYear = homeContext.year;
  }
  return row;
}

// ── Estado de emisión "LISTO" por mes ────────────────────────────────────────
// El check no es una propiedad de la pieza, sino un hecho mensual: indica si la
// pieza se ha emitido en un mes concreto. Por eso se guarda en un mapa
// row.listoByMonth = { "YYYY-MM": true, ... } y NO en un único booleano.
// En el Excel se persiste en la propia columna LISTO como la lista de meses
// emitidos (p. ej. "2026-06;2026-07"), sin columnas nuevas ni filas duplicadas.
function monthKeyOf(month, year) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function monthKeyFor(ctx = currentCalendarContext) {
  return monthKeyOf(ctx.month, ctx.year);
}

function ctxFromMonthKey(key) {
  const match = /^(\d{4})-(\d{2})$/.exec(`${key ?? ""}`);
  if (!match) {
    return currentCalendarContext;
  }
  return { year: Number.parseInt(match[1], 10), month: Number.parseInt(match[2], 10) };
}

function getRowListo(row, ctx = currentCalendarContext) {
  const map = row?.listoByMonth;
  if (!map || typeof map !== "object") {
    return false;
  }
  return !!map[monthKeyFor(ctx)];
}

function setRowListo(row, value, ctx = currentCalendarContext) {
  if (!row) {
    return;
  }
  if (!row.listoByMonth || typeof row.listoByMonth !== "object") {
    row.listoByMonth = {};
  }
  const key = monthKeyFor(ctx);
  if (value) {
    row.listoByMonth[key] = true;
  } else {
    delete row.listoByMonth[key];
  }
}

function rowHasAnyListo(row) {
  const map = row?.listoByMonth;
  if (!map || typeof map !== "object") {
    return false;
  }
  return Object.keys(map).some((key) => map[key]);
}

// Serializa el mapa de meses emitidos para guardarlo en la celda LISTO del Excel.
function encodeListoByMonth(map) {
  if (!map || typeof map !== "object") {
    return "";
  }
  return Object.keys(map)
    .filter((key) => map[key])
    .sort()
    .join(";");
}

function isLegacyListoTruthy(rawValue) {
  if (rawValue === true) {
    return true;
  }
  const text = `${rawValue ?? ""}`.trim().toLowerCase();
  return ["true", "1", "x", "si", "sí", "verdadero"].includes(text);
}

// Reconstruye row.listoByMonth desde el valor crudo de la celda LISTO del Excel.
// Debe llamarse DESPUÉS de haber parseado las fechas de la fila (necesita el rango).
//  - Formato nuevo: lista de meses emitidos "YYYY-MM;YYYY-MM" → esos meses.
//  - Formato antiguo (booleano/texto "verdadero"): como antes el check se
//    compartía entre todos los meses de la pieza, se marca emitido en todo su
//    rango para preservar el estado visible previo (sin perder datos).
function applyImportedListo(row, rawValue) {
  const map = {};
  const tokens = `${rawValue ?? ""}`.match(/\d{4}-\d{2}/g);
  if (tokens && tokens.length) {
    tokens.forEach((token) => { map[token] = true; });
    row.listoByMonth = map;
    return;
  }

  if (isLegacyListoTruthy(rawValue)) {
    const range = getRowRange(row);
    if (range) {
      let year = range.startDate.getFullYear();
      let month = range.startDate.getMonth() + 1;
      const endYear = range.endDate.getFullYear();
      const endMonth = range.endDate.getMonth() + 1;
      while (year < endYear || (year === endYear && month <= endMonth)) {
        map[monthKeyOf(month, year)] = true;
        month += 1;
        if (month > 12) { month = 1; year += 1; }
      }
    } else {
      map[monthKeyOf(row.homeMonth, row.homeYear)] = true;
    }
  }

  row.listoByMonth = map;
}

function normalizeMaxSimultaneous(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const parsed = Number.parseInt(value, 10);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : DEFAULT_MAX_SIMULTANEOUS;
}

function createBlock({
  id,
  blockType,
  headerColor,
  maxSimultaneous = DEFAULT_MAX_SIMULTANEOUS,
  rows = null,
  collapsed = false,
}) {
  const resolvedBlockType = `${blockType ?? ""}`.trim();
  const resolvedRows = Array.isArray(rows) && rows.length
    ? rows
    : [newRowForBlock(resolvedBlockType)];
  
  if (resolvedRows.length === 1 && !(Array.isArray(rows) && rows.length)) {
    resolvedRows[0]._autoPlaceholder = true;
  }

  return {
    id,
    blockType: resolvedBlockType,
    headerColor,
    maxSimultaneous: normalizeMaxSimultaneous(maxSimultaneous),
    collapsed: !!collapsed,
    rows: resolvedRows,
  };
}

function createSeparatorBlock({ id, label, color = "#ffffff" }) {
  return {
    id,
    blockType: `${label ?? ""}`.trim(),
    headerColor: color,
    maxSimultaneous: null,
    collapsed: false,
    rows: [],
    isSeparator: true,
  };
}

function getMaxSimultaneousLabel(maxSimultaneous) {
  const normalizedMax = normalizeMaxSimultaneous(maxSimultaneous);
  return Number.isInteger(normalizedMax) ? `(Máx. ${normalizedMax} simultáneas)` : "";
}

function areAllBlocksCollapsed() {
  const collapsibleBlocks = blocks.filter((block) => !block.isSeparator);
  return collapsibleBlocks.length > 0 && collapsibleBlocks.every((block) => block.collapsed);
}

function setAllBlocksCollapsed(collapsed) {
  blocks = blocks.map((block) => (block.isSeparator ? block : { ...block, collapsed }));
}

function scrollToTopAfterGlobalCollapse(root) {
  const monthBodyWrapper = root?.querySelector(".month-block__body-scroll-wrapper");
  const rightBodyScroll = root?.querySelector("#right-body-scroll");

  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: "auto" });
    monthBodyWrapper?.scrollTo({ top: 0, behavior: "auto" });
    rightBodyScroll?.scrollTo({ top: 0, behavior: "auto" });
  };

  scrollToTop();
  window.requestAnimationFrame(scrollToTop);
}

function updateGlobalCollapseButtonState() {
  const button = document.getElementById(GLOBAL_COLLAPSE_BUTTON_ID);
  if (!button) {
    return;
  }

  const allCollapsed = areAllBlocksCollapsed();
  button.textContent = allCollapsed ? "+" : "−";
  button.setAttribute("aria-expanded", allCollapsed ? "false" : "true");
  button.setAttribute("aria-label", allCollapsed ? "Desplegar todos los bloques" : "Plegar todos los bloques");
}

function createDefaultBlocks() {
  return [
    createBlock({ id: "block-1", blockType: "Promo 20", headerColor: "#8fb596", maxSimultaneous: 5 }),
    createBlock({ id: "block-2", blockType: "Promo 20", headerColor: "#e8cd8e", maxSimultaneous: 5 }),
    createBlock({ id: "block-3", blockType: "Promo 40", headerColor: "#8fb596", maxSimultaneous: 5 }),
    createBlock({ id: "block-4", blockType: "Promo 40", headerColor: "#e8cd8e", maxSimultaneous: 5 }),
    createBlock({ id: "block-5", blockType: "Otras Duraciones", headerColor: "#8fb596", maxSimultaneous: 5 }),
    createBlock({ id: "block-6", blockType: "Otras duraciones", headerColor: "#e8cd8e", maxSimultaneous: 5 }),
    createBlock({ id: "block-7", blockType: "Combo", headerColor: "#8fb596", maxSimultaneous: 1 }),
    createBlock({ id: "block-8", blockType: "Bumper", headerColor: "#8fb596", maxSimultaneous: 8 }),
    createBlock({ id: "block-9", blockType: "Bumper", headerColor: "#e8cd8e", maxSimultaneous: 8 }),
    createBlock({ id: "block-10", blockType: "ID", headerColor: "#8fb596", maxSimultaneous: 1 }),
    createBlock({ id: "block-11", blockType: "Pasos a Publi", headerColor: "#8fb596", maxSimultaneous: 5 }),
    createBlock({ id: "block-12", blockType: "pasos a Publi", headerColor: "#e8cd8e", maxSimultaneous: 5 }),
    createBlock({ id: "block-13", blockType: "Intruso", headerColor: "#8fb596", maxSimultaneous: 10 }),
    createBlock({ id: "block-14", blockType: "Loop protección Pop-Ups", headerColor: "#8fb596", maxSimultaneous: null }),
    createSeparatorBlock({ id: "separator-1", label: "OTROS CANALES" }),
    createBlock({ id: "block-15", blockType: "Canales LaLiga", headerColor: "#e8cd8e", maxSimultaneous: null }),
    createBlock({ id: "block-16", blockType: "Canales Golf", headerColor: "#e8cd8e", maxSimultaneous: null }),
    createBlock({ id: "block-17", blockType: "Canales Caza y Pesca", headerColor: "#e8cd8e", maxSimultaneous: null }),
    createSeparatorBlock({ id: "separator-2", label: "VOD" }),
    createBlock({ id: "block-18", blockType: "Arranque", headerColor: "#8fb596", maxSimultaneous: 1 }),
    createBlock({ id: "block-19", blockType: "Loop", headerColor: "#e8cd8e", maxSimultaneous: 1 }),
    createBlock({ id: "block-20", blockType: "Pre Roll", headerColor: "#8fb596", maxSimultaneous: 5 }),
    createBlock({ id: "block-21", blockType: "Pre Roll", headerColor: "#e8cd8e", maxSimultaneous: 5 }),
    createSeparatorBlock({ id: "separator-3", label: "FREEMIUM" }),
    createBlock({ id: "block-22", blockType: "Otras Duraciones", headerColor: "#aa87c6", maxSimultaneous: null }),
    createBlock({ id: "block-23", blockType: "Bumper", headerColor: "#c7a8e5", maxSimultaneous: null }),
    createBlock({ id: "block-24", blockType: "Pasos a Publi", headerColor: "#aa87c6", maxSimultaneous: null }),
    createBlock({ id: "block-25", blockType: "Intruso", headerColor: "#c7a8e5", maxSimultaneous: null }),
    createBlock({ id: "block-26", blockType: "Pre Roll", headerColor: "#aa87c6", maxSimultaneous: null }),
    createSeparatorBlock({ id: "separator-4", label: "UPSELL" }),
    createBlock({ id: "block-27", blockType: "Promo 20", headerColor: "#aa87c6", maxSimultaneous: null }),
    createBlock({ id: "block-32", blockType: "Promo 40", headerColor: "#c7a8e5", maxSimultaneous: null }),
    createBlock({ id: "block-33", blockType: "Otras Duraciones", headerColor: "#aa87c6", maxSimultaneous: null }),
    createBlock({ id: "block-28", blockType: "Bumper", headerColor: "#c7a8e5", maxSimultaneous: null }),
    createBlock({ id: "block-29", blockType: "Pasos a Publi", headerColor: "#aa87c6", maxSimultaneous: null }),
    createBlock({ id: "block-30", blockType: "Intruso", headerColor: "#c7a8e5", maxSimultaneous: null }),
    createBlock({ id: "block-34", blockType: "Pre Roll", headerColor: "#aa87c6", maxSimultaneous: null }),
    createBlock({ id: "block-31", blockType: "Loop", headerColor: "#c7a8e5", maxSimultaneous: null }),
  ];
}

let blocks = createDefaultBlocks();
let contextMenu = { open: false, x: 0, y: 0, blockIndex: -1, rowIndex: -1 };
let menuElement = null;
let selectedCell = null;
let sortState = { key: null, dir: "asc" };
let selectedCellState = null;
let editingCell = null;
let titleOverlayLayer = null;
let genreMenuElement = null;
let fillHandleElement = null;
let fillDragState = null;
const DRAG_THRESHOLD_PX = 6;
let copyAntsElement = null;
let copyRange = null;
let copyRangeBlockIndex = null;
let dragSelectState = {
  pointerDown: false,
  isDragSelect: false,
  anchorCell: null,
  anchorCol: null,
  anchorBlockIndex: null,
  anchorRow: null,
  downX: 0,
  downY: 0,
};
let dragSelection = null;
let shiftSelectAnchor = null;
let suppressNextGridClick = false;
let genreTypeBuffer = "";
let genreTypeBufferTimestamp = 0;
const GENRE_TYPE_BUFFER_TIMEOUT_MS = 700;
const TOAST_DURATION_MS = 3200;
const HISTORY_LIMIT = 200;
const HISTORY_GROUP_WINDOW_MS = 650;
let toastElement = null;
let toastHideTimer = null;
let deleteConfirmElement = null;
let deleteConfirmState = null;
let currentCalendarContext = { ...DEFAULT_CALENDAR_CONTEXT };
let undoStack = [];
let redoStack = [];
let pendingHistoryAction = null;
let pendingHistoryCommitTimer = null;
let activeHistoryAction = null;
let isApplyingHistory = false;
let searchQuery = "";
let preSearchCollapseState = null;

function cloneRowData(row) {
  return {
    ...row,
    // Copia profunda del mapa de emisión por mes para que las filas clonadas
    // (duplicar, deshacer/rehacer) no compartan el mismo objeto.
    listoByMonth: { ...(row?.listoByMonth || {}) },
  };
}

function cloneRows(rows) {
  return Array.isArray(rows) ? rows.map((row) => cloneRowData(row)) : [];
}

function getCellByMeta(meta) {
  if (!meta) {
    return null;
  }

  return document.querySelector(
    `[data-block-index="${meta.blockIndex}"][data-row-index="${meta.rowIndex}"][data-column-key="${meta.columnKey}"]`
  );
}

function getCellMetaFromRowKey(rowKey, columnKey) {
  if (!rowKey || !columnKey) {
    return null;
  }

  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
    const block = blocks[blockIndex];
    if (!block?.rows?.length) {
      continue;
    }

    const rowIndex = block.rows.findIndex((row) => row?.rowKey === rowKey);
    if (rowIndex >= 0) {
      return { blockIndex, rowIndex, columnKey };
    }
  }

  return null;
}

function getPrimaryCellForHistory(options = {}) {
  if (options.primaryCell) {
    return options.primaryCell;
  }

  const meta = selectedCell ? getCellMeta(selectedCell) : null;
  if (!meta) {
    return null;
  }

  const row = blocks[meta.blockIndex]?.rows?.[meta.rowIndex];
  return { ...meta, rowKey: row?.rowKey || null };
}

function createHistoryAction(type, options = {}) {
  return {
    type,
    patches: [],
    timestamp: Date.now(),
    primaryCell: getPrimaryCellForHistory(options),
    groupKey: options.groupKey || null,
  };
}

function clearPendingHistoryTimer() {
  if (pendingHistoryCommitTimer) {
    window.clearTimeout(pendingHistoryCommitTimer);
    pendingHistoryCommitTimer = null;
  }
}

function finalizeHistoryAction(action) {
  if (!action || !action.patches.length) {
    return;
  }

  undoStack.push(action);
  if (undoStack.length > HISTORY_LIMIT) {
    undoStack = undoStack.slice(undoStack.length - HISTORY_LIMIT);
  }
  redoStack = [];
}

function commitPendingHistoryAction() {
  if (!pendingHistoryAction || isApplyingHistory) {
    return;
  }

  finalizeHistoryAction(pendingHistoryAction);
  pendingHistoryAction = null;
  clearPendingHistoryTimer();
}

function schedulePendingHistoryCommit() {
  clearPendingHistoryTimer();
  pendingHistoryCommitTimer = window.setTimeout(() => {
    commitPendingHistoryAction();
  }, HISTORY_GROUP_WINDOW_MS);
}

function ensureActiveHistoryAction(type, options = {}) {
  if (isApplyingHistory) {
    return null;
  }

  const forceIsolated = !!options.forceIsolated;
  const groupKey = options.groupKey || null;

  if (forceIsolated) {
    commitPendingHistoryAction();
    const isolated = createHistoryAction(type, options);
    activeHistoryAction = isolated;
    return isolated;
  }

  if (!pendingHistoryAction) {
    pendingHistoryAction = createHistoryAction(type, options);
    schedulePendingHistoryCommit();
    return pendingHistoryAction;
  }

  const sameType = pendingHistoryAction.type === type;
  const sameGroup = pendingHistoryAction.groupKey && groupKey && pendingHistoryAction.groupKey === groupKey;
  if (sameType && sameGroup) {
    schedulePendingHistoryCommit();
    return pendingHistoryAction;
  }

  commitPendingHistoryAction();
  pendingHistoryAction = createHistoryAction(type, options);
  schedulePendingHistoryCommit();
  return pendingHistoryAction;
}

function closeActiveHistoryAction() {
  if (!activeHistoryAction || isApplyingHistory) {
    activeHistoryAction = null;
    return;
  }

  finalizeHistoryAction(activeHistoryAction);
  activeHistoryAction = null;
}

function withHistoryAction(type, options, callback) {
  const action = ensureActiveHistoryAction(type, { ...options, forceIsolated: true });
  if (!action) {
    return callback?.();
  }

  try {
    return callback?.();
  } finally {
    closeActiveHistoryAction();
  }
}

function addPatchToCurrentAction(patch, options = {}) {
  if (isApplyingHistory || !patch) {
    return;
  }

  const current = activeHistoryAction || ensureActiveHistoryAction(options.type || "edit", options);
  if (!current) {
    return;
  }

  current.patches.push(patch);
  // Cualquier cambio de datos editado activa la red de seguridad local.
  markDirty();
}

function applyPatch(patch, direction) {
  if (!patch) {
    return;
  }

  if (patch.type === "setCell") {
    const value = direction === "forward" ? patch.after : patch.before;
    const block = blocks[patch.blockIndex];
    const row = block?.rows?.[patch.rowIndex];
    if (!row) {
      return;
    }

    const normalized = parseCellValue(patch.columnKey, value);
    if (patch.columnKey === "title") {
      row.title = normalized;
    } else if (patch.columnKey === "listo") {
      setRowListo(row, normalized, patch.monthKey ? ctxFromMonthKey(patch.monthKey) : currentCalendarContext);
    } else if (DATE_COLUMNS.has(patch.columnKey)) {
      applyDateCellValue(row, patch.columnKey, normalized, { preserveRawOnInvalid: true });
    } else if (patch.columnKey === "genre") {
      row.genre = normalized;
    } else if (patch.columnKey === "id") {
      row.id = normalized;
    }
    return;
  }

  if (patch.type === "insertRows") {
    const block = blocks[patch.blockIndex];
    if (!block) {
      return;
    }

    const nextRows = [...block.rows];
    if (direction === "forward") {
      nextRows.splice(patch.atIndex, 0, ...cloneRows(patch.rows));
    } else {
      nextRows.splice(patch.atIndex, patch.rows.length);
    }
    blocks[patch.blockIndex] = { ...block, rows: nextRows };
    return;
  }

  if (patch.type === "deleteRows") {
    const block = blocks[patch.blockIndex];
    if (!block) {
      return;
    }

    const nextRows = [...block.rows];
    if (direction === "forward") {
      nextRows.splice(patch.atIndex, patch.rows.length);
  if (!nextRows.length) {
    const fallback = newRowForBlock(block.blockType, currentCalendarContext);
    fallback._autoPlaceholder = true;
    nextRows.push(fallback);
  }
    } else {
      nextRows.splice(patch.atIndex, 0, ...cloneRows(patch.rows));
    }
    blocks[patch.blockIndex] = { ...block, rows: nextRows };
  }
}

function restoreHistoryFocus(action) {
  const fallbackMeta = action?.primaryCell || null;
  let cell = null;

  if (fallbackMeta?.rowKey) {
    const resolved = getCellMetaFromRowKey(fallbackMeta.rowKey, fallbackMeta.columnKey);
    cell = getCellByMeta(resolved);
  }

  if (!cell && fallbackMeta?.blockIndex !== undefined) {
    cell = getCellByMeta(fallbackMeta);
  }

  if (cell) {
    setSelectedCell(cell);
    focusCellWithoutEditing(cell);
  }
}

function runHistoryAction(action, direction) {
  if (!action?.patches?.length) {
    return;
  }

  commitPendingHistoryAction();
  clearPendingHistoryTimer();

  isApplyingHistory = true;
  try {
    const patches = direction === "undo" ? [...action.patches].reverse() : action.patches;
    const patchDirection = direction === "undo" ? "backward" : "forward";
    patches.forEach((patch) => applyPatch(patch, patchDirection));
  } finally {
    isApplyingHistory = false;
  }

  renderRows();
  restoreHistoryFocus(action);
  // Deshacer/rehacer también cambia los datos → autoguardado local.
  markDirty();
}

function undoLastAction() {
  commitPendingHistoryAction();
  const action = undoStack.pop();
  if (!action) {
    return;
  }

  runHistoryAction(action, "undo");
  redoStack.push(action);
}

function redoLastAction() {
  commitPendingHistoryAction();
  const action = redoStack.pop();
  if (!action) {
    return;
  }

  runHistoryAction(action, "redo");
  undoStack.push(action);
}

function createSetCellPatch(meta, before, after) {
  return {
    type: "setCell",
    blockIndex: meta.blockIndex,
    rowIndex: meta.rowIndex,
    rowKey: meta.rowKey,
    columnKey: meta.columnKey,
    // Para LISTO (estado por mes) guardamos el mes al que aplica el cambio, de
    // modo que deshacer/rehacer afecte al mes correcto aunque se haya navegado.
    monthKey: meta.monthKey,
    before,
    after,
  };
}

function getMonthTitleText(month, year) {
  const monthName = MONTH_NAMES_ES[month - 1] || "";
  return `${monthName.toUpperCase()} ${year}`;
}

function getMonthAriaLabel(month, year) {
  const monthName = MONTH_NAMES_ES[month - 1] || "";
  const readableName = `${monthName.charAt(0).toUpperCase()}${monthName.slice(1)}`;
  return `MonthBlockGrid ${readableName} ${year}`;
}

function normalizeMonthLabel(value) {
  return `${value ?? ""}`
    .trim()
    .toLocaleLowerCase("es-ES")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "");
}

function resolveCalendarContextFromTitle(titleText) {
  const normalizedTitle = `${titleText ?? ""}`.trim();
  const match = normalizedTitle.match(/^([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)\s+(\d{4})$/u);
  if (!match) {
    return { ...DEFAULT_CALENDAR_CONTEXT };
  }

  const monthLabel = normalizeMonthLabel(match[1]);
  const year = Number.parseInt(match[2], 10);
  const month = MONTH_LABEL_TO_NUMBER[monthLabel];

  if (!Number.isInteger(month) || !Number.isInteger(year)) {
    return { ...DEFAULT_CALENDAR_CONTEXT };
  }

  return {
    month,
    year,
    daysInMonth: daysInMonth(month, year),
  };
}

function updateCalendarContext(root = document) {
  const titleElement = root.querySelector?.(".panel-layout__month-title");
  currentCalendarContext = resolveCalendarContextFromTitle(titleElement?.textContent);
  return currentCalendarContext;
}

function applyCalendarContextToView(root = document) {
  const titleElement = root.querySelector?.(".panel-layout__month-title");
  const monthBlock = root.querySelector?.(".month-block");
  const dayHeaderCells = root.querySelectorAll?.("#right-header-track .day-cell") || [];
  const dayRows = root.querySelectorAll?.("#right-body .day-row") || [];

  if (titleElement) {
    titleElement.textContent = getMonthTitleText(currentCalendarContext.month, currentCalendarContext.year);
  }

  if (monthBlock) {
    monthBlock.setAttribute("aria-label", getMonthAriaLabel(currentCalendarContext.month, currentCalendarContext.year));
  }

  dayHeaderCells.forEach((cell, index) => {
    const day = index + 1;
    updateDayHeaderCell(cell, day, currentCalendarContext);
    cell.classList.toggle("inactive", day > currentCalendarContext.daysInMonth);
  });

  dayRows.forEach((row) => {
    [...row.children].forEach((cell, index) => {
      const day = index + 1;
      cell.classList.toggle("inactive", day > currentCalendarContext.daysInMonth);
    });
  });

  root.dispatchEvent(new CustomEvent("calendar:month-change", {
    bubbles: true,
    detail: { ...currentCalendarContext },
  }));

  renderRows();

  syncFillHandlePosition();
  syncCopyAntsPosition();
  updateGlobalCollapseButtonState();
}

function getWeekdayInitial(day, calendarContext = currentCalendarContext) {
  const weekdayIndex = new Date(calendarContext.year, calendarContext.month - 1, day).getDay();
  return WEEKDAY_INITIALS_ES[(weekdayIndex + 6) % 7] || "";
}

function updateDayHeaderCell(cell, day, calendarContext = currentCalendarContext) {
  if (!cell) {
    return;
  }

  cell.setAttribute("data-day", String(day));
  cell.innerHTML = `
    <span class="day-cell__weekday" aria-hidden="true">${getWeekdayInitial(day, calendarContext)}</span>
    <span class="day-cell__number">${day}</span>
  `;
}

function shiftCalendarMonth(deltaMonths, root = document) {
  if (!Number.isInteger(deltaMonths) || deltaMonths === 0) {
    return;
  }

  const date = new Date(currentCalendarContext.year, currentCalendarContext.month - 1 + deltaMonths, 1);
  currentCalendarContext = {
    month: date.getMonth() + 1,
    year: date.getFullYear(),
    daysInMonth: daysInMonth(date.getMonth() + 1, date.getFullYear()),
  };
  applyCalendarContextToView(root);
}

function attachMonthNavigation(root) {
  const navArrows = root.querySelectorAll(".panel-layout__month-nav-arrow");
  if (!navArrows.length) {
    return;
  }

  navArrows.forEach((arrow, index) => {
    const delta = index === 0 ? -1 : 1;
    arrow.setAttribute("role", "button");
    arrow.setAttribute("tabindex", "0");
    arrow.setAttribute("aria-label", delta < 0 ? "Mes anterior" : "Mes siguiente");

    arrow.addEventListener("click", () => {
      shiftCalendarMonth(delta, root);
    });

    arrow.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }
      event.preventDefault();
      shiftCalendarMonth(delta, root);
    });
  });
}

function getContextSelectionTarget(blockIndex, rowIndex) {
  if (
    dragSelection
    && Number.isInteger(dragSelection.blockIndex)
    && dragSelection.blockIndex === blockIndex
    && Number.isInteger(dragSelection.r1)
    && Number.isInteger(dragSelection.r2)
  ) {
    const startRow = Math.min(dragSelection.r1, dragSelection.r2);
    const endRow = Math.max(dragSelection.r1, dragSelection.r2);
    if (rowIndex >= startRow && rowIndex <= endRow) {
      // Contar filas reales visibles en el rango, no el span aritmético
      const block = blocks[blockIndex];
      const orderedRows = getOrderedRowsForMonth(block, currentCalendarContext);
      const rowsInRange = orderedRows.filter(
        (item) => item.sourceIndex >= startRow && item.sourceIndex <= endRow
      );
      return {
        blockIndex,
        startRow,
        endRow,
        count: Math.max(1, rowsInRange.length),
      };
    }
  }

  return getContextRowTarget(blockIndex, rowIndex);
}

function getDeleteTarget(preferredBlockIndex = null, preferredRowIndex = null) {
  if (Number.isInteger(preferredBlockIndex) && Number.isInteger(preferredRowIndex)) {
    const contextTarget = getContextSelectionTarget(preferredBlockIndex, preferredRowIndex);
    if (contextTarget) {
      return contextTarget;
    }
  }

  if (
    dragSelection
    && Number.isInteger(dragSelection.blockIndex)
    && Number.isInteger(dragSelection.r1)
    && Number.isInteger(dragSelection.r2)
  ) {
    if (preferredBlockIndex === null || dragSelection.blockIndex === preferredBlockIndex) {
      const startRow = Math.min(dragSelection.r1, dragSelection.r2);
      const endRow = Math.max(dragSelection.r1, dragSelection.r2);
      return {
        blockIndex: dragSelection.blockIndex,
        startRow,
        endRow,
        count: endRow - startRow + 1,
      };
    }
  }

  const activeMeta = getCellMeta(selectedCell);
  if (!activeMeta) {
    return null;
  }

  if (preferredBlockIndex !== null && activeMeta.blockIndex !== preferredBlockIndex) {
    return null;
  }

  return {
    blockIndex: activeMeta.blockIndex,
    startRow: activeMeta.rowIndex,
    endRow: activeMeta.rowIndex,
    count: 1,
  };
}

function canDeleteRows(preferredBlockIndex = null, preferredRowIndex = null) {
  return !!getDeleteTarget(preferredBlockIndex, preferredRowIndex);
}

function getDuplicateTarget(preferredBlockIndex = null, preferredRowIndex = null) {
  return getDeleteTarget(preferredBlockIndex, preferredRowIndex);
}

function getContextRowTarget(blockIndex, rowIndex) {
  const block = blocks[blockIndex];
  if (!block?.rows?.length || !Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= block.rows.length) {
    return null;
  }

  return {
    blockIndex,
    startRow: rowIndex,
    endRow: rowIndex,
    count: 1,
  };
}

function getInsertTargetFromContext(blockIndex, rowIndex) {
  return getContextSelectionTarget(blockIndex, rowIndex);
}

function copyRowDataInto(sourceRow, targetRow) {
  if (!sourceRow || !targetRow) {
    return targetRow;
  }

  const { rowKey, ...sourceData } = sourceRow;
  return {
    ...targetRow,
    ...sourceData,
    // Copia profunda para no compartir el mapa de emisión con la fila origen.
    listoByMonth: { ...(sourceData.listoByMonth || {}) },
    rowKey: targetRow.rowKey,
  };
}

function duplicateRowsAroundSelection(direction, preferredBlockIndex = null, preferredRowIndex = null) {
  const duplicateTarget = getDuplicateTarget(preferredBlockIndex, preferredRowIndex);
  if (!duplicateTarget || (direction !== "above" && direction !== "below")) {
    return;
  }

  const { blockIndex, startRow, endRow, count } = duplicateTarget;
  const block = blocks[blockIndex];
  if (!block?.rows?.length) {
    return;
  }

  withHistoryAction("duplicate", { groupKey: `duplicate:${blockIndex}:${startRow}:${direction}` }, () => {
    const nextRows = [...block.rows];
    const sourceRows = nextRows.slice(startRow, endRow + 1).map((row) => ({ ...row }));

    const targetStart = direction === "above" ? startRow : endRow + 1;
    const rowsToInsert = Array.from({ length: count }, () => newRowForBlock(block.blockType, currentCalendarContext));
    nextRows.splice(targetStart, 0, ...rowsToInsert);

    sourceRows.forEach((sourceRow, offset) => {
      const targetIndex = targetStart + offset;
      nextRows[targetIndex] = copyRowDataInto(sourceRow, nextRows[targetIndex]);
    });

    blocks[blockIndex] = { ...block, rows: nextRows };
    addPatchToCurrentAction({ type: "insertRows", blockIndex, atIndex: targetStart, rows: cloneRows(rowsToInsert) }, { type: "duplicate" });
    sourceRows.forEach((sourceRow, offset) => {
      const row = nextRows[targetStart + offset];
      if (!row) {
        return;
      }
      ["listo", "title", "startDate", "endDate", "genre", "id"].forEach((columnKey) => {
        const before = getCellRawValue(row, columnKey);
        const after = getCellRawValue(sourceRow, columnKey);
        if (before !== after) {
          addPatchToCurrentAction(createSetCellPatch({ blockIndex, rowIndex: targetStart + offset, rowKey: row.rowKey, columnKey }, before, after), { type: "duplicate" });
        }
      });
    });

    renderRows();
  });
}

let refreshDeleteControlsScheduled = false;
function refreshDeleteControlsNow() {
  document.querySelectorAll('.gutter-icon-btn[data-action="delete-rows"]').forEach((button) => {
    const blockIndex = Number.parseInt(button.dataset.blockIndex, 10);
    const enabled = canDeleteRows(Number.isNaN(blockIndex) ? null : blockIndex);
    button.disabled = !enabled;
    button.classList.toggle("is-disabled", !enabled);
  });

  if (menuElement?.classList.contains("open")) {
    updateContextMenuDeleteState();
  }
}

// Coalesced + idle-deferred wrapper. Selection changes call this on every
// click; running the querySelectorAll + class toggles synchronously was
// costing a few ms on each click on older machines. Idle scheduling means
// it runs after the click is visually acknowledged.
function refreshDeleteControls() {
  if (refreshDeleteControlsScheduled) return;
  refreshDeleteControlsScheduled = true;
  const run = () => {
    refreshDeleteControlsScheduled = false;
    refreshDeleteControlsNow();
  };
  if (typeof window.requestIdleCallback === "function") {
    window.requestIdleCallback(run, { timeout: 200 });
  } else {
    setTimeout(run, 0);
  }
}

function getTitleOverlayLayer() {
  if (titleOverlayLayer?.isConnected) {
    return titleOverlayLayer;
  }

  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) {
    return null;
  }

  const layer = document.createElement("div");
  layer.className = "title-edit-overlay-layer";
  gridRoot.appendChild(layer);
  titleOverlayLayer = layer;
  return titleOverlayLayer;
}

function getToastElement() {
  if (toastElement?.isConnected) {
    return toastElement;
  }

  const toast = document.createElement("div");
  toast.className = "grid-toast";
  document.body.appendChild(toast);
  toastElement = toast;
  return toastElement;
}

function showGridToast(message) {
  if (!message) {
    return;
  }

  const toast = getToastElement();
  toast.textContent = message;
  toast.classList.add("is-visible");

  if (toastHideTimer) {
    window.clearTimeout(toastHideTimer);
  }

  toastHideTimer = window.setTimeout(() => {
    toast.classList.remove("is-visible");
  }, TOAST_DURATION_MS);
}

function closeDeleteConfirmModal({ shouldRestoreFocus = true } = {}) {
  if (!deleteConfirmElement) {
    deleteConfirmState = null;
    return;
  }

  deleteConfirmElement.classList.remove("open");
  deleteConfirmElement.setAttribute("aria-hidden", "true");
  document.body.classList.remove("delete-modal-open");

  const triggerElement = deleteConfirmState?.triggerElement;
  deleteConfirmState = null;

  if (shouldRestoreFocus && triggerElement?.isConnected) {
    triggerElement.focus({ preventScroll: true });
  }
}

function handleDeleteConfirmKeydown(event) {
  if (!deleteConfirmState || !deleteConfirmElement?.classList.contains("open")) {
    return;
  }

  if (event.key === "Escape") {
    event.preventDefault();
    closeDeleteConfirmModal();
    return;
  }

  if (event.key !== "Tab") {
    return;
  }

  const focusable = [...deleteConfirmElement.querySelectorAll('button:not([disabled])')];
  if (!focusable.length) {
    event.preventDefault();
    return;
  }

  const currentIndex = focusable.indexOf(document.activeElement);
  const direction = event.shiftKey ? -1 : 1;
  const nextIndex = currentIndex === -1
    ? 0
    : (currentIndex + direction + focusable.length) % focusable.length;

  event.preventDefault();
  focusable[nextIndex].focus();
}

function ensureDeleteConfirmElement() {
  if (deleteConfirmElement) {
    return deleteConfirmElement;
  }

  const overlay = document.createElement("div");
  overlay.className = "delete-confirm-overlay";
  overlay.setAttribute("aria-hidden", "true");
  overlay.innerHTML = `
    <div class="delete-confirm-modal" role="dialog" aria-modal="true" aria-labelledby="delete-confirm-title">
      <p id="delete-confirm-title" class="delete-confirm-modal__text"></p>
      <div class="delete-confirm-modal__actions">
        <button type="button" class="delete-confirm-modal__btn" data-action="ok">OK</button>
        <button type="button" class="delete-confirm-modal__btn" data-action="cancel">Cancelar</button>
      </div>
    </div>
  `;

  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      event.preventDefault();
      event.stopPropagation();
    }
  });

  overlay.addEventListener("keydown", handleDeleteConfirmKeydown);

  overlay.addEventListener("click", (event) => {
    const action = event.target.closest("[data-action]")?.dataset?.action;
    if (!action) {
      return;
    }

    if (action === "cancel") {
      closeDeleteConfirmModal();
      return;
    }

    if (action === "ok") {
      const target = deleteConfirmState?.target;
      closeDeleteConfirmModal({ shouldRestoreFocus: false });
      executeDeleteRows(target);
    }
  });

  document.body.appendChild(overlay);
  deleteConfirmElement = overlay;
  return deleteConfirmElement;
}

function openDeleteConfirmModal(target, triggerElement = document.activeElement) {
  if (!target) {
    return;
  }

  closeContextMenu();

  const overlay = ensureDeleteConfirmElement();
  const title = overlay.querySelector(".delete-confirm-modal__text");
  title.textContent = `Vas a eliminar ${target.count} filas`;

  deleteConfirmState = {
    target,
    triggerElement: triggerElement instanceof HTMLElement ? triggerElement : null,
  };

  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  document.body.classList.add("delete-modal-open");

  const okButton = overlay.querySelector('[data-action="ok"]');
  okButton?.focus({ preventScroll: true });
}

function setSelectedCell(cell) {
  if (IS_VIEWER_MODE) { return; }
  if (selectedCell && selectedCell !== cell && selectedCell.isConnected) {
    selectedCell.classList.remove("is-selected");
  }

  selectedCell = cell;

  if (selectedCell?.dataset?.rowId && selectedCell?.dataset?.columnKey) {
    selectedCellState = {
      rowId: selectedCell.dataset.rowId,
      columnKey: selectedCell.dataset.columnKey,
    };
  } else {
    selectedCellState = null;
  }

  if (selectedCell?.isConnected) {
    selectedCell.classList.add("is-selected");

    const gridRoot = document.querySelector(".month-block__body-grid");
    if (gridRoot && !editingCell) {
      gridRoot.focus({ preventScroll: true });
    }
  }

  syncFillHandlePosition();
  refreshDeleteControls();
}

function isSelectedCellState(row, columnKey) {
  return selectedCellState?.rowId === row.rowKey && selectedCellState?.columnKey === columnKey;
}

function getCellMeta(cell) {
  if (!cell?.dataset) {
    return null;
  }

  const blockIndex = Number.parseInt(cell.dataset.blockIndex, 10);
  const rowIndex = Number.parseInt(cell.dataset.rowIndex, 10);
  const columnKey = cell.dataset.columnKey;

  if (Number.isNaN(blockIndex) || Number.isNaN(rowIndex) || !columnKey) {
    return null;
  }

  return { blockIndex, rowIndex, columnKey };
}

function getRowByCell(cell) {
  const meta = getCellMeta(cell);
  if (!meta) {
    return null;
  }

  const block = blocks[meta.blockIndex];
  const row = block?.rows?.[meta.rowIndex];
  if (!row) {
    return null;
  }

  return { meta, row };
}

function getColumnByKey(columnKey) {
  return columns.find((column) => column.key === columnKey) || null;
}

function getCopyRangeValues(selection) {
  if (!selection || copyRangeBlockIndex === null) {
    return [];
  }

  const sourceBlock = blocks[copyRangeBlockIndex];
  if (!sourceBlock?.rows?.length) {
    return [];
  }

  const orderedRows = getOrderedRowsForMonth(sourceBlock, currentCalendarContext);
  return orderedRows
    .filter((item) => item.sourceIndex >= selection.r1 && item.sourceIndex <= selection.r2)
    .map((item) => getCellRawValue(item.row, selection.col));
}

function resolveVerticalPasteValues({ rangeSize, clipboardText }) {
  const normalizedClipboard = `${clipboardText || ""}`.replace(/\r\n/g, "\n");
  const shouldUseCopyRange =
    !!copyRange
    && copyRangeBlockIndex !== null
    && normalizedClipboard === buildCopyTextFromSelection(copyRange);

  if (shouldUseCopyRange) {
    const sourceValues = getCopyRangeValues(copyRange);
    if (sourceValues.length === 1) {
      return Array.from({ length: rangeSize }, () => sourceValues[0]);
    }

    if (sourceValues.length > 1) {
      return sourceValues;
    }
  }

  const clipboardLines = normalizedClipboard.split("\n");
  if (clipboardLines.length > 1 && clipboardLines[clipboardLines.length - 1] === "") {
    clipboardLines.pop();
  }

  const normalizedLines = clipboardLines.map((line) => line.split("\t")[0]);
  if (!normalizedLines.length) {
    return [];
  }

  if (normalizedLines.length === 1) {
    return Array.from({ length: rangeSize }, () => normalizedLines[0]);
  }

  return normalizedLines;
}

function parseCellValue(columnKey, rawValue) {
  const column = getColumnByKey(columnKey);
  const textValue = `${rawValue ?? ""}`;

  if (!column) {
    return textValue;
  }

  if (column.type === "checkbox") {
    const normalized = textValue.trim().toLowerCase();
    return ["true", "1", "x", "si", "sí"].includes(normalized);
  }

  if (columnKey === "title") {
    return textValue.slice(0, 100);
  }

  if (column?.cellType === "select") {
    const trimmedInput = textValue.trim();
    if (!trimmedInput) {
      return "";
    }

    const normalizedInput = trimmedInput.toLocaleLowerCase();
    const matchedOption = column.options?.find((option) => option.toLocaleLowerCase() === normalizedInput);
    // Preserve unrecognised values (e.g. legacy genres from older Excel exports)
    // so that data never gets silently wiped during import / paste.
    return matchedOption || trimmedInput;
  }

  if (DATE_COLUMNS.has(columnKey)) {
    return `${rawValue ?? ""}`;
  }

  return textValue;
}

function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function daysInMonth(month, year) {
  if (month === 2) {
    return isLeapYear(year) ? 29 : 28;
  }
  if ([4, 6, 9, 11].includes(month)) {
    return 30;
  }
  return 31;
}

function formatDateDisplay(day, month, year) {
  const dd = String(day).padStart(2, "0");
  const mm = String(month).padStart(2, "0");
  const yy = String(year).slice(-2);
  return `${dd}/${mm}/${yy}`;
}

function formatDateISO(day, month, year) {
  const dd = String(day).padStart(2, "0");
  const mm = String(month).padStart(2, "0");
  return `${year}-${mm}-${dd}`;
}

function normalizeHeaderToken(value) {
  return `${value ?? ""}`
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizeBlockToken(value) {
  return normalizeHeaderToken(value)
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function findExcelHeaderMatch(sourceHeader, candidates) {
  const normalizedSource = normalizeHeaderToken(sourceHeader);
  if (!normalizedSource) {
    return false;
  }

  return candidates.some((candidate) => {
    const normalizedCandidate = normalizeHeaderToken(candidate);
    if (!normalizedCandidate) {
      return false;
    }

    if (normalizedSource === normalizedCandidate) {
      return true;
    }

    return normalizedSource.startsWith(`${normalizedCandidate} `)
      || normalizedSource.endsWith(` ${normalizedCandidate}`)
      || normalizedSource.includes(` ${normalizedCandidate} `);
  });
}

function mapExcelColumns(headerRow) {
  const mapping = {};
  const normalizedHeaders = Array.isArray(headerRow) ? headerRow : [];

  Object.entries(EXCEL_COLUMN_ALIASES).forEach(([columnKey, aliases]) => {
    const index = normalizedHeaders.findIndex((header) => findExcelHeaderMatch(header, aliases));
    if (index >= 0) {
      mapping[columnKey] = index;
    }
  });

  const blockIndex = normalizedHeaders.findIndex((header) => findExcelHeaderMatch(header, EXCEL_BLOCK_HEADER_CANDIDATES));
  if (blockIndex >= 0) {
    mapping.blockType = blockIndex;
  }

  return mapping;
}

function excelDateNumberToDisplay(value) {
  if (!window.XLSX?.SSF?.parse_date_code || typeof value !== "number") {
    return null;
  }

  const parsed = window.XLSX.SSF.parse_date_code(value);
  if (!parsed || !parsed.y || !parsed.m || !parsed.d) {
    return null;
  }

  return formatDateDisplay(parsed.d, parsed.m, parsed.y);
}

function normalizeImportedCellValue(columnKey, rawValue) {
  if (!DATE_COLUMNS.has(columnKey)) {
    return `${rawValue ?? ""}`;
  }

  if (typeof rawValue === "number") {
    return excelDateNumberToDisplay(rawValue) || `${rawValue}`;
  }

  return `${rawValue ?? ""}`;
}

function isImportedRowEmpty(rowValues, mapping) {
  const trackedKeys = ["title", "startDate", "endDate", "genre", "id", "listo", "blockType"];
  return trackedKeys.every((key) => {
    const index = mapping[key];
    if (!Number.isInteger(index)) {
      return true;
    }

    return !`${rowValues[index] ?? ""}`.trim();
  });
}

function blockTypeMatchesNormalized(candidate, normalizedBlock) {
  if (candidate === normalizedBlock) {
    return true;
  }

  // Word-overlap fuzzy match: at least 60% of the shorter set of words must match.
  // This prevents short names like "LOOP" from matching "LOOPS PROTECCION POP UPS".
  const srcWords = normalizedBlock.split(" ");
  const candWords = candidate.split(" ");
  const srcSet = new Set(srcWords);
  const candSet = new Set(candWords);
  const matchCount = candWords.filter((w) => srcSet.has(w)).length
    + srcWords.filter((w) => candSet.has(w)).length;
  const minWords = Math.min(srcWords.length, candWords.length);
  return matchCount >= Math.ceil(minWords * 1.2);
}

function findTargetBlockIndex(rawBlockType, occurrence = 1) {
  const normalizedBlock = normalizeBlockToken(rawBlockType);
  if (!normalizedBlock) {
    return blocks.findIndex((block) => !block.isSeparator);
  }

  // Find the Nth (1-based) block whose normalized type matches.
  // First pass: exact matches.
  let count = 0;
  let lastExact = -1;
  for (let i = 0; i < blocks.length; i += 1) {
    if (blocks[i].isSeparator) {
      continue;
    }
    if (normalizeBlockToken(blocks[i].blockType) === normalizedBlock) {
      count += 1;
      lastExact = i;
      if (count === occurrence) {
        return i;
      }
    }
  }
  if (lastExact >= 0) {
    return lastExact;
  }

  // Second pass: word-overlap fuzzy matches.
  count = 0;
  let lastFuzzy = -1;
  for (let i = 0; i < blocks.length; i += 1) {
    if (blocks[i].isSeparator) {
      continue;
    }
    if (blockTypeMatchesNormalized(normalizeBlockToken(blocks[i].blockType), normalizedBlock)) {
      count += 1;
      lastFuzzy = i;
      if (count === occurrence) {
        return i;
      }
    }
  }
  return lastFuzzy;
}

function findExcelHeaderRow(matrix) {
  if (!Array.isArray(matrix)) {
    return { headerRowIndex: -1, headerRow: [] };
  }

  for (let index = 0; index < matrix.length; index += 1) {
    const row = Array.isArray(matrix[index]) ? matrix[index] : [];
    const mapping = mapExcelColumns(row);
    const hasTitleColumn = Number.isInteger(mapping.title);
    const hasAtLeastOneDateColumn = Number.isInteger(mapping.startDate) || Number.isInteger(mapping.endDate);
    if (hasTitleColumn && hasAtLeastOneDateColumn) {
      return { headerRowIndex: index, headerRow: row };
    }
  }

  return { headerRowIndex: -1, headerRow: [] };
}

function inferIdColumnIndexFromHeader(headerRow, mapping) {
  if (Number.isInteger(mapping.id) || !Number.isInteger(mapping.genre)) {
    return mapping.id;
  }

  const nextIndex = mapping.genre + 1;
  const rawHeader = `${headerRow[nextIndex] ?? ""}`.trim();
  return rawHeader ? mapping.id : nextIndex;
}

function isExcelTemplateBlockHeader(rowValues, mapping) {
  const blockCellIndex = Number.isInteger(mapping.blockType)
    ? mapping.blockType
    : (Number.isInteger(mapping.listo) ? mapping.listo : 0);
  const candidateLabel = `${rowValues[blockCellIndex] ?? ""}`.trim();
  if (!candidateLabel) {
    return false;
  }

  const hasDataInMainColumns = ["title", "startDate", "endDate", "genre", "id"].some((key) => {
    const index = mapping[key];
    if (!Number.isInteger(index)) {
      return false;
    }

    return !!`${rowValues[index] ?? ""}`.trim();
  });

  return !hasDataInMainColumns;
}

function importRowsFromExcelMatrix(matrix, options = {}) {
  const { silent = false } = options;
  if (!Array.isArray(matrix) || matrix.length < 2) {
    if (!silent) showGridToast("El archivo no contiene datos para importar");
    return;
  }

  const { headerRowIndex, headerRow } = findExcelHeaderRow(matrix);
  if (headerRowIndex < 0) {
    if (!silent) showGridToast("Faltan columnas obligatorias: TÍTULO, INICIO VIG o FIN VIG");
    return;
  }

  const dataRows = matrix.slice(headerRowIndex + 1);
  const mapping = mapExcelColumns(headerRow);
  mapping.id = inferIdColumnIndexFromHeader(headerRow, mapping);

  const hasTitleColumn = Number.isInteger(mapping.title);
  const hasAtLeastOneDateColumn = Number.isInteger(mapping.startDate) || Number.isInteger(mapping.endDate);
  if (!hasTitleColumn || !hasAtLeastOneDateColumn) {
    if (!silent) showGridToast("Faltan columnas obligatorias: TÍTULO, INICIO VIG o FIN VIG");
    return;
  }

  let importedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;
  let currentBlockIndex = blocks.findIndex((block) => !block.isSeparator);

  // Track how many times each normalized block type name has appeared as a header
  // so that the 2nd "PROMO 20" header maps to the 2nd "Promo 20" block, etc.
  const blockTypeOccurrenceCount = {};

  dataRows.forEach((rowValues) => {
    if (!Array.isArray(rowValues) || isImportedRowEmpty(rowValues, mapping)) {
      skippedCount += 1;
      return;
    }

    if (isExcelTemplateBlockHeader(rowValues, mapping)) {
      const rawBlockType = Number.isInteger(mapping.blockType)
        ? rowValues[mapping.blockType]
        : rowValues[Number.isInteger(mapping.listo) ? mapping.listo : 0];
      const normalizedRaw = normalizeBlockToken(rawBlockType);
      blockTypeOccurrenceCount[normalizedRaw] = (blockTypeOccurrenceCount[normalizedRaw] || 0) + 1;
      const occurrence = blockTypeOccurrenceCount[normalizedRaw];
      const resolvedBlockIndex = findTargetBlockIndex(rawBlockType, occurrence);
      if (resolvedBlockIndex >= 0) {
        currentBlockIndex = resolvedBlockIndex;
      }
      skippedCount += 1;
      return;
    }

    const rawBlockType = Number.isInteger(mapping.blockType) ? rowValues[mapping.blockType] : "";
    const blockIndex = `${rawBlockType ?? ""}`.trim() ? findTargetBlockIndex(rawBlockType) : currentBlockIndex;
    if (!Number.isInteger(blockIndex) || blockIndex < 0) {
      skippedCount += 1;
      return;
    }

    const targetBlock = blocks[blockIndex];
    const importedRow = newRowForBlock(targetBlock.blockType, currentCalendarContext);
    importedRow.homeMonth = currentCalendarContext.month;
    importedRow.homeYear = currentCalendarContext.year;

    let rawListoValue;
    Object.keys(EXCEL_COLUMN_ALIASES).forEach((columnKey) => {
      const sourceIndex = mapping[columnKey];
      if (!Number.isInteger(sourceIndex)) {
        return;
      }

      const normalizedValue = normalizeImportedCellValue(columnKey, rowValues[sourceIndex]);

      if (DATE_COLUMNS.has(columnKey)) {
        applyDateCellValue(importedRow, columnKey, normalizedValue);
        return;
      }

      if (columnKey === "actualizado") {
        const normalized = `${normalizedValue ?? ""}`.trim().toLowerCase();
        importedRow.actualizado = ["true", "1", "x", "si", "sí", "actualizado"].includes(normalized);
        return;
      }

      if (columnKey === "rowKey") {
        const persistedKey = `${normalizedValue ?? ""}`.trim();
        if (persistedKey) {
          // Reuse the stable identity from disk so cross-session edits to the
          // same logical row can be matched during merge-on-save.
          importedRow.rowKey = persistedKey;
        }
        return;
      }

      if (columnKey === "listo") {
        // El estado de emisión se procesa tras el bucle, cuando ya se han
        // parseado las fechas (necesarias para migrar datos antiguos al rango).
        rawListoValue = rowValues[sourceIndex];
        return;
      }

      const parsedValue = parseCellValue(columnKey, normalizedValue);
      importedRow[columnKey] = parsedValue;
    });

    applyImportedListo(importedRow, rawListoValue);

    const hasDateErrors = !!importedRow.startDateError || !!importedRow.endDateError || !!importedRow.dateRangeError;
    if (hasDateErrors) {
      invalidCount += 1;
      return;
    }

    if (targetBlock.rows.every((r) => r._autoPlaceholder)) {
      targetBlock.rows.length = 0;
    }
    targetBlock.rows.push(importedRow);
    importedCount += 1;
  });

  if (!silent) renderRows();

  if (!silent) {
    const summary = [`${importedCount} fila(s) importada(s)`];
    if (invalidCount) {
      summary.push(`${invalidCount} descartada(s) por fecha no válida`);
    }
    if (skippedCount) {
      summary.push(`${skippedCount} vacía(s) o sin bloque destino`);
    }
    showGridToast(summary.join(" · "));
  }
}

function resolveContextFromSheetName(sheetName) {
  const normalized = `${sheetName ?? ""}`.trim();
  // Accept 3 or 4-digit years to handle typos like "FEBRERO 206" instead of "FEBRERO 2026"
  const match = normalized.match(/^([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)\s+(\d{3,4})$/u);
  if (!match) {
    return null;
  }

  const monthLabel = normalizeMonthLabel(match[1]);
  const month = MONTH_LABEL_TO_NUMBER[monthLabel];
  if (!Number.isInteger(month)) {
    return null;
  }

  let year = Number.parseInt(match[2], 10);
  if (year < 1000) {
    // 3-digit year is a typo (e.g. "206" → 2026): fall back to current year
    year = new Date().getFullYear();
  }

  return {
    month,
    year,
    daysInMonth: new Date(year, month, 0).getDate(),
  };
}

function handleExcelFileSelection(event) {
  const input = event?.currentTarget;
  const file = input?.files?.[0];
  if (!file) {
    return;
  }

  if (!window.XLSX) {
    showGridToast("No se pudo cargar la librería de Excel");
    input.value = "";
    return;
  }

  const reader = new FileReader();
  reader.onload = (loadEvent) => {
    try {
      const workbook = window.XLSX.read(loadEvent.target?.result, { type: "array", cellDates: false });
      const sheetsWithData = workbook.SheetNames.filter((name) => workbook.Sheets[name]);
      if (!sheetsWithData.length) {
        showGridToast("No se encontró ninguna hoja en el archivo");
        return;
      }

      const savedContext = currentCalendarContext;
      let totalImported = 0;

      sheetsWithData.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const matrix = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
        const sheetContext = resolveContextFromSheetName(sheetName);
        if (sheetContext) {
          currentCalendarContext = sheetContext;
        }
        importRowsFromExcelMatrix(matrix);
      });

      currentCalendarContext = savedContext;
    } catch (error) {
      showGridToast("No se pudo importar el Excel. Revisa el formato del archivo");
    } finally {
      input.value = "";
    }
  };

  reader.onerror = () => {
    showGridToast("No se pudo leer el archivo seleccionado");
    input.value = "";
  };

  reader.readAsArrayBuffer(file);
}

function attachExcelImportControls(root) {
  const toolbarInner = root.querySelector(".panel-layout__toolbar-inner");
  const importButton = root.querySelector(".import-excel-btn");
  if (!toolbarInner || !importButton) {
    return;
  }

  if (!ENABLE_EXCEL_IMPORT) {
    importButton.disabled = true;
    importButton.title = "Importación Excel desactivada por configuración";
    return;
  }

  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".xlsx,.xls";
  fileInput.className = "import-excel-input";
  fileInput.style.display = "none";
  fileInput.addEventListener("change", handleExcelFileSelection);
  toolbarInner.appendChild(fileInput);

  importButton.addEventListener("click", () => {
    fileInput.click();
  });
}

async function buildExcelEdicionBuffer(srcBlocks = blocks) {
  if (!window.ExcelJS) {
    throw new Error("ExcelJS no cargado");
  }

  // Mapa de colores UI → colores Excel (ARGB)
  const HEADER_COLOR_MAP = {
    "#8fb596": "FF70AD47", // verde
    "#e8cd8e": "FFFFC000", // amarillo
    "#aa87c6": "FFAA87C6", // púrpura oscuro
    "#c7a8e5": "FFC7A8E5", // púrpura claro
  };
  const COLOR_RED_SEP   = "FFC00000";
  const COLOR_BLUE_HDR  = "FF4472C4";
  const COLOR_WHITE     = "FFFFFFFF";
  const COLOR_BLACK     = "FF000000";

  function toArgb(hexColor) {
    return HEADER_COLOR_MAP[hexColor?.toLowerCase()] || "FFD9D9D9";
  }

  function applyHeaderStyle(cell, bgArgb, textArgb = COLOR_WHITE) {
    cell.font = { bold: true, color: { argb: textArgb } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgArgb } };
  }

  // Recopilar meses con datos
  const monthsMap = new Map();
  srcBlocks.forEach((block) => {
    if (block.isSeparator) return;
    block.rows.forEach((row) => {
      if (isPlaceholderRow(row)) return;
      const key = `${row.homeYear}-${String(row.homeMonth).padStart(2, "0")}`;
      if (!monthsMap.has(key)) {
        monthsMap.set(key, { month: row.homeMonth, year: row.homeYear });
      }
    });
  });

  if (monthsMap.size === 0) {
    const { month, year } = currentCalendarContext;
    monthsMap.set(`${year}-${String(month).padStart(2, "0")}`, { month, year });
  }

  const sortedMonths = [...monthsMap.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, ctx]) => ctx);

  const workbook = new ExcelJS.Workbook();

  sortedMonths.forEach(({ month, year }) => {
    const monthName = MONTH_NAMES_ES[month - 1].toUpperCase();
    const ws = workbook.addWorksheet(`${monthName} ${year}`);

    ws.columns = [
      { width: 8  },  // LISTO
      { width: 55 },  // TITULO
      { width: 12 },  // INICIO VIG
      { width: 12 },  // FIN VIG
      { width: 18 },  // GENERO
      { width: 14 },  // ID
      { width: 12 },  // ACTUALIZADO
      { width: 28 },  // ROW_KEY (stable identity used for merge-on-save)
    ];
    // ROW_KEY is technical metadata, hide it from human readers.
    ws.getColumn(8).hidden = true;

    // — Fila de cabecera principal (azul, negrita, blanco) —
    const headerRow = ws.addRow(["LISTO", "TITULO", "INICIO VIG", "FIN VIG", "GENERO", "ID", "ACTUALIZADO", "ROW_KEY"]);
    headerRow.eachCell((cell) => applyHeaderStyle(cell, COLOR_BLUE_HDR));
    headerRow.commit();

    // — Bloques —
    srcBlocks.forEach((block) => {
      const blockLabel  = block.blockType.toUpperCase();
      const isSep       = block.isSeparator;
      const isRedSep    = isSep && (blockLabel === "OTROS CANALES" || blockLabel === "VOD" || blockLabel === "FREEMIUM" || blockLabel === "UPSELL");
      const bgArgb      = isRedSep ? COLOR_RED_SEP : toArgb(block.headerColor);
      const textArgb    = COLOR_WHITE;

      // Cabecera de bloque (celda A fusionada A:H)
      const blockHeaderRow = ws.addRow([blockLabel, null, null, null, null, null, null, null]);
      const rowNum = blockHeaderRow.number;
      ws.mergeCells(`A${rowNum}:H${rowNum}`);
      applyHeaderStyle(blockHeaderRow.getCell(1), bgArgb, textArgb);
      blockHeaderRow.commit();

      if (isSep) return;

      // Filas de datos del mes
      const monthRows = block.rows.filter(
        (row) => row.homeMonth === month && row.homeYear === year && !isPlaceholderRow(row)
      );

      monthRows.forEach((row) => {
        const dataRow = ws.addRow([
          encodeListoByMonth(row.listoByMonth) || null,
          row.title       || null,
          row.startDateText || null,
          row.endDateText   || null,
          row.genre       ? row.genre.toUpperCase() : null,
          row.id          || null,
          !!row.actualizado,
          row.rowKey      || null,
        ]);
        // Forzar texto en columnas de fecha
        dataRow.getCell(3).numFmt = "@";
        dataRow.getCell(4).numFmt = "@";
        // ROW_KEY: store as plain text, no formula/number coercion.
        dataRow.getCell(8).numFmt = "@";
        dataRow.commit();
      });
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, sheetCount: sortedMonths.length };
}

// =============================================================================
// MERGE-ON-SAVE
//
// Snapshot of the data at the moment of the last successful Drive load /
// save. Used to compute the local delta when the user pulls the Guardar
// button so that concurrent edits from other sessions are preserved.
//
// null while no baseline exists (cold boot before first Drive load) — in
// that case saveToGoogleDrive falls back to a plain overwrite.
// =============================================================================
let loadedSnapshot = null;

// Row fields that participate in the per-cell diff. listoByMonth is handled
// separately because it is an object map and needs structural comparison.
const DIFFABLE_ROW_FIELDS = ["title", "genre", "id", "startDateText", "endDateText", "actualizado"];
const DIFFABLE_DATE_FIELDS = new Set(["startDateText", "endDateText"]);
const DIFFABLE_FIELD_DEFAULTS = {
  title: "",
  genre: "",
  id: "",
  startDateText: "",
  endDateText: "",
  actualizado: false,
};

function normalizeDiffableField(row, field) {
  const v = row?.[field];
  if (v === undefined || v === null) {
    return DIFFABLE_FIELD_DEFAULTS[field];
  }
  return v;
}

function deepCloneBlocks(srcBlocks) {
  return JSON.parse(JSON.stringify(srcBlocks));
}

// Parse an XLSX buffer into a NEW blocks structure, without touching the live
// `blocks` array or the live UI. Reuses the existing import code by swapping
// the global state around the parse call.
function parseBufferToBlocks(buffer) {
  if (!window.XLSX) {
    throw new Error("XLSX no cargado");
  }
  const workbook = window.XLSX.read(buffer, { type: "array", cellDates: false });
  const sheetsWithData = workbook.SheetNames.filter((name) => workbook.Sheets[name]);
  if (!sheetsWithData.length) {
    return createDefaultBlocks();
  }

  const liveBlocks = blocks;
  const liveContext = { ...currentCalendarContext };

  blocks = createDefaultBlocks();
  let parsedBlocks;

  try {
    sheetsWithData.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const matrix = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      const sheetContext = resolveContextFromSheetName(sheetName);
      if (sheetContext) {
        currentCalendarContext = sheetContext;
      }
      importRowsFromExcelMatrix(matrix, { silent: true });
    });
    parsedBlocks = blocks;
  } finally {
    blocks = liveBlocks;
    currentCalendarContext = liveContext;
  }

  return parsedBlocks;
}

// Build a map rowKey → { row, blockId } for every non-placeholder row in
// `srcBlocks`. Skips rows without a rowKey (defensive — shouldn't happen with
// session-prefixed identifiers in place).
function indexRowsByKey(srcBlocks) {
  const map = new Map();
  srcBlocks.forEach((block) => {
    if (!block?.rows) return;
    block.rows.forEach((row) => {
      if (!row || row._autoPlaceholder) return;
      if (!row.rowKey) return;
      map.set(row.rowKey, { row, blockId: block.id });
    });
  });
  return map;
}

// Compute the difference between `snapshot` (what the user pulled from Drive)
// and `current` (what they have on screen). Returns three buckets:
//   cellChanges    — per-cell edits to existing rows
//   newRows        — rows the user created locally
//   deletedRowKeys — rows the user deleted locally
function computeRowDelta(snapshot, current) {
  const snapshotByKey = indexRowsByKey(snapshot);
  const currentByKey = indexRowsByKey(current);

  const cellChanges = [];
  const newRows = [];
  const deletedRowKeys = [];

  for (const rowKey of snapshotByKey.keys()) {
    if (!currentByKey.has(rowKey)) {
      deletedRowKeys.push(rowKey);
    }
  }

  for (const [rowKey, currEntry] of currentByKey) {
    const snapEntry = snapshotByKey.get(rowKey);
    if (!snapEntry) {
      newRows.push({
        rowSnapshot: JSON.parse(JSON.stringify(currEntry.row)),
        blockId: currEntry.blockId,
      });
      continue;
    }

    const snapRow = snapEntry.row;
    const currRow = currEntry.row;

    DIFFABLE_ROW_FIELDS.forEach((field) => {
      const before = normalizeDiffableField(snapRow, field);
      const after = normalizeDiffableField(currRow, field);
      if (before !== after) {
        cellChanges.push({ rowKey, field, value: after });
      }
    });

    // listoByMonth is an object; compare via its canonical serialization.
    const snapListo = encodeListoByMonth(snapRow.listoByMonth || {});
    const currListo = encodeListoByMonth(currRow.listoByMonth || {});
    if (snapListo !== currListo) {
      cellChanges.push({
        rowKey,
        field: "listoByMonth",
        value: JSON.parse(JSON.stringify(currRow.listoByMonth || {})),
      });
    }
  }

  return { cellChanges, newRows, deletedRowKeys };
}

// Apply a previously computed delta to `targetBlocks` (typically the freshly
// fetched remote state). Local edits win on conflicts (Política A): if both
// you and another session edited the same cell, yours overrides on save.
function applyRowDelta(targetBlocks, delta) {
  const { cellChanges, newRows, deletedRowKeys } = delta;

  // 1. Deletes first so subsequent index lookups are accurate.
  if (deletedRowKeys.length) {
    const deleteSet = new Set(deletedRowKeys);
    targetBlocks.forEach((block) => {
      if (!block?.rows) return;
      block.rows = block.rows.filter((row) => !row || !deleteSet.has(row.rowKey));
    });
  }

  // 2. Cell changes — find row by rowKey and overwrite the field.
  const targetByKey = indexRowsByKey(targetBlocks);
  cellChanges.forEach(({ rowKey, field, value }) => {
    const entry = targetByKey.get(rowKey);
    if (!entry) {
      // Row was deleted remotely; we cannot apply changes to it.
      return;
    }
    const row = entry.row;
    if (DIFFABLE_DATE_FIELDS.has(field)) {
      const dateKey = field === "startDateText" ? "startDate" : "endDate";
      applyDateCellValue(row, dateKey, value);
    } else if (field === "listoByMonth") {
      row.listoByMonth = { ...(value || {}) };
    } else {
      row[field] = value;
    }
  });

  // 3. New rows — append to the matching block (by stable id).
  newRows.forEach(({ rowSnapshot, blockId }) => {
    const targetBlock = targetBlocks.find((b) => b && b.id === blockId && !b.isSeparator);
    if (!targetBlock || !Array.isArray(targetBlock.rows)) return;
    if (targetBlock.rows.every((r) => r && r._autoPlaceholder)) {
      targetBlock.rows.length = 0;
    }
    targetBlock.rows.push(rowSnapshot);
  });
}

// `collapsed` is a per-user UI flag not persisted in Excel. Preserve it from
// the live blocks onto the freshly-parsed remote structure before adopting it.
function preserveBlockCollapsedState(targetBlocks, sourceBlocks) {
  const collapsedById = new Map();
  sourceBlocks.forEach((b) => {
    if (b?.id) collapsedById.set(b.id, !!b.collapsed);
  });
  targetBlocks.forEach((b) => {
    if (b?.id && collapsedById.has(b.id)) {
      b.collapsed = collapsedById.get(b.id);
    }
  });
}

// =============================================================================
// PRESENCE INDICATOR
//
// Best-effort "is someone else editing right now?" signal. Each browser
// session publishes a heartbeat key into the Excel's Drive appProperties
// every PRESENCE_HEARTBEAT_MS milliseconds. The same session also polls the
// bag every PRESENCE_POLL_MS to count how many *other* sessions have been
// active in the last PRESENCE_STALE_MS — that count drives a small badge
// next to the search box.
//
// Purely informational: there is no lock and no blocking. Users decide for
// themselves whether they want to edit while others are present. The merge
// logic still keeps everyone's changes consistent regardless.
// =============================================================================
const PRESENCE_KEY_PREFIX = "presence_";
const PRESENCE_HEARTBEAT_MS = 20_000;
const PRESENCE_POLL_MS = 6_000;
// While saving we beat much faster so the saving flag never goes stale before
// the upload finishes — drive uploads can briefly exceed the normal heartbeat
// interval and we don't want other sessions to evict our flag mid-save.
const PRESENCE_SAVE_HEARTBEAT_MS = 4_000;
const PRESENCE_STALE_MS = 60_000;
// A saving flag is considered current if its epoch is within this window.
// Slightly more lenient than presence staleness because saves take a few
// seconds and we want the lock to hold for the full operation.
const PRESENCE_SAVING_FLAG_MAX_AGE_S = 25;
const EDITOR_NAME_STORAGE_KEY = `panelControlEditorName:${window.PANEL_CONFIG?.GOOGLE_DRIVE_FILE_ID || "default"}`;
const EDITOR_NAME_MAX_LENGTH = 24;
// Short session id used as the appProperties key. Drive limits each key to
// ~124 chars and the bag to ~30 entries, so we keep ids compact.
const PRESENCE_SESSION_ID = `${Math.random().toString(36).slice(2, 8)}${Date.now().toString(36).slice(-4)}`;
let presenceHeartbeatTimer = null;
let presencePollTimer = null;
let presenceSaveHeartbeatTimer = null;
let presenceStarted = false;
let editorName = "";
// Flags driving the "someone is saving" UI lock. Toggled by saveToGoogleDrive
// (own save) and by presencePoll (someone else's save detected via Drive).
let selfIsSaving = false;
let otherIsSaving = false;

function presenceElement() {
  return document.getElementById("presence-indicator");
}

function readStoredEditorName() {
  if (!hasLocalStorage()) return "";
  try {
    const raw = window.localStorage.getItem(EDITOR_NAME_STORAGE_KEY);
    return `${raw ?? ""}`.trim().slice(0, EDITOR_NAME_MAX_LENGTH);
  } catch (_) {
    return "";
  }
}

function persistEditorName(name) {
  if (!hasLocalStorage()) return;
  try {
    window.localStorage.setItem(EDITOR_NAME_STORAGE_KEY, name);
  } catch (_) { /* ignore */ }
}

// Sanitise a free-text alias for both storage and Drive transport. Strips the
// reserved separator we use to encode (epoch|name) inside one appProperties
// value, collapses whitespace, hard-limits length.
function sanitiseEditorName(raw) {
  return `${raw ?? ""}`
    .replace(/\|/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, EDITOR_NAME_MAX_LENGTH);
}

// Inject the (one-off) modal asking the user for a display name. Returns a
// Promise that resolves with the chosen sanitised string. Cannot be dismissed
// without typing something — we need *some* identifier, even if cheeky.
function promptForEditorName({ initialValue = "", reason = "primera" } = {}) {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "editor-name-overlay";
    overlay.innerHTML = `
      <div class="editor-name-card" role="dialog" aria-modal="true" aria-labelledby="editor-name-title">
        <h2 id="editor-name-title">${reason === "primera" ? "¿Cómo te llamas?" : "Cambiar tu nombre"}</h2>
        <p>Tu nombre aparecerá al pasar el ratón sobre el indicador de presencia, para que los demás editores sepan quién está dentro. Puede ser tu nombre, un alias o lo que quieras.</p>
        <input type="text" class="editor-name-input" maxlength="${EDITOR_NAME_MAX_LENGTH}" placeholder="Tu nombre o alias…" autocomplete="off" />
        <div class="editor-name-actions">
          <button type="button" class="editor-name-submit" disabled>Continuar</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const input = overlay.querySelector(".editor-name-input");
    const submit = overlay.querySelector(".editor-name-submit");

    input.value = initialValue;
    submit.disabled = !sanitiseEditorName(input.value);

    input.addEventListener("input", () => {
      submit.disabled = !sanitiseEditorName(input.value);
    });

    const finish = () => {
      const clean = sanitiseEditorName(input.value);
      if (!clean) return;
      overlay.remove();
      resolve(clean);
    };

    submit.addEventListener("click", finish);
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        finish();
      }
    });

    // Focus the input after the modal has rendered.
    setTimeout(() => input.focus(), 30);
  });
}

async function ensureEditorName({ force = false } = {}) {
  if (!force) {
    const stored = readStoredEditorName();
    if (stored) {
      editorName = stored;
      return editorName;
    }
  }
  const fresh = await promptForEditorName({
    initialValue: force ? editorName : "",
    reason: force ? "cambiar" : "primera",
  });
  editorName = fresh;
  persistEditorName(editorName);
  return editorName;
}

function formatPresenceTooltip(otherNames, savingNames = []) {
  const meLabel = editorName || "Tú";
  const named = otherNames.filter(Boolean);
  const anonCount = otherNames.length - named.length;
  const parts = [meLabel, ...named];
  if (anonCount === 1) parts.push("(1 sin nombre)");
  if (anonCount > 1) parts.push(`(${anonCount} sin nombre)`);
  const base = parts.join(", ");
  if (selfIsSaving) {
    return `${base} · guardando…`;
  }
  if (savingNames && savingNames.length) {
    const who = savingNames.filter(Boolean);
    if (who.length === 1) {
      return `${base} · ${who[0]} está guardando…`;
    }
    if (who.length > 1) {
      return `${base} · ${who.join(", ")} están guardando…`;
    }
    return `${base} · otra sesión está guardando…`;
  }
  return base;
}

function setPresenceTooltipText(text) {
  const el = presenceElement();
  if (!el) return;
  const slot = el.querySelector(".presence-indicator__tooltip");
  if (slot) slot.textContent = text;
}

function renderPresenceState({ otherNames = [], savingNames = [], offline = false, loading = false } = {}) {
  const el = presenceElement();
  if (!el) return;
  if (loading) {
    el.dataset.state = "loading";
    el.querySelector(".presence-indicator__count").textContent = "…";
    setPresenceTooltipText("Comprobando…");
    return;
  }
  if (offline) {
    el.dataset.state = "offline";
    el.querySelector(".presence-indicator__count").textContent = "?";
    setPresenceTooltipText("Sin conexión");
    return;
  }
  const otherCount = otherNames.length;
  const total = otherCount + 1;
  el.querySelector(".presence-indicator__count").textContent = String(total);
  if (otherCount === 0) {
    el.dataset.state = "alone";
  } else if (otherCount === 1) {
    el.dataset.state = "paired";
  } else {
    el.dataset.state = "busy";
  }
  setPresenceTooltipText(formatPresenceTooltip(otherNames, savingNames));
}

// Toggle the global "someone is saving" UI lock: disables the Save button and
// pulses the presence badge whenever any session (this one or another) has
// the saving flag set. Safe to call repeatedly; it's idempotent.
function updateSavingUIState() {
  const someoneSaving = selfIsSaving || otherIsSaving;
  const btn = document.getElementById("save-drive-btn");
  if (btn) {
    if (someoneSaving) {
      btn.disabled = true;
      btn.dataset.savingLock = "1";
      btn.textContent = "Guardando…";
    } else if (btn.dataset.savingLock === "1") {
      // Only restore from our own lock — don't fight whatever the caller had
      // already done (e.g. a manual save in flight already toggled disabled).
      btn.disabled = false;
      delete btn.dataset.savingLock;
      btn.textContent = "GUARDAR";
    }
  }
  const badge = presenceElement();
  if (badge) {
    badge.classList.toggle("is-saving", someoneSaving);
  }
}

// Encode the heartbeat value as `${epoch}|${name}` (idle) or
// `${epoch}|${name}|saving` (during a Drive upload). Old-format payloads
// (epoch only, no separator) are still recognised on read so a rolling deploy
// doesn't break the count.
function encodePresenceValue(nowSec, name, isSaving = false) {
  const cleanName = sanitiseEditorName(name);
  if (isSaving) {
    return cleanName ? `${nowSec}|${cleanName}|saving` : `${nowSec}||saving`;
  }
  return cleanName ? `${nowSec}|${cleanName}` : `${nowSec}`;
}

function decodePresenceValue(raw) {
  const text = `${raw ?? ""}`;
  const parts = text.split("|");
  const epoch = Number.parseInt(parts[0], 10);
  const name = (parts[1] || "").trim();
  const isSaving = parts[2] === "saving";
  return {
    epoch: Number.isInteger(epoch) ? epoch : null,
    name,
    isSaving,
  };
}

async function presenceHeartbeat() {
  if (!window.GoogleDrive?.patchAppProperties) return;
  try {
    const nowSec = Math.floor(Date.now() / 1000);
    await window.GoogleDrive.patchAppProperties({
      [`${PRESENCE_KEY_PREFIX}${PRESENCE_SESSION_ID}`]: encodePresenceValue(nowSec, editorName, selfIsSaving),
    });
  } catch (err) {
    console.warn("[presence] heartbeat failed:", err);
  }
}

async function presencePoll() {
  if (!window.GoogleDrive?.getAppProperties) return;
  try {
    const bag = await window.GoogleDrive.getAppProperties();
    const nowSec = Math.floor(Date.now() / 1000);
    const staleSec = PRESENCE_STALE_MS / 1000;
    const otherNames = [];
    const savingNames = [];
    const staleKeysToEvict = {};
    let anyOtherSaving = false;
    Object.keys(bag).forEach((key) => {
      if (!key.startsWith(PRESENCE_KEY_PREFIX)) return;
      const sessionId = key.slice(PRESENCE_KEY_PREFIX.length);
      const { epoch, name, isSaving } = decodePresenceValue(bag[key]);
      if (epoch === null) return;
      const ageSec = nowSec - epoch;
      if (ageSec > staleSec) {
        // Opportunistically evict stale entries so the bag does not grow
        // unbounded over months of crashed tabs.
        staleKeysToEvict[key] = null;
        return;
      }
      if (sessionId !== PRESENCE_SESSION_ID) {
        otherNames.push(name || "");
        if (isSaving && ageSec <= PRESENCE_SAVING_FLAG_MAX_AGE_S) {
          anyOtherSaving = true;
          savingNames.push(name || "");
        }
      }
    });
    renderPresenceState({ otherNames, savingNames });
    if (anyOtherSaving !== otherIsSaving) {
      otherIsSaving = anyOtherSaving;
      updateSavingUIState();
    }
    if (Object.keys(staleKeysToEvict).length) {
      // Fire-and-forget; if it fails, next poll will try again.
      window.GoogleDrive.patchAppProperties(staleKeysToEvict).catch(() => {});
    }
  } catch (err) {
    console.warn("[presence] poll failed:", err);
    renderPresenceState({ offline: true });
  }
}

async function startPresenceTracking() {
  if (presenceStarted || IS_VIEWER_MODE) return;
  if (!window.GoogleDrive?.patchAppProperties) return;
  presenceStarted = true;
  renderPresenceState({ loading: true });
  // Make sure we have an identity before the first heartbeat, so other
  // sessions immediately see the name rather than a fleeting anonymous tick.
  await ensureEditorName();
  // Initial heartbeat + poll right away so the user sees feedback fast.
  presenceHeartbeat().then(presencePoll);
  presenceHeartbeatTimer = setInterval(presenceHeartbeat, PRESENCE_HEARTBEAT_MS);
  presencePollTimer = setInterval(presencePoll, PRESENCE_POLL_MS);
  // Double-click on the badge reopens the name prompt — single click stays
  // inert because the badge is informative, not actionable.
  const el = presenceElement();
  if (el) {
    el.addEventListener("dblclick", async () => {
      await ensureEditorName({ force: true });
      // Push the new name to Drive right away and refresh the local tooltip.
      presenceHeartbeat().then(presencePoll);
    });
  }
  // Best-effort cleanup of our own entry when the tab closes. Drive may
  // ignore the request if the browser is mid-teardown, in which case the
  // entry will simply go stale on its own within PRESENCE_STALE_MS.
  window.addEventListener("beforeunload", () => {
    try {
      window.GoogleDrive.patchAppProperties({
        [`${PRESENCE_KEY_PREFIX}${PRESENCE_SESSION_ID}`]: null,
      });
    } catch (_) { /* ignore */ }
  });
}

// =============================================================================
// CHANGE HISTORY ("Últimos cambios")
//
// A sidecar JSON file in Drive (PANEL_CONTROL_HISTORIAL.json) accumulates a
// rolling log of edits. Every successful save derives entries from the local
// delta and prepends them to the file. A side panel renders the most recent
// HISTORY_MAX_ENTRIES (=500) entries with a click-to-jump affordance into
// the grid below.
//
// File discovery: id is stored in the Excel's appProperties under key
// `historyFileId` so every session of the shared account finds it instantly.
// Cold boot falls back to `findJsonFileByName` and finally creates a new
// file if none exists.
// =============================================================================
const HISTORY_FILE_NAME = "PANEL_CONTROL_HISTORIAL.json";
const HISTORY_APP_PROP_KEY = "historyFileId";
const HISTORY_MAX_ENTRIES = 500;
const HISTORY_FLASH_MS = 2200;

let historyFileId = null;
let historyEntries = [];
let historyLoaded = false;
let historyLoading = false;
let historyPanelOpen = false;

// Map internal column keys to user-facing labels for the history cards.
const HISTORY_COLUMN_LABEL = {
  listo: "Listo",
  title: "Título",
  startDate: "Inicio Vig.",
  endDate: "Fin Vig.",
  genre: "Género",
  id: "ID",
  actualizado: "Actualizado",
  listoByMonth: "Listo",
};

function formatHistoryColumn(columnKey) {
  return HISTORY_COLUMN_LABEL[columnKey] || columnKey;
}

function formatHistoryRelativeTime(ts) {
  const then = new Date(ts);
  if (Number.isNaN(then.getTime())) return "";
  const diffMs = Date.now() - then.getTime();
  const diffMin = Math.floor(diffMs / 60_000);
  if (diffMin < 1) return "ahora mismo";
  if (diffMin < 60) return `hace ${diffMin} min`;
  const diffHr = Math.floor(diffMin / 60);
  if (diffHr < 24) return `hace ${diffHr} h`;
  const now = new Date();
  const sameDay = then.toDateString() === now.toDateString();
  if (sameDay) return then.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
  const yesterday = new Date(now);
  yesterday.setDate(now.getDate() - 1);
  if (then.toDateString() === yesterday.toDateString()) {
    return `ayer, ${then.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" })}`;
  }
  return then.toLocaleString("es-ES", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });
}

function formatHistoryValue(value) {
  if (value === null || value === undefined || value === "") return "(vacío)";
  if (typeof value === "boolean") return value ? "Sí" : "No";
  if (typeof value === "object") {
    // listoByMonth maps look like {"YYYY-MM": true, ...}
    const months = Object.keys(value).filter((k) => value[k]).sort();
    return months.length ? months.join(", ") : "(vacío)";
  }
  return String(value);
}

// Derive a stable initial-circle colour from an editor name so each user has
// a consistent looking avatar across sessions.
function colourForEditorName(name) {
  const palette = ["#2eb84e", "#f1ae15", "#1a73e8", "#9b59b6", "#e8442b", "#16a085", "#d35400", "#7f8c8d"];
  if (!name) return palette[palette.length - 1];
  let h = 0;
  for (let i = 0; i < name.length; i += 1) h = (h * 31 + name.charCodeAt(i)) | 0;
  return palette[Math.abs(h) % palette.length];
}

// Verify that a candidate history file actually belongs to *this* Excel.
// We embed the Excel's FILE_ID into the history file's content; anything
// without that marker, or with a mismatching one, is rejected so dev and
// prod (which share the same Drive account) cannot accidentally share a
// single history file.
async function isHistoryFileForThisExcel(fileId, excelFileId) {
  try {
    const doc = await window.GoogleDrive.readJsonFile(fileId);
    return doc?.excelFileId === excelFileId;
  } catch (_) {
    return false;
  }
}

// Locate the history file once per session. Priority:
//   1. Excel.appProperties.historyFileId (fast, cross-session) — validated
//   2. Drive search by HISTORY_FILE_NAME — validated
//   3. Create a fresh file with excelFileId marker, save the id into
//      appProperties.
async function ensureHistoryFile() {
  if (historyFileId) return historyFileId;
  if (!window.GoogleDrive?.getAppProperties) return null;
  const excelFileId = window.PANEL_CONFIG?.GOOGLE_DRIVE_FILE_ID || "";

  try {
    const appProperties = await window.GoogleDrive.getAppProperties();
    const cachedId = appProperties?.[HISTORY_APP_PROP_KEY];
    if (cachedId && await isHistoryFileForThisExcel(cachedId, excelFileId)) {
      historyFileId = cachedId;
      return historyFileId;
    }
    // Cached id is missing, dead, or pointing at a different Excel's
    // history (e.g. dev/prod cross-contamination from when they shared
    // a name). Fall through to discovery / fresh creation.
  } catch (_) { /* ignore — fall through */ }

  try {
    const found = await window.GoogleDrive.findJsonFileByName(HISTORY_FILE_NAME);
    if (found && await isHistoryFileForThisExcel(found, excelFileId)) {
      historyFileId = found;
      await window.GoogleDrive.patchAppProperties({ [HISTORY_APP_PROP_KEY]: found });
      return historyFileId;
    }
  } catch (_) { /* ignore — fall through to creation */ }

  try {
    const created = await window.GoogleDrive.createJsonFile(HISTORY_FILE_NAME, {
      version: 1,
      excelFileId,
      entries: [],
    });
    historyFileId = created;
    await window.GoogleDrive.patchAppProperties({ [HISTORY_APP_PROP_KEY]: created });
    return historyFileId;
  } catch (err) {
    console.warn("[history] could not create history file:", err);
    return null;
  }
}

async function loadHistory() {
  if (historyLoading) return;
  historyLoading = true;
  try {
    const fileId = await ensureHistoryFile();
    if (!fileId) return;
    const doc = await window.GoogleDrive.readJsonFile(fileId);
    historyEntries = Array.isArray(doc?.entries) ? doc.entries : [];
    historyLoaded = true;
    if (historyPanelOpen) renderHistoryPanelContents();
  } catch (err) {
    console.warn("[history] could not load history:", err);
  } finally {
    historyLoading = false;
  }
}

function formatHomeMonthLabel(homeMonth, homeYear) {
  if (!Number.isInteger(homeMonth) || !Number.isInteger(homeYear)) return "";
  const name = MONTH_NAMES_ES[homeMonth - 1];
  if (!name) return "";
  return `${name.toUpperCase()} ${homeYear}`;
}

// Build human-friendly entries from a save delta. We use the row's current
// title as a snapshot so the entry is meaningful even if the row title is
// edited later or the row is deleted.
function buildHistoryEntriesFromDelta(delta, srcBlocks, editor, baselineSnapshot) {
  const nowIso = new Date().toISOString();
  const out = [];

  // Build a quick lookup from rowKey → { rowTitle, blockType, monthLabel }.
  const rowMeta = new Map();
  srcBlocks.forEach((block) => {
    block.rows?.forEach((row) => {
      if (!row?.rowKey || row._autoPlaceholder) return;
      rowMeta.set(row.rowKey, {
        title: row.title || "(sin título)",
        blockType: block.blockType || "",
        monthLabel: formatHomeMonthLabel(row.homeMonth, row.homeYear),
      });
    });
  });
  // For deleted rows we look back at the snapshot.
  const snapshotMeta = new Map();
  baselineSnapshot?.forEach((block) => {
    block.rows?.forEach((row) => {
      if (!row?.rowKey || row._autoPlaceholder) return;
      snapshotMeta.set(row.rowKey, {
        title: row.title || "(sin título)",
        blockType: block.blockType || "",
        monthLabel: formatHomeMonthLabel(row.homeMonth, row.homeYear),
      });
    });
  });

  // Cell edits — we need before/after, which the delta itself doesn't carry
  // for cellChanges (it has only `after`). Derive `before` from the snapshot.
  const snapshotRows = new Map();
  baselineSnapshot?.forEach((block) => {
    block.rows?.forEach((row) => {
      if (row?.rowKey) snapshotRows.set(row.rowKey, row);
    });
  });

  // Also keep raw month / year so the panel can navigate the calendar when
  // the user clicks a card whose row lives in a different month.
  const rowMonthByKey = new Map();
  srcBlocks.forEach((block) => {
    block.rows?.forEach((row) => {
      if (!row?.rowKey) return;
      rowMonthByKey.set(row.rowKey, { homeMonth: row.homeMonth, homeYear: row.homeYear });
    });
  });
  baselineSnapshot?.forEach((block) => {
    block.rows?.forEach((row) => {
      if (!row?.rowKey || rowMonthByKey.has(row.rowKey)) return;
      rowMonthByKey.set(row.rowKey, { homeMonth: row.homeMonth, homeYear: row.homeYear });
    });
  });

  delta.cellChanges?.forEach(({ rowKey, field, value }) => {
    const meta = rowMeta.get(rowKey) || snapshotMeta.get(rowKey) || {};
    const snapRow = snapshotRows.get(rowKey);
    const before = snapRow ? snapRow[field === "startDateText" ? "startDateText" : field === "endDateText" ? "endDateText" : field] : undefined;
    const rowMonth = rowMonthByKey.get(rowKey) || {};
    out.push({
      ts: nowIso,
      editor: editor || "Anónimo",
      kind: "cell",
      rowKey,
      rowTitle: meta.title || "",
      blockType: meta.blockType || "",
      monthLabel: meta.monthLabel || "",
      homeMonth: rowMonth.homeMonth,
      homeYear: rowMonth.homeYear,
      column: field,
      before: formatHistoryValue(before),
      after: formatHistoryValue(value),
    });
  });

  delta.newRows?.forEach(({ rowSnapshot, blockId }) => {
    const block = srcBlocks.find((b) => b.id === blockId);
    out.push({
      ts: nowIso,
      editor: editor || "Anónimo",
      kind: "add",
      rowKey: rowSnapshot?.rowKey || "",
      rowTitle: rowSnapshot?.title || "(sin título)",
      blockType: block?.blockType || "",
      monthLabel: formatHomeMonthLabel(rowSnapshot?.homeMonth, rowSnapshot?.homeYear),
      homeMonth: rowSnapshot?.homeMonth,
      homeYear: rowSnapshot?.homeYear,
    });
  });

  delta.deletedRowKeys?.forEach((rowKey) => {
    const meta = snapshotMeta.get(rowKey) || {};
    const rowMonth = rowMonthByKey.get(rowKey) || {};
    out.push({
      ts: nowIso,
      editor: editor || "Anónimo",
      kind: "delete",
      rowKey,
      rowTitle: meta.title || "(sin título)",
      blockType: meta.blockType || "",
      monthLabel: meta.monthLabel || "",
      homeMonth: rowMonth.homeMonth,
      homeYear: rowMonth.homeYear,
    });
  });

  return out;
}

// Fire-and-forget append after a successful save. We re-read the file so two
// concurrent saves don't trample each other's appends; under heavy contention
// the latest writer still wins, but the loss is at most a few entries from
// the *history*, not from the data itself.
async function appendHistoryEntries(newEntries) {
  if (!newEntries.length) return;
  const fileId = await ensureHistoryFile();
  if (!fileId) return;
  const excelFileId = window.PANEL_CONFIG?.GOOGLE_DRIVE_FILE_ID || "";
  try {
    let current;
    try {
      current = await window.GoogleDrive.readJsonFile(fileId);
    } catch (_) {
      current = { version: 1, excelFileId, entries: [] };
    }
    const existing = Array.isArray(current?.entries) ? current.entries : [];
    const merged = [...newEntries, ...existing].slice(0, HISTORY_MAX_ENTRIES);
    await window.GoogleDrive.writeJsonFile(fileId, {
      version: 1,
      excelFileId,
      entries: merged,
    });
    historyEntries = merged;
    historyLoaded = true;
    if (historyPanelOpen) renderHistoryPanelContents();
  } catch (err) {
    console.warn("[history] could not append entries:", err);
  }
}

// =============================================================================
// HISTORY PANEL UI
// =============================================================================

function ensureHistoryPanelElement() {
  let panel = document.getElementById("history-panel");
  if (panel) return panel;
  panel = document.createElement("aside");
  panel.id = "history-panel";
  panel.className = "history-panel";
  panel.setAttribute("aria-label", "Últimos cambios");
  panel.innerHTML = `
    <header class="history-panel__header">
      <h2 class="history-panel__title">Últimos cambios</h2>
      <div class="history-panel__actions">
        <button type="button" class="history-panel__refresh" aria-label="Refrescar lista">↻</button>
        <button type="button" class="history-panel__close" aria-label="Cerrar panel">✕</button>
      </div>
    </header>
    <div class="history-panel__list" id="history-panel-list">
      <div class="history-panel__loading">Cargando…</div>
    </div>
  `;
  document.body.appendChild(panel);
  panel.querySelector(".history-panel__refresh").addEventListener("click", () => {
    historyLoaded = false;
    loadHistory();
  });
  panel.querySelector(".history-panel__close").addEventListener("click", () => {
    closeHistoryPanel();
  });
  panel.querySelector(".history-panel__list").addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;

    // "Ir a la fila" button — takes precedence over the toggle behaviour so
    // expanding remains separate from navigating.
    const jumpBtn = target.closest(".history-card__jump");
    if (jumpBtn) {
      const rowKey = jumpBtn.dataset.rowKey;
      if (!rowKey) return;
      const targetMonth = Number.parseInt(jumpBtn.dataset.homeMonth, 10);
      const targetYear = Number.parseInt(jumpBtn.dataset.homeYear, 10);
      flashRowByKey(rowKey, {
        targetMonth: Number.isInteger(targetMonth) ? targetMonth : null,
        targetYear: Number.isInteger(targetYear) ? targetYear : null,
      });
      return;
    }

    // Otherwise click anywhere on the card → accordion toggle.
    const card = target.closest(".history-card");
    if (!card) return;
    const wasExpanded = card.classList.contains("is-expanded");
    // Collapse any other expanded card first (single-expand accordion).
    panel.querySelectorAll(".history-card.is-expanded").forEach((c) => {
      c.classList.remove("is-expanded");
    });
    if (!wasExpanded) card.classList.add("is-expanded");
  });
  return panel;
}

function renderHistoryPanelContents() {
  const panel = ensureHistoryPanelElement();
  const list = panel.querySelector(".history-panel__list");
  if (!list) return;
  if (!historyLoaded && historyLoading) {
    list.innerHTML = `<div class="history-panel__loading">Cargando…</div>`;
    return;
  }
  if (!historyEntries.length) {
    list.innerHTML = `<div class="history-panel__empty">Aún no hay cambios registrados.</div>`;
    return;
  }
  const html = historyEntries.map((entry) => historyCardHtml(entry)).join("");
  list.innerHTML = html;
}

function historyCardHtml(entry) {
  const initial = (entry.editor || "?").trim().slice(0, 1).toUpperCase();
  const colour = colourForEditorName(entry.editor);
  const when = formatHistoryRelativeTime(entry.ts);
  const blockText = entry.blockType ? entry.blockType.toUpperCase() : "";
  // Prefix the block label with "MES AÑO -" when we know which calendar
  // month the row lives in. Entries logged before this change won't have it
  // and will fall back gracefully to just the block name.
  const blockLabel = entry.monthLabel
    ? `${entry.monthLabel} – ${blockText}`
    : blockText;
  const rowTitle = entry.rowTitle || "(sin título)";

  // Summary line — visible both collapsed and expanded. Minimal info so the
  // user can recognise the change without having to open the card.
  let summary;
  if (entry.kind === "add") {
    summary = `<div class="history-card__summary history-card__summary--add">+ Añadió <strong>${escapeHtml(rowTitle)}</strong></div>`;
  } else if (entry.kind === "delete") {
    summary = `<div class="history-card__summary history-card__summary--delete">− Borró <strong>${escapeHtml(rowTitle)}</strong></div>`;
  } else {
    summary = `<div class="history-card__summary">Editó <strong>${escapeHtml(rowTitle)}</strong></div>`;
  }

  // Detail section — hidden until the card is expanded. For cell edits shows
  // column + before/after. Always ends with the "Ir a la fila" button.
  let details = "";
  if (entry.kind === "cell") {
    const col = formatHistoryColumn(entry.column);
    details = `
      <div class="history-card__meta">${escapeHtml(col)}</div>
      <div class="history-card__before">◯ ${escapeHtml(entry.before ?? "(vacío)")}</div>
      <div class="history-card__after">● ${escapeHtml(entry.after ?? "(vacío)")}</div>
    `;
  }
  const jumpBtn = `
    <div class="history-card__details-footer">
      <button type="button"
              class="history-card__jump"
              data-row-key="${escapeAttr(entry.rowKey || "")}"
              data-home-month="${escapeAttr(Number.isInteger(entry.homeMonth) ? entry.homeMonth : "")}"
              data-home-year="${escapeAttr(Number.isInteger(entry.homeYear) ? entry.homeYear : "")}">
        Ir a la fila
        <svg viewBox="0 0 24 24" width="11" height="11" aria-hidden="true" focusable="false">
          <path d="M5 12h14M13 5l7 7-7 7" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </button>
    </div>
  `;

  return `
    <article class="history-card"
             data-row-key="${escapeAttr(entry.rowKey || "")}"
             data-home-month="${escapeAttr(Number.isInteger(entry.homeMonth) ? entry.homeMonth : "")}"
             data-home-year="${escapeAttr(Number.isInteger(entry.homeYear) ? entry.homeYear : "")}">
      <div class="history-card__avatar" style="background:${colour}">${escapeHtml(initial)}</div>
      <div class="history-card__body">
        <header class="history-card__header">
          <span class="history-card__name">${escapeHtml(entry.editor || "Anónimo")}</span>
          <span class="history-card__when">${escapeHtml(when)}</span>
        </header>
        ${blockLabel ? `<div class="history-card__block">${escapeHtml(blockLabel)}</div>` : ""}
        ${summary}
        <div class="history-card__details">
          ${details}
          ${jumpBtn}
        </div>
      </div>
    </article>
  `;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}

function openHistoryPanel() {
  if (historyPanelOpen) return;
  historyPanelOpen = true;
  const panel = ensureHistoryPanelElement();
  panel.classList.add("is-open");
  document.body.classList.add("history-panel-open");
  // Show whatever we have cached immediately, then refetch in the background
  // so opening the panel always reflects the latest saved changes from any
  // session — no need for the user to press the manual refresh button.
  renderHistoryPanelContents();
  loadHistory();
}

function closeHistoryPanel() {
  historyPanelOpen = false;
  const panel = document.getElementById("history-panel");
  if (panel) panel.classList.remove("is-open");
  document.body.classList.remove("history-panel-open");
}

function toggleHistoryPanel() {
  if (historyPanelOpen) closeHistoryPanel();
  else openHistoryPanel();
}

// Internal: do the scroll + flash, assuming the row is already in the current
// month view. Returns true on success, false if the row isn't in the DOM.
function flashRowInCurrentView(rowKey) {
  const leftBody = document.getElementById("left-body");
  const rightBody = document.getElementById("right-body");
  if (!leftBody || !rightBody) return false;
  const target = leftBody.querySelector(`[data-row-id="${CSS.escape(rowKey)}"]`);
  if (!target) return false;
  const leftRow = target.closest(".left-row");
  if (!leftRow) return false;
  leftRow.scrollIntoView({ behavior: "smooth", block: "center" });
  leftRow.classList.add("is-flash-highlight");
  const allLeftRows = Array.from(leftBody.children);
  const idx = allLeftRows.indexOf(leftRow);
  const dayRow = idx >= 0 ? rightBody.children[idx] : null;
  dayRow?.classList.add("is-flash-highlight");
  setTimeout(() => {
    leftRow.classList.remove("is-flash-highlight");
    dayRow?.classList.remove("is-flash-highlight");
  }, HISTORY_FLASH_MS);
  return true;
}

// Scroll to the row whose rowKey matches and briefly flash it. If the entry
// has a target month/year and we're currently on a different one, navigate
// the calendar first, then flash once the new month finishes rendering.
function flashRowByKey(rowKey, { targetMonth = null, targetYear = null } = {}) {
  if (!rowKey) return;
  // Fast path: already showing the right month, or the entry has no month
  // info (legacy entries from before this field existed).
  const needsMonthShift =
    Number.isInteger(targetMonth) && Number.isInteger(targetYear)
    && (targetMonth !== currentCalendarContext.month || targetYear !== currentCalendarContext.year);

  if (!needsMonthShift) {
    if (!flashRowInCurrentView(rowKey)) {
      showGridToast("Esa fila ya no existe en el panel actual");
    }
    return;
  }

  // Navigate the calendar to the target month, then wait for the re-render to
  // complete before locating the row and flashing it.
  currentCalendarContext = {
    month: targetMonth,
    year: targetYear,
    daysInMonth: daysInMonth(targetMonth, targetYear),
  };
  applyCalendarContextToView(document);
  // Two RAFs: first one lets the new month layout settle, second guarantees
  // the new rows are in the DOM by the time we look for them.
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      if (!flashRowInCurrentView(rowKey)) {
        showGridToast("Esa fila ya no existe en el panel actual");
      }
    });
  });
}

async function saveToGoogleDrive() {
  if (!window.GoogleDrive) {
    showGridToast("Servicio no disponible");
    return;
  }
  if (selfIsSaving) {
    // Defensive: button is disabled while we save, but pre-empt anything that
    // managed to slip through anyway.
    return;
  }

  // Pre-check: read the presence bag once before doing anything so we can
  // refuse the save if somebody else has just started theirs. This catches
  // the race that periodic polling alone cannot cover (their flag may have
  // been written between two of our polls).
  if (window.GoogleDrive.getAppProperties) {
    try {
      const bag = await window.GoogleDrive.getAppProperties();
      const nowSec = Math.floor(Date.now() / 1000);
      let conflictName = "";
      for (const key of Object.keys(bag)) {
        if (!key.startsWith(PRESENCE_KEY_PREFIX)) continue;
        const sessionId = key.slice(PRESENCE_KEY_PREFIX.length);
        if (sessionId === PRESENCE_SESSION_ID) continue;
        const { epoch, name, isSaving } = decodePresenceValue(bag[key]);
        if (!isSaving || epoch === null) continue;
        if (nowSec - epoch > PRESENCE_SAVING_FLAG_MAX_AGE_S) continue;
        conflictName = name || "Otra sesión";
        break;
      }
      if (conflictName) {
        showGridToast(`${conflictName} está guardando ahora mismo · espera unos segundos y reintenta`);
        // Reflect the conflict in the UI immediately even before the next poll
        // (which would have caught it on its own within PRESENCE_POLL_MS).
        if (!otherIsSaving) {
          otherIsSaving = true;
          updateSavingUIState();
        }
        return;
      }
    } catch (err) {
      // Non-fatal: proceed without the pre-check. The merge logic still keeps
      // us safe from data loss.
      console.warn("[save] pre-check failed, proceeding:", err);
    }
  }

  selfIsSaving = true;
  updateSavingUIState();
  // Push our saving flag immediately so other sessions see it on their next
  // poll (≤ PRESENCE_POLL_MS ahead) or pre-check.
  await presenceHeartbeat();
  // Boost heartbeat cadence so the flag never goes stale during long saves.
  if (presenceSaveHeartbeatTimer) clearInterval(presenceSaveHeartbeatTimer);
  presenceSaveHeartbeatTimer = setInterval(presenceHeartbeat, PRESENCE_SAVE_HEARTBEAT_MS);

  showGridToast("Guardando...");
  try {
    // Fast-path: never loaded from Drive in this session → no merge possible,
    // just overwrite. Happens only on the first save after a draft restore
    // before Drive has been polled.
    if (!loadedSnapshot) {
      const { buffer, sheetCount } = await buildExcelEdicionBuffer();
      await window.GoogleDrive.saveXlsxBuffer(buffer);
      loadedSnapshot = deepCloneBlocks(blocks);
      clearDraft();
      showGridToast(`Guardado · ${sheetCount} hoja(s)`);
      return;
    }

    // Merge path: compute local diff vs. snapshot, pull latest from Drive,
    // replay diff on top, upload merged result, adopt as new state.
    const localDelta = computeRowDelta(loadedSnapshot, blocks);
    // Snapshot before we mutate it — needed to derive "before" values for
    // the history entries we'll append after the save succeeds.
    const historyBaseline = loadedSnapshot;
    const historySrcBlocks = blocks;
    const remoteBuffer = await window.GoogleDrive.loadXlsxBuffer({ useAuth: true });
    const remoteBlocks = parseBufferToBlocks(remoteBuffer);
    applyRowDelta(remoteBlocks, localDelta);
    preserveBlockCollapsedState(remoteBlocks, blocks);

    const { buffer, sheetCount } = await buildExcelEdicionBuffer(remoteBlocks);
    await window.GoogleDrive.saveXlsxBuffer(buffer);

    blocks = remoteBlocks;
    loadedSnapshot = deepCloneBlocks(blocks);
    clearDraft();
    renderRows();
    showGridToast(`Guardado · ${sheetCount} hoja(s)`);

    // Fire-and-forget: log the edit deltas into the change-history file so
    // other editors (and this one) can review them later via the side panel.
    const newHistoryEntries = buildHistoryEntriesFromDelta(
      localDelta, historySrcBlocks, editorName, historyBaseline
    );
    if (newHistoryEntries.length) {
      appendHistoryEntries(newHistoryEntries).catch(() => {});
    }
  } catch (err) {
    console.error("saveToGoogleDrive error:", err);
    showGridToast("Error al guardar · tus cambios siguen a salvo en este equipo");
  } finally {
    // Always release the lock: clear the flag locally, kill the boosted
    // heartbeat timer, and push one final heartbeat so other sessions stop
    // showing the "guardando" overlay within their next poll.
    if (presenceSaveHeartbeatTimer) {
      clearInterval(presenceSaveHeartbeatTimer);
      presenceSaveHeartbeatTimer = null;
    }
    selfIsSaving = false;
    updateSavingUIState();
    // Fire-and-forget; if the network is gone the entry will become stale and
    // eventually be ignored by other sessions on its own.
    presenceHeartbeat().catch(() => {});
  }
}

// Shift an ISO date (YYYY-MM-DD) by `deltaDays` and return it formatted as
// DD/MM/YY — same display format the rest of the panel uses. Returns "" if
// the input is missing or not parseable, so callers can fall through with
// `|| null`.
function shiftDateTextByDays(isoString, deltaDays) {
  if (!isoString || typeof isoString !== "string") return "";
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoString);
  if (!match) return "";
  const year = Number.parseInt(match[1], 10);
  const month = Number.parseInt(match[2], 10);
  const day = Number.parseInt(match[3], 10);
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return "";
  const d = new Date(year, month - 1, day);
  d.setDate(d.getDate() + deltaDays);
  return formatDateDisplay(d.getDate(), d.getMonth() + 1, d.getFullYear());
}

async function exportExcelAplicativo() {
  if (!window.ExcelJS) {
    showGridToast("No se pudo cargar la librería ExcelJS");
    return;
  }

  const TIPO_MAP = {
    "promo 20":                   "Promo",
    "promo 40":                   "Promo",
    "otras duraciones":           "Otras Duraciones",
    "combo":                      "Combo",
    "bumper":                     "Bumper",
    "id":                         "ID",
    "pasos a publi":              "Paso a Publi",
    "intruso":                    "Intruso",
    "loop proteccion pop-ups":    "Loop",
    "loop protección pop-ups":    "Loop",
    "canales laliga":             "Promo",
    "canales golf":               "Promo",
    "canales caza y pesca":       "Promo",
    "arranque":                   "Arranque",
    "loop":                       "Loop",
    "pre roll":                   "Preroll",
  };

  function getTipo(blockType) {
    const key = (blockType || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/\p{Diacritic}/gu, "");
    return TIPO_MAP[key] || blockType;
  }

  const COLOR_BLUE_HDR = "FF4472C4";
  const COLOR_WHITE    = "FFFFFFFF";

  function applyBlueHeader(cell) {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_BLUE_HDR } };
    cell.font = { bold: true, color: { argb: COLOR_WHITE } };
  }

  const COL_DEFS = [
    { header: "titulo",                    width: 60,  hidden: false },
    { header: "fecha_entrega",             width: 14,  hidden: false },
    { header: "fecha_programacion",        width: 14,  hidden: false },
    { header: "tipo",                      width: 18,  hidden: false },
    { header: "requiere_guion",            width: 14,  hidden: false },
    { header: "jefe_guion",                width: 20,  hidden: true  },
    { header: "fecha_limite_guion",        width: 14,  hidden: true  },
    { header: "requiere_realiza",          width: 14,  hidden: true  },
    { header: "jefe_realiza",              width: 20,  hidden: true  },
    { header: "fecha_entrega_realiza",     width: 14,  hidden: true  },
    { header: "requiere_diseno",           width: 14,  hidden: true  },
    { header: "jefe_diseno",               width: 20,  hidden: true  },
    { header: "fecha_limite_diseno",       width: 14,  hidden: true  },
    { header: "requiere_ambientacion",     width: 14,  hidden: true  },
    { header: "jefe_ambientacion",         width: 20,  hidden: true  },
    { header: "fecha_limite_ambientacion", width: 14,  hidden: true  },
    { header: "requiere_locucion",         width: 14,  hidden: true  },
    { header: "jefe_locucion",             width: 20,  hidden: true  },
    { header: "fecha_limite_locucion",     width: 14,  hidden: true  },
    { header: "material",                  width: 14,  hidden: false },
    { header: "observacion_material",      width: 24,  hidden: false },
    { header: "link",                      width: 30,  hidden: false },
    { header: "EE",                        width: 10,  hidden: false },
    { header: "calificacion",              width: 14,  hidden: false },
    { header: "etiquetas",                 width: 14,  hidden: false },
    { header: "etiquetas",                 width: 14,  hidden: true  },
    { header: "etiquetas",                 width: 14,  hidden: false },
    { header: "notas",                     width: 24,  hidden: true  },
    { header: "genero",                    width: 14,  hidden: false },
  ];

  const { month, year } = currentCalendarContext;

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Hoja1");

  // Anchos y columnas ocultas
  ws.columns = COL_DEFS.map((col) => ({ width: col.width }));
  COL_DEFS.forEach((col, i) => {
    if (col.hidden) ws.getColumn(i + 1).hidden = true;
  });

  // Fila de cabecera azul
  const headerRow = ws.addRow(COL_DEFS.map((col) => col.header));
  headerRow.eachCell((cell) => applyBlueHeader(cell));
  headerRow.commit();

  // Todas las filas de datos del mes, sin separadores de bloque.
  // Filtramos por la fecha REAL de inicio de vigencia (no por el mes donde
  // vive la fila en el panel) para evitar que una misma promo aparezca en
  // dos meses de Aplicativo distintos cuando se cruza entre meses.
  const isInExportMonth = (row) => {
    if (!row?.startDateISO) return false;
    const m = /^(\d{4})-(\d{2})-/.exec(row.startDateISO);
    if (!m) return false;
    return Number.parseInt(m[1], 10) === year && Number.parseInt(m[2], 10) === month;
  };
  blocks.forEach((block) => {
    if (block.isSeparator) return;

    const monthRows = block.rows.filter(
      (row) => !isPlaceholderRow(row) && isInExportMonth(row)
    );

    monthRows.forEach((row) => {
      const titulo = `${block.blockType.toUpperCase()} - ${row.title || ""}`.trim();
      const tipo   = getTipo(block.blockType);

      const values = Array(COL_DEFS.length).fill(null);
      values[0]  = titulo;
      // fecha_entrega: la fecha de inicio de vigencia MENOS una semana,
      // porque la entrega del material al aplicativo debe ir 7 días por
      // delante de la fecha de inicio en el panel.
      values[1]  = shiftDateTextByDays(row.startDateISO, -7) || null;
      // fecha_programacion: la fecha de inicio de vigencia (antes era el fin).
      values[2]  = row.startDateText || null;
      values[3]  = tipo;
      values[19] = "OK";        // material
      values[23] = "no lleva";  // calificacion
      values[28] = row.genre         || null;

      const dataRow = ws.addRow(values);
      dataRow.getCell(2).numFmt = "@";
      dataRow.getCell(3).numFmt = "@";
      dataRow.commit();
    });
  });

  // Generar y descargar
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a   = document.createElement("a");
  a.href    = url;

  // Nombre: Carga_Marzo_2026
  const monthNameCap = MONTH_NAMES_ES[month - 1].charAt(0).toUpperCase()
    + MONTH_NAMES_ES[month - 1].slice(1);
  a.download = `Carga_${monthNameCap}_${year}.xlsx`;

  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  showGridToast(`Excel Aplicativo exportado · ${monthNameCap} ${year}`);
}

function attachExcelExportControls(root) {
  const exportBtn = root.querySelector(".export-excel-btn");
  if (exportBtn) {
    exportBtn.addEventListener("click", () => exportExcelAplicativo());
  }

  if (IS_VIEWER_MODE) return;

  const saveDriveBtn = root.querySelector("#save-drive-btn");
  if (saveDriveBtn) {
    // The button's disabled state is fully driven by updateSavingUIState now,
    // which mirrors the (self / other) saving flags in real time. No manual
    // disable/enable here — that would race with the global state.
    saveDriveBtn.addEventListener("click", () => { saveToGoogleDrive(); });
  }

  const historyToggleBtn = root.querySelector("#history-toggle-btn");
  if (historyToggleBtn) {
    historyToggleBtn.addEventListener("click", () => { toggleHistoryPanel(); });
  }
}

function parseISODateValue(value) {
  const normalized = `${value ?? ""}`.trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    return null;
  }

  const [yearText, monthText, dayText] = normalized.split("-");
  const year = Number.parseInt(yearText, 10);
  const month = Number.parseInt(monthText, 10);
  const day = Number.parseInt(dayText, 10);
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return null;
  }

  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(month, year)) {
    return null;
  }

  return new Date(year, month - 1, day);
}

// Día 1 del mes de origen de la fila (inicio implícito cuando no hay fecha de inicio).
function getRowHomeMonthStart(row) {
  const homeMonth = Number.isInteger(row?.homeMonth) ? row.homeMonth : currentCalendarContext.month;
  const homeYear = Number.isInteger(row?.homeYear) ? row.homeYear : currentCalendarContext.year;
  return new Date(homeYear, homeMonth - 1, 1);
}

// True si la fila tiene fecha de fin pero NO fecha de inicio: el inicio se
// interpreta implícitamente como el día 1 de su mes de origen y la entrada se
// propaga a los meses siguientes hasta la fecha de fin.
function rowHasImplicitStart(row) {
  if (!row) {
    return false;
  }
  if (parseISODateValue(row.startDateISO)) {
    return false;
  }
  return !!parseISODateValue(row.endDateISO);
}

function getRowRange(row) {
  if (!row) {
    return null;
  }

  let startDate = parseISODateValue(row.startDateISO);
  const endDate = parseISODateValue(row.endDateISO);
  if (!endDate) {
    return null;
  }

  // Sin fecha de inicio pero con fecha de fin → inicio implícito = día 1 del mes de origen.
  if (!startDate) {
    startDate = getRowHomeMonthStart(row);
  }

  if (endDate < startDate) {
    return null;
  }

  return { startDate, endDate };
}

function normalizeForSearch(text) {
  return `${text ?? ""}`
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "");
}

function rowMatchesSearch(row, normalizedQuery) {
  if (!normalizedQuery) return true;
  return normalizeForSearch(row.title).includes(normalizedQuery);
}

function blockHasMatchForSearch(block, normalizedQuery) {
  if (!normalizedQuery || block.isSeparator) return true;
  return block.rows.some(
    (row) => !isPlaceholderRow(row) && rowMatchesSearch(row, normalizedQuery)
  );
}

function applySearch(rawQuery) {
  const normalized = normalizeForSearch(rawQuery);

  if (normalized && !preSearchCollapseState) {
    preSearchCollapseState = new Map(blocks.map((b) => [b.id, b.collapsed]));
  }

  if (normalized) {
    blocks = blocks.map((block) => {
      if (block.isSeparator) return block;
      const hasMatch = blockHasMatchForSearch(block, normalized);
      return { ...block, collapsed: !hasMatch };
    });
  } else if (preSearchCollapseState) {
    blocks = blocks.map((block) => ({
      ...block,
      collapsed: preSearchCollapseState.get(block.id) ?? block.collapsed,
    }));
    preSearchCollapseState = null;
  }

  searchQuery = normalized;
  renderRows();
}

function attachSearchControls(root) {
  const wrapper = root.querySelector(".search-box-wrapper");
  const input   = root.querySelector(".search-box-input");
  const clearBtn = root.querySelector(".search-box-clear");
  if (!input) return;

  input.addEventListener("input", () => {
    wrapper?.classList.toggle("has-value", !!input.value);
    applySearch(input.value);
  });

  clearBtn?.addEventListener("click", () => {
    input.value = "";
    wrapper?.classList.remove("has-value");
    applySearch("");
    input.focus();
  });
}

function isPlaceholderRow(row) {
  if (!row) {
    return false;
  }

  const hasDateRange = !!getRowRange(row);
  if (hasDateRange) {
    return false;
  }

  return !rowHasAnyListo(row)
    && !`${row.title ?? ""}`.trim()
    && !`${row.genre ?? ""}`.trim()
    && !`${row.id ?? ""}`.trim()
    && !`${row.startDateText ?? ""}`.trim()
    && !`${row.endDateText ?? ""}`.trim()
    && !`${row.startDateISO ?? ""}`.trim()
    && !`${row.endDateISO ?? ""}`.trim();
}

function getVisibleMonthRange(calendarContext) {
  return {
    startDate: new Date(calendarContext.year, calendarContext.month - 1, 1),
    endDate: new Date(calendarContext.year, calendarContext.month, 0),
  };
}

function intersectsVisibleMonth(rowRange, monthRange) {
  if (!rowRange || !monthRange) {
    return false;
  }

  return rowRange.startDate <= monthRange.endDate && rowRange.endDate >= monthRange.startDate;
}

function getSortValue(row, key) {
  if (key === "startDate") { return row.startDateISO || "\uffff"; }
  if (key === "endDate")   { return row.endDateISO   || "\uffff"; }
  if (key === "listo")     { return getRowListo(row) ? 0 : 1; }
  return `${row[key] ?? ""}`.toLocaleLowerCase("es-ES");
}

function getOrderedRowsForMonth(block, calendarContext) {
  const monthRange = getVisibleMonthRange(calendarContext);

  const indexedRows = block.rows.map((row, sourceIndex) => {
    if (!Number.isInteger(row.homeMonth) || !Number.isInteger(row.homeYear)) {
      row.homeMonth = calendarContext.month;
      row.homeYear = calendarContext.year;
    }

    const rowRange = getRowRange(row);
    const isVisibleInMonth = intersectsVisibleMonth(rowRange, monthRange);
    const isUnscheduledInMonth =
      !rowRange &&
      row.homeMonth === calendarContext.month &&
      row.homeYear === calendarContext.year;
    const isVisibleInCurrentMonth = isVisibleInMonth || isUnscheduledInMonth;
    const isInheritedInMonth =
      !!rowRange &&
      isVisibleInMonth &&
      (
        rowRange.startDate.getMonth() + 1 !== calendarContext.month ||
        rowRange.startDate.getFullYear() !== calendarContext.year
      );

    return {
      row,
      sourceIndex,
      rowRange,
      isVisibleInCurrentMonth,
      isInheritedInMonth,
    };
  });

const visibleRows = indexedRows.filter((item) => item.isVisibleInCurrentMonth);

const nonPlaceholders = visibleRows.filter((item) => !item.row._autoPlaceholder);
const placeholders = visibleRows.filter((item) => item.row._autoPlaceholder);

  if (sortState.key) {
    nonPlaceholders.sort((left, right) => {
      const a = getSortValue(left.row, sortState.key);
      const b = getSortValue(right.row, sortState.key);
      const cmp =
        typeof a === "number" && typeof b === "number"
          ? a - b
          : `${a}`.localeCompare(`${b}`, "es-ES", { numeric: true });
      if (cmp !== 0) {
        return sortState.dir === "asc" ? cmp : -cmp;
      }
      return left.sourceIndex - right.sourceIndex;
    });
  }

  const allowedPlaceholders = nonPlaceholders.length > 0
    ? placeholders.filter((item) => !isPlaceholderRow(item.row))
    : placeholders;
  const orderedVisibleRows = [...nonPlaceholders, ...allowedPlaceholders];

  if (orderedVisibleRows.length) {
    return orderedVisibleRows;
  }

  // Fallback: bloque vacío → crear fila placeholder
  const fallbackRow = newRowForBlock(block.blockType, calendarContext);
  fallbackRow._autoPlaceholder = true;
  block.rows.push(fallbackRow);

  return [{
    row: fallbackRow,
    sourceIndex: block.rows.length - 1,
    rowRange: null,
    isVisibleInCurrentMonth: true,
    isInheritedInMonth: false,
  }];
}

function getBlockDailyCounts(block, calendarContext) {
  const counts = new Array(32).fill(0);
  const orderedRows = getOrderedRowsForMonth(block, calendarContext);
  const monthStart = new Date(calendarContext.year, calendarContext.month - 1, 1);
  const monthEnd = new Date(calendarContext.year, calendarContext.month, 0);

  orderedRows.forEach(({ rowRange }) => {
    if (!rowRange) {
      return;
    }

    const visibleStart = rowRange.startDate < monthStart ? monthStart : rowRange.startDate;
    const visibleEnd = rowRange.endDate > monthEnd ? monthEnd : rowRange.endDate;
    if (visibleEnd < visibleStart) {
      return;
    }

    for (let day = visibleStart.getDate(); day <= visibleEnd.getDate(); day += 1) {
      counts[day] += 1;
    }
  });

  return counts;
}

function parseDateInput(
  text,
  defaultMonth = currentCalendarContext.month,
  defaultYear = currentCalendarContext.year,
) {
  const originalText = `${text ?? ""}`;
  const parser = window.PanelDateUtils?.parseDatePartsFromText;
  const parsed = typeof parser === "function"
    ? parser(originalText, { defaultMonth, defaultYear })
    : null;

  if (!parsed) {
    const trimmed = originalText.trim();
    if (!trimmed) {
      return { ok: true, display: "", iso: null, error: null };
    }

    return { ok: false, display: originalText, iso: null, error: "Fecha no válida (DD/MM/YY)" };
  }

  if (parsed.ok && parsed.empty) {
    return { ok: true, display: "", iso: null, error: null };
  }

  if (!parsed.ok) {
    return { ok: false, display: originalText, iso: null, error: parsed.error || "Fecha no válida (DD/MM/YY)" };
  }

  return {
    ok: true,
    display: formatDateDisplay(parsed.day, parsed.month, parsed.year),
    iso: formatDateISO(parsed.day, parsed.month, parsed.year),
    error: null,
  };
}

function getDateFieldNames(columnKey) {
  return columnKey === "startDate"
    ? { textField: "startDateText", isoField: "startDateISO" }
    : { textField: "endDateText", isoField: "endDateISO" };
}

function renderDateCell(cell, row, columnKey) {
  const { textField } = getDateFieldNames(columnKey);
  const displayValue = row[textField] || "";
  const errorMessage = row[`${columnKey}Error`] || row.dateRangeError || "";
  cell.textContent = displayValue;
  cell.title = errorMessage;
  cell.classList.toggle("has-error", !!errorMessage);
}

function validateRowDateRange(row, { notify = false } = {}) {
  if (!row) {
    return { ok: true, error: null };
  }

  const hasParseError = !!row.startDateError || !!row.endDateError;
  if (hasParseError) {
    row.dateRangeError = null;
    return { ok: false, error: null };
  }

  const startDate = parseISODateValue(row.startDateISO);
  const endDate = parseISODateValue(row.endDateISO);
  if (!startDate || !endDate) {
    row.dateRangeError = null;
    return { ok: true, error: null };
  }

  if (endDate < startDate) {
    const rangeError = "La fecha de fin no puede ser anterior a la fecha de inicio";
    const hadRangeError = row.dateRangeError === rangeError;
    row.dateRangeError = rangeError;
    if (notify && !hadRangeError) {
      showGridToast(rangeError);
    }
    return { ok: false, error: rangeError };
  }

  row.dateRangeError = null;
  return { ok: true, error: null };
}

function validateAllRowsDateRanges() {
  blocks.forEach((block) => {
    block.rows.forEach((row) => {
      validateRowDateRange(row);
    });
  });
}

function applyDateCellValue(row, columnKey, rawValue, { preserveRawOnInvalid = true } = {}) {
  const { textField, isoField } = getDateFieldNames(columnKey);
  const parsed = parseDateInput(rawValue);
  if (parsed.ok) {
    row[textField] = parsed.display;
    row[isoField] = parsed.iso;
    row[`${columnKey}Error`] = null;
    validateRowDateRange(row, { notify: true });
    return { ok: true, display: row[textField], iso: row[isoField] };
  }

  row[textField] = preserveRawOnInvalid ? `${rawValue ?? ""}` : "";
  row[isoField] = null;
  row[`${columnKey}Error`] = parsed.error;
  validateRowDateRange(row);
  return { ok: false, display: row[textField], iso: null, error: parsed.error };
}

function setCellValue(cell, rawValue, historyOptions = {}) {
  const rowData = getRowByCell(cell);
  if (!rowData) {
    return null;
  }

  const { row, meta } = rowData;
  const before = getCellRawValue(row, meta.columnKey);
  const parsedValue = parseCellValue(meta.columnKey, rawValue);

  if (meta.columnKey === "title") {
    row.title = parsedValue;
    const titleText = cell.querySelector(".title-cell__text");
    if (titleText) {
      titleText.textContent = row.title;
      titleText.title = row.title;
    }
  } else if (meta.columnKey === "listo") {
    setRowListo(row, parsedValue);
    const checkbox = cell.querySelector('input[type="checkbox"]');
    if (checkbox) {
      checkbox.checked = getRowListo(row);
    }
    // Sincronizar el checkbox general del bloque: marcado solo si TODAS las
    // filas visibles del mes están marcadas (al desmarcar una, se desactiva).
    const blockIndex = Number.parseInt(cell.dataset.blockIndex, 10);
    if (!Number.isNaN(blockIndex) && blocks[blockIndex]) {
      const block = blocks[blockIndex];
      const realRows = getOrderedRowsForMonth(block, currentCalendarContext)
        .filter((item) => item.isVisibleInCurrentMonth && !item.row._autoPlaceholder)
        .map((item) => item.row);
      const allListo = realRows.length > 0 && realRows.every((r) => getRowListo(r));
      const leftBody = document.getElementById("left-body");
      const blockCheckbox = leftBody?.querySelector(
        `.left-row.group[data-block-index="${blockIndex}"] .listo-checkbox`
      );
      if (blockCheckbox) {
        blockCheckbox.checked = allListo;
      }
    }
  } else if (DATE_COLUMNS.has(meta.columnKey)) {
    applyDateCellValue(row, meta.columnKey, parsedValue);
    renderDateCell(cell, row, meta.columnKey);
  } else if (meta.columnKey === "genre") {
    row.genre = parsedValue;
    cell.textContent = row.genre;
  } else if (meta.columnKey === "id") {
    row.id = parsedValue;
    cell.textContent = row.id;
  }

  const after = getCellRawValue(row, meta.columnKey);
  if (before !== after) {
    const patchMeta = { ...meta, rowKey: row.rowKey };
    if (meta.columnKey === "listo") {
      patchMeta.monthKey = monthKeyFor();
    }
    addPatchToCurrentAction(
      createSetCellPatch(patchMeta, before, after),
      {
        type: historyOptions.type || "edit",
        groupKey: historyOptions.groupKey || `${meta.blockIndex}:${meta.rowIndex}:${meta.columnKey}`,
      }
    );
  }

  return { row, meta };
}

function focusCellEditor(cell) {
  if (!cell) {
    return;
  }

  const columnKey = cell.dataset.columnKey;
  if (columnKey === "title" && typeof cell.openEditMode === "function") {
    cell.openEditMode({ keepContent: true });
    return;
  }

  if (columnKey === "listo") {
    const checkbox = cell.querySelector('input[type="checkbox"]');
    if (checkbox) {
      checkbox.focus();
    }
    return;
  }

  if (DATE_COLUMNS.has(columnKey) && typeof cell.openEditMode === "function") {
    cell.openEditMode({ keepContent: true });
    return;
  }

  if (columnKey === "id" && typeof cell.openEditMode === "function") {
    cell.openEditMode({ keepContent: true });
  }
}

function moveSelectionDownWithinBlock(cell) {
  const meta = getCellMeta(cell);
  if (!meta) {
    return { moved: false, cell };
  }

  const block = blocks[meta.blockIndex];
  const nextRowIndex = meta.rowIndex + 1;
  if (!block || nextRowIndex >= block.rows.length) {
    return { moved: false, cell };
  }

  const nextCell = document.querySelector(
    `[data-block-index="${meta.blockIndex}"][data-row-index="${nextRowIndex}"][data-column-key="${meta.columnKey}"]`
  );

  if (!nextCell) {
    return { moved: false, cell };
  }

  setSelectedCell(nextCell);
  return { moved: true, cell: nextCell };
}

function getAdjacentCellByArrow(cell, key) {
  const meta = getCellMeta(cell);
  if (!meta) {
    return null;
  }

  if (key === "ArrowLeft" || key === "ArrowRight") {
    const row = cell.parentElement;
    if (!row) {
      return null;
    }

    const rowCells = [...row.querySelectorAll("[data-column-key]")];
    const currentIndex = rowCells.indexOf(cell);
    if (currentIndex < 0) {
      return null;
    }

    const delta = key === "ArrowLeft" ? -1 : 1;
    return rowCells[currentIndex + delta] || null;
  }

  if (key === "ArrowUp" || key === "ArrowDown") {
    const block = blocks[meta.blockIndex];
    if (!block) {
      return null;
    }

    const delta = key === "ArrowUp" ? -1 : 1;
    const nextRowIndex = meta.rowIndex + delta;
    if (nextRowIndex < 0 || nextRowIndex >= block.rows.length) {
      return null;
    }

    return document.querySelector(
      `[data-block-index="${meta.blockIndex}"][data-row-index="${nextRowIndex}"][data-column-key="${meta.columnKey}"]`
    );
  }

  return null;
}

function isCellVisible(cell) {
  if (!cell) {
    return false;
  }

  const styles = window.getComputedStyle(cell);
  return styles.display !== "none" && styles.visibility !== "hidden";
}

function getRowEditableColumnKeys(rowElement) {
  if (!rowElement) {
    return [];
  }

  return columns
    .filter((column) => column.editable !== false && column.visible !== false)
    .map((column) => column.key)
    .filter((columnKey) => {
      const rowCell = rowElement.querySelector(`[data-column-key="${columnKey}"]`);
      return isCellVisible(rowCell);
    });
}

function getNextTabCell(cell, direction) {
  const meta = getCellMeta(cell);
  if (!meta) {
    return null;
  }

  const block = blocks[meta.blockIndex];
  if (!block) {
    return null;
  }

  const currentRow = cell.parentElement;
  const currentRowColumns = getRowEditableColumnKeys(currentRow);
  if (!currentRowColumns.length) {
    return null;
  }

  const currentColumnIndex = currentRowColumns.indexOf(meta.columnKey);
  if (currentColumnIndex < 0) {
    return null;
  }

  const nextColumnIndex = currentColumnIndex + direction;
  if (nextColumnIndex >= 0 && nextColumnIndex < currentRowColumns.length) {
    return currentRow.querySelector(`[data-column-key="${currentRowColumns[nextColumnIndex]}"]`);
  }

  let nextRowIndex = meta.rowIndex + direction;
  while (nextRowIndex >= 0 && nextRowIndex < block.rows.length) {
    const nextRow = document.querySelector(
      `[data-block-index="${meta.blockIndex}"][data-row-index="${nextRowIndex}"]`
    )?.parentElement;

    const nextRowColumns = getRowEditableColumnKeys(nextRow);
    if (nextRowColumns.length) {
      const targetColumnKey = direction > 0 ? nextRowColumns[0] : nextRowColumns[nextRowColumns.length - 1];
      const nextCell = nextRow?.querySelector(`[data-column-key="${targetColumnKey}"]`);
      if (nextCell) {
        return nextCell;
      }
    }

    nextRowIndex += direction;
  }

  return cell;
}

function focusCellWithoutEditing(cell) {
  if (!cell || editingCell) {
    return;
  }

  requestAnimationFrame(() => {
    if (!editingCell) {
      const gridRoot = document.querySelector(".month-block__body-grid");
      gridRoot?.focus({ preventScroll: true });
    }
  });
}

function isEditingElement(element) {
  if (!element) {
    return false;
  }

  const tagName = element.tagName?.toLowerCase();
  return tagName === "input" || tagName === "textarea" || tagName === "select" || element.isContentEditable;
}

function ensureGenreMenuElement() {
  if (genreMenuElement?.isConnected) {
    return genreMenuElement;
  }

  genreMenuElement = document.createElement("div");
  genreMenuElement.className = "genre-dropdown-menu";
  genreMenuElement.setAttribute("role", "listbox");
  document.body.appendChild(genreMenuElement);
  return genreMenuElement;
}

function getCellRawValue(row, columnKey) {
  if (!row) {
    return "";
  }

  if (columnKey === "listo") {
    return getRowListo(row) ? "true" : "";
  }

  if (columnKey === "title") {
    return row.title || "";
  }

  if (columnKey === "genre") {
    return row.genre || "";
  }

  if (columnKey === "id") {
    return row.id || "";
  }

  if (DATE_COLUMNS.has(columnKey)) {
    const { textField } = getDateFieldNames(columnKey);
    return row[textField] || "";
  }

  return "";
}

function computeFillValue(masterValue, targetOffset, columnKey) {
  const normalizedValue = `${masterValue ?? ""}`;
  if (!normalizedValue) {
    return normalizedValue;
  }

  if (DATE_COLUMNS.has(columnKey) || columnKey === "genre") {
    return normalizedValue;
  }

  const seriesMatch = normalizedValue.match(/^(.*?)(\d+)$/);
  if (!seriesMatch) {
    return normalizedValue;
  }

  const prefix = seriesMatch[1];
  const numberText = seriesMatch[2];
  const nextNumber = Number.parseInt(numberText, 10) + targetOffset;
  const paddedNumber = String(nextNumber).padStart(numberText.length, "0");
  return `${prefix}${paddedNumber}`;
}

function ensureFillHandleElement() {
  if (fillHandleElement?.isConnected) {
    return fillHandleElement;
  }

  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) {
    return null;
  }

  fillHandleElement = document.createElement("button");
  fillHandleElement.type = "button";
  fillHandleElement.className = "fill-handle";
  fillHandleElement.setAttribute("aria-label", "Autorrelleno hacia abajo");
  fillHandleElement.setAttribute("tabindex", "-1");
  fillHandleElement.addEventListener("pointerdown", startFillDrag);
  gridRoot.appendChild(fillHandleElement);
  return fillHandleElement;
}

function ensureCopyAntsElement() {
  if (copyAntsElement?.isConnected) {
    return copyAntsElement;
  }

  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) {
    return null;
  }

  copyAntsElement = document.createElement("div");
  copyAntsElement.className = "copy-ants";
  copyAntsElement.setAttribute("aria-hidden", "true");
  gridRoot.appendChild(copyAntsElement);
  return copyAntsElement;
}

function syncCopyAntsPosition() {
  const ants = ensureCopyAntsElement();
  if (!ants) {
    return;
  }

  if (!copyRange || copyRangeBlockIndex === null) {
    ants.classList.remove("is-visible");
    return;
  }

  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) {
    ants.classList.remove("is-visible");
    return;
  }

  const topCell = document.querySelector(
    `[data-block-index="${copyRangeBlockIndex}"][data-row-index="${copyRange.r1}"][data-column-key="${copyRange.col}"]`
  );
  const bottomCell = document.querySelector(
    `[data-block-index="${copyRangeBlockIndex}"][data-row-index="${copyRange.r2}"][data-column-key="${copyRange.col}"]`
  );

  if (!topCell || !bottomCell) {
    ants.classList.remove("is-visible");
    return;
  }

  const rootRect = gridRoot.getBoundingClientRect();
  const topRect = topCell.getBoundingClientRect();
  const bottomRect = bottomCell.getBoundingClientRect();

  ants.style.left = `${topRect.left - rootRect.left}px`;
  ants.style.top = `${topRect.top - rootRect.top}px`;
  ants.style.width = `${topRect.width}px`;
  ants.style.height = `${bottomRect.bottom - topRect.top}px`;
  ants.classList.add("is-visible");
}

function setCopyRange(nextRange, blockIndex = null) {
  if (!nextRange) {
    copyRange = null;
    copyRangeBlockIndex = null;
    syncCopyAntsPosition();
    return;
  }

  copyRange = {
    col: nextRange.col,
    r1: nextRange.r1,
    r2: nextRange.r2,
  };
  copyRangeBlockIndex = blockIndex;
  syncCopyAntsPosition();
}

function getCopySelection() {
  if (dragSelection) {
    return {
      blockIndex: dragSelection.blockIndex,
      col: dragSelection.col,
      r1: dragSelection.r1,
      r2: dragSelection.r2,
    };
  }

  const activeMeta = getCellMeta(selectedCell);
  if (!activeMeta) {
    return null;
  }

  return {
    blockIndex: activeMeta.blockIndex,
    col: activeMeta.columnKey,
    r1: activeMeta.rowIndex,
    r2: activeMeta.rowIndex,
  };
}

function buildCopyTextFromSelection(selection) {
  const block = blocks[selection.blockIndex];
  if (!block) {
    return "";
  }

  const orderedRows = getOrderedRowsForMonth(block, currentCalendarContext);
  const values = orderedRows
    .filter((item) => item.sourceIndex >= selection.r1 && item.sourceIndex <= selection.r2)
    .map((item) => getCellRawValue(item.row, selection.col));

  return values.join("\n");
}

function copyTextToClipboard(text) {
  const fallbackCopy = () => {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    textarea.style.pointerEvents = "none";
    textarea.style.left = "-9999px";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
  };

  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(text).catch(() => {
      fallbackCopy();
    });
    return;
  }

  fallbackCopy();
}

function clearFillPreview() {
  document.querySelectorAll(".left-row > div[data-column-key].is-fill-preview").forEach((cell) => {
    cell.classList.remove("is-fill-preview");
  });
}

function clearDragSelectionPreview() {
  document.querySelectorAll(".left-row > div[data-column-key].is-drag-selected").forEach((cell) => {
    cell.classList.remove("is-drag-selected");
  });
}

function renderDragSelectionPreview(selection) {
  clearDragSelectionPreview();

  if (!selection) {
    refreshDeleteControls();
    return;
  }

  for (let rowIndex = selection.r1; rowIndex <= selection.r2; rowIndex += 1) {
    const cell = document.querySelector(
      `[data-block-index="${selection.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${selection.col}"]`
    );
    if (cell) {
      cell.classList.add("is-drag-selected");
    }
  }

  refreshDeleteControls();
}

function getCellFromPointer(event) {
  const directCell = event.target?.closest?.("[data-column-key]");
  if (directCell) {
    return directCell;
  }

  const hoveredCells = document.elementsFromPoint(event.clientX, event.clientY)
    .map((element) => element.closest?.("[data-column-key]"))
    .filter(Boolean);

  return hoveredCells[0] || null;
}

function updateDragSelectionFromPointer(event) {
  if (!dragSelectState.pointerDown || !dragSelectState.isDragSelect) {
    return;
  }

  const hoverCell = getCellFromPointer(event);
  const hoverMeta = getCellMeta(hoverCell);
  if (!hoverMeta || hoverMeta.blockIndex !== dragSelectState.anchorBlockIndex) {
    return;
  }

  const r1 = Math.min(dragSelectState.anchorRow, hoverMeta.rowIndex);
  const r2 = Math.max(dragSelectState.anchorRow, hoverMeta.rowIndex);
  dragSelection = {
    blockIndex: dragSelectState.anchorBlockIndex,
    col: dragSelectState.anchorCol,
    r1,
    r2,
  };

  renderDragSelectionPreview(dragSelection);
}

function resetDragSelectState() {
  dragSelectState = {
    pointerDown: false,
    isDragSelect: false,
    anchorCell: null,
    anchorCol: null,
    anchorBlockIndex: null,
    anchorRow: null,
    downX: 0,
    downY: 0,
  };
}

function handleGridPointerDown(event) {
  if (IS_VIEWER_MODE) { return; }
  if (event.button !== 0 || fillDragState || editingCell) {
    return;
  }

  if (event.target.closest(".fill-handle")) {
    return;
  }

  const cell = event.target.closest(".left-row > div[data-column-key]");
  if (!cell) {
    return;
  }

  const meta = getCellMeta(cell);
  if (!meta) {
    return;
  }
// Shift+Click → selección de rango
  if (event.shiftKey && !event.ctrlKey && !event.metaKey) {
    event.preventDefault();

    if (
      shiftSelectAnchor &&
      shiftSelectAnchor.blockIndex === meta.blockIndex &&
      shiftSelectAnchor.columnKey === meta.columnKey
    ) {
      // Hay ancla en el mismo bloque y columna → extender rango
      const block = blocks[meta.blockIndex];
      const orderedRows = getOrderedRowsForMonth(block, currentCalendarContext);

      const anchorVisIdx = orderedRows.findIndex(
        (item) => item.sourceIndex === shiftSelectAnchor.anchorSourceIndex
      );
      const targetVisIdx = orderedRows.findIndex(
        (item) => item.sourceIndex === meta.rowIndex
      );

      if (anchorVisIdx >= 0 && targetVisIdx >= 0) {
        const minVis = Math.min(anchorVisIdx, targetVisIdx);
        const maxVis = Math.max(anchorVisIdx, targetVisIdx);
        const selectedSourceIndices = orderedRows
          .slice(minVis, maxVis + 1)
          .map((item) => item.sourceIndex);

        dragSelection = {
          blockIndex: meta.blockIndex,
          col: meta.columnKey,
          r1: Math.min(...selectedSourceIndices),
          r2: Math.max(...selectedSourceIndices),
        };

        setSelectedCell(cell);
        renderDragSelectionPreview(dragSelection);
      }
   } else {
      // No hay ancla compatible → establecer ancla en la celda actual
      const elseBlock = blocks[meta.blockIndex];
      const elseOrderedRows = getOrderedRowsForMonth(elseBlock, currentCalendarContext);
      const elseVisIdx = Math.max(0, elseOrderedRows.findIndex(
        (item) => item.sourceIndex === meta.rowIndex
      ));
      shiftSelectAnchor = {
        blockIndex: meta.blockIndex,
        columnKey: meta.columnKey,
        anchorSourceIndex: meta.rowIndex,
        anchorVisibleIndex: elseVisIdx,
        activeVisibleIndex: elseVisIdx,
      };
      
      dragSelection = null;
      clearDragSelectionPreview();
      setSelectedCell(cell);
    }
    return;
  }
 const anchorBlock = blocks[meta.blockIndex];
  const anchorOrderedRows = getOrderedRowsForMonth(anchorBlock, currentCalendarContext);
  const anchorVisIdx = Math.max(0, anchorOrderedRows.findIndex(
    (item) => item.sourceIndex === meta.rowIndex
  ));

  shiftSelectAnchor = {
    blockIndex: meta.blockIndex,
    columnKey: meta.columnKey,
    anchorSourceIndex: meta.rowIndex,
    anchorVisibleIndex: anchorVisIdx,
    activeVisibleIndex: anchorVisIdx,
  };
  dragSelection = null;
  clearDragSelectionPreview();

  dragSelectState = {
    pointerDown: true,
    isDragSelect: false,
    anchorCell: cell,
    anchorCol: meta.columnKey,
    anchorBlockIndex: meta.blockIndex,
    anchorRow: meta.rowIndex,
    downX: event.clientX,
    downY: event.clientY,
  };
}

function handleGridPointerMove(event) {
  if (!dragSelectState.pointerDown || fillDragState) {
    return;
  }

  if (!dragSelectState.isDragSelect) {
    const dx = event.clientX - dragSelectState.downX;
    const dy = event.clientY - dragSelectState.downY;
    const distance = Math.hypot(dx, dy);
    if (distance <= DRAG_THRESHOLD_PX) {
      return;
    }

    dragSelectState.isDragSelect = true;
    dragSelection = {
      blockIndex: dragSelectState.anchorBlockIndex,
      col: dragSelectState.anchorCol,
      r1: dragSelectState.anchorRow,
      r2: dragSelectState.anchorRow,
    };
    renderDragSelectionPreview(dragSelection);
  }

  updateDragSelectionFromPointer(event);
}

function handleGridPointerUp() {
  if (!dragSelectState.pointerDown) {
    return;
  }

  if (dragSelectState.isDragSelect) {
    suppressNextGridClick = true;
    setSelectedCell(dragSelectState.anchorCell);
    setTimeout(() => {
      suppressNextGridClick = false;
    }, 0);
  }

  resetDragSelectState();
}

function handleGridPointerCancel() {
  if (!dragSelectState.pointerDown) {
    return;
  }
  resetDragSelectState();
}

function handleGridClickCapture(event) {
  if (!suppressNextGridClick) {
    return;
  }

  if (event.target.closest(".left-row > div[data-column-key]")) {
    event.preventDefault();
    event.stopPropagation();
    suppressNextGridClick = false;
  }
}

// =============================================================================
// Delegated cell event handlers
//
// Per-cell click / dblclick / focus listeners used to be attached inside each
// attach*Cell function. With ~50 rows × 5 column types × 3 listeners that was
// 750+ registrations on every renderRows(). These three functions live at the
// grid root and dispatch based on dataset.columnKey, eliminating that bloat.
// =============================================================================
function findCellFromEvent(event) {
  const target = event.target;
  if (!(target instanceof Element)) return null;
  // Only react to direct children of .left-row (the actual cells); the .day-row
  // cells live in the gantt, which has its own handlers.
  const cell = target.closest(".left-row > div[data-column-key]");
  return cell || null;
}

function handleGridDelegatedClick(event) {
  const cell = findCellFromEvent(event);
  if (!cell) return;

  const wasSelected = selectedCell === cell;
  const columnKey = cell.dataset.columnKey;

  setSelectedCell(cell);

  if (columnKey === "listo") {
    // Clicks on the checkbox input itself are handled by its own change event
    // — don't double-toggle when the user clicked the input.
    if (event.target.tagName === "INPUT") return;
    if (typeof cell.toggleListo === "function") {
      cell.toggleListo();
    }
    return;
  }

  // Genre opens edit mode on a click against an already-selected cell. Mirrors
  // the previous per-cell behaviour where the second click on the dropdown
  // expands the menu.
  if (columnKey === "genre" && wasSelected && typeof cell.openEditMode === "function") {
    cell.openEditMode({ keepContent: true });
  }
}

function handleGridDelegatedDblClick(event) {
  const cell = findCellFromEvent(event);
  if (!cell) return;
  if (typeof cell.openEditMode === "function") {
    setSelectedCell(cell);
    cell.openEditMode({ keepContent: true });
  }
}

function handleGridDelegatedFocusIn(event) {
  const cell = findCellFromEvent(event);
  if (!cell) return;
  setSelectedCell(cell);
}

function updateFillPreview(masterMeta, targetRowIndex) {
  clearFillPreview();
  if (targetRowIndex <= masterMeta.rowIndex) {
    return;
  }

  for (let rowIndex = masterMeta.rowIndex + 1; rowIndex <= targetRowIndex; rowIndex += 1) {
    const cell = document.querySelector(
      `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${masterMeta.columnKey}"]`
    );
    if (cell) {
      cell.classList.add("is-fill-preview");
    }
  }
}

function getFillTargetRowIndexFromPointer(event, masterMeta) {
  const cells = document.elementsFromPoint(event.clientX, event.clientY)
    .map((element) => element.closest?.("[data-column-key]"))
    .filter(Boolean);

  const matchedCell = cells.find(
    (cell) => cell.dataset.blockIndex === String(masterMeta.blockIndex) && cell.dataset.columnKey === masterMeta.columnKey
  );

  if (matchedCell) {
    const nextIndex = Number.parseInt(matchedCell.dataset.rowIndex, 10);
    return Number.isNaN(nextIndex) ? masterMeta.rowIndex : nextIndex;
  }

  const block = blocks[masterMeta.blockIndex];
  const lastRowIndex = Math.max(0, (block?.rows?.length || 1) - 1);
  const lastCell = document.querySelector(
    `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${lastRowIndex}"][data-column-key="${masterMeta.columnKey}"]`
  );

  if (lastCell && event.clientY > lastCell.getBoundingClientRect().bottom) {
    return lastRowIndex;
  }

  return masterMeta.rowIndex;
}

function applyFillDown(masterMeta, targetRowIndex) {
  if (targetRowIndex <= masterMeta.rowIndex) {
    return;
  }

  const masterCell = document.querySelector(
    `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${masterMeta.rowIndex}"][data-column-key="${masterMeta.columnKey}"]`
  );
  const masterData = masterCell ? getRowByCell(masterCell) : null;
  if (!masterData) {
    return;
  }

  const masterValue = getCellRawValue(masterData.row, masterMeta.columnKey);
  withHistoryAction("fill", { groupKey: `fill:${masterMeta.blockIndex}:${masterMeta.columnKey}:${masterMeta.rowIndex}` }, () => {
    for (let rowIndex = masterMeta.rowIndex + 1; rowIndex <= targetRowIndex; rowIndex += 1) {
      const targetCell = document.querySelector(
        `[data-block-index="${masterMeta.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${masterMeta.columnKey}"]`
      );
      if (!targetCell) {
        continue;
      }

      const offset = rowIndex - masterMeta.rowIndex;
      setCellValue(targetCell, computeFillValue(masterValue, offset, masterMeta.columnKey), { type: "fill", groupKey: `fill:${masterMeta.blockIndex}:${masterMeta.columnKey}:${masterMeta.rowIndex}` });
    }
  });
}

function stopFillDrag(applyChanges) {
  if (!fillDragState) {
    return;
  }

  const { pointerId, masterMeta, previewRowIndex } = fillDragState;
  const handle = ensureFillHandleElement();
  if (handle && pointerId !== null && pointerId !== undefined) {
    handle.releasePointerCapture?.(pointerId);
  }

  document.removeEventListener("pointermove", handleFillDragMove);
  document.removeEventListener("pointerup", handleFillDragEnd);
  document.removeEventListener("pointercancel", handleFillDragCancel);

  clearFillPreview();
  fillDragState = null;

  if (applyChanges) {
    applyFillDown(masterMeta, previewRowIndex);
  }

  syncFillHandlePosition();
}

function handleFillDragMove(event) {
  if (!fillDragState) {
    return;
  }

  const nextTarget = getFillTargetRowIndexFromPointer(event, fillDragState.masterMeta);
  const clampedTarget = Math.max(fillDragState.masterMeta.rowIndex, nextTarget);
  fillDragState.previewRowIndex = clampedTarget;
  updateFillPreview(fillDragState.masterMeta, clampedTarget);
}

function handleFillDragEnd(event) {
  event.preventDefault();
  stopFillDrag(true);
}

function handleFillDragCancel() {
  stopFillDrag(false);
}

function startFillDrag(event) {
  if (IS_VIEWER_MODE) { return; }
  if (event.button !== 0 || !selectedCell || editingCell) {
    return;
  }

  const masterMeta = getCellMeta(selectedCell);
  if (!masterMeta) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  const handle = ensureFillHandleElement();
  handle?.setPointerCapture?.(event.pointerId);

  fillDragState = {
    pointerId: event.pointerId,
    masterMeta,
    previewRowIndex: masterMeta.rowIndex,
  };

  document.addEventListener("pointermove", handleFillDragMove);
  document.addEventListener("pointerup", handleFillDragEnd);
  document.addEventListener("pointercancel", handleFillDragCancel);
}

function syncFillHandlePosition() {
  const handle = ensureFillHandleElement();
  if (!handle) {
    return;
  }

  if (!selectedCell || editingCell || fillDragState || dragSelectState.isDragSelect || !selectedCell.isConnected) {
    handle.classList.remove("is-visible");
    return;
  }

  const gridRoot = document.querySelector(".month-block__body-grid");
  if (!gridRoot) {
    handle.classList.remove("is-visible");
    return;
  }

  const cellRect = selectedCell.getBoundingClientRect();
  const rootRect = gridRoot.getBoundingClientRect();
  handle.style.left = `${cellRect.right - rootRect.left - 5}px`;
  handle.style.top = `${cellRect.bottom - rootRect.top - 5}px`;
  handle.classList.add("is-visible");
}

function handleGridEnterKey(event) {
  const isArrowNavigationKey = ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key);
  const isPrintableKey = event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey;
  const hasSelectedCell = !!selectedCell && !!getCellMeta(selectedCell);
  const keyLower = event.key.toLowerCase();
  const isUndoShortcut = (event.ctrlKey || event.metaKey) && keyLower === "z" && !event.shiftKey;
  const isRedoShortcut = ((event.ctrlKey || event.metaKey) && keyLower === "z" && event.shiftKey)
    || (event.ctrlKey && !event.metaKey && keyLower === "y");

  if (isUndoShortcut || isRedoShortcut) {
    if (IS_VIEWER_MODE) { return; }
    const activeElement = document.activeElement;
    const editingNative = isEditingElement(activeElement) && !(activeElement?.classList?.contains("editor-overlay"));
    if (editingNative) {
      return;
    }

    event.preventDefault();
    if (editingCell) {
      editingCell.cancel?.();
    }
    if (isUndoShortcut) {
      undoLastAction();
    } else {
      redoLastAction();
    }
    return;
  }
  
  if (editingCell) {
    if (editingCell.type === "select" && typeof editingCell.handleKeyDown === "function") {
      const handled = editingCell.handleKeyDown(event);
      if (handled) {
        return;
      }
    }

    if (event.key === "Tab") {
      event.preventDefault();
      const currentCell = editingCell.cell;
      editingCell.commit();
      const nextCell = getNextTabCell(currentCell, event.shiftKey ? -1 : 1);
      if (nextCell) {
        setSelectedCell(nextCell);
        focusCellWithoutEditing(nextCell);
      }
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      const currentCell = editingCell.cell;
      editingCell.commit();
      const nextSelection = moveSelectionDownWithinBlock(currentCell);
      focusCellWithoutEditing(nextSelection.cell);
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      editingCell.cancel();
      focusCellWithoutEditing(selectedCell);
      return;
    }

    if (isArrowNavigationKey && editingCell.type !== "select") {
      event.preventDefault();
      const currentCell = editingCell.cell;
      editingCell.commit();
      const nextCell = getAdjacentCellByArrow(currentCell, event.key);
      if (nextCell) {
        setSelectedCell(nextCell);
        focusCellWithoutEditing(nextCell);
      }
      return;
    }
    
    return;
  }

  if (!hasSelectedCell) {
    if (event.key === "Escape" && copyRange) {
      setCopyRange(null);
      event.preventDefault();
    }
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
    if (isEditingElement(document.activeElement)) {
      return;
    }

    const nextCopySelection = getCopySelection();
    if (!nextCopySelection) {
      return;
    }

    setCopyRange(
      {
        col: nextCopySelection.col,
        r1: nextCopySelection.r1,
        r2: nextCopySelection.r2,
      },
      nextCopySelection.blockIndex
    );
    const clipboardText = buildCopyTextFromSelection(nextCopySelection);
    copyTextToClipboard(clipboardText);
    event.preventDefault();
    return;
  }

  if (event.key === "Escape") {
    let handledEscape = false;

    if (copyRange) {
      setCopyRange(null);
      handledEscape = true;
    }

    if (selectedCell) {
      shiftSelectAnchor = null;
      setSelectedCell(null);
      dragSelection = null;
      clearDragSelectionPreview();
      handledEscape = true;
    }

    if (handledEscape) {
      event.preventDefault();
      return;
    }
  }

  if (event.key === "Tab") {
    event.preventDefault();
    const nextCell = getNextTabCell(selectedCell, event.shiftKey ? -1 : 1);
    if (!nextCell) {
      return;
    }

    setSelectedCell(nextCell);
    focusCellWithoutEditing(nextCell);
    return;
  }

if (isArrowNavigationKey) {
    if (isEditingElement(document.activeElement)) {
      return;
    }

if (event.shiftKey && (event.key === "ArrowDown" || event.key === "ArrowUp")) {
      const meta = getCellMeta(selectedCell);
      if (!meta) return;
      event.preventDefault();

      const block = blocks[meta.blockIndex];
      if (!block) return;

      const orderedRows = getOrderedRowsForMonth(block, currentCalendarContext);
      if (!orderedRows.length) return;

      if (
        !shiftSelectAnchor
        || shiftSelectAnchor.blockIndex !== meta.blockIndex
        || shiftSelectAnchor.columnKey !== meta.columnKey
        || shiftSelectAnchor.anchorSourceIndex !== meta.rowIndex
      ) {
        const anchorVisIdx = orderedRows.findIndex((item) => item.sourceIndex === meta.rowIndex);
        if (anchorVisIdx < 0) return;
        shiftSelectAnchor = {
          blockIndex: meta.blockIndex,
          columnKey: meta.columnKey,
          anchorSourceIndex: meta.rowIndex,
          anchorVisibleIndex: anchorVisIdx,
          activeVisibleIndex: anchorVisIdx,
        };
      }

      const delta = event.key === "ArrowDown" ? 1 : -1;
      shiftSelectAnchor.activeVisibleIndex = Math.max(
        0,
        Math.min(orderedRows.length - 1, shiftSelectAnchor.activeVisibleIndex + delta)
      );

      const minVis = Math.min(shiftSelectAnchor.anchorVisibleIndex, shiftSelectAnchor.activeVisibleIndex);
      const maxVis = Math.max(shiftSelectAnchor.anchorVisibleIndex, shiftSelectAnchor.activeVisibleIndex);
      const selectedSourceIndices = orderedRows.slice(minVis, maxVis + 1).map((item) => item.sourceIndex);

      dragSelection = {
        blockIndex: meta.blockIndex,
        col: meta.columnKey,
        r1: Math.min(...selectedSourceIndices),
        r2: Math.max(...selectedSourceIndices),
      };

      renderDragSelectionPreview(dragSelection);
      return;
    }

    shiftSelectAnchor = null;

    const nextCell = getAdjacentCellByArrow(selectedCell, event.key);
    if (!nextCell) {
      return;
    }

    event.preventDefault();
    setSelectedCell(nextCell);
    focusCellWithoutEditing(nextCell);
    return;
  }

  if (event.key === "Enter") {
    event.preventDefault();
    if (selectedCell.dataset.columnKey === "genre" && typeof selectedCell.openEditMode === "function") {
      selectedCell.openEditMode({ keepContent: true });
      return;
    }

    const nextSelection = moveSelectionDownWithinBlock(selectedCell);
    focusCellWithoutEditing(nextSelection.cell);
    return;
  }

  if (event.key === "F2" && typeof selectedCell.openEditMode === "function") {
    event.preventDefault();
    selectedCell.openEditMode({ keepContent: true });
    return;
  }

  // Space on a listo checkbox cell toggles it. Used to be a per-cell keydown
  // listener; lives here now because per-cell listeners were a significant
  // chunk of the listener count we trimmed.
  if ((event.key === " " || event.key === "Spacebar")
      && selectedCell.dataset.columnKey === "listo"
      && typeof selectedCell.toggleListo === "function") {
    event.preventDefault();
    selectedCell.toggleListo();
    return;
  }

    if ((event.ctrlKey || event.metaKey) && (event.key === "Delete" || event.key === "Backspace")) {
    if (IS_VIEWER_MODE) { return; }
    if (isEditingElement(document.activeElement)) {
      return;
    }

    const target = getDeleteTarget();
    if (!target) {
      return;
    }

    event.preventDefault();
    openDeleteConfirmModal(target, selectedCell);
    return;
  }

  if (isPrintableKey && typeof selectedCell.openEditMode === "function") {
    if (IS_VIEWER_MODE) { return; }
    const column = getColumnByKey(selectedCell.dataset.columnKey);
    if (column?.cellType === "select") {
      event.preventDefault();
      const now = Date.now();
      const normalizedKey = event.key.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase();
      genreTypeBuffer = now - genreTypeBufferTimestamp <= GENRE_TYPE_BUFFER_TIMEOUT_MS
        ? `${genreTypeBuffer}${normalizedKey}`
        : normalizedKey;
      genreTypeBufferTimestamp = now;

      const matchedOption = column.options?.find((option) => {
        const normalizedOption = option.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase();
        return normalizedOption.startsWith(genreTypeBuffer);
      });

      if (matchedOption) {
        setCellValue(selectedCell, matchedOption, { type: "edit", groupKey: `${selectedCell.dataset.blockIndex}:${selectedCell.dataset.rowIndex}:${selectedCell.dataset.columnKey}` });
      }
      return;
    }

    event.preventDefault();
    selectedCell.openEditMode({ replaceWith: event.key });
    return;
  }

if ((event.key === "Delete" || event.key === "Backspace") && selectedCell) {
    if (IS_VIEWER_MODE) { return; }
    const selectedRowIndex = Number.parseInt(selectedCell.dataset.rowIndex, 10);
    const hasVerticalRangeSelection =
      !!dragSelection
      && dragSelection.blockIndex === Number.parseInt(selectedCell.dataset.blockIndex, 10)
      && dragSelection.col === selectedCell.dataset.columnKey
      && dragSelection.r2 > dragSelection.r1
      && selectedRowIndex >= dragSelection.r1
      && selectedRowIndex <= dragSelection.r2;

    if (hasVerticalRangeSelection && !editingCell && !isEditingElement(document.activeElement)) {
      withHistoryAction("clear", { groupKey: `clear:${dragSelection.blockIndex}:${dragSelection.col}` }, () => {
        for (let rowIndex = dragSelection.r1; rowIndex <= dragSelection.r2; rowIndex += 1) {
          const targetCell = document.querySelector(
            `[data-block-index="${dragSelection.blockIndex}"][data-row-index="${rowIndex}"][data-column-key="${dragSelection.col}"]`
          );
          if (targetCell) {
            setCellValue(targetCell, "", { type: "clear", groupKey: `clear:${dragSelection.blockIndex}:${dragSelection.col}` });
          }
        }
      });

      event.preventDefault();
      return;
    }

    const rowData = getRowByCell(selectedCell);
    if (!rowData) {
      return;
    }

    event.preventDefault();
    withHistoryAction("clear", { groupKey: `clear:${selectedCell.dataset.blockIndex}:${selectedCell.dataset.rowIndex}:${selectedCell.dataset.columnKey}` }, () => {
      setCellValue(selectedCell, "", { type: "clear", groupKey: `clear:${selectedCell.dataset.blockIndex}:${selectedCell.dataset.rowIndex}:${selectedCell.dataset.columnKey}` });
    });
    
    focusCellEditor(selectedCell);
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
    return;
  }
}

function handleGridPaste(event) {
  if (IS_VIEWER_MODE) { return; }
  if (!selectedCell || editingCell) return;

  const pastedText = event.clipboardData?.getData("text/plain") || "";
  const clipboardLines = pastedText
    .split(/\r?\n/)
    .filter((line, index, all) => line !== "" || index < all.length - 1);
  if (!clipboardLines.length) return;

  event.preventDefault();

  const selectedMeta = getCellMeta(selectedCell);
  if (!selectedMeta) return;

  const block = blocks[selectedMeta.blockIndex];
  if (!block?.rows?.length) return;

  // Fuente de verdad: solo filas visibles en el mes actual
  const orderedRows = getOrderedRowsForMonth(block, currentCalendarContext);

  // ¿El selectedCell está DENTRO del rango de dragSelection activo?
  const selectedSourceIndex = selectedMeta.rowIndex;
  const isInsideDragSelection =
    !!dragSelection
    && dragSelection.blockIndex === selectedMeta.blockIndex
    && dragSelection.col === selectedMeta.columnKey
    && dragSelection.r2 > dragSelection.r1
    && selectedSourceIndex >= dragSelection.r1
    && selectedSourceIndex <= dragSelection.r2;

  withHistoryAction("paste", { groupKey: "paste" }, () => {

    if (isInsideDragSelection && !isEditingElement(document.activeElement)) {
      // ── PATH A: Pegar sobre un rango seleccionado ──────────────────────
      // Filas visibles dentro del rango (por sourceIndex, no aritmética ciega)
      const visibleInRange = orderedRows.filter(
        (item) => item.sourceIndex >= dragSelection.r1 && item.sourceIndex <= dragSelection.r2
      );
      const rangeSize = visibleInRange.length;

      const pasteValues = resolveVerticalPasteValues({ rangeSize, clipboardText: pastedText });
      if (!pasteValues.length) return;

      let targetRowKeys = visibleInRange.map((item) => item.row?.rowKey).filter(Boolean);

      // Insertar filas extra si el portapapeles tiene más líneas que el rango
      const rowsToInsert = Math.max(0, pasteValues.length - rangeSize);
      if (rowsToInsert > 0) {
        const lastSourceIndex = Math.max(...visibleInRange.map((i) => i.sourceIndex));
        insertRows(dragSelection.blockIndex, lastSourceIndex + 1, rowsToInsert, { historyType: "paste" });
        const updatedBlock = blocks[dragSelection.blockIndex];
        const newRows = updatedBlock?.rows?.slice(lastSourceIndex + 1, lastSourceIndex + 1 + rowsToInsert) || [];
        targetRowKeys = targetRowKeys.concat(newRows.map((r) => r?.rowKey).filter(Boolean));
      }

      const maxPaste = Math.min(pasteValues.length, targetRowKeys.length);
      for (let i = 0; i < maxPaste; i++) {
        const meta = getCellMetaFromRowKey(targetRowKeys[i], dragSelection.col);
        const cell = meta ? getCellByMeta(meta) : null;
        if (cell) setCellValue(cell, pasteValues[i], { type: "paste", groupKey: "paste" });
      }

      renderRows();
      return;
    }

    // ── PATH B: Pegar desde celda ancla hacia abajo ────────────────────
    // Localizar la celda ancla en las filas visibles por rowKey (no sourceIndex)
    const anchorRowId = selectedCell?.dataset?.rowId || null;
    let anchorVisIdx = orderedRows.findIndex((item) => item.row?.rowKey === anchorRowId);
    if (anchorVisIdx < 0) anchorVisIdx = 0;

const pasteCount = clipboardLines.length;

    const anchorIsPlaceholder = !!orderedRows[anchorVisIdx]?.row?._autoPlaceholder;

    let targetRowKeys = anchorIsPlaceholder
      ? []
      : orderedRows
          .slice(anchorVisIdx, anchorVisIdx + pasteCount)
          .filter((item) => !item.row._autoPlaceholder)
          .map((item) => item.row?.rowKey)
          .filter(Boolean);

    // Insertar filas nuevas si faltan destinos
    if (pasteCount > targetRowKeys.length) {
      const missing = pasteCount - targetRowKeys.length;
      // Siempre al final de las filas visibles del mes, no en medio
      const lastVisSourceIndex = orderedRows.length > 0
        ? Math.max(...orderedRows.map((i) => i.sourceIndex))
        : block.rows.length - 1;
      insertRows(selectedMeta.blockIndex, lastVisSourceIndex + 1, missing, { historyType: "paste" });
      const updatedBlock = blocks[selectedMeta.blockIndex];
      const newRows = updatedBlock?.rows?.slice(lastVisSourceIndex + 1, lastVisSourceIndex + 1 + missing) || [];
      targetRowKeys = targetRowKeys.concat(newRows.map((r) => r?.rowKey).filter(Boolean));
    }

    const maxPaste = clipboardLines.length > 1
      ? Math.min(clipboardLines.length, targetRowKeys.length)
      : 1;

    for (let i = 0; i < maxPaste; i++) {
      const meta = getCellMetaFromRowKey(targetRowKeys[i], selectedMeta.columnKey);
      const cell = meta ? getCellByMeta(meta) : null;
      if (cell) setCellValue(cell, clipboardLines[i], { type: "paste", groupKey: "paste" });
    }

    renderRows();

    // Mover foco a la última celda pegada en columnas de fecha
    if (DATE_COLUMNS.has(selectedMeta.columnKey) && maxPaste > 0) {
      const lastMeta = getCellMetaFromRowKey(targetRowKeys[maxPaste - 1], selectedMeta.columnKey);
      const lastCell = lastMeta ? getCellByMeta(lastMeta) : null;
      if (lastCell) {
        setSelectedCell(lastCell);
        focusCellWithoutEditing(lastCell);
      }
    }
  });
}

function attachDateCell(cell, row, columnKey) {
  cell.classList.add("date-cell");
  cell.tabIndex = 0;

  const render = () => renderDateCell(cell, row, columnKey);

  const openEditMode = ({ replaceWith, keepContent = false } = {}) => {
    if (IS_VIEWER_MODE) { return; }
    if (editingCell?.cell === cell) {
      return;
    }

    setCopyRange(null);

    cell.classList.add("is-editing");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "date-cell__input editor-overlay is-editing";
    const currentText = row[getDateFieldNames(columnKey).textField] || "";
    input.value = keepContent ? currentText : (replaceWith ?? currentText);
    cell.textContent = "";
    cell.appendChild(input);

    const cleanup = () => {
      if (editingCell?.cell === cell) {
        editingCell = null;
      }
      cell.classList.remove("is-editing");
      render();
      syncFillHandlePosition();
    };

    const commit = () => {
      setCellValue(cell, input.value, { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
      cleanup();
    };

    const cancel = () => {
      cleanup();
    };

    input.addEventListener("blur", commit, { once: true });

    editingCell = {
      cell,
      input,
      commit: () => input.blur(),
      cancel,
    };
    syncFillHandlePosition();
    
    requestAnimationFrame(() => {
      input.focus({ preventScroll: true });
      const end = input.value.length;
      input.setSelectionRange(end, end);
    });
  };

  cell.openEditMode = openEditMode;
  // click / dblclick / focus handled by delegated grid handlers.
  render();
}

function attachGenreCell(cell, row) {
  cell.classList.add("genre-cell");
  cell.tabIndex = 0;

  const render = () => {
    cell.textContent = row.genre || "";
  };

  const openEditMode = ({ keepContent = false, replaceWith } = {}) => {
    if (IS_VIEWER_MODE) { return; }
    if (editingCell?.cell === cell) {
      return;
    }

    setCopyRange(null);

    const column = getColumnByKey("genre");
    if (!column) {
      return;
    }

    const menu = ensureGenreMenuElement();
    cell.classList.add("is-editing");
    const currentValue = keepContent ? row.genre || "" : (replaceWith ?? row.genre ?? "");
    let highlightedIndex = Math.max(0, column.options.findIndex((option) => option === currentValue));
    const originalValue = row.genre || "";
    let cancelled = false;

    const commit = () => {
      if (!cancelled) {
        setCellValue(cell, row.genre, { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
      }
      cleanup();
    };

    const cancel = () => {
      cancelled = true;
      row.genre = originalValue;
      cleanup();
    };

    const renderOptions = () => {
      menu.innerHTML = "";
      column.options.forEach((option, index) => {
        const optionElement = document.createElement("button");
        optionElement.type = "button";
        optionElement.className = "genre-dropdown-menu__option";
        if (option === currentValue) {
          optionElement.classList.add("is-selected");
        }
        if (index === highlightedIndex) {
          optionElement.classList.add("is-highlighted");
        }
        optionElement.textContent = option;
        optionElement.setAttribute("role", "option");
        optionElement.setAttribute("aria-selected", option === currentValue ? "true" : "false");
        optionElement.addEventListener("mousedown", (event) => event.preventDefault());
        optionElement.addEventListener("click", () => {
          row.genre = option;
          commit();
        });
        menu.appendChild(optionElement);
      });
    };

    const positionMenu = () => {
      const cellRect = cell.getBoundingClientRect();
      const menuWidth = Math.max(0, cellRect.width - 2);
      menu.style.left = `${cellRect.left}px`;
      menu.style.top = `${cellRect.bottom - 1}px`;
      menu.style.width = `${menuWidth}px`;
      menu.style.maxWidth = `${menuWidth}px`;
      menu.classList.add("open");
    };

    const cleanup = () => {
      if (editingCell?.cell === cell) {
        editingCell = null;
      }
      document.removeEventListener("mousedown", handlePointerDownOutside);
      menu.classList.remove("open");
      window.removeEventListener("resize", positionMenu);
      cell.classList.remove("is-editing");
      render();
      syncFillHandlePosition();
    };

    const handleKeyDown = (event) => {
      if (event.key === "ArrowDown") {
        event.preventDefault();
        highlightedIndex = Math.min(column.options.length - 1, highlightedIndex + 1);
        renderOptions();
        return true;
      }

      if (event.key === "ArrowUp") {
        event.preventDefault();
        highlightedIndex = Math.max(0, highlightedIndex - 1);
        renderOptions();
        return true;
      }

      if (event.key === "Enter") {
        event.preventDefault();
        row.genre = column.options[highlightedIndex] || "";
        commit();
        const nextSelection = moveSelectionDownWithinBlock(cell);
        focusCellWithoutEditing(nextSelection.cell);
        return true;
      }

      if (event.key === "Escape") {
        event.preventDefault();
        cancel();
        focusCellWithoutEditing(cell);
        return true;
      }

      return false;
    };

    const handlePointerDownOutside = (event) => {
      if (!menu.contains(event.target) && !cell.contains(event.target)) {
        commit();
      }
    };

    renderOptions();
    positionMenu();
    window.addEventListener("resize", positionMenu);
    document.addEventListener("mousedown", handlePointerDownOutside);

    editingCell = {
      cell,
      input: menu,
      type: "select",
      commit,
      cancel,
      handleKeyDown,
    };
    syncFillHandlePosition();
  };

  cell.openEditMode = openEditMode;
  // click (which on genre also opens edit if already selected), focus
  // delegated to grid root handlers — see handleGridDelegatedClick.
  render();
}
function insertRow(blockIndex, atIndex) {
  insertRows(blockIndex, atIndex, 1);
}

function insertRows(blockIndex, atIndex, count = 1, options = {}) {
  if (!Number.isInteger(count) || count <= 0) {
    return [];
  }

  const block = blocks[blockIndex];
  if (!block) {
    return [];
  }

  const nextRows = [...block.rows];
  const rowsToInsert = Array.from({ length: count }, () => newRowForBlock(block.blockType, currentCalendarContext));
  nextRows.splice(atIndex, 0, ...rowsToInsert);
  blocks[blockIndex] = { ...block, rows: nextRows };

  addPatchToCurrentAction(
    {
      type: "insertRows",
      blockIndex,
      atIndex,
      rows: cloneRows(rowsToInsert),
    },
    { type: options.historyType || "rows", groupKey: options.groupKey || `insert:${blockIndex}:${atIndex}` }
  );

  if (options.render !== false) {
    renderRows();
  }

  return rowsToInsert;
}

function deleteRowsInBlock(blockIndex, startRow, endRow, options = {}) {
  const block = blocks[blockIndex];
  if (!block?.rows?.length) {
    return null;
  }

  const safeStart = Math.max(0, Math.min(startRow, endRow));
  const safeEnd = Math.min(block.rows.length - 1, Math.max(startRow, endRow));
  if (safeEnd < safeStart) {
    return null;
  }

  const hasStructuralRow = block.rows
    .slice(safeStart, safeEnd + 1)
    .some((row) => row?.isHeader || row?.isStructural);
  if (hasStructuralRow) {
    return null;
  }

  const removeCount = safeEnd - safeStart + 1;
  const removedRows = cloneRows(block.rows.slice(safeStart, safeEnd + 1));
  const nextRows = [...block.rows];
  nextRows.splice(safeStart, removeCount);

  if (!nextRows.length) {
    nextRows.push(newRowForBlock(block.blockType, currentCalendarContext));
  }

  blocks[blockIndex] = { ...block, rows: nextRows };

  addPatchToCurrentAction(
    {
      type: "deleteRows",
      blockIndex,
      atIndex: safeStart,
      rows: removedRows,
    },
    { type: options.historyType || "rows", groupKey: options.groupKey || `delete:${blockIndex}:${safeStart}` }
  );

  return {
    removedStart: safeStart,
    removedEnd: safeEnd,
    removeCount,
    lastRowIndex: nextRows.length - 1,
  };
}

function createSelectionState(blockIndex, rowIndex, columnKey) {
  return {
    blockIndex,
    rowIndex,
    columnKey,
  };
}

function normalizeSelectionAfterDelete(blockIndex, deleteInfo) {
  const block = blocks[blockIndex];
  const activeMeta = getCellMeta(selectedCell);

  if (dragSelection && dragSelection.blockIndex === blockIndex) {
    const selectionStartsBeforeDelete = dragSelection.r1 < deleteInfo.removedStart;
    const selectionEndsBeforeDelete = dragSelection.r2 < deleteInfo.removedStart;
    const selectionStartsAfterDelete = dragSelection.r1 > deleteInfo.removedEnd;

    if (selectionEndsBeforeDelete) {
      // Keep selection as-is.
    } else if (selectionStartsAfterDelete) {
      dragSelection = {
        ...dragSelection,
        r1: Math.max(0, dragSelection.r1 - deleteInfo.removeCount),
        r2: Math.max(0, dragSelection.r2 - deleteInfo.removeCount),
      };
    } else {
      dragSelection = null;
    }

    if (selectionStartsBeforeDelete && dragSelection && dragSelection.r2 < dragSelection.r1) {
      dragSelection = null;
    }
  }

  if (copyRange && copyRangeBlockIndex === blockIndex) {
    const intersects = !(copyRange.r2 < deleteInfo.removedStart || copyRange.r1 > deleteInfo.removedEnd);
    if (intersects) {
      setCopyRange(null);
    } else if (copyRange.r1 > deleteInfo.removedEnd) {
      setCopyRange(
        {
          ...copyRange,
          r1: Math.max(0, copyRange.r1 - deleteInfo.removeCount),
          r2: Math.max(0, copyRange.r2 - deleteInfo.removeCount),
        },
        copyRangeBlockIndex
      );
    }
  }

  let nextSelection = null;
  if (activeMeta && activeMeta.blockIndex === blockIndex) {
    if (activeMeta.rowIndex < deleteInfo.removedStart) {
      nextSelection = createSelectionState(blockIndex, activeMeta.rowIndex, activeMeta.columnKey);
    } else if (activeMeta.rowIndex > deleteInfo.removedEnd) {
      nextSelection = createSelectionState(
        blockIndex,
        Math.max(0, activeMeta.rowIndex - deleteInfo.removeCount),
        activeMeta.columnKey
      );
    } else {
      nextSelection = createSelectionState(
        blockIndex,
        Math.min(deleteInfo.removedStart, deleteInfo.lastRowIndex),
        activeMeta.columnKey
      );
    }
  } else if (selectedCellState) {
    nextSelection = null;
  }

  renderRows();

  if (nextSelection) {
    const nextCell = document.querySelector(
      `[data-block-index="${nextSelection.blockIndex}"][data-row-index="${nextSelection.rowIndex}"][data-column-key="${nextSelection.columnKey}"]`
    );
    if (nextCell) {
      setSelectedCell(nextCell);
      focusCellWithoutEditing(nextCell);
    } else {
      setSelectedCell(null);
    }
  } else {
    setSelectedCell(null);
  }

  renderDragSelectionPreview(dragSelection);
}

function executeDeleteRows(target) {
  if (!target) {
    return;
  }

  withHistoryAction("delete-rows", { groupKey: `delete:${target.blockIndex}:${target.startRow}:${target.endRow}` }, () => {
    const deleteInfo = deleteRowsInBlock(target.blockIndex, target.startRow, target.endRow, { historyType: "delete-rows" });
    if (!deleteInfo) {
      return;
    }

    normalizeSelectionAfterDelete(target.blockIndex, deleteInfo);
  });
}

function ensureContextMenuElement() {
  if (menuElement) {
    return menuElement;
  }

  menuElement = document.createElement("div");
  menuElement.className = "context-menu";
  menuElement.setAttribute("role", "menu");
  menuElement.innerHTML = `
    <button type="button" class="context-menu__item" data-action="above" role="menuitem">Añadir Filas encima</button>
    <button type="button" class="context-menu__item" data-action="below" role="menuitem">Añadir Filas debajo</button>
    <div class="context-menu__divider" role="separator"></div>
    <button type="button" class="context-menu__item" data-action="duplicate-above" role="menuitem">Duplicar filas encima</button>
    <button type="button" class="context-menu__item" data-action="duplicate-below" role="menuitem">Duplicar filas debajo</button>
    <div class="context-menu__divider" role="separator"></div>
    <button type="button" class="context-menu__item" data-action="delete" role="menuitem">Eliminar filas</button>
    <div class="context-menu__divider" role="separator"></div>
    <button type="button" class="context-menu__item" data-action="actualizado" role="menuitem">Marcar Actualizado</button>
  `;

  menuElement.addEventListener("click", (event) => {
    const clickTarget = event.target instanceof Element ? event.target : null;
    const target = clickTarget ? clickTarget.closest("[data-action]") : null;
    if (!target || !contextMenu.open) {
      return;
    }

    const blockIndex = Number.parseInt(contextMenu.blockIndex, 10);
    const rowIndex = Number.parseInt(contextMenu.rowIndex, 10);
    if (!Number.isInteger(blockIndex) || !Number.isInteger(rowIndex)) {
      closeContextMenu();
      return;
    }

    if (target.dataset.action === "above") {
      const insertTarget = getInsertTargetFromContext(blockIndex, rowIndex);
      if (!insertTarget) {
        closeContextMenu();
        return;
      }
      withHistoryAction("insert-rows", { groupKey: `insert:${blockIndex}:${insertTarget.startRow}` }, () => {
        insertRows(blockIndex, insertTarget.startRow, insertTarget.count, { historyType: "insert-rows" });
      });
      closeContextMenu();
      return;
    }

    if (target.dataset.action === "below") {
      const insertTarget = getInsertTargetFromContext(blockIndex, rowIndex);
      if (!insertTarget) {
        closeContextMenu();
        return;
      }
      withHistoryAction("insert-rows", { groupKey: `insert:${blockIndex}:${insertTarget.endRow + 1}` }, () => {
        insertRows(blockIndex, insertTarget.endRow + 1, insertTarget.count, { historyType: "insert-rows" });
      });
      closeContextMenu();
      return;
    }
    
    if (target.dataset.action === "duplicate-above") {
      duplicateRowsAroundSelection("above", blockIndex, rowIndex);
      closeContextMenu();
      return;
    }

    if (target.dataset.action === "duplicate-below") {
      duplicateRowsAroundSelection("below", blockIndex, rowIndex);
      closeContextMenu();
      return;
    }

    if (target.dataset.action === "delete") {
      const deleteTarget = getDeleteTarget(blockIndex, rowIndex);
      if (!deleteTarget) {
        return;
      }
      openDeleteConfirmModal(deleteTarget);
      closeContextMenu();
      return;
    }

    if (target.dataset.action === "actualizado") {
      toggleRowActualizado(blockIndex, rowIndex);
      closeContextMenu();
    }
  });

  document.body.appendChild(menuElement);
  return menuElement;
}

function updateContextMenuDeleteState() {
  if (!menuElement) {
    return;
  }

  const deleteItem = menuElement.querySelector('[data-action="delete"]');
  if (deleteItem) {
    const enabled = canDeleteRows(contextMenu.blockIndex, contextMenu.rowIndex);
    deleteItem.disabled = !enabled;
    deleteItem.classList.toggle("is-disabled", !enabled);
  }

  const actualizadoItem = menuElement.querySelector('[data-action="actualizado"]');
  if (actualizadoItem) {
    const row = blocks[contextMenu.blockIndex]?.rows?.[contextMenu.rowIndex];
    const enabled = !!row && !row._autoPlaceholder;
    actualizadoItem.disabled = !enabled;
    actualizadoItem.classList.toggle("is-disabled", !enabled);
    actualizadoItem.textContent = row?.actualizado ? "Desmarcar Actualizado" : "Marcar Actualizado";
  }
}

function toggleRowActualizado(blockIndex, rowIndex) {
  const block = blocks[blockIndex];
  const row = block?.rows?.[rowIndex];
  if (!row || row._autoPlaceholder) {
    return;
  }
  row.actualizado = !row.actualizado;
  renderRows();
}

function handleOutsidePointer(event) {
  if (menuElement && !menuElement.contains(event.target)) {
    closeContextMenu();
  }
}

function handleMenuEscape(event) {
  if (event.key === "Escape") {
    closeContextMenu();
  }
}

function closeContextMenu() {
  contextMenu = { open: false, x: 0, y: 0, blockIndex: -1, rowIndex: -1 };
  if (menuElement) {
    menuElement.classList.remove("open");
  }
  document.removeEventListener("mousedown", handleOutsidePointer);
  document.removeEventListener("keydown", handleMenuEscape);
}

function openContextMenu(event, blockIndex, rowIndex) {
  if (IS_VIEWER_MODE) { event.preventDefault(); return; }
  event.preventDefault();

  contextMenu = {
    open: true,
    x: event.clientX,
    y: event.clientY,
    blockIndex,
    rowIndex,
  };

  const menu = ensureContextMenuElement();
  menu.classList.add("open");

  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const EDGE_PADDING_PX = 8;

  const menuRect = menu.getBoundingClientRect();
  const maxLeft = Math.max(EDGE_PADDING_PX, viewportWidth - menuRect.width - EDGE_PADDING_PX);
  const maxTop = Math.max(EDGE_PADDING_PX, viewportHeight - menuRect.height - EDGE_PADDING_PX);

  const safeLeft = Math.min(contextMenu.x, maxLeft);
  const safeTop = Math.min(contextMenu.y, maxTop);

  menu.style.left = `${Math.max(EDGE_PADDING_PX, safeLeft)}px`;
  menu.style.top = `${Math.max(EDGE_PADDING_PX, safeTop)}px`;

  updateContextMenuDeleteState();
  
  document.addEventListener("mousedown", handleOutsidePointer);
  document.addEventListener("keydown", handleMenuEscape);
}

function createLeftRow({ group = false, cells = [], onToggleCollapse = null, collapsed = false, showToggle = group } = {}) {
  const leftRow = document.createElement("div");
  leftRow.className = `left-row ${group ? "group" : ""}`;

  for (let i = 0; i < 7; i++) {
    const cell = document.createElement("div");

    if (i === 0) {
      cell.classList.add("gutter");
      if (group && showToggle) {
        const toggleBtn = document.createElement("button");
        toggleBtn.className = "gutter-icon-btn gutter-icon-btn--collapse";
        toggleBtn.type = "button";
        toggleBtn.setAttribute("aria-label", collapsed ? "Desplegar bloque" : "Plegar bloque");
        toggleBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
        toggleBtn.textContent = collapsed ? "+" : "−";
        toggleBtn.addEventListener("click", () => {
          if (typeof onToggleCollapse === "function") {
            onToggleCollapse();
          }
        });

        cell.appendChild(toggleBtn);
      }
    } else if (group && i === 2 && cells[i] && typeof cells[i] === "object") {
      const leftText = document.createElement("span");
      leftText.textContent = cells[i].left || "";
      const rightText = document.createElement("span");
      rightText.textContent = cells[i].right || "";
      cell.classList.add("group-title-cell");
      cell.append(leftText, rightText);
    } else {
      cell.textContent = cells[i] || "";
    }

    leftRow.appendChild(cell);
  }

  return leftRow;
}

function createDayRow(group = false) {
  const dayRow = document.createElement("div");
  dayRow.className = `day-row ${group ? "group" : ""}`;
  const totalActiveDays = currentCalendarContext.daysInMonth;
  
  for (let day = 1; day <= 31; day++) {
    const dayCell = document.createElement("div");
    dayCell.className = `day-cell ${day > totalActiveDays ? "inactive" : ""}`;
    dayRow.appendChild(dayCell);
  }

  return dayRow;
}

function attachListoCheckbox(cell, row) {
  cell.classList.add("checkbox-cell");
  cell.textContent = "";
  cell.tabIndex = IS_VIEWER_MODE ? -1 : 0;

  const input = document.createElement("input");
  input.type = "checkbox";
  input.className = "listo-checkbox";
  input.checked = getRowListo(row);
  input.setAttribute("aria-label", "Marcar LISTO");

  if (IS_VIEWER_MODE) {
    input.disabled = true;
    input.style.pointerEvents = "none";
    cell.appendChild(input);
    return;
  }

  const toggleListo = (nextValue = !getRowListo(row)) => {
    if (IS_VIEWER_MODE) { return; }
    setCellValue(cell, nextValue ? "true" : "", { type: "toggle", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:listo` });
    input.checked = getRowListo(row);
  };

  // The checkbox input itself still owns its own change listener — native
  // form controls need it. Cell-level click / focus / space-key handling is
  // now delegated at the grid root via handleGridDelegatedClick et al.
  input.addEventListener("change", () => {
    toggleListo(input.checked);
  });

  // Exposed so the delegated grid click handler can toggle from clicks that
  // miss the input itself.
  cell.toggleListo = toggleListo;

  cell.appendChild(input);
}

function attachBlockListoCheckbox(cell, block) {
  cell.classList.add("checkbox-cell", "checkbox-cell--group-toggle");
  cell.textContent = "";

  const input = document.createElement("input");
  input.type = "checkbox";
  input.className = "listo-checkbox";
  input.setAttribute("aria-label", `Marcar LISTO para todo el bloque ${block.blockType}`);

  const syncInputState = () => {
    const visibleRows = getOrderedRowsForMonth(block, currentCalendarContext)
      .filter((item) => item.isVisibleInCurrentMonth && !item.row._autoPlaceholder)
      .map((item) => item.row);
    input.checked = visibleRows.length > 0 && visibleRows.every((row) => getRowListo(row));
  };

  const toggleBlock = () => {
    if (IS_VIEWER_MODE) { return; }
    // El check de bloque solo afecta a las piezas visibles en el mes actual, y
    // marca/desmarca su estado de emisión de ESE mes (no el de otros meses).
    const realRows = getOrderedRowsForMonth(block, currentCalendarContext)
      .filter((item) => item.isVisibleInCurrentMonth && !item.row._autoPlaceholder)
      .map((item) => item.row);
    const targetValue = !(realRows.length > 0 && realRows.every((row) => getRowListo(row)));
    const visibleKeys = new Set(realRows.map((row) => row.rowKey));
    const monthKey = monthKeyFor();
    withHistoryAction("toggle", { groupKey: `toggle-block:${block.id}` }, () => {
      block.rows.forEach((row, rowIndex) => {
        if (!visibleKeys.has(row.rowKey)) { return; }
        const before = getCellRawValue(row, "listo");
        const after = targetValue ? "true" : "";
        if (before !== after) {
          addPatchToCurrentAction(createSetCellPatch({ blockIndex: blocks.findIndex((candidate) => candidate.id === block.id), rowIndex, rowKey: row.rowKey, columnKey: "listo", monthKey }, before, after), { type: "toggle", groupKey: `toggle-block:${block.id}` });
          setRowListo(row, targetValue);
        }
      });
      renderRows();
    });
  };

  input.addEventListener("click", (event) => {
    event.stopPropagation();
    toggleBlock();
  });

  input.addEventListener("keydown", (event) => {
    if (event.key === " " || event.key === "Spacebar") {
      event.preventDefault();
      toggleBlock();
    }
  });

  syncInputState();
  cell.appendChild(input);
}

function attachTitleCell(cell, row) {
  cell.classList.add("title-cell");
  let isEditing = false;

  const placeCaretAtEnd = (editorEl) => {
    if (!editorEl || !editorEl.isConnected) {
      return;
    }

    if (editorEl instanceof HTMLInputElement || editorEl instanceof HTMLTextAreaElement) {
      const end = editorEl.value.length;
      editorEl.setSelectionRange(end, end);
      return;
    }

    if (!editorEl.isContentEditable) {
      return;
    }

    const selection = window.getSelection();
    if (!selection) {
      return;
    }

    const range = document.createRange();
    range.selectNodeContents(editorEl);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  };

  const focusTitleEditor = (editorEl) => {
    if (!editorEl || !editorEl.isConnected) {
      return;
    }

    editorEl.focus({ preventScroll: true });
    placeCaretAtEnd(editorEl);
  };
  
  const renderReadMode = () => {
    isEditing = false;
    cell.classList.remove("is-editing");    
    cell.textContent = "";
    const text = document.createElement("span");
    text.className = "title-cell__text";
    text.textContent = row.title || "";
    text.title = row.title || "";
    cell.appendChild(text);
  };

  const openEditMode = ({ replaceWith, keepContent = false } = {}) => {
    if (IS_VIEWER_MODE) {
      if (editingCell?.input) { editingCell.input.focus(); }
      return;
    }
    if (isEditing) {
      if (editingCell?.input) {
        editingCell.input.focus();
      }
      return;
    }

    setCopyRange(null);

    isEditing = true;
    cell.classList.add("is-editing");
    const overlayLayer = getTitleOverlayLayer();
    if (!overlayLayer) {
      return;
    }

    const input = document.createElement("input");
    input.type = "text";
    input.className = "title-cell__input editor-overlay is-editing";
    input.disabled = false;
    input.readOnly = false;
    input.maxLength = 100;
    input.value = replaceWith !== undefined ? replaceWith : row.title || "";
    if (keepContent) {
      input.value = row.title || "";
    }
    const originalValue = row.title || "";
    let cancelled = false;

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");

    const updateOverlayPosition = () => {
      const gridRoot = overlayLayer.parentElement;
      if (!gridRoot) {
        return;
      }

      const cellRect = cell.getBoundingClientRect();
      const rootRect = gridRoot.getBoundingClientRect();
      const styles = window.getComputedStyle(cell);
      const horizontalPadding = Number.parseFloat(styles.paddingLeft || "0") + Number.parseFloat(styles.paddingRight || "0") + 24;
      const fontWeight = window.getComputedStyle(input).fontWeight || styles.fontWeight;
      const fontSize = window.getComputedStyle(input).fontSize || styles.fontSize;
      const fontFamily = window.getComputedStyle(input).fontFamily || styles.fontFamily;

      let measuredWidth = cellRect.width;
      if (context) {
        context.font = `${fontWeight} ${fontSize} ${fontFamily}`;
        const textWidth = context.measureText(input.value || " ").width;
        measuredWidth = textWidth + horizontalPadding;
      }

      const maxWidth = Math.max(cellRect.width, rootRect.right - cellRect.left - 2);
      const width = Math.min(maxWidth, Math.max(cellRect.width, measuredWidth));

      input.style.left = `${cellRect.left - rootRect.left}px`;
      input.style.top = `${cellRect.top - rootRect.top}px`;
      input.style.width = `${width}px`;
      input.style.height = `${cellRect.height}px`;
    };

    const cleanupEditingState = () => {
      if (editingCell?.cell === cell) {
        editingCell = null;
      }
      syncFillHandlePosition();
    };

    const commit = () => {
      if (cancelled) {
        return;
      }

      const nextValue = (input.value || "").slice(0, 100);
      setCellValue(cell, nextValue, { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
      input.remove();
      window.removeEventListener("resize", updateOverlayPosition);
      cleanupEditingState();
      renderReadMode();
    };

    const cancel = () => {
      cancelled = true;
      row.title = originalValue;
      input.remove();
      window.removeEventListener("resize", updateOverlayPosition);
      cleanupEditingState();
      renderReadMode();
    };

    input.addEventListener("input", () => {
      if (input.value.length > 100) {
        input.value = input.value.slice(0, 100);
      }
    });

    input.addEventListener("blur", commit, { once: true });

    editingCell = {
      cell,
      input,
      commit: () => input.blur(),
      cancel,
    };
    syncFillHandlePosition();
    
    overlayLayer.appendChild(input);
    window.addEventListener("resize", updateOverlayPosition);
    updateOverlayPosition();

    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        focusTitleEditor(input);
      });
    });
  };

  cell.openEditMode = openEditMode;
  // click / dblclick / focus handled by delegated grid handlers.
  renderReadMode();
}

function attachIdTextCell(cell, row) {
  cell.classList.add("text-cell");
  cell.tabIndex = 0;

  const renderReadMode = () => {
    cell.classList.remove("is-editing");
    cell.textContent = row.id || "";
  };

  const openEditMode = ({ replaceWith, keepContent = false } = {}) => {
    if (IS_VIEWER_MODE) { return; }
    if (editingCell?.cell === cell) {
      return;
    }

    setCopyRange(null);

    cell.classList.add("is-editing");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "date-cell__input editor-overlay is-editing";
    const currentText = row.id || "";
    input.value = keepContent ? currentText : (replaceWith ?? currentText);
    cell.textContent = "";
    cell.appendChild(input);

    const cleanup = () => {
      if (editingCell?.cell === cell) {
        editingCell = null;
      }
      renderReadMode();
      syncFillHandlePosition();
    };

    const commit = () => {
      setCellValue(cell, input.value || "", { type: "edit", groupKey: `${cell.dataset.blockIndex}:${cell.dataset.rowIndex}:${cell.dataset.columnKey}` });
      cleanup();
    };

    const cancel = () => {
      cleanup();
    };

    input.addEventListener("blur", commit, { once: true });

    editingCell = {
      cell,
      input,
      commit: () => input.blur(),
      cancel,
    };
    syncFillHandlePosition();
    
    requestAnimationFrame(() => {
      input.focus({ preventScroll: true });
      const end = input.value.length;
      input.setSelectionRange(end, end);
    });
  };

  cell.openEditMode = openEditMode;

  // click / dblclick / focus handled by delegated grid handlers.
  renderReadMode();
}

function renderMonthBlockGrid(root) {
  if (IS_VIEWER_MODE) { document.body.classList.add("is-viewer-mode"); }
  const monthTitle = getMonthTitleText(DEFAULT_CALENDAR_CONTEXT.month, DEFAULT_CALENDAR_CONTEXT.year);
  const monthAriaLabel = getMonthAriaLabel(DEFAULT_CALENDAR_CONTEXT.month, DEFAULT_CALENDAR_CONTEXT.year);
  
  root.innerHTML = `
    <section class="panel-layout" aria-label="Panel de control M+">
      <header class="panel-layout__top-header">
        <img src="assets/img/cabecera_panel.svg" alt="Panel de control M+" />
      </header>

      <div class="panel-layout__month-strip" aria-label="Selector de mes">
        <div class="panel-layout__month-strip-inner">
          <div class="panel-layout__month-title">${monthTitle}</div>
          <div class="panel-layout__month-nav">
            <span class="panel-layout__month-nav-arrow">‹</span>
            <span class="panel-layout__month-nav-arrow">›</span>
          </div>
        </div>
      </div>

      <div class="panel-layout__toolbar" aria-label="Acciones del panel">
        <div class="panel-layout__toolbar-inner">
          ${IS_VIEWER_MODE ? `` : `
          <button type="button" class="save-drive-btn" id="save-drive-btn">GUARDAR</button>
          `}
          <button type="button" class="export-excel-btn export-excel-btn--viewer" data-export="aplicativo">EXPORTAR APLICATIVO</button>
          <div class="search-box-wrapper">
            <span class="search-box-icon" aria-hidden="true">⌕</span>
            <input type="text" class="search-box-input" placeholder="Buscar título..." autocomplete="off" aria-label="Buscar en el panel" />
            <button type="button" class="search-box-clear" aria-label="Limpiar búsqueda">✕</button>
          </div>
          ${IS_VIEWER_MODE ? `` : `
          <button type="button" class="history-toggle-btn" id="history-toggle-btn" aria-label="Últimos cambios">
            <img src="assets/img/ic_historial.svg" alt="" width="18" height="18" />
          </button>
          <div class="presence-indicator" id="presence-indicator" data-state="loading" role="status" aria-live="polite">
            <span class="presence-indicator__count">…</span>
            <span class="presence-indicator__tooltip" aria-hidden="true"></span>
          </div>
          `}
        </div>
      </div>

      <section class="month-block" aria-label="${monthAriaLabel}">
        <header class="month-block__header">
          <div class="left-header" id="left-header"></div>
          <div class="right-header-scroll" id="right-header-scroll">
            <div class="right-header-track" id="right-header-track"></div>
          </div>
        </header>
        <div class="month-block__body">
          <div class="month-block__body-scroll-wrapper">
            <div class="month-block__body-grid" tabindex="0" aria-label="Grid de planificación">
              <div class="left-grid" id="left-body"></div>
              <div class="right-body-scroll" id="right-body-scroll">
                <div id="right-body"></div>
              </div>
            </div>
          </div>
        </div>
      </section>

    </section>
  `;

  const leftHeader = root.querySelector("#left-header");
  const globalCollapseButton = document.createElement("button");
  globalCollapseButton.id = GLOBAL_COLLAPSE_BUTTON_ID;
  globalCollapseButton.className = "left-header-global-toggle gutter-icon-btn gutter-icon-btn--collapse";
  globalCollapseButton.type = "button";
  globalCollapseButton.addEventListener("click", () => {
    const allCollapsed = areAllBlocksCollapsed();
    setAllBlocksCollapsed(!allCollapsed);
    renderRows();
    scrollToTopAfterGlobalCollapse(root);
  });

  leftHeader.appendChild(globalCollapseButton);

  headers.forEach((label, index) => {
    const columnKey = columns[index].key;
    const cell = document.createElement("div");
    cell.className = "left-header-sortable";
    cell.dataset.sortKey = columnKey;
    cell.title = `Ordenar por ${label}`;

    const labelSpan = document.createElement("span");
    labelSpan.textContent = label;
    cell.appendChild(labelSpan);

    const arrow = document.createElement("span");
    arrow.className = "sort-arrow";
    cell.appendChild(arrow);

    cell.addEventListener("click", () => {
      if (sortState.key === columnKey) {
        sortState = sortState.dir === "asc"
          ? { key: columnKey, dir: "desc" }
          : { key: null, dir: "asc" };
      } else {
        sortState = { key: columnKey, dir: "asc" };
      }
      updateSortHeaderIndicators();
      renderRows();
    });

    leftHeader.appendChild(cell);
  });
  updateSortHeaderIndicators();
  updateGlobalCollapseButtonState();
  
  const calendarContext = updateCalendarContext(root);
  const dayHeader = root.querySelector("#right-header-track");
  for (let day = 1; day <= 31; day++) {
    const cell = document.createElement("div");
    cell.className = `day-cell ${day > calendarContext.daysInMonth ? "inactive" : ""}`;
    updateDayHeaderCell(cell, day, calendarContext);
    dayHeader.appendChild(cell);
  }

  renderRows();
  attachMonthNavigation(root);

  const gridRoot = root.querySelector(".month-block__body-grid");
  gridRoot?.addEventListener("keydown", handleGridEnterKey);
  gridRoot?.addEventListener("paste", handleGridPaste);
  gridRoot?.addEventListener("pointerdown", handleGridPointerDown);
  document.addEventListener("pointermove", handleGridPointerMove);
  document.addEventListener("pointerup", handleGridPointerUp);
  document.addEventListener("pointercancel", handleGridPointerCancel);
  gridRoot?.addEventListener("click", handleGridClickCapture, true);
  // Delegated cell-level handlers — replaces the per-cell click / dblclick /
  // focus listeners that used to be attached inside attachTitleCell,
  // attachDateCell, attachGenreCell, attachIdTextCell, attachListoCheckbox.
  // Cuts the listener count by ~3-4× the cell count on each renderRows().
  gridRoot?.addEventListener("click", handleGridDelegatedClick);
  gridRoot?.addEventListener("dblclick", handleGridDelegatedDblClick);
  gridRoot?.addEventListener("focusin", handleGridDelegatedFocusIn);
  if (!IS_VIEWER_MODE) { ensureFillHandleElement(); }
  attachExcelExportControls(root);
  attachSearchControls(root);
  
  const rightBodyScroll = gridRoot?.querySelector("#right-body-scroll");
  rightBodyScroll?.addEventListener("scroll", () => {
    syncFillHandlePosition();
    syncCopyAntsPosition();
  });
  window.addEventListener("resize", () => {
    syncFillHandlePosition();
    syncCopyAntsPosition();
  });
}

function updateSortHeaderIndicators() {
  const leftHeader = document.getElementById("left-header");
  if (!leftHeader) { return; }
  leftHeader.querySelectorAll(".left-header-sortable").forEach((cell) => {
    const key = cell.dataset.sortKey;
    const arrow = cell.querySelector(".sort-arrow");
    const isActive = sortState.key === key;
    cell.classList.toggle("sort-active", isActive);
    if (arrow) {
      arrow.textContent = isActive ? (sortState.dir === "asc" ? " ↑" : " ↓") : "";
    }
  });
}

function renderRows() {
  updateCalendarContext(document);
  document.querySelector(".panel-layout")
    ?.classList.toggle("search-is-active", !!searchQuery);
  validateAllRowsDateRanges();
  const leftBody = document.getElementById("left-body");
  const rightBody = document.getElementById("right-body");
  const gridRoot = document.querySelector(".month-block__body-grid");
  
  leftBody.innerHTML = "";
  rightBody.innerHTML = "";
  selectedCell = null;
  clearFillPreview();
  renderDragSelectionPreview(dragSelection);
  
  blocks.forEach((block, blockIndex) => {
    if (block.isSeparator) {
      const separatorLeftRow = createLeftRow({ group: true, showToggle: false });
      const separatorDayRow = createDayRow();

      separatorLeftRow.classList.add("left-row--section-separator");
      separatorDayRow.classList.add("day-row--section-separator");
      separatorLeftRow.dataset.separatorLabel = block.blockType.toUpperCase();
      separatorDayRow.dataset.separatorLabel = block.blockType.toUpperCase();
      
      if (block.headerColor) {
        separatorLeftRow.style.setProperty("--group-bg", block.headerColor);
        separatorDayRow.style.setProperty("--group-bg", block.headerColor);
      }

      leftBody.appendChild(separatorLeftRow);
      rightBody.appendChild(separatorDayRow);
      return;
    }

    const groupLeftRow = createLeftRow({
      group: true,
      cells: ["", "", { left: block.blockType.toUpperCase(), right: getMaxSimultaneousLabel(block.maxSimultaneous) }, "", "", "", ""],
      collapsed: !!block.collapsed,
      onToggleCollapse: () => {
        blocks[blockIndex] = { ...block, collapsed: !block.collapsed };
        renderRows();
      },
    });
    const groupDayRow = createDayRow(true);
    // Anclaje estable para localizar el check general del bloque (el índice
    // incluye separadores, así que no se puede deducir contando filas de grupo).
    groupLeftRow.dataset.blockIndex = String(blockIndex);

    if (block.headerColor) {
      groupLeftRow.style.setProperty("--group-bg", block.headerColor);
      groupDayRow.style.setProperty("--group-bg", block.headerColor);
    }

        attachBlockListoCheckbox(groupLeftRow.children[1], block);

if (searchQuery && !blockHasMatchForSearch(block, searchQuery)) {
      groupLeftRow.classList.add("is-search-dim");
      groupDayRow.classList.add("is-search-dim");
    }

    leftBody.appendChild(groupLeftRow);
    rightBody.appendChild(groupDayRow);

    const blockDailyCounts = getBlockDailyCounts(block, currentCalendarContext);
    groupDayRow.dataset.blockDailyCounts = blockDailyCounts.slice(1).join(",");

    if (block.collapsed) {
      return;
    }

    const orderedRows = getOrderedRowsForMonth(block, currentCalendarContext);
    orderedRows.forEach(({ row, sourceIndex }) => {
      if (IS_VIEWER_MODE && row._autoPlaceholder) { return; }
      const leftRow = createLeftRow();
      const dayRow = createDayRow();

      if (row.actualizado) {
        leftRow.classList.add("is-actualizado");
        dayRow.classList.add("is-actualizado");
      }

      attachListoCheckbox(leftRow.children[1], row);
      attachTitleCell(leftRow.children[2], row);

      leftRow.children[1].dataset.blockIndex = String(blockIndex);
      leftRow.children[1].dataset.rowIndex = String(sourceIndex);
      leftRow.children[1].dataset.rowId = row.rowKey;
      leftRow.children[1].dataset.columnKey = "listo";
      if (isSelectedCellState(row, "listo")) {
        selectedCell = leftRow.children[1];
        selectedCell.classList.add("is-selected");
      }

      leftRow.children[2].dataset.blockIndex = String(blockIndex);
      leftRow.children[2].dataset.rowIndex = String(sourceIndex);
      leftRow.children[2].dataset.rowId = row.rowKey;
      leftRow.children[2].dataset.columnKey = "title";
      leftRow.children[2].tabIndex = 0;
      if (isSelectedCellState(row, "title")) {
        selectedCell = leftRow.children[2];
        selectedCell.classList.add("is-selected");
      }

      attachDateCell(leftRow.children[3], row, "startDate");
      leftRow.children[3].dataset.blockIndex = String(blockIndex);
      leftRow.children[3].dataset.rowIndex = String(sourceIndex);
      leftRow.children[3].dataset.rowId = row.rowKey;
      leftRow.children[3].dataset.columnKey = "startDate";

      // Inicio implícito (sin fecha de inicio + con fecha de fin): el gantt usa
      // este día 1 del mes de origen como ancla para propagar y dibujar la barra.
      if (rowHasImplicitStart(row)) {
        const homeStart = getRowHomeMonthStart(row);
        leftRow.dataset.implicitStart =
          `${homeStart.getFullYear()}-${String(homeStart.getMonth() + 1).padStart(2, "0")}-01`;
      } else {
        delete leftRow.dataset.implicitStart;
      }
      if (isSelectedCellState(row, "startDate")) {
        selectedCell = leftRow.children[3];
        selectedCell.classList.add("is-selected");
      }

      attachDateCell(leftRow.children[4], row, "endDate");
      leftRow.children[4].dataset.blockIndex = String(blockIndex);
      leftRow.children[4].dataset.rowIndex = String(sourceIndex);
      leftRow.children[4].dataset.rowId = row.rowKey;
      leftRow.children[4].dataset.columnKey = "endDate";
      if (isSelectedCellState(row, "endDate")) {
        selectedCell = leftRow.children[4];
        selectedCell.classList.add("is-selected");
      }

      attachGenreCell(leftRow.children[5], row);
      leftRow.children[5].dataset.blockIndex = String(blockIndex);
      leftRow.children[5].dataset.rowIndex = String(sourceIndex);
      leftRow.children[5].dataset.rowId = row.rowKey;
      leftRow.children[5].dataset.columnKey = "genre";
      if (isSelectedCellState(row, "genre")) {
        selectedCell = leftRow.children[5];
        selectedCell.classList.add("is-selected");
      }

      attachIdTextCell(leftRow.children[6], row);
      leftRow.children[6].dataset.blockIndex = String(blockIndex);
      leftRow.children[6].dataset.rowIndex = String(sourceIndex);
      leftRow.children[6].dataset.rowId = row.rowKey;
      leftRow.children[6].dataset.columnKey = "id";
      if (isSelectedCellState(row, "id")) {
        selectedCell = leftRow.children[6];
        selectedCell.classList.add("is-selected");
      }

leftRow.addEventListener("contextmenu", (event) => openContextMenu(event, blockIndex, sourceIndex));
      dayRow.addEventListener("contextmenu", (event) => openContextMenu(event, blockIndex, sourceIndex));

      if (searchQuery && !rowMatchesSearch(row, searchQuery)) {
        leftRow.classList.add("is-search-dim");
        dayRow.classList.add("is-search-dim");
      }

      leftBody.appendChild(leftRow);
      rightBody.appendChild(dayRow);
    });
  });

  syncFillHandlePosition();
  syncCopyAntsPosition();
  updateGlobalCollapseButtonState();
}

function formatDraftTimestamp(ms) {
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) {
    return "";
  }
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// Si quedó trabajo local sin sincronizar de la última sesión, ofrecer restaurarlo.
function maybeOfferDraftRecovery() {
  if (IS_VIEWER_MODE) {
    return;
  }
  const draft = readDraft();
  if (!draft) {
    return;
  }
  showDraftRecoveryBanner(draft);
}

function showDraftRecoveryBanner(draft) {
  document.getElementById("draft-recovery-banner")?.remove();

  const banner = document.createElement("div");
  banner.id = "draft-recovery-banner";
  banner.setAttribute("role", "alert");
  banner.style.cssText = [
    "position:fixed", "left:50%", "top:16px", "transform:translateX(-50%)",
    "z-index:10000", "max-width:560px", "width:calc(100% - 32px)",
    "background:#fff7ed", "border:1px solid #f59e0b", "border-radius:10px",
    "box-shadow:0 8px 30px rgba(0,0,0,0.25)", "padding:14px 16px",
    "font-family:system-ui,-apple-system,sans-serif", "color:#7c2d12",
    "display:flex", "align-items:center", "gap:12px", "flex-wrap:wrap",
  ].join(";");

  const text = document.createElement("div");
  text.style.cssText = "flex:1;min-width:220px;font-size:14px;line-height:1.4";
  const when = formatDraftTimestamp(draft.savedAt);
  text.innerHTML =
    `<strong>Cambios sin guardar de tu última sesión</strong><br>` +
    `Se detectó trabajo local${when ? ` (${when})` : ""} que no llegó a guardarse en Drive. ` +
    `¿Quieres restaurarlo?`;

  const restoreBtn = document.createElement("button");
  restoreBtn.type = "button";
  restoreBtn.textContent = "Restaurar";
  restoreBtn.style.cssText =
    "background:#ea580c;color:#fff;border:0;padding:8px 16px;border-radius:6px;" +
    "font-size:14px;font-weight:600;cursor:pointer";

  const discardBtn = document.createElement("button");
  discardBtn.type = "button";
  discardBtn.textContent = "Descartar";
  discardBtn.style.cssText =
    "background:transparent;color:#7c2d12;border:1px solid #d6b08c;padding:8px 16px;" +
    "border-radius:6px;font-size:14px;font-weight:600;cursor:pointer";

  restoreBtn.addEventListener("click", () => {
    try {
      blocks = draft.blocks;
      renderRows();
      // Sigue habiendo cambios sin sincronizar con Drive: mantener el borrador.
      hasUnsavedChanges = true;
      scheduleDraftAutosave();
      showGridToast("Trabajo restaurado · recuerda Guardar en Drive");
    } catch (err) {
      console.error("[draft] error al restaurar:", err);
      showGridToast("No se pudo restaurar el borrador");
    }
    banner.remove();
  });

  discardBtn.addEventListener("click", () => {
    clearDraft();
    banner.remove();
  });

  banner.append(text, restoreBtn, discardBtn);
  document.body.appendChild(banner);
}

async function autoLoadFromDrive() {
  // Nunca sobrescribir cambios locales sin guardar con la versión de Drive.
  if (initialDriveLoadDone && hasUnsavedChanges) {
    console.warn("[autoLoad] omitido: hay cambios sin guardar en memoria");
    return;
  }
  if (!window.XLSX) {
    showGridToast("No se pudo cargar la librería de Excel");
    return;
  }
  if (!window.GoogleDrive && !IS_VIEWER_MODE) {
    showGridToast("Servicio no disponible");
    return;
  }

  showGridToast("Cargando datos...");

  try {
    let buffer;
    if (window.GoogleDrive) {
      buffer = await window.GoogleDrive.loadXlsxBuffer({ useAuth: !IS_VIEWER_MODE });
    } else {
      // Fallback para visor sin gdrive.js: fetch directo con API key
      const FILE_ID = window.PANEL_CONFIG?.GOOGLE_DRIVE_FILE_ID;
      const API_KEY = window.PANEL_CONFIG?.GOOGLE_API_KEY;
      if (!FILE_ID || !API_KEY) {
        throw new Error("PANEL_CONFIG missing");
      }
      const url = `https://www.googleapis.com/drive/v3/files/${FILE_ID}?alt=media&key=${API_KEY}`;
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      buffer = await response.arrayBuffer();
    }

    const workbook = window.XLSX.read(buffer, { type: "array", cellDates: false });
    const sheetsWithData = workbook.SheetNames.filter((name) => workbook.Sheets[name]);
    if (!sheetsWithData.length) {
      showGridToast("No se encontró ninguna hoja en el archivo");
      return;
    }

    const savedContext = { ...currentCalendarContext };

    sheetsWithData.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const matrix = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      const sheetContext = resolveContextFromSheetName(sheetName);
      if (sheetContext) {
        currentCalendarContext = sheetContext;
      }
      importRowsFromExcelMatrix(matrix);
    });

    currentCalendarContext = savedContext;
    applyCalendarContextToView(document);

    initialDriveLoadDone = true;
    // Snapshot the freshly loaded state — this is the baseline against which
    // local edits will be diffed on the next save (merge-on-save).
    loadedSnapshot = deepCloneBlocks(blocks);
    // Start broadcasting our heartbeat and polling for other sessions now
    // that Drive is reachable and authenticated.
    startPresenceTracking();
    // Pre-fetch the change history so the panel is instant when first opened.
    if (!IS_VIEWER_MODE) {
      loadHistory().catch(() => { /* swallow — handled inside */ });
    }
    // La versión de Drive es ahora el estado de referencia; cualquier borrador
    // local representa trabajo posterior sin sincronizar → ofrecer recuperarlo.
    maybeOfferDraftRecovery();

  } catch (err) {
    showGridToast("No se pudo cargar el Excel");
    console.error("autoLoadFromDrive error:", err);
  }
}

renderMonthBlockGrid(document.getElementById("app"));

if (IS_VIEWER_MODE) {
  autoLoadFromDrive();
} else {
  // Editor: requiere sign-in OAuth antes de cargar datos.
  if (window.GoogleDrive?.isSignedIn?.()) {
    autoLoadFromDrive();
  } else {
    document.addEventListener("gdrive:signedin", autoLoadFromDrive, { once: true });
    if (window.GoogleDrive?.showGate) {
      window.GoogleDrive.showGate();
    } else {
      // GIS aún no disponible — esperar y reintentar
      const waitInterval = setInterval(() => {
        if (window.GoogleDrive?.showGate) {
          clearInterval(waitInterval);
          window.GoogleDrive.showGate();
        }
      }, 100);
      setTimeout(() => clearInterval(waitInterval), 10000);
    }
  }
}
