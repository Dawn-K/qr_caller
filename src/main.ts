import QRCode from "qrcode";
import * as XLSX from "xlsx";

type NoticeLevel = "info" | "success" | "error";
type CallStatus = "未处理" | "已同意" | "已拒绝";
type AppPermissionState =
  | "idle"
  | "explaining"
  | "requesting"
  | "granted"
  | "denied_blocked"
  | "unsupported_blocked";
type FilePermissionMode = "read" | "readwrite";
type FilePermissionResult = "granted" | "denied" | "prompt";

const STATUS_COLUMN = "状态";
const STATUS_ORDER: CallStatus[] = ["未处理", "已同意", "已拒绝"];

interface CallerRow {
  id: number;
  position: number;
  values: Record<string, string>;
  status: CallStatus;
}

interface WorkbookState {
  workbook: XLSX.WorkBook;
  sheetName: string;
}

interface WritableSink {
  write(data: ArrayBuffer): Promise<void>;
  close(): Promise<void>;
}

interface FilePermissionDescriptor {
  mode: FilePermissionMode;
}

interface PermissionWritableHandle {
  name?: string;
  getFile(): Promise<File>;
  createWritable(): Promise<WritableSink>;
  queryPermission?: (descriptor: FilePermissionDescriptor) => Promise<FilePermissionResult>;
  requestPermission?: (descriptor: FilePermissionDescriptor) => Promise<FilePermissionResult>;
}

interface FilePickerWindow extends Window {
  showOpenFilePicker?: (options?: {
    excludeAcceptAllOption?: boolean;
    multiple?: boolean;
    types?: Array<{ description?: string; accept: Record<string, string[]> }>;
  }) => Promise<unknown[]>;
}

interface AppState {
  workbookState: WorkbookState | null;
  fileName: string;
  fileHandle: PermissionWritableHandle | null;
  lastHandle: PermissionWritableHandle | null;
  headers: string[];
  rows: CallerRow[];
  phoneHeader: string | null;
  filter: CallStatus;
  currentRowId: number | null;
  saveTimer: number | null;
  isSaving: boolean;
  qrToken: number;
  permissionState: AppPermissionState;
  permissionMessage: string;
}

const state: AppState = {
  workbookState: null,
  fileName: "",
  fileHandle: null,
  lastHandle: null,
  headers: [],
  rows: [],
  phoneHeader: null,
  filter: "未处理",
  currentRowId: null,
  saveTimer: null,
  isSaving: false,
  qrToken: 0,
  permissionState: "idle",
  permissionMessage: "",
};

const openFileBtn = getEl<HTMLButtonElement>("open-file-btn");
const fileNameEl = getEl<HTMLParagraphElement>("file-name");
const statusTabs = getEl<HTMLElement>("status-tabs");
const saveHint = getEl<HTMLParagraphElement>("save-hint");
const notice = getEl<HTMLDivElement>("notice");
const permissionExplainer = getEl<HTMLElement>("permission-explainer");
const blockedBanner = getEl<HTMLElement>("blocked-banner");
const blockedTitle = getEl<HTMLHeadingElement>("blocked-title");
const blockedMessage = getEl<HTMLParagraphElement>("blocked-message");
const retryAuthBtn = getEl<HTMLButtonElement>("retry-auth-btn");
const emptyState = getEl<HTMLElement>("empty-state");
const emptyTitle = getEl<HTMLHeadingElement>("empty-title");
const emptyDesc = getEl<HTMLParagraphElement>("empty-desc");
const cardView = getEl<HTMLElement>("card-view");
const listView = getEl<HTMLElement>("list-view");
const statusBadge = getEl<HTMLSpanElement>("status-badge");
const cardTitle = getEl<HTMLHeadingElement>("card-title");
const phoneText = getEl<HTMLDivElement>("phone-text");
const fieldList = getEl<HTMLDListElement>("field-list");
const qrCanvas = getEl<HTMLCanvasElement>("qr-canvas");
const prevBtn = getEl<HTMLButtonElement>("prev-btn");
const nextBtn = getEl<HTMLButtonElement>("next-btn");
const acceptBtn = getEl<HTMLButtonElement>("accept-btn");
const rejectBtn = getEl<HTMLButtonElement>("reject-btn");
const listTitle = getEl<HTMLHeadingElement>("list-title");
const listHeadRow = getEl<HTMLTableRowElement>("list-head-row");
const listBody = getEl<HTMLTableSectionElement>("list-body");

openFileBtn.addEventListener("click", () => {
  void handleOpenAuthorization();
});
retryAuthBtn.addEventListener("click", () => {
  void retryAuthorization();
});
prevBtn.addEventListener("click", () => moveCurrentCard(-1));
nextBtn.addEventListener("click", () => moveCurrentCard(1));
acceptBtn.addEventListener("click", () => setCurrentStatus("已同意"));
rejectBtn.addEventListener("click", () => setCurrentStatus("已拒绝"));

initializePermissionGate();
render();

function getEl<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (node === null) {
    throw new Error(`Missing element: ${id}`);
  }
  return node as T;
}

function setNotice(message: string, level: NoticeLevel = "info"): void {
  notice.textContent = message;
  notice.className = `notice ${level}`;
}

function setNoticeVisibility(visible: boolean): void {
  notice.classList.toggle("hidden", !visible);
}

function setEmptyMessage(title: string, description: string): void {
  emptyTitle.textContent = title;
  emptyDesc.textContent = description;
}

function updateSaveHint(text: string): void {
  saveHint.textContent = text;
}

function initializePermissionGate(): void {
  if (!isFileAccessSupported()) {
    state.permissionState = "unsupported_blocked";
    state.permissionMessage = "当前浏览器不支持文件读写授权。请使用最新版 Chrome 或 Edge。";
    updateSaveHint("当前浏览器不支持本功能。请切换到 Chrome/Edge。");
    setNotice("浏览器不支持 File System Access API。", "error");
    return;
  }
  state.permissionState = "explaining";
  setNotice("请先阅读权限说明，再点击“授权并打开 Excel”。", "info");
  updateSaveHint("授权成功后，状态会自动写回原 Excel 文件。");
}

function isFileAccessSupported(): boolean {
  const pickerWindow = window as FilePickerWindow;
  return typeof pickerWindow.showOpenFilePicker === "function";
}

function normalizeStatus(raw: string): CallStatus {
  const text = raw.trim();
  if (text === "已同意" || text === "同意" || text === "接受") {
    return "已同意";
  }
  if (text === "已拒绝" || text === "拒绝") {
    return "已拒绝";
  }
  return "未处理";
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function normalizePhone(phone: string): string {
  return phone.replace(/\s+/g, "").replace(/[\-()]/g, "").trim();
}

function dedupeHeaders(rawHeaders: string[]): string[] {
  const counter = new Map<string, number>();
  return rawHeaders.map((header, index) => {
    const base = header === "" ? `列${index + 1}` : header;
    const seen = counter.get(base) ?? 0;
    counter.set(base, seen + 1);
    return seen === 0 ? base : `${base}_${seen + 1}`;
  });
}

function detectPhoneHeader(headers: string[]): string | null {
  const keywords = ["电话号码", "手机号", "手机", "电话", "phone", "tel", "mobile"];
  for (const header of headers) {
    const normalized = header.toLowerCase().replace(/\s+/g, "");
    if (keywords.some((keyword) => normalized.includes(keyword))) {
      return header;
    }
  }
  return null;
}

function getRowsByStatus(status: CallStatus): CallerRow[] {
  return state.rows.filter((row) => row.status === status);
}

function getCurrentUnprocessedRows(): CallerRow[] {
  return getRowsByStatus("未处理");
}

function ensureCurrentRowForFilter(): void {
  if (state.filter !== "未处理") {
    return;
  }
  const rows = getCurrentUnprocessedRows();
  if (rows.length === 0) {
    state.currentRowId = null;
    return;
  }
  if (state.currentRowId === null || rows.every((row) => row.id !== state.currentRowId)) {
    const firstRow = rows[0];
    state.currentRowId = firstRow === undefined ? null : firstRow.id;
  }
}

function createWorkbookSheet(): XLSX.WorkSheet {
  const headers = [...state.headers, STATUS_COLUMN];
  const sortedRows = [...state.rows].sort((a, b) => a.position - b.position);
  const body = sortedRows.map((row) => {
    return headers.map((header) => (header === STATUS_COLUMN ? row.status : row.values[header] ?? ""));
  });
  const matrix: Array<Array<string>> = [headers, ...body];
  return XLSX.utils.aoa_to_sheet(matrix);
}

function upsertSheetInWorkbook(): WorkbookState | null {
  if (state.workbookState === null) {
    return null;
  }
  const { workbook, sheetName } = state.workbookState;
  workbook.Sheets[sheetName] = createWorkbookSheet();
  if (!workbook.SheetNames.includes(sheetName)) {
    workbook.SheetNames.unshift(sheetName);
  }
  return state.workbookState;
}

async function handleOpenAuthorization(): Promise<void> {
  if (state.permissionState === "requesting" || state.permissionState === "unsupported_blocked") {
    return;
  }
  setNotice("请在弹窗中选择 Excel 文件并授权读写权限。", "info");
  await authorizeAndOpen(false);
}

async function retryAuthorization(): Promise<void> {
  if (state.permissionState !== "denied_blocked") {
    return;
  }
  setNotice("正在重新申请读写权限。", "info");
  await authorizeAndOpen(true);
}

async function authorizeAndOpen(reuseLastHandle: boolean): Promise<void> {
  state.permissionState = "requesting";
  render();

  try {
    let handle: PermissionWritableHandle | null = reuseLastHandle ? state.lastHandle : null;

    if (handle === null) {
      handle = await pickWritableHandle();
      if (handle === null) {
        state.permissionState = state.workbookState === null ? "explaining" : "granted";
        setNotice("已取消选择文件。", "info");
        render();
        return;
      }
      state.lastHandle = handle;
    }

    const granted = await ensureReadWritePermission(handle);
    if (!granted) {
      blockByDenied("你已拒绝文件读写权限。请点击“重新授权”，授权成功后才能继续。", true);
      return;
    }

    const file = await handle.getFile();
    const loaded = await loadWorkbook(file, handle);
    if (!loaded) {
      state.permissionState = "granted";
      render();
      return;
    }

    state.permissionState = "granted";
    state.permissionMessage = "";
    setNotice(`读取成功：${state.rows.length} 条记录。`, "success");
    render();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (isAbortError(message)) {
      state.permissionState = state.workbookState === null ? "explaining" : "granted";
      setNotice("已取消授权。", "info");
      render();
      return;
    }
    blockByDenied(`权限申请失败：${message}`, false);
  }
}

async function pickWritableHandle(): Promise<PermissionWritableHandle | null> {
  const pickerWindow = window as FilePickerWindow;
  if (typeof pickerWindow.showOpenFilePicker !== "function") {
    return null;
  }
  const handles = await pickerWindow.showOpenFilePicker({
    excludeAcceptAllOption: false,
    multiple: false,
    types: [
      {
        description: "Excel 文件",
        accept: {
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          "application/vnd.ms-excel": [".xls"],
        },
      },
    ],
  });

  const candidate = handles[0];
  if (!isPermissionWritableHandle(candidate)) {
    throw new Error("浏览器没有返回可写文件句柄。");
  }
  return candidate;
}

function isPermissionWritableHandle(value: unknown): value is PermissionWritableHandle {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const maybe = value as Partial<PermissionWritableHandle>;
  return typeof maybe.getFile === "function" && typeof maybe.createWritable === "function";
}

async function ensureReadWritePermission(handle: PermissionWritableHandle): Promise<boolean> {
  const descriptor: FilePermissionDescriptor = { mode: "readwrite" };
  if (typeof handle.queryPermission === "function") {
    const existing = await handle.queryPermission(descriptor);
    if (existing === "granted") {
      return true;
    }
  }
  if (typeof handle.requestPermission === "function") {
    const requested = await handle.requestPermission(descriptor);
    return requested === "granted";
  }
  return false;
}

function blockByDenied(message: string, fromDeny: boolean): void {
  const unifiedMessage = "你已拒绝或撤销文件读写权限。请点击“重新授权”，授权成功后才能继续。";
  state.permissionState = "denied_blocked";
  state.permissionMessage = unifiedMessage;
  state.fileHandle = null;
  updateSaveHint("权限未授予，系统已阻断。请完成重新授权。");
  setNotice(fromDeny ? "读写权限被拒绝，请点击“重新授权”。" : message, "error");
  render();
}

function isAbortError(message: string): boolean {
  return message.includes("AbortError") || message.includes("aborted") || message.includes("The user aborted");
}

async function loadWorkbook(file: File, handle: PermissionWritableHandle): Promise<boolean> {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    if (sheetName === undefined) {
      throw new Error("工作簿没有可用工作表");
    }
    const sheet = workbook.Sheets[sheetName];
    if (sheet === undefined) {
      throw new Error("无法读取第一个工作表");
    }

    const rows = XLSX.utils.sheet_to_json<Array<unknown>>(sheet, {
      header: 1,
      raw: false,
      defval: "",
    });
    const headerRow = rows[0] ?? [];
    if (headerRow.length === 0) {
      throw new Error("Excel 第一行表头为空");
    }

    const dedupedHeaders = dedupeHeaders(headerRow.map(normalizeCell));
    const statusIndex = dedupedHeaders.findIndex((header) => header === STATUS_COLUMN);
    const dataHeaders = dedupedHeaders.filter((header) => header !== STATUS_COLUMN);
    const phoneHeader = detectPhoneHeader(dataHeaders);
    if (phoneHeader === null) {
      throw new Error("未找到电话号码列（例如 电话号码/手机号/phone）");
    }

    const parsedRows: CallerRow[] = [];
    let idCounter = 1;
    for (let index = 1; index < rows.length; index += 1) {
      const sourceRow = rows[index] ?? [];
      const values: Record<string, string> = {};
      for (const header of dataHeaders) {
        const headerIndex = dedupedHeaders.indexOf(header);
        const cell = headerIndex >= 0 ? sourceRow[headerIndex] : "";
        values[header] = normalizeCell(cell);
      }

      const rawStatus = statusIndex >= 0 ? normalizeCell(sourceRow[statusIndex]) : "";
      const status = normalizeStatus(rawStatus);
      const hasContent = dataHeaders.some((header) => values[header] !== "") || rawStatus !== "";
      if (!hasContent) {
        continue;
      }

      parsedRows.push({
        id: idCounter,
        position: idCounter,
        values,
        status,
      });
      idCounter += 1;
    }

    state.workbookState = { workbook, sheetName };
    state.fileName = file.name;
    state.fileHandle = handle;
    state.lastHandle = handle;
    state.headers = dataHeaders;
    state.rows = parsedRows;
    state.phoneHeader = phoneHeader;
    state.filter = "未处理";
    state.currentRowId = parsedRows.find((row) => row.status === "未处理")?.id ?? null;
    updateSaveHint("已连接原文件，状态将自动保存。");
    return true;
  } catch (error) {
    const text = error instanceof Error ? error.message : "未知错误";
    setNotice(`Excel 解析失败：${text}`, "error");
    updateSaveHint("文件读取失败，请检查格式后重新授权并打开。");
    return false;
  }
}

function renderTabs(): void {
  const shouldShowTabs = state.permissionState === "granted" && state.rows.length > 0;
  statusTabs.classList.toggle("hidden", !shouldShowTabs);
  if (!shouldShowTabs) {
    statusTabs.innerHTML = "";
    return;
  }

  statusTabs.innerHTML = "";
  for (const status of STATUS_ORDER) {
    const count = getRowsByStatus(status).length;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `tab-btn ${state.filter === status ? "active" : ""}`.trim();
    button.innerHTML = `<span>${status}</span><span>${count}</span>`;
    button.addEventListener("click", () => {
      if (state.permissionState !== "granted") {
        return;
      }
      state.filter = status;
      ensureCurrentRowForFilter();
      render();
    });
    statusTabs.appendChild(button);
  }
}

function renderPermissionGate(): void {
  const isUnsupported = state.permissionState === "unsupported_blocked";
  const isRequesting = state.permissionState === "requesting";
  const isDeniedBlocked = state.permissionState === "denied_blocked";

  permissionExplainer.classList.toggle("hidden", state.permissionState === "granted");
  setNoticeVisibility(!isDeniedBlocked);

  openFileBtn.disabled = isUnsupported || isRequesting;
  openFileBtn.textContent = isRequesting ? "授权中..." : "授权并打开 Excel";

  blockedBanner.classList.toggle("hidden", !isDeniedBlocked);
  blockedTitle.textContent = "需要读写权限";
  blockedMessage.textContent = state.permissionMessage;
  retryAuthBtn.classList.toggle("hidden", false);
  retryAuthBtn.disabled = isRequesting;
}

function setViewMode(mode: "empty" | "card" | "list" | "none"): void {
  emptyState.classList.toggle("hidden", mode !== "empty");
  cardView.classList.toggle("hidden", mode !== "card");
  listView.classList.toggle("hidden", mode !== "list");
}

function renderCard(row: CallerRow, positionInFilter: number, totalInFilter: number): void {
  const statusClass = row.status === "已同意" ? "status-accepted" : row.status === "已拒绝" ? "status-rejected" : "status-pending";
  statusBadge.textContent = row.status;
  statusBadge.className = `status-badge ${statusClass}`;
  cardTitle.textContent = `客户信息 (${positionInFilter + 1}/${totalInFilter})`;

  const phoneHeader = state.phoneHeader;
  const rawPhone = phoneHeader === null ? "" : row.values[phoneHeader] ?? "";
  const normalizedPhone = normalizePhone(rawPhone);
  phoneText.textContent = normalizedPhone === "" ? "电话号码为空" : normalizedPhone;

  fieldList.innerHTML = "";
  for (const header of state.headers) {
    const dt = document.createElement("dt");
    dt.textContent = header;
    const dd = document.createElement("dd");
    dd.textContent = row.values[header] ?? "";
    fieldList.append(dt, dd);
  }

  prevBtn.classList.toggle("hidden", positionInFilter === 0);
  nextBtn.classList.toggle("hidden", positionInFilter === totalInFilter - 1);

  acceptBtn.disabled = false;
  rejectBtn.disabled = false;

  state.qrToken += 1;
  const token = state.qrToken;
  void renderQr(normalizedPhone, token);
}

async function renderQr(phone: string, token: number): Promise<void> {
  const context = qrCanvas.getContext("2d");
  if (context === null) {
    return;
  }
  context.clearRect(0, 0, qrCanvas.width, qrCanvas.height);
  if (phone === "") {
    context.fillStyle = "#8b93a4";
    context.font = "16px Source Sans 3";
    context.textAlign = "center";
    context.fillText("无可用号码", qrCanvas.width / 2, qrCanvas.height / 2);
    return;
  }

  try {
    await QRCode.toCanvas(qrCanvas, `tel:${phone}`, {
      margin: 1,
      width: qrCanvas.width,
      color: { dark: "#10131f", light: "#ffffff" },
    });
  } catch {
    if (token !== state.qrToken) {
      return;
    }
    context.fillStyle = "#8b93a4";
    context.font = "16px Source Sans 3";
    context.textAlign = "center";
    context.fillText("二维码生成失败", qrCanvas.width / 2, qrCanvas.height / 2);
  }
}

function renderList(): void {
  listTitle.textContent = `${state.filter}列表`;
  const rows = getRowsByStatus(state.filter);
  const columns: string[] = [];
  if (state.phoneHeader !== null) {
    columns.push(state.phoneHeader);
  }
  for (const header of state.headers) {
    if (!columns.includes(header) && columns.length < 4) {
      columns.push(header);
    }
  }

  listHeadRow.innerHTML = "";
  for (const header of ["序号", ...columns]) {
    const th = document.createElement("th");
    th.textContent = header;
    listHeadRow.appendChild(th);
  }

  listBody.innerHTML = "";
  for (const [index, row] of rows.entries()) {
    const tr = document.createElement("tr");
    const serial = document.createElement("td");
    serial.textContent = String(index + 1);
    tr.appendChild(serial);
    for (const header of columns) {
      const td = document.createElement("td");
      td.textContent = row.values[header] ?? "";
      tr.appendChild(td);
    }
    listBody.appendChild(tr);
  }
}

function render(): void {
  renderPermissionGate();
  renderTabs();

  const hasWorkbook = state.workbookState !== null;
  fileNameEl.textContent = hasWorkbook ? `当前文件：${state.fileName}` : "未打开文件";

  if (state.permissionState === "unsupported_blocked") {
    setEmptyMessage("当前浏览器不支持此功能", "请使用最新版 Chrome 或 Edge 后重试。");
    setViewMode("empty");
    return;
  }

  if (state.permissionState === "denied_blocked") {
    setViewMode("none");
    return;
  }

  if (state.permissionState !== "granted") {
    setEmptyMessage("请先授权并打开 Excel", "为了把状态写回原文件，需要读写权限。文件仅在本地处理，不会上传。");
    setViewMode("empty");
    return;
  }

  if (!hasWorkbook) {
    setEmptyMessage("请授权并选择 Excel 文件", "请选择包含电话号码列的数据文件。状态列默认使用“状态”。");
    setViewMode("empty");
    return;
  }

  if (state.rows.length === 0) {
    setEmptyMessage("当前文件没有可展示数据", "请检查表头与数据行是否完整。状态列可留空。");
    setViewMode("empty");
    return;
  }

  if (state.filter === "未处理") {
    ensureCurrentRowForFilter();
    const unprocessedRows = getCurrentUnprocessedRows();
    if (unprocessedRows.length === 0 || state.currentRowId === null) {
      setEmptyMessage("未处理已全部完成", "你可以切换到“已同意/已拒绝”查看结果，或重新打开文件。");
      setViewMode("empty");
      return;
    }
    const index = unprocessedRows.findIndex((row) => row.id === state.currentRowId);
    const safeIndex = index === -1 ? 0 : index;
    const row = unprocessedRows[safeIndex];
    if (row === undefined) {
      setEmptyMessage("未处理数据异常", "请重新授权并打开文件后重试。");
      setViewMode("empty");
      return;
    }
    state.currentRowId = row.id;
    renderCard(row, safeIndex, unprocessedRows.length);
    setViewMode("card");
    return;
  }

  renderList();
  setViewMode("list");
}

function moveCurrentCard(delta: -1 | 1): void {
  if (state.permissionState !== "granted" || state.filter !== "未处理" || state.currentRowId === null) {
    return;
  }
  const rows = getCurrentUnprocessedRows();
  const index = rows.findIndex((row) => row.id === state.currentRowId);
  if (index < 0) {
    return;
  }
  const nextIndex = index + delta;
  if (nextIndex < 0 || nextIndex >= rows.length) {
    return;
  }
  const nextRow = rows[nextIndex];
  if (nextRow === undefined) {
    return;
  }
  state.currentRowId = nextRow.id;
  render();
}

function setCurrentStatus(nextStatus: CallStatus): void {
  if (state.permissionState !== "granted" || state.filter !== "未处理" || state.currentRowId === null) {
    return;
  }
  const currentRow = state.rows.find((row) => row.id === state.currentRowId);
  if (currentRow === undefined) {
    return;
  }

  currentRow.status = nextStatus;
  scheduleAutoSave();

  const nextRow = state.rows.find((row) => row.position > currentRow.position && row.status === "未处理")
    ?? state.rows.find((row) => row.status === "未处理")
    ?? null;

  if (nextRow === null) {
    state.currentRowId = null;
    setNotice("当前是最后一条未处理记录，已全部处理完成。", "info");
  } else {
    state.currentRowId = nextRow.id;
    setNotice(`已标记为${nextStatus}，自动切换到下一条。`, "success");
  }

  render();
}

function scheduleAutoSave(): void {
  if (state.permissionState !== "granted") {
    return;
  }
  if (state.fileHandle === null) {
    blockByDenied("写入权限不可用，请重新授权后继续。", false);
    return;
  }
  if (state.saveTimer !== null) {
    window.clearTimeout(state.saveTimer);
  }
  updateSaveHint("状态已变更，准备自动保存...");
  state.saveTimer = window.setTimeout(() => {
    void saveWorkbookToOriginal();
  }, 500);
}

async function saveWorkbookToOriginal(): Promise<void> {
  if (state.fileHandle === null || state.workbookState === null || state.permissionState !== "granted") {
    return;
  }
  if (state.isSaving) {
    return;
  }

  state.isSaving = true;
  try {
    upsertSheetInWorkbook();
    const buffer = XLSX.write(state.workbookState.workbook, {
      type: "array",
      bookType: "xlsx",
    }) as ArrayBuffer;
    const writable = await state.fileHandle.createWritable();
    await writable.write(buffer);
    await writable.close();
    updateSaveHint(`已自动保存：${new Date().toLocaleTimeString()}`);
  } catch {
    blockByDenied("自动保存失败，可能权限已被撤销。请点击“重新授权”恢复。", false);
  } finally {
    state.isSaving = false;
    if (state.saveTimer !== null) {
      window.clearTimeout(state.saveTimer);
      state.saveTimer = null;
    }
  }
}
