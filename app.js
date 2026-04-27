const palette = [
  "#2563eb", "#0ea5e9", "#14b8a6", "#22c55e", "#84cc16",
  "#eab308", "#f59e0b", "#f97316", "#ef4444", "#ec4899",
  "#a855f7", "#8b5cf6", "#64748b", "#0891b2", "#10b981"
];

const state = {
  displayMode: "plan",
  rowMode: "assignee",
  selectedDepartments: new Set(),
  selectedEpics: new Set(),
  selectedAssignees: new Set(),
  hideEmpty: true,
  records: [],
  originalRecords: [],
  selectedTaskId: null,
  source: null,
  fileName: "planning-viewer.csv",
  user: {
    name: "",
    clientId: ""
  },
  selectedWorkspaceOption: "",
  availableWorkspaces: [],
  remoteToasts: [],
  remoteFlashTaskIds: new Set(),
  sync: {
    client: null,
    channel: null,
    workspaceId: "",
    connected: false,
    applyingRemote: false,
    refreshTimer: null
  }
};

const dragState = {
  active: false,
  mode: null,
  recordId: null,
  startX: 0,
  startY: 0,
  startDate: null,
  endDate: null,
  originalAssignee: null,
  pointerId: null,
  moved: false
};

const monthMap = {
  jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06", juni: "06",
  jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
};
const monthShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const deptSearch = document.getElementById("deptSearch");
const epicSearch = document.getElementById("epicSearch");
const assigneeSearch = document.getElementById("assigneeSearch");
const deptList = document.getElementById("deptList");
const epicList = document.getElementById("epicList");
const assigneeList = document.getElementById("assigneeList");
const activeChips = document.getElementById("activeChips");
const board = document.getElementById("board");
const summary = document.getElementById("summary");
const changeStatus = document.getElementById("changeStatus");
const syncStatus = document.getElementById("syncStatus");
const taskModal = document.getElementById("taskModal");
const taskModalContent = document.getElementById("taskModalContent");
const taskModalBackdrop = document.getElementById("taskModalBackdrop");
const collaboratorName = document.getElementById("collaboratorName");
const currentWorkspaceLabel = document.getElementById("currentWorkspaceLabel");
const remoteToastStack = document.getElementById("remoteToastStack");
const workspaceModal = document.getElementById("workspaceModal");
const workspaceModalBackdrop = document.getElementById("workspaceModalBackdrop");
const workspaceUserNameInput = document.getElementById("workspaceUserName");
const workspaceList = document.getElementById("workspaceList");
const newWorkspaceIdInput = document.getElementById("newWorkspaceId");
const newWorkspaceCsvInput = document.getElementById("newWorkspaceCsv");
const workspaceModalNote = document.getElementById("workspaceModalNote");
const identityModal = document.getElementById("identityModal");
const identityModalBackdrop = document.getElementById("identityModalBackdrop");
const identityNameInput = document.getElementById("identityName");

const userStorageKey = "planning-viewer-user";
const workspaceStorageKey = "planning-viewer-workspace";
const defaultSyncConfig = {
  url: "https://xbzvlwbmdtqwdzoxnnxa.supabase.co",
  anonKey: "sb_publishable_qDXHe21m5T_OP0urEu95EA_5WBxqOYs",
  workspaceId: ""
};

function createClientId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `client-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeLabel(value) {
  const trimmed = String(value == null ? "" : value).trim();
  return trimmed || "-";
}

function hydrateUserIdentity() {
  let stored = null;
  try {
    stored = JSON.parse(localStorage.getItem(userStorageKey) || "null");
  } catch (_error) {
    localStorage.removeItem(userStorageKey);
  }

  state.user.clientId = stored && stored.clientId ? stored.clientId : createClientId();
  state.user.name = stored && stored.name ? String(stored.name).trim() : "";
  renderIdentity();
  persistUserIdentity(state.user.name);
}

function persistUserIdentity(name) {
  state.user.name = String(name || "").trim();
  if (!state.user.clientId) state.user.clientId = createClientId();
  localStorage.setItem(userStorageKey, JSON.stringify({
    name: state.user.name,
    clientId: state.user.clientId
  }));
  renderIdentity();
}

function renderIdentity() {
  collaboratorName.textContent = state.user.name || "Not set";
  identityNameInput.value = state.user.name;
  workspaceUserNameInput.value = state.user.name;
  currentWorkspaceLabel.textContent = state.sync.workspaceId || "Not selected";
}

function openIdentityModal(forceFocus) {
  identityModal.hidden = false;
  identityNameInput.value = state.user.name;
  if (forceFocus) {
    window.setTimeout(() => identityNameInput.focus(), 30);
  }
}

function persistWorkspaceSelection(workspaceId) {
  localStorage.setItem(workspaceStorageKey, workspaceId || "");
}

function hydrateWorkspaceSelection() {
  const stored = localStorage.getItem(workspaceStorageKey);
  if (stored) {
    state.sync.workspaceId = stored.trim();
  }
  renderIdentity();
}

function renderWorkspaceList() {
  if (!state.availableWorkspaces.length) {
    workspaceList.innerHTML = '<div class="editor-help">No published workspaces yet. Create one to start the shared board.</div>';
    return;
  }

  workspaceList.innerHTML = state.availableWorkspaces.map(workspace => `
    <button class="workspace-option ${workspace.id === state.selectedWorkspaceOption ? "active" : ""}" type="button" data-workspace-option="${escapeHtml(workspace.id)}">
      <div class="workspace-name">${escapeHtml(workspace.id)}</div>
      <div class="workspace-meta">${escapeHtml(workspace.file_name || "Shared planning board")}</div>
      <div class="workspace-meta">Updated ${escapeHtml(new Date(workspace.updated_at).toLocaleString("en-GB"))}</div>
    </button>
  `).join("");

  workspaceList.querySelectorAll("[data-workspace-option]").forEach(button => {
    button.addEventListener("click", () => {
      state.selectedWorkspaceOption = button.dataset.workspaceOption || "";
      renderWorkspaceList();
    });
  });
}

function setWorkspaceModalNote(message, tone = "info") {
  workspaceModalNote.textContent = message;
  workspaceModalNote.style.background = tone === "error" ? "#fef2f2" : "#eff6ff";
  workspaceModalNote.style.color = tone === "error" ? "#b91c1c" : "#1d4ed8";
}

function openWorkspaceModal() {
  workspaceModal.hidden = false;
  workspaceUserNameInput.value = state.user.name;
  newWorkspaceIdInput.value = "";
  newWorkspaceCsvInput.value = "";
  state.selectedWorkspaceOption = state.sync.workspaceId || state.selectedWorkspaceOption || (state.availableWorkspaces[0] ? state.availableWorkspaces[0].id : "");
  renderWorkspaceList();
  setWorkspaceModalNote(state.availableWorkspaces.length
    ? "Pick an existing workspace or create a new one."
    : "No shared workspaces found yet. Create a new one to begin.");
}

function closeWorkspaceModal() {
  workspaceModal.hidden = true;
}

async function fetchAvailableWorkspaces() {
  if (!state.sync.client) return [];
  const { data, error } = await state.sync.client
    .from("planning_workspaces")
    .select("id, file_name, updated_at")
    .order("updated_at", { ascending: false });

  if (error) {
    setWorkspaceModalNote(`Could not load workspaces: ${error.message}`, "error");
    return [];
  }

  state.availableWorkspaces = data || [];
  if (!state.selectedWorkspaceOption && state.availableWorkspaces[0]) {
    state.selectedWorkspaceOption = state.availableWorkspaces[0].id;
  }
  renderWorkspaceList();
  return state.availableWorkspaces;
}

async function activateWorkspace(workspaceId, options = {}) {
  const { loadExisting = true, notifyChanges = false, parsedData = null } = options;
  state.sync.workspaceId = workspaceId;
  state.selectedWorkspaceOption = workspaceId;
  persistWorkspaceSelection(workspaceId);
  renderIdentity();
  subscribeToWorkspace(workspaceId);

  if (parsedData) {
    applyParsedData(parsedData);
    await publishCurrentBoard();
    setSyncStatus(`Workspace "${workspaceId}" is ready and linked to ${parsedData.fileName}.`, "connected");
  } else if (loadExisting) {
    await loadWorkspaceFromRemote(false, notifyChanges);
  } else {
    await loadDefaultCsv();
    setSyncStatus(`Workspace "${workspaceId}" is ready.`, "connected");
  }
}

async function submitJoinWorkspace() {
  const name = workspaceUserNameInput.value.trim();
  const workspaceId = state.selectedWorkspaceOption;
  if (!name) {
    setWorkspaceModalNote("Add your name before joining a workspace.", "error");
    workspaceUserNameInput.focus();
    return;
  }
  if (!workspaceId) {
    setWorkspaceModalNote("Choose an existing workspace first.", "error");
    return;
  }
  persistUserIdentity(name);
  await activateWorkspace(workspaceId, { loadExisting: true, notifyChanges: false });
  closeWorkspaceModal();
}

async function submitCreateWorkspace() {
  const name = workspaceUserNameInput.value.trim();
  const workspaceId = newWorkspaceIdInput.value.trim();
  if (!name) {
    setWorkspaceModalNote("Add your name before creating a workspace.", "error");
    workspaceUserNameInput.focus();
    return;
  }
  if (!workspaceId) {
    setWorkspaceModalNote("Enter a workspace ID to create a new board.", "error");
    newWorkspaceIdInput.focus();
    return;
  }
  const file = newWorkspaceCsvInput.files && newWorkspaceCsvInput.files[0];
  if (!file) {
    setWorkspaceModalNote("Upload the CSV that belongs to this new workspace.", "error");
    newWorkspaceCsvInput.focus();
    return;
  }
  const text = await file.text();
  const parsed = parseCsvText(text, file.name);
  if (!parsed) {
    setWorkspaceModalNote("Could not parse this CSV. Please use the same sprint rehearsal export format.", "error");
    return;
  }
  persistUserIdentity(name);
  await activateWorkspace(workspaceId, { loadExisting: false, notifyChanges: false, parsedData: parsed });
  closeWorkspaceModal();
}

function closeIdentityModal() {
  identityModal.hidden = true;
}

function saveIdentityFromInput() {
  const name = identityNameInput.value.trim();
  if (!name) {
    identityNameInput.focus();
    return;
  }
  persistUserIdentity(name);
  closeIdentityModal();
}

function ensureUserIdentity() {
  if (state.user.name) return true;
  openIdentityModal(true);
  setSyncStatus("Add your name before connecting so teammates can see who changed what.", "dirty");
  return false;
}

function queueRemoteToast(toast) {
  if (!toast || state.remoteToasts.length >= 3) return;
  const id = `toast-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const nextToast = { id, ...toast };
  state.remoteToasts.push(nextToast);
  renderRemoteToasts();
  window.setTimeout(() => {
    state.remoteToasts = state.remoteToasts.filter(item => item.id !== id);
    renderRemoteToasts();
  }, 5000);
}

function renderRemoteToasts() {
  remoteToastStack.innerHTML = state.remoteToasts.map(toast => `
    <div class="toast-card">
      <div class="toast-title">${escapeHtml(toast.title)}</div>
      <div class="toast-body">${escapeHtml(toast.body)}</div>
    </div>
  `).join("");
}

function flashRemoteTask(recordId) {
  if (!recordId) return;
  state.remoteFlashTaskIds.add(recordId);
  renderView();
  window.setTimeout(() => {
    state.remoteFlashTaskIds.delete(recordId);
    renderView();
  }, 1800);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function cloneRecords(records) {
  return records.map(record => ({
    ...record,
    days: [...record.days]
  }));
}

function cloneRows(rows) {
  return rows.map(row => [...row]);
}

function getSyncConfig() {
  return {
    url: defaultSyncConfig.url,
    anonKey: defaultSyncConfig.anonKey,
    workspaceId: state.sync.workspaceId || defaultSyncConfig.workspaceId
  };
}

function setSyncStatus(message, tone) {
  syncStatus.textContent = message;
  syncStatus.classList.remove("dirty", "connected");
  if (tone === "dirty") syncStatus.classList.add("dirty");
  if (tone === "connected") syncStatus.classList.add("connected");
}

function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === "," && !inQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") i++;
      row.push(current);
      rows.push(row);
      row = [];
      current = "";
      continue;
    }

    current += ch;
  }

  if (current.length > 0 || row.length > 0) {
    row.push(current);
    rows.push(row);
  }

  return rows;
}

function csvEscape(value) {
  const stringValue = value == null ? "" : String(value);
  if (/[",\n]/.test(stringValue)) {
    return `"${stringValue.replaceAll('"', '""')}"`;
  }
  return stringValue;
}

function parseDate(value) {
  const s = String(value).trim();
  const dash = s.match(/^(\d{1,2})-([A-Za-z]+)-(\d{2})$/);
  if (dash) {
    const dd = dash[1].padStart(2, "0");
    const mm = monthMap[dash[2].toLowerCase()];
    const yy = "20" + dash[3];
    return mm ? `${yy}-${mm}-${dd}` : null;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return null;
}

function parseIsoDateParts(dateStr) {
  const match = String(dateStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return {
    year: Number(match[1]),
    month: Number(match[2]),
    day: Number(match[3])
  };
}

function createLocalDate(dateStr) {
  const parts = parseIsoDateParts(dateStr);
  if (!parts) return null;
  return new Date(parts.year, parts.month - 1, parts.day);
}

function formatIsoDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatTemplateDate(dateStr) {
  const date = createLocalDate(dateStr);
  if (!date) return "";
  const day = String(date.getDate()).padStart(2, "0");
  const month = monthShort[date.getMonth()];
  const year = String(date.getFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
}

function normalizeBusinessDate(dateStr, direction) {
  const date = createLocalDate(dateStr);
  if (!date) return null;
  while (date.getDay() === 0 || date.getDay() === 6) {
    date.setDate(date.getDate() + (direction === "backward" ? -1 : 1));
  }
  return formatIsoDate(date);
}

function addBusinessDays(startStr, count) {
  let remaining = Math.max(1, Number(count) || 1) - 1;
  const date = createLocalDate(startStr);
  if (!date) return startStr;
  while (remaining > 0) {
    date.setDate(date.getDate() + 1);
    const day = date.getDay();
    if (day !== 0 && day !== 6) remaining--;
  }
  return formatIsoDate(date);
}

function shiftBusinessDate(dateStr, delta) {
  const date = createLocalDate(dateStr);
  if (!date || !delta) return dateStr;
  const direction = delta > 0 ? 1 : -1;
  let remaining = Math.abs(delta);
  while (remaining > 0) {
    date.setDate(date.getDate() + direction);
    const day = date.getDay();
    if (day !== 0 && day !== 6) remaining--;
  }
  return formatIsoDate(date);
}

function dateRangeBusiness(startStr, endStr) {
  const out = [];
  const start = createLocalDate(startStr);
  const end = createLocalDate(endStr);
  if (!start || !end) return out;

  const cursor = new Date(start);
  while (cursor <= end) {
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) out.push(formatIsoDate(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return out;
}

function enrichRecord(record) {
  return {
    ...record,
    department: normalizeLabel(record.department),
    assignee: normalizeLabel(record.assignee),
    updatedByName: record.updatedByName || "",
    updatedByClient: record.updatedByClient || "",
    days: dateRangeBusiness(record.start, record.end)
  };
}

function parseAnchorDate(value, yearHint) {
  const match = String(value).trim().match(/^(\d{1,2})\s+([A-Za-z]+)/);
  if (!match) return null;
  const month = monthMap[match[2].toLowerCase()];
  if (!month) return null;
  return `${yearHint}-${month}-${match[1].padStart(2, "0")}`;
}

function buildHeaderDates(anchorDate, count) {
  const dates = [];
  let current = normalizeBusinessDate(anchorDate, "forward");
  while (dates.length < count && current) {
    dates.push(current);
    current = addBusinessDays(current, 2);
  }
  return dates;
}

function parseCsvRowsData(rows, fileName) {
  if (rows.length < 4) return null;
  const header = rows[2];
  const colIndex = {};
  header.forEach((heading, index) => { colIndex[heading] = index; });

  const deptIdx = colIndex["Department"];
  const epicIdx = colIndex["Epic"];
  const taskIdx = colIndex["Task Summary"];
  const assigneeIdx = colIndex["Assignee\n(use email)"];
  const startIdx = colIndex["Start"];
  const endIdx = colIndex["End"];
  const dayIdx = colIndex["Day"];

  if ([deptIdx, epicIdx, taskIdx, assigneeIdx, startIdx, endIdx, dayIdx].some(index => index === undefined)) {
    return null;
  }

  const dateColumnIndices = [];
  for (let i = dayIdx + 1; i < header.length; i++) {
    dateColumnIndices.push(i);
  }

  const yearHintRow = rows.slice(3).find(row => parseDate(row[startIdx]));
  const yearHint = yearHintRow ? parseDate(yearHintRow[startIdx]).slice(0, 4) : String(new Date().getFullYear());
  const anchorDate = parseAnchorDate(rows[0][8], yearHint);
  const dateColumnDates = buildHeaderDates(anchorDate, dateColumnIndices.length);
  const dateIndexMap = new Map();
  dateColumnDates.forEach((dateStr, index) => dateIndexMap.set(dateStr, dateColumnIndices[index]));

  const records = [];
  for (let rowIndex = 3; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const department = (row[deptIdx] || "").trim();
    const epic = (row[epicIdx] || "").trim();
    const task = (row[taskIdx] || "").trim();
    const assignee = (row[assigneeIdx] || "").trim();
    const start = parseDate(row[startIdx] || "");
    const end = parseDate(row[endIdx] || "");

    if (!epic || !task || !start || !end) continue;

    records.push(enrichRecord({
      id: `row-${rowIndex}`,
      rowIndex,
      department,
      epic,
      task,
      assignee,
      start,
      end
    }));
  }

  return {
    records,
    source: {
      rows,
      colIndex,
      dateColumnIndices,
      dateColumnDates,
      dateIndexMap
    },
    fileName: fileName || "planning-viewer.csv"
  };
}

function parseCsvText(text, fileName) {
  const rows = parseCsvRows(text);
  if (rows.length && rows[rows.length - 1].every(cell => cell === "")) rows.pop();
  return parseCsvRowsData(rows, fileName);
}

function getDepartments() {
  return Array.from(new Set(state.records.map(record => record.department))).sort();
}

function getEpics() {
  return Array.from(new Set(state.records.map(record => record.epic))).sort();
}

function getAssignees() {
  return Array.from(new Set(state.records.map(record => record.assignee))).sort();
}

function getColorMap() {
  const departments = getDepartments();
  const map = {};
  departments.forEach((department, index) => {
    map[department] = palette[index % palette.length];
  });
  return map;
}

function getDateCols() {
  if (state.records.length === 0) return [];
  let min = state.records[0].start;
  let max = state.records[0].end;
  state.records.forEach(record => {
    if (record.start < min) min = record.start;
    if (record.end > max) max = record.end;
  });
  return dateRangeBusiness(min, max);
}

function getOriginalRecord(recordId) {
  return state.originalRecords.find(record => record.id === recordId) || null;
}

function getRecord(recordId) {
  return state.records.find(record => record.id === recordId) || null;
}

function isTaskDirty(record) {
  const original = getOriginalRecord(record.id) || record;
  if (!original) return false;
  return record.assignee !== original.assignee || record.start !== original.start || record.end !== original.end;
}

function getDirtyRecords() {
  return state.records.filter(isTaskDirty);
}

function toWorkspaceTaskRows(records, workspaceId) {
  return records.map(record => ({
    id: record.id,
    workspace_id: workspaceId,
    row_index: record.rowIndex,
    department: record.department,
    epic: record.epic,
    task: record.task,
    assignee: record.assignee,
    start_date: record.start,
    end_date: record.end,
    updated_by_name: record.updatedByName || state.user.name || "",
    updated_by_client: record.updatedByClient || state.user.clientId || ""
  }));
}

function fromWorkspaceTaskRows(taskRows) {
  return taskRows
    .slice()
    .sort((a, b) => a.row_index - b.row_index)
    .map(row => enrichRecord({
      id: row.id,
      rowIndex: row.row_index,
      department: row.department,
      epic: row.epic,
      task: row.task,
      assignee: row.assignee,
      start: row.start_date,
      end: row.end_date,
      updatedByName: row.updated_by_name || "",
      updatedByClient: row.updated_by_client || ""
    }));
}

function buildRecordMap(records) {
  return new Map(records.map(record => [record.id, record]));
}

function describeFieldChange(label, before, after) {
  return `${label} from ${before} to ${after}`;
}

function getRecordChangeDetails(previous, next) {
  const details = [];
  if (previous.start !== next.start) details.push(describeFieldChange("start date", previous.start, next.start));
  if (previous.end !== next.end) details.push(describeFieldChange("end date", previous.end, next.end));
  if (previous.assignee !== next.assignee) details.push(describeFieldChange("assignee", previous.assignee, next.assignee));
  if (previous.department !== next.department) details.push(describeFieldChange("department", previous.department, next.department));
  return details;
}

function applyRemoteChanges(previousRecords, nextRecords, workspaceRow, notifyChanges) {
  const previousMap = buildRecordMap(previousRecords);
  const nextMap = buildRecordMap(nextRecords);
  const workspaceActorName = workspaceRow && workspaceRow.last_published_by_name ? workspaceRow.last_published_by_name : "A teammate";
  const workspaceActorClient = workspaceRow && workspaceRow.last_published_by_client ? workspaceRow.last_published_by_client : "";
  const changedTaskIds = [];

  if (!notifyChanges) return changedTaskIds;

  nextRecords.forEach(record => {
    const previous = previousMap.get(record.id);
    const actorName = record.updatedByName || workspaceActorName;
    const actorClient = record.updatedByClient || workspaceActorClient;

    if (!previous) {
      if (actorClient && actorClient === state.user.clientId) return;
      queueRemoteToast({
        title: `${actorName} added task ${record.task}.`,
        body: `${record.task} is now assigned to ${record.assignee} from ${record.start} to ${record.end}.`
      });
      changedTaskIds.push(record.id);
      return;
    }

    const changes = getRecordChangeDetails(previous, record);
    if (changes.length === 0) return;
    if (actorClient && actorClient === state.user.clientId) return;

    queueRemoteToast({
      title: `${actorName} edited task ${record.task}.`,
      body: changes.join(", ")
    });
    changedTaskIds.push(record.id);
  });

  previousRecords.forEach(record => {
    if (nextMap.has(record.id)) return;
    if (workspaceActorClient && workspaceActorClient === state.user.clientId) return;
    queueRemoteToast({
      title: `${workspaceActorName} deleted task ${record.task}.`,
      body: `${record.task} was removed from the shared board.`
    });
  });

  return changedTaskIds;
}

function clearSyncChannel() {
  if (state.sync.channel) {
    state.sync.channel.unsubscribe();
    state.sync.channel = null;
  }
}

function disconnectSync() {
  clearSyncChannel();
  state.sync.client = null;
  state.sync.connected = false;
  setSyncStatus("Live sync is temporarily unavailable.", "dirty");
}

function hasSupabaseSdk() {
  return Boolean(window.supabase && typeof window.supabase.createClient === "function");
}

async function connectSync() {
  const config = getSyncConfig();
  if (!hasSupabaseSdk()) {
    setSyncStatus("Supabase client library is not available in this browser session.", "dirty");
    return;
  }
  if (!config.url || !config.anonKey) {
    setSyncStatus("Supabase is missing its background configuration.", "dirty");
    return;
  }

  clearSyncChannel();
  state.sync.client = window.supabase.createClient(config.url, config.anonKey);
  state.sync.connected = true;
  setSyncStatus("Live sync is ready in the background. Choose a workspace to start collaborating.", "connected");
}

function scheduleRemoteRefresh() {
  if (!state.sync.connected || !state.sync.workspaceId) return;
  if (state.sync.refreshTimer) clearTimeout(state.sync.refreshTimer);
  state.sync.refreshTimer = setTimeout(() => {
    loadWorkspaceFromRemote(true);
  }, 150);
}

function subscribeToWorkspace(workspaceId) {
  if (!state.sync.client) return;
  clearSyncChannel();

  state.sync.channel = state.sync.client
    .channel(`planning-workspace:${workspaceId}`)
    .on("postgres_changes", {
      event: "*",
      schema: "public",
      table: "planning_tasks",
      filter: `workspace_id=eq.${workspaceId}`
    }, scheduleRemoteRefresh)
    .on("postgres_changes", {
      event: "*",
      schema: "public",
      table: "planning_workspaces",
      filter: `id=eq.${workspaceId}`
    }, scheduleRemoteRefresh)
    .subscribe();
}

async function publishCurrentBoard() {
  if (!state.sync.connected || !state.sync.client || !state.sync.workspaceId) {
    setSyncStatus("Choose a workspace before publishing the board.", "dirty");
    return;
  }
  if (!state.source || state.records.length === 0) {
    setSyncStatus("Upload a CSV before publishing a shared workspace.", "dirty");
    return;
  }

  const workspaceId = state.sync.workspaceId;
  const client = state.sync.client;

  const { error: workspaceError } = await client
    .from("planning_workspaces")
    .upsert({
      id: workspaceId,
      file_name: state.fileName,
      csv_rows: state.source.rows,
      last_published_by_name: state.user.name,
      last_published_by_client: state.user.clientId
    });
  if (workspaceError) {
    setSyncStatus(`Could not publish workspace metadata: ${workspaceError.message}`, "dirty");
    return;
  }

  const { error: deleteError } = await client
    .from("planning_tasks")
    .delete()
    .eq("workspace_id", workspaceId);
  if (deleteError) {
    setSyncStatus(`Could not replace workspace tasks: ${deleteError.message}`, "dirty");
    return;
  }

  const { error: insertError } = await client
    .from("planning_tasks")
    .insert(toWorkspaceTaskRows(state.records, workspaceId));
  if (insertError) {
    setSyncStatus(`Could not publish tasks: ${insertError.message}`, "dirty");
    return;
  }

  setSyncStatus(`Published ${state.records.length} tasks to workspace "${workspaceId}".`, "connected");
  await fetchAvailableWorkspaces();
}

async function loadWorkspaceFromRemote(silent, notifyChanges = true) {
  if (!state.sync.connected || !state.sync.client || !state.sync.workspaceId) {
    if (!silent) setSyncStatus("Choose a workspace before loading shared data.", "dirty");
    return;
  }

  const workspaceId = state.sync.workspaceId;
  const client = state.sync.client;

  const { data: workspaceRow, error: workspaceError } = await client
    .from("planning_workspaces")
    .select("id, file_name, csv_rows, last_published_by_name, last_published_by_client")
    .eq("id", workspaceId)
    .maybeSingle();
  if (workspaceError) {
    if (!silent) setSyncStatus(`Could not load workspace metadata: ${workspaceError.message}`, "dirty");
    return;
  }
  if (!workspaceRow) {
    if (!silent) setSyncStatus(`Workspace "${workspaceId}" does not exist yet. Publish a board first.`, "dirty");
    return;
  }

  const { data: taskRows, error: taskError } = await client
    .from("planning_tasks")
    .select("id, workspace_id, row_index, department, epic, task, assignee, start_date, end_date, updated_by_name, updated_by_client")
    .eq("workspace_id", workspaceId);
  if (taskError) {
    if (!silent) setSyncStatus(`Could not load workspace tasks: ${taskError.message}`, "dirty");
    return;
  }

  const parsed = parseCsvRowsData(workspaceRow.csv_rows, workspaceRow.file_name);
  if (!parsed) {
    if (!silent) setSyncStatus("The workspace metadata could not be parsed back into a CSV board.", "dirty");
    return;
  }

  const previousRecords = cloneRecords(state.records);
  const nextRecords = fromWorkspaceTaskRows(taskRows || []);
  const changedTaskIds = applyRemoteChanges(previousRecords, nextRecords, workspaceRow, notifyChanges && previousRecords.length > 0);

  state.sync.applyingRemote = true;
  state.source = parsed.source;
  state.fileName = parsed.fileName;
  state.records = nextRecords;
  state.originalRecords = cloneRecords(state.records);
  if (state.selectedTaskId && !getRecord(state.selectedTaskId)) state.selectedTaskId = null;
  state.sync.applyingRemote = false;
  renderAll();
  changedTaskIds.forEach(flashRemoteTask);
  if (!silent) setSyncStatus(`Loaded workspace "${workspaceId}" with ${state.records.length} tasks.`, "connected");
}

async function persistTaskToWorkspace(record) {
  if (!record || !state.sync.connected || !state.sync.client || !state.sync.workspaceId || state.sync.applyingRemote) {
    return;
  }

  record.updatedByName = state.user.name;
  record.updatedByClient = state.user.clientId;
  const { error } = await state.sync.client
    .from("planning_tasks")
    .upsert({
      id: record.id,
      workspace_id: state.sync.workspaceId,
      row_index: record.rowIndex,
      department: record.department,
      epic: record.epic,
      task: record.task,
      assignee: record.assignee,
      start_date: record.start,
      end_date: record.end,
      updated_by_name: record.updatedByName,
      updated_by_client: record.updatedByClient
    });
  if (error) {
    setSyncStatus(`Could not sync task "${record.task}": ${error.message}`, "dirty");
    return;
  }

  setSyncStatus(`Live sync connected to "${state.sync.workspaceId}". Last synced task: ${record.task}.`, "connected");
}

function updateStats(filtered) {
  document.getElementById("statTasks").textContent = filtered.length.toLocaleString();
  document.getElementById("statDepts").textContent = new Set(filtered.map(record => record.department)).size.toLocaleString();
  document.getElementById("statEpics").textContent = new Set(filtered.map(record => record.epic)).size.toLocaleString();
  document.getElementById("statAssignees").textContent = new Set(filtered.map(record => record.assignee)).size.toLocaleString();
}

function addOptions(container, values, selectedSet, showColor) {
  const colorMap = getColorMap();
  container.innerHTML = "";
  values.forEach(value => {
    const label = document.createElement("label");
    label.className = "option";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = selectedSet.has(value);
    input.addEventListener("change", () => {
      if (input.checked) selectedSet.add(value);
      else selectedSet.delete(value);
      renderAll();
    });
    label.appendChild(input);

    if (showColor) {
      const swatch = document.createElement("span");
      swatch.className = "swatch";
      swatch.style.background = colorMap[value] || "#94a3b8";
      label.appendChild(swatch);
    }

    const text = document.createElement("span");
    text.textContent = value;
    label.appendChild(text);
    container.appendChild(label);
  });
}

function renderOptions() {
  const dTerm = deptSearch.value.trim().toLowerCase();
  const eTerm = epicSearch.value.trim().toLowerCase();
  const aTerm = assigneeSearch.value.trim().toLowerCase();

  addOptions(deptList, getDepartments().filter(value => value.toLowerCase().includes(dTerm)), state.selectedDepartments, true);
  addOptions(epicList, getEpics().filter(value => value.toLowerCase().includes(eTerm)), state.selectedEpics, false);
  addOptions(assigneeList, getAssignees().filter(value => value.toLowerCase().includes(aTerm)), state.selectedAssignees, false);
}

function getFilteredRecords() {
  return state.records.filter(record => {
    const deptOk = state.selectedDepartments.size === 0 || state.selectedDepartments.has(record.department);
    const epicOk = state.selectedEpics.size === 0 || state.selectedEpics.has(record.epic);
    const assigneeOk = state.selectedAssignees.size === 0 || state.selectedAssignees.has(record.assignee);
    return deptOk && epicOk && assigneeOk;
  });
}

function renderChips() {
  activeChips.innerHTML = "";

  const viewChip = document.createElement("span");
  viewChip.className = "chip";
  viewChip.textContent = state.displayMode === "plan" ? "View: Plan" : "View: Epic summary";
  activeChips.appendChild(viewChip);

  if (state.displayMode === "plan") {
    const rowsChip = document.createElement("span");
    rowsChip.className = "chip";
    rowsChip.textContent = state.rowMode === "assignee" ? "Rows: Assignee" : "Rows: Epic";
    activeChips.appendChild(rowsChip);
  }

  const dirtyCount = getDirtyRecords().length;
  if (dirtyCount > 0) {
    const dirtyChip = document.createElement("span");
    dirtyChip.className = "chip";
    dirtyChip.textContent = `${dirtyCount} edited`;
    activeChips.appendChild(dirtyChip);
  }

  if (!state.hideEmpty) {
    const emptyChip = document.createElement("span");
    emptyChip.className = "chip";
    emptyChip.textContent = "Showing empty rows";
    activeChips.appendChild(emptyChip);
  }

  [...state.selectedDepartments].sort().forEach(value => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = `Dept: ${value}`;
    activeChips.appendChild(chip);
  });

  [...state.selectedEpics].sort().forEach(value => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = `Epic: ${value}`;
    activeChips.appendChild(chip);
  });

  [...state.selectedAssignees].sort().forEach(value => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = `Assignee: ${value}`;
    activeChips.appendChild(chip);
  });
}

function dayHeader(dateStr) {
  const date = createLocalDate(dateStr);
  if (!date) return "";
  const num = date.toLocaleDateString("en-GB", { day: "2-digit" });
  const month = date.toLocaleDateString("en-GB", { month: "short" });
  const weekday = date.toLocaleDateString("en-GB", { weekday: "short" });
  return `<div class="day-num">${num}</div><div class="day-sub">${weekday} • ${month}</div>`;
}

function buildRows(filtered, dateCols) {
  const rowMap = new Map();

  filtered.forEach(record => {
    const key = state.rowMode === "assignee" ? record.assignee : record.epic;
    if (!rowMap.has(key)) {
      rowMap.set(key, {
        name: key,
        tasks: [],
        departments: new Set(),
        assignees: new Set()
      });
    }
    const row = rowMap.get(key);
    row.tasks.push(record);
    row.departments.add(record.department);
    row.assignees.add(record.assignee);
  });

  const rows = Array.from(rowMap.values()).map(row => {
    row.tasks.sort((a, b) => a.start.localeCompare(b.start) || a.end.localeCompare(b.end) || a.task.localeCompare(b.task));
    row.sub = state.rowMode === "assignee"
      ? Array.from(row.departments).sort().join(", ")
      : `Assignees: ${Array.from(row.assignees).sort().join(", ")}`;

    const lanes = [];
    row.tasks.forEach(task => {
      const startIdx = dateCols.indexOf(task.days[0]);
      const endIdx = dateCols.indexOf(task.days[task.days.length - 1]);
      const spanStart = startIdx < 0 ? 0 : startIdx;
      const spanEnd = endIdx < 0 ? spanStart : endIdx;

      let lane = 0;
      while (true) {
        if (!lanes[lane]) {
          lanes[lane] = [];
          break;
        }
        const overlap = lanes[lane].some(existing => !(spanEnd < existing.startIdx || spanStart > existing.endIdx));
        if (!overlap) break;
        lane++;
      }

      const placed = { ...task, startIdx: spanStart, endIdx: spanEnd, lane };
      lanes[lane].push(placed);
      task._placed = placed;
    });

    row.laneCount = Math.max(1, lanes.length);
    row.placedTasks = row.tasks.map(task => task._placed);
    row.firstDay = row.placedTasks.length ? row.placedTasks[0].start : "9999-12-31";
    return row;
  });

  rows.sort((a, b) => a.firstDay.localeCompare(b.firstDay) || a.name.localeCompare(b.name));
  return state.hideEmpty ? rows.filter(row => row.placedTasks.length > 0) : rows;
}

function taskBarHtml(task, colorMap) {
  const left = task.startIdx * 150;
  const width = ((task.endIdx - task.startIdx + 1) * 150) - 8;
  const top = 8 + (task.lane * 58);
  const color = colorMap[task.department] || "#64748b";
  const sub = state.rowMode === "assignee"
    ? `${escapeHtml(task.epic)} • ${escapeHtml(task.start)} to ${escapeHtml(task.end)}`
    : `${escapeHtml(task.assignee)} • ${escapeHtml(task.start)} to ${escapeHtml(task.end)}`;
  const classes = [
    "task-bar",
    isTaskDirty(task) ? "dirty" : "",
    state.remoteFlashTaskIds.has(task.id) ? "remote-flash" : "",
    task.id === state.selectedTaskId ? "selected" : ""
  ].filter(Boolean).join(" ");
  const dirtyMark = isTaskDirty(task) ? '<span class="task-dirty-mark"></span>' : "";

  return `<div class="${classes}" data-task-id="${escapeHtml(task.id)}" style="left:${left}px; width:${width}px; top:${top}px; background:${color};" title="${escapeHtml(`${task.task} | ${sub}`)}"><span class="resize-handle left" data-handle="start"></span><span class="resize-handle right" data-handle="end"></span>${dirtyMark}<div class="task-title">${escapeHtml(task.task)}</div><div class="task-sub">${sub}</div></div>`;
}

function syncTaskBarVisual(record, element, mode) {
  const dateCols = getDateCols();
  const startIdx = dateCols.indexOf(record.days[0]);
  const endIdx = dateCols.indexOf(record.days[record.days.length - 1]);
  if (startIdx < 0 || endIdx < 0) return;
  element.style.left = `${startIdx * 150}px`;
  element.style.width = `${((endIdx - startIdx + 1) * 150) - 8}px`;
  element.classList.toggle("dragging", mode === "move");
  element.classList.toggle("resizing-start", mode === "resize-start");
  element.classList.toggle("resizing-end", mode === "resize-end");
}

function clearDragVisuals(element) {
  if (!element) return;
  element.classList.remove("dragging", "resizing-start", "resizing-end");
}

function openTaskModal(recordId) {
  state.selectedTaskId = recordId;
  renderTaskModal();
  taskModal.hidden = false;
}

function closeTaskModal() {
  taskModal.hidden = true;
}

function handleTaskPointerMove(event, element) {
  if (!dragState.active) return;
  const record = getRecord(dragState.recordId);
  if (!record) return;

  const deltaX = event.clientX - dragState.startX;
  const deltaY = event.clientY - dragState.startY;
  const threshold = dragState.mode === "move" ? 10 : 6;
  if (!dragState.moved && Math.abs(deltaX) < threshold && Math.abs(deltaY) < threshold) {
    return;
  }
  dragState.moved = true;

  const deltaDays = Math.round((event.clientX - dragState.startX) / 150);
  let nextStart = dragState.startDate;
  let nextEnd = dragState.endDate;

  if (dragState.mode === "move") {
    nextStart = shiftBusinessDate(dragState.startDate, deltaDays);
    nextEnd = shiftBusinessDate(dragState.endDate, deltaDays);
  } else if (dragState.mode === "resize-start") {
    nextStart = shiftBusinessDate(dragState.startDate, deltaDays);
    if (nextStart > nextEnd) nextStart = nextEnd;
  } else if (dragState.mode === "resize-end") {
    nextEnd = shiftBusinessDate(dragState.endDate, deltaDays);
    if (nextEnd < nextStart) nextEnd = nextStart;
  }

  record.start = nextStart;
  record.end = nextEnd;
  record.days = dateRangeBusiness(record.start, record.end);
  state.selectedTaskId = record.id;
  syncTaskBarVisual(record, element, dragState.mode);
}

function finishTaskPointerDrag(element) {
  if (!dragState.active) return;
  const shouldOpenModal = dragState.mode === "move" && !dragState.moved;
  const changedRecord = dragState.moved ? getRecord(dragState.recordId) : null;
  dragState.active = false;
  dragState.mode = null;
  const recordId = dragState.recordId;
  dragState.recordId = null;
  dragState.pointerId = null;
  dragState.moved = false;
  clearDragVisuals(element);
  if (changedRecord) {
    changedRecord.updatedByName = state.user.name;
    changedRecord.updatedByClient = state.user.clientId;
  }
  renderAll();
  if (changedRecord) persistTaskToWorkspace(changedRecord);
  if (shouldOpenModal && recordId) openTaskModal(recordId);
}

function startTaskPointerDrag(event, mode) {
  const bar = event.currentTarget.closest(".task-bar");
  if (!bar) return;
  const record = getRecord(bar.dataset.taskId);
  if (!record) return;
  event.preventDefault();
  state.selectedTaskId = record.id;
  dragState.active = true;
  dragState.mode = mode;
  dragState.recordId = record.id;
  dragState.startX = event.clientX;
  dragState.startY = event.clientY;
  dragState.startDate = record.start;
  dragState.endDate = record.end;
  dragState.originalAssignee = record.assignee;
  dragState.pointerId = event.pointerId;
  dragState.moved = false;
  bar.setPointerCapture(event.pointerId);
  syncTaskBarVisual(record, bar, mode);

  const onMove = moveEvent => {
    if (!dragState.active || moveEvent.pointerId !== dragState.pointerId) return;
    handleTaskPointerMove(moveEvent, bar);
  };
  const onEnd = endEvent => {
    if (endEvent.pointerId !== dragState.pointerId) return;
    bar.removeEventListener("pointermove", onMove);
    bar.removeEventListener("pointerup", onEnd);
    bar.removeEventListener("pointercancel", onEnd);
    finishTaskPointerDrag(bar);
  };

  bar.addEventListener("pointermove", onMove);
  bar.addEventListener("pointerup", onEnd);
  bar.addEventListener("pointercancel", onEnd);
}

function bindTaskBarEvents() {
  board.querySelectorAll(".task-bar").forEach(node => {
    node.addEventListener("pointerdown", event => {
      const handle = event.target.closest(".resize-handle");
      if (handle) {
        startTaskPointerDrag(event, handle.dataset.handle === "start" ? "resize-start" : "resize-end");
        return;
      }
      startTaskPointerDrag(event, "move");
    });
  });
}

function renderBoard() {
  const filtered = getFilteredRecords();
  const dateCols = getDateCols();
  const colorMap = getColorMap();

  updateStats(filtered);
  renderChips();

  if (filtered.length === 0 || dateCols.length === 0) {
    board.innerHTML = state.records.length === 0
      ? '<div class="empty-state">Choose or create a workspace to start planning.</div>'
      : '<div class="empty-state">No rows match the current filters.</div>';
    return;
  }

  const rows = buildRows(filtered, dateCols);
  const firstColTitle = state.rowMode === "assignee" ? "Assignee" : "Epic";

  let html = '<div class="board-inner">';
  html += `<div class="header-row"><div class="header-left">${firstColTitle}</div><div class="header-days">`;
  dateCols.forEach(dateStr => {
    html += `<div class="day-head">${dayHeader(dateStr)}</div>`;
  });
  html += "</div></div>";

  rows.forEach(row => {
    const rowHeight = Math.max(74, 8 + (row.laneCount * 58));
    const timelineWidth = dateCols.length * 150;
    html += '<div class="board-row">';
    html += `<div class="row-label"><div class="row-title">${escapeHtml(row.name)}</div><div class="row-sub">${escapeHtml(row.sub)}</div></div>`;
    html += `<div class="timeline" style="width:${timelineWidth}px; min-width:${timelineWidth}px; height:${rowHeight}px;">`;
    if (row.placedTasks.length === 0) {
      html += '<div class="empty-row">·</div>';
    } else {
      row.placedTasks.forEach(task => { html += taskBarHtml(task, colorMap); });
    }
    html += "</div></div>";
  });

  html += "</div>";
  board.innerHTML = html;
  bindTaskBarEvents();
}

function renderSummary() {
  const filtered = getFilteredRecords();
  updateStats(filtered);
  renderChips();

  if (filtered.length === 0) {
    summary.innerHTML = state.records.length === 0
      ? '<div class="empty-state">Choose or create a workspace to see the summary.</div>'
      : '<div class="empty-state">No rows match the current filters.</div>';
    return;
  }

  const epicMap = new Map();
  filtered.forEach(record => {
    if (!epicMap.has(record.epic)) {
      epicMap.set(record.epic, {
        epic: record.epic,
        count: 0,
        departments: new Set(),
        assignees: new Set()
      });
    }
    const row = epicMap.get(record.epic);
    row.count += 1;
    row.departments.add(record.department);
    row.assignees.add(record.assignee);
  });

  const rows = Array.from(epicMap.values()).sort((a, b) => b.count - a.count || a.epic.localeCompare(b.epic));
  const maxCount = Math.max(...rows.map(row => row.count), 1);

  let html = '<div class="summary-head"><div>Epic</div><div>Departments</div><div>Assignees</div><div>Tasks</div></div>';
  rows.forEach(row => {
    const width = Math.max(12, Math.round((row.count / maxCount) * 100));
    html += '<div class="summary-row">';
    html += `<div><div class="summary-epic">${escapeHtml(row.epic)}</div><div class="summary-meta">${row.count} tasks in current filter</div></div>`;
    html += `<div class="summary-meta">${escapeHtml(Array.from(row.departments).sort().join(", "))}</div>`;
    html += `<div class="summary-meta">${escapeHtml(Array.from(row.assignees).sort().join(", "))}</div>`;
    html += `<div><div class="summary-count">${row.count}</div><div class="summary-bar"><div class="summary-bar-fill" style="width:${width}%;"></div></div></div>`;
    html += "</div>";
  });

  summary.innerHTML = html;
}

function renderStatus() {
  const dirtyCount = getDirtyRecords().length;
  const selected = getRecord(state.selectedTaskId);
  changeStatus.classList.toggle("dirty", dirtyCount > 0);

  if (dirtyCount === 0) {
    changeStatus.textContent = selected
      ? "Drag the selected task to move it. Drag the left or right edge to resize it. Click a task to open the edit overlay."
      : "Drag a task bar to move dates. Drag its edges to resize. Click a task to open the edit overlay.";
    return;
  }

  changeStatus.textContent = `${dirtyCount} task${dirtyCount === 1 ? "" : "s"} changed from the original CSV. Download to export the edited sheet or use Undo CSV changes to revert everything.`;
}

function updateTask(recordId, changes) {
  const record = getRecord(recordId);
  if (!record) return;

  if (typeof changes.assignee === "string") record.assignee = normalizeLabel(changes.assignee);
  if (typeof changes.department === "string") record.department = normalizeLabel(changes.department);
  if (changes.start) record.start = normalizeBusinessDate(changes.start, "forward") || record.start;
  if (changes.end) record.end = normalizeBusinessDate(changes.end, "backward") || record.end;

  if (changes.syncEndIfNeeded && record.end < record.start) record.end = record.start;

  if (changes.workingDays) {
    const start = normalizeBusinessDate(record.start, "forward") || record.start;
    record.start = start;
    record.end = addBusinessDays(start, changes.workingDays);
  } else if (record.end < record.start) {
    record.end = record.start;
  }

  record.days = dateRangeBusiness(record.start, record.end);
  record.updatedByName = state.user.name;
  record.updatedByClient = state.user.clientId;
  renderAll();
  if (!taskModal.hidden) renderTaskModal();
  persistTaskToWorkspace(record);
}

function renderTaskModal() {
  const record = getRecord(state.selectedTaskId);
  if (!record) {
    taskModalContent.innerHTML = '<div class="editor-empty"><strong>Task editor</strong><br />Select a task on the board to edit it.</div>';
    return;
  }

  const original = getOriginalRecord(record.id);
  const dirty = isTaskDirty(record);
  taskModalContent.innerHTML = `
    <div class="section-title">Selected task</div>
    <div class="editor-head">
      <div>
        <div class="editor-title">${escapeHtml(record.task)}</div>
        <div class="editor-meta">${escapeHtml(record.department)} • ${escapeHtml(record.epic)}</div>
      </div>
      <div class="editor-badge ${dirty ? "dirty" : ""}">${dirty ? "Edited" : "Original"}</div>
    </div>
    <div class="input-grid single">
      <div class="editor-field">
        <label for="editAssignee">Assignee</label>
        <select id="editAssignee" class="editor-input">
          ${getAssignees().map(name => `<option value="${escapeHtml(name)}" ${name === record.assignee ? "selected" : ""}>${escapeHtml(name)}</option>`).join("")}
          ${getAssignees().includes(record.assignee) ? "" : `<option value="${escapeHtml(record.assignee)}" selected>${escapeHtml(record.assignee)}</option>`}
        </select>
      </div>
    </div>
    <div class="editor-help">Current: ${escapeHtml(record.assignee)} • ${escapeHtml(record.start)} to ${escapeHtml(record.end)} (${record.days.length} working day${record.days.length === 1 ? "" : "s"})</div>
    <div class="editor-help">Original: ${escapeHtml(original.assignee)} • ${escapeHtml(original.start)} to ${escapeHtml(original.end)} (${original.days.length} working day${original.days.length === 1 ? "" : "s"})</div>
    <div class="input-grid">
      <div class="editor-field">
        <label for="editStart">Start date</label>
        <input id="editStart" class="editor-input" type="date" value="${escapeHtml(record.start)}" />
      </div>
      <div class="editor-field">
        <label for="editEnd">End date</label>
        <input id="editEnd" class="editor-input" type="date" value="${escapeHtml(record.end)}" />
      </div>
    </div>
    <div class="input-grid">
      <div class="editor-field">
        <label for="editDays">Working days</label>
        <input id="editDays" class="editor-input" type="number" min="1" step="1" value="${record.days.length}" />
      </div>
      <div class="editor-field">
        <label>&nbsp;</label>
        <div class="editor-help">Drag on the board to update dates and duration. Download the CSV when you are happy with the changes.</div>
      </div>
    </div>
    <div class="editor-actions">
      <button class="action light" id="closeTaskModal">Close</button>
    </div>
  `;

  document.getElementById("editAssignee").addEventListener("change", event => {
    updateTask(record.id, { assignee: event.target.value });
  });

  document.getElementById("editStart").addEventListener("change", event => {
    updateTask(record.id, { start: event.target.value, syncEndIfNeeded: true });
  });

  document.getElementById("editEnd").addEventListener("change", event => {
    updateTask(record.id, { end: event.target.value });
  });

  document.getElementById("editDays").addEventListener("change", event => {
    updateTask(record.id, { workingDays: Math.max(1, Number(event.target.value) || 1) });
  });

  document.getElementById("closeTaskModal").addEventListener("click", closeTaskModal);
}

function renderView() {
  if (state.displayMode === "summary") {
    board.hidden = true;
    summary.hidden = false;
    renderSummary();
    return;
  }

  summary.hidden = true;
  board.hidden = false;
  renderBoard();
}

function renderAll() {
  renderOptions();
  renderView();
  renderStatus();
  if (!taskModal.hidden) renderTaskModal();
}

function resetFilters() {
  state.displayMode = "plan";
  state.rowMode = "assignee";
  state.selectedDepartments.clear();
  state.selectedEpics.clear();
  state.selectedAssignees.clear();
  state.hideEmpty = true;
  deptSearch.value = "";
  epicSearch.value = "";
  assigneeSearch.value = "";
  document.getElementById("toggleHideEmpty").checked = true;
  document.getElementById("displayPlan").classList.add("active");
  document.getElementById("displaySummary").classList.remove("active");
  document.getElementById("viewAssignee").classList.add("active");
  document.getElementById("viewEpic").classList.remove("active");
  renderAll();
}

function resetChanges() {
  state.records = cloneRecords(state.originalRecords);
  if (state.selectedTaskId && !getRecord(state.selectedTaskId)) state.selectedTaskId = null;
  closeTaskModal();
  renderAll();
}

function setMode(mode) {
  state.rowMode = mode;
  document.getElementById("viewAssignee").classList.toggle("active", mode === "assignee");
  document.getElementById("viewEpic").classList.toggle("active", mode === "epic");
  renderAll();
}

function setDisplayMode(mode) {
  state.displayMode = mode;
  document.getElementById("displayPlan").classList.toggle("active", mode === "plan");
  document.getElementById("displaySummary").classList.toggle("active", mode === "summary");
  renderAll();
}

function buildExportRows() {
  if (!state.source) return null;
  const rows = cloneRows(state.source.rows);

  state.records.forEach(record => {
    const row = rows[record.rowIndex] || [];
    const { colIndex, dateColumnIndices, dateIndexMap } = state.source;

    row[colIndex["Department"]] = record.department;
    row[colIndex["Assignee\n(use email)"]] = record.assignee;
    row[colIndex["Start"]] = formatTemplateDate(record.start);
    row[colIndex["End"]] = formatTemplateDate(record.end);
    row[colIndex["Day"]] = String(record.days.length);

    dateColumnIndices.forEach(index => { row[index] = ""; });
    record.days.forEach(day => {
      const columnIndex = dateIndexMap.get(day);
      if (columnIndex !== undefined) row[columnIndex] = "1";
    });

    rows[record.rowIndex] = row;
  });

  return rows;
}

function downloadCsv() {
  const rows = buildExportRows();
  if (!rows) {
    alert("No CSV source is loaded yet.");
    return;
  }

  const csvText = rows.map(row => row.map(csvEscape).join(",")).join("\r\n");
  const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  const suffix = getDirtyRecords().length > 0 ? " - edited" : "";
  anchor.download = state.fileName.replace(/\.csv$/i, "") + suffix + ".csv";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function applyParsedData(parsed) {
  state.source = parsed.source;
  state.fileName = parsed.fileName;
  state.records = cloneRecords(parsed.records);
  state.originalRecords = cloneRecords(parsed.records);
  state.selectedTaskId = null;
  resetFilters();
}

async function loadDefaultCsv() {
  state.source = null;
  state.fileName = "planning-viewer.csv";
  state.records = [];
  state.originalRecords = [];
  state.selectedTaskId = null;
  renderAll();
}

function bindEvents() {
  deptSearch.addEventListener("input", renderOptions);
  epicSearch.addEventListener("input", renderOptions);
  assigneeSearch.addEventListener("input", renderOptions);

  document.getElementById("deptAll").addEventListener("click", () => { getDepartments().forEach(value => state.selectedDepartments.add(value)); renderAll(); });
  document.getElementById("deptClear").addEventListener("click", () => { state.selectedDepartments.clear(); renderAll(); });
  document.getElementById("epicAll").addEventListener("click", () => { getEpics().forEach(value => state.selectedEpics.add(value)); renderAll(); });
  document.getElementById("epicClear").addEventListener("click", () => { state.selectedEpics.clear(); renderAll(); });
  document.getElementById("assigneeAll").addEventListener("click", () => { getAssignees().forEach(value => state.selectedAssignees.add(value)); renderAll(); });
  document.getElementById("assigneeClear").addEventListener("click", () => { state.selectedAssignees.clear(); renderAll(); });

  document.getElementById("toggleHideEmpty").addEventListener("change", event => {
    state.hideEmpty = event.target.checked;
    renderAll();
  });

  document.getElementById("displayPlan").addEventListener("click", () => setDisplayMode("plan"));
  document.getElementById("displaySummary").addEventListener("click", () => setDisplayMode("summary"));
  document.getElementById("viewAssignee").addEventListener("click", () => setMode("assignee"));
  document.getElementById("viewEpic").addEventListener("click", () => setMode("epic"));
  document.getElementById("resetFiltersBtn").addEventListener("click", resetFilters);
  document.getElementById("undoChangesBtn").addEventListener("click", resetChanges);
  document.getElementById("downloadBtn").addEventListener("click", downloadCsv);
  document.getElementById("switchWorkspaceBtn").addEventListener("click", async () => {
    await fetchAvailableWorkspaces();
    openWorkspaceModal();
  });
  taskModalBackdrop.addEventListener("click", closeTaskModal);
  document.getElementById("editIdentityBtn").addEventListener("click", () => openIdentityModal(true));
  document.getElementById("saveIdentityBtn").addEventListener("click", saveIdentityFromInput);
  document.getElementById("cancelIdentityBtn").addEventListener("click", closeIdentityModal);
  identityModalBackdrop.addEventListener("click", closeIdentityModal);
  identityNameInput.addEventListener("keydown", event => {
    if (event.key === "Enter") saveIdentityFromInput();
  });
  workspaceModalBackdrop.addEventListener("click", () => {
    if (state.sync.workspaceId) closeWorkspaceModal();
  });
  document.getElementById("joinWorkspaceBtn").addEventListener("click", submitJoinWorkspace);
  document.getElementById("createWorkspaceBtn").addEventListener("click", submitCreateWorkspace);
  workspaceUserNameInput.addEventListener("keydown", event => {
    if (event.key === "Enter") submitJoinWorkspace();
  });
  newWorkspaceIdInput.addEventListener("keydown", event => {
    if (event.key === "Enter") submitCreateWorkspace();
  });
}

async function init() {
  hydrateUserIdentity();
  hydrateWorkspaceSelection();
  bindEvents();
  await connectSync();
  await fetchAvailableWorkspaces();
  await loadDefaultCsv();
  if (state.user.name && state.sync.workspaceId) {
    const known = state.availableWorkspaces.some(workspace => workspace.id === state.sync.workspaceId);
    if (known) {
      await activateWorkspace(state.sync.workspaceId, { loadExisting: true, notifyChanges: false });
      return;
    }
  }
  openWorkspaceModal();
}

init();
