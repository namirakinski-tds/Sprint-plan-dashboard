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
  fileName: "planning-viewer.csv"
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
  pointerId: null
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
const csvFile = document.getElementById("csvFile");
const changeStatus = document.getElementById("changeStatus");
const taskEditor = document.getElementById("taskEditor");

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

function parseCsvText(text, fileName) {
  const rows = parseCsvRows(text);
  if (rows.length && rows[rows.length - 1].every(cell => cell === "")) rows.pop();
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

    if (!department || !epic || !task || !assignee || !start || !end) continue;

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
  const original = getOriginalRecord(record.id);
  if (!original) return false;
  return record.assignee !== original.assignee || record.start !== original.start || record.end !== original.end;
}

function getDirtyRecords() {
  return state.records.filter(isTaskDirty);
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

function handleTaskPointerMove(event, element) {
  if (!dragState.active) return;
  const record = getRecord(dragState.recordId);
  if (!record) return;

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
  dragState.active = false;
  dragState.mode = null;
  dragState.recordId = null;
  dragState.pointerId = null;
  clearDragVisuals(element);
  renderAll();
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
    board.innerHTML = '<div class="empty-state">No rows match the current filters.</div>';
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
    summary.innerHTML = '<div class="empty-state">No rows match the current filters.</div>';
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
      ? "Drag the selected task to move it. Drag the left or right edge to resize it. Change assignee from the board-side task panel."
      : "Drag a task bar to move dates. Drag its edges to resize. Change assignee from the board-side task panel.";
    return;
  }

  changeStatus.textContent = `${dirtyCount} task${dirtyCount === 1 ? "" : "s"} changed from the original CSV. Download to export the edited sheet or use Undo CSV changes to revert everything.`;
}

function renderEditor() {
  const record = getRecord(state.selectedTaskId);
  if (!record) {
    taskEditor.innerHTML = '<div class="editor-empty"><strong>Board editing</strong><br />Move tasks directly on the timeline. Drag the bar to shift dates, drag the left or right edge to change duration, and select a task to change its assignee here.</div>';
    return;
  }

  const original = getOriginalRecord(record.id);
  const dirty = isTaskDirty(record);
  taskEditor.innerHTML = `
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
    <div class="editor-help">Drag on the board to update dates and duration. Use this selector to change assignee. Download the CSV when you are happy with the changes.</div>
  `;

  document.getElementById("editAssignee").addEventListener("change", event => {
    const nextAssignee = event.target.value;
    record.assignee = nextAssignee;
    renderAll();
  });
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
  renderEditor();
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
  const response = await fetch("./default-template.csv");
  if (!response.ok) throw new Error("Default CSV template not found");
  const text = await response.text();
  const parsed = parseCsvText(text, "default-template.csv");
  if (!parsed) throw new Error("Default CSV template could not be parsed");
  applyParsedData(parsed);
}

function bindEvents() {
  csvFile.addEventListener("change", async event => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const text = await file.text();
    const parsed = parseCsvText(text, file.name);
    if (!parsed) {
      alert("Could not parse this CSV. Please use the same sprint rehearsal export format.");
      return;
    }
    applyParsedData(parsed);
  });

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
}

async function init() {
  bindEvents();
  try {
    await loadDefaultCsv();
  } catch (error) {
    board.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
    changeStatus.textContent = "Upload your sprint CSV to start editing.";
  }
}

init();
