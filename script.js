/* =========================
   GLOBAL STATE
========================= */

let workbook = null;
let worksheet = null;
let allRows = [];
let headers = [];

let selectedShift = null;
let selectedDateLabel = "";

let taskStatuses = {};        // key: taskId -> good|monitor|escalate
let escalationNotes = {};     // key: taskId -> { issue, rootCause, remarks }
let currentEscalationTask = null;

/* =========================
   LOAD EXCEL FROM GITHUB
========================= */

document.addEventListener("DOMContentLoaded", loadExcelTemplate);

async function loadExcelTemplate() {
  try {
    const response = await fetch("./data/Weekday_Monitoring Template.xlsx");
    if (!response.ok) throw new Error("Fetch failed");

    const arrayBuffer = await response.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });

    workbook = wb;
    worksheet = wb.Sheets[wb.SheetNames[0]];

    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    headers = json[0];
    allRows = json.slice(1);

    buildShiftFilters();
    buildExportButtons();
  } catch (err) {
    console.error(err);
    document.getElementById("output").innerHTML =
      "<p style='color:red;text-align:center;'>Failed to load monitoring template</p>";
  }
}

/* =========================
   SHIFT FILTER
========================= */

function buildShiftFilters() {
  const shiftCol = headers.indexOf("Shift");
  const shifts = [...new Set(allRows.map(r => r[shiftCol]).filter(Boolean))];

  const container = document.getElementById("shiftFilter");
  container.innerHTML = "";

  shifts.forEach(shift => {
    const btn = document.createElement("button");
    btn.textContent = shift;
    btn.onclick = () => selectShift(shift, btn);
    container.appendChild(btn);
  });
}

function selectShift(shift, btn) {
  selectedShift = shift;

  document.querySelectorAll("#shiftFilter button").forEach(b =>
    b.classList.remove("active")
  );
  btn.classList.add("active");

  renderTasks();
}

/* =========================
   TASK RENDER
========================= */

function renderTasks() {
  const shiftCol = headers.indexOf("Shift");
  const taskCol = headers.indexOf("Task");

  const container = document.getElementById("output");
  container.innerHTML = "";

  allRows.forEach((row, idx) => {
    if (row[shiftCol] !== selectedShift) return;

    const taskId = `${selectedShift}_${idx}`;
    const status = taskStatuses[taskId] || "";

    const div = document.createElement("div");
    div.className = "task-row";

    div.innerHTML = `
      <span class="task-name">${row[taskCol]}</span>
      <button class="good ${status === "good" ? "active" : ""}" onclick="setStatus('${taskId}','good')">Good</button>
      <button class="monitor ${status === "monitor" ? "active" : ""}" onclick="setStatus('${taskId}','monitor')">Monitor</button>
      <button class="escalate ${status === "escalate" ? "active" : ""}" onclick="openEscalation('${taskId}')">Escalate</button>
    `;

    container.appendChild(div);
  });
}

function setStatus(taskId, status) {
  taskStatuses[taskId] = status;
  renderTasks();
}

/* =========================
   ESCALATION MODAL
========================= */

function openEscalation(taskId) {
  currentEscalationTask = taskId;
  taskStatuses[taskId] = "escalate";

  const notes = escalationNotes[taskId] || {};
  document.getElementById("noteIssue").value = notes.issue || "";
  document.getElementById("noteRootCause").value = notes.rootCause || "";
  document.getElementById("noteRemarks").value = notes.remarks || "";

  document.getElementById("escalationModal").style.display = "block";
  renderTasks();
}

function closeModal() {
  document.getElementById("escalationModal").style.display = "none";
}

function saveEscalationNotes() {
  escalationNotes[currentEscalationTask] = {
    issue: document.getElementById("noteIssue").value,
    rootCause: document.getElementById("noteRootCause").value,
    remarks: document.getElementById("noteRemarks").value
  };
  closeModal();
}

/* =========================
   EXPORT BUTTONS
========================= */

function buildExportButtons() {
  document.getElementById("exportButtons").innerHTML = `
    <button onclick="exportShift()">Export Shift</button>
    <button onclick="exportEscalations()">Export Escalations</button>
    <button onclick="resetAll()">Reset Status</button>
  `;
}

/* =========================
   EXPORT SHIFT
========================= */

function exportShift() {
  if (!selectedShift) return alert("Select a shift first");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Monitoring");

  ws.addRow(headers);

  allRows.forEach((row, idx) => {
    if (row[headers.indexOf("Shift")] !== selectedShift) return;

    const taskId = `${selectedShift}_${idx}`;
    const status = taskStatuses[taskId] || "";

    const r = ws.addRow(row);
    const cell = r.getCell(headers.length);

    if (status === "good") cell.fill = fill("00FF00");
    if (status === "monitor") cell.fill = fill("FFFF00");
    if (status === "escalate") cell.fill = fill("FF0000");
  });

  downloadExcel(wb, buildFileName());
}

/* =========================
   EXPORT ESCALATIONS
========================= */

function exportEscalations() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Escalations");

  ws.addRow(["Shift", "Task", "Issue", "Root Cause", "Remarks"]);

  allRows.forEach((row, idx) => {
    const taskId = `${row[headers.indexOf("Shift")]}_${idx}`;
    if (taskStatuses[taskId] !== "escalate") return;

    const notes = escalationNotes[taskId] || {};
    ws.addRow([
      row[headers.indexOf("Shift")],
      row[headers.indexOf("Task")],
      notes.issue || "",
      notes.rootCause || "",
      notes.remarks || ""
    ]);
  });

  downloadExcel(wb, `Escalations_${buildFileName()}`);
}

/* =========================
   HELPERS
========================= */

function fill(color) {
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color }
  };
}

function buildFileName() {
  return `MTDay_Shift_${selectedShift.replace(/\s+/g, "_")}.xlsx`;
}

function downloadExcel(wb, filename) {
  wb.xlsx.writeBuffer().then(buf => {
    const blob = new Blob([buf]);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
  });
}

function resetAll() {
  if (!confirm("Reset all statuses?")) return;
  taskStatuses = {};
  escalationNotes = {};
  renderTasks();
}
