/*********************************************************
 * GLOBAL STATE
 *********************************************************/
let allDataRows = [];
let headers = [];
let filteredRows = [];
let currentIndex = 0;
let selectedShift = "All";

let taskStatuses = {};
let escalationNotes = {};
let currentEscalationKey = "";

/*********************************************************
 * CURRENT TIMES
 *********************************************************/
function updateCurrentTimes() {
  const now = new Date();
  const manila = now.toLocaleString("en-US", { timeZone: "Asia/Manila" });
  const mt = now.toLocaleString("en-US", { timeZone: "America/Denver" });

  document.getElementById("currentTimes").innerHTML =
    `<b>Manila:</b> ${manila} | <b>MT:</b> ${mt}`;
}
setInterval(updateCurrentTimes, 1000);
updateCurrentTimes();

/*********************************************************
 * DATE (MT SAFE)
 *********************************************************/
function getMTFileDate() {
  const p = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "numeric"
  }).formatToParts(new Date());

  const g = t => p.find(x => x.type === t)?.value;
  return `${g("month")}_${g("day")}_${g("year")}`;
}

/*********************************************************
 * EXCEL TIME
 *********************************************************/
function excelTimeToString(value) {
  if (typeof value === "number") {
    const sec = Math.round(value * 86400);
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    return new Date(0, 0, 0, h, m).toLocaleTimeString("en-US", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true
    });
  }
  return value || "";
}

/*********************************************************
 * LOAD TEMPLATE
 *********************************************************/
async function loadMonitoringTemplate() {
  try {
    const res = await fetch("./data/Weekday_Monitoring Template.xlsx");
    if (!res.ok) throw new Error();

    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    headers = rows[0];
    allDataRows = rows.slice(1).map((r, i) => ({ __rowId: i, data: r }));
    filteredRows = allDataRows;

    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
    renderCurrentCard();
  } catch {
    output.innerHTML = "<b style='color:red'>Failed to load monitoring template</b>";
  }
}
loadMonitoringTemplate();

/*********************************************************
 * NAVIGATION
 *********************************************************/
function buildNavigation() {
  navigation.innerHTML = `
    <button onclick="prevCard()">Previous</button>
    <span id="cardCounter"></span>
    <button onclick="nextCard()">Next</button>
  `;
}
function prevCard() {
  if (currentIndex > 0) currentIndex--, renderCurrentCard();
}
function nextCard() {
  if (currentIndex < filteredRows.length - 1) currentIndex++, renderCurrentCard();
}

/*********************************************************
 * RENDER CARD
 *********************************************************/
function renderCurrentCard() {
  const obj = filteredRows[currentIndex];
  if (!obj) return;

  const row = obj.data;
  const id = obj.__rowId;

  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${row[0]} — ${excelTimeToString(row[1])} Manila | ${excelTimeToString(row[2])} MT
      </div>
  `;

  for (let i = 3; i < headers.length; i++) {
    if (!row[i]) continue;
    const app = headers[i];
    const key = `${row[0]}-${app}-${id}`;
    const s = taskStatuses[key] || "";

    html += `
      <div class="task ${s === "escalate" ? "escalated" : ""}">
        <strong>${app}</strong>
        <span>${row[i]}</span>
        <div class="task-buttons">
          <button class="${s === "good" ? "active-good" : ""}"
            onclick="setStatus('${key}','good')">Good</button>
          <button class="${s === "monitor" ? "active-monitor" : ""}"
            onclick="setStatus('${key}','monitor')">Monitor</button>
          <button class="${s === "escalate" ? "active-escalate" : ""}"
            onclick="setStatus('${key}','escalate',true)">Escalate</button>
        </div>
      </div>
    `;
  }

  output.innerHTML = html + "</div>";
  cardCounter.innerText = `Shift ${currentIndex + 1} of ${filteredRows.length}`;
}

/*********************************************************
 * STATUS + MODAL
 *********************************************************/
function setStatus(key, status, open = false) {
  taskStatuses[key] = status;

  if (status === "escalate" && open) {
    currentEscalationKey = key;
    const n = escalationNotes[key] || {};
    noteIssue.value = n.issue || "";
    noteRootCause.value = n.rootCause || "";
    noteRemarks.value = n.remarks || "";
    escalationModal.style.display = "block";
  }
  renderCurrentCard();
}

function closeModal() {
  escalationModal.style.display = "none";
}
function saveEscalationNotes() {
  escalationNotes[currentEscalationKey] = {
    issue: noteIssue.value,
    rootCause: noteRootCause.value,
    remarks: noteRemarks.value
  };
  closeModal();
}

/*********************************************************
 * SHIFT FILTER
 *********************************************************/
function buildShiftFilter() {
  const shifts = [...new Set(allDataRows.map(r => r.data[0]))];
  shiftFilter.innerHTML =
    `<button class="${selectedShift === "All" ? "active-filter" : ""}"
      onclick="filterByShift('All')">All Shift</button>` +
    shifts.map(s =>
      `<button class="${selectedShift === s ? "active-filter" : ""}"
        onclick="filterByShift('${s}')">${s}</button>`
    ).join("");
}
function filterByShift(s) {
  selectedShift = s;
  filteredRows = s === "All" ? allDataRows : allDataRows.filter(r => r.data[0] === s);
  currentIndex = 0;
  buildShiftFilter();
  renderCurrentCard();
}

/*********************************************************
 * EXPORT BUTTONS
 *********************************************************/
function buildExportButtons() {
  exportButtons.innerHTML = `
    <button onclick="exportShiftData()">Export Shift</button>
    <button onclick="exportEscalatedTasks()">Export Escalation</button>
    <button onclick="resetAllStatuses()">Reset Status</button>
  `;
}

/*********************************************************
 * EXPORT SHIFT (FIXED)
 *********************************************************/
async function exportShiftData() {
  if (!window.ExcelJS) return alert("ExcelJS not loaded");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Shift");

  ws.addRow(headers);

  filteredRows.forEach((obj, r) => {
    const row = [...obj.data];
    row[1] = excelTimeToString(row[1]);
    row[2] = excelTimeToString(row[2]);
    ws.addRow(row);

    for (let c = 3; c < headers.length; c++) {
      const key = `${row[0]}-${headers[c]}-${obj.__rowId}`;
      const cell = ws.getRow(r + 2).getCell(c + 1);
      const s = taskStatuses[key];

      if (s === "good")
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
      if (s === "monitor")
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      if (s === "escalate")
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
    }
  });

  const name = `${getMTFileDate()}_${selectedShift}.xlsx`;
  const buf = await wb.xlsx.writeBuffer();
  download(buf, name);
}

/*********************************************************
 * EXPORT ESCALATION
 *********************************************************/
async function exportEscalatedTasks() {
  if (!window.ExcelJS) return alert("ExcelJS not loaded");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Escalations");

  ws.addRow(["Shift", "App", "Task", "Issue", "Root Cause", "Remarks"]);

  Object.keys(taskStatuses).forEach(k => {
    if (taskStatuses[k] !== "escalate") return;
    const [shift, app] = k.split("-");
    const note = escalationNotes[k] || {};
    ws.addRow([shift, app, "", note.issue, note.rootCause, note.remarks]);
  });

  const buf = await wb.xlsx.writeBuffer();
  download(buf, `${getMTFileDate()}_Escalations.xlsx`);
}

/*********************************************************
 * DOWNLOAD
 *********************************************************/
function download(buf, name) {
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
}

/*********************************************************
 * RESET
 *********************************************************/
function resetAllStatuses() {
  if (!confirm("Reset all statuses and notes?")) return;
  taskStatuses = {};
  escalationNotes = {};
  renderCurrentCard();
}
