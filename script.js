/*********************************************************
 * GLOBAL STATE
 *********************************************************/
let headers = [];
let allDataRows = [];
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

  const manila = now.toLocaleString("en-US", {
    timeZone: "Asia/Manila",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  });

  const mountain = now.toLocaleString("en-US", {
    timeZone: "America/Denver",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  });

  document.getElementById("currentTimes").innerHTML = `
    <strong>Manila:</strong> ${manila} |
    <strong>MT:</strong> ${mountain}
  `;
}
setInterval(updateCurrentTimes, 1000);
updateCurrentTimes();

/*********************************************************
 * EXCEL HELPERS
 *********************************************************/
function excelTimeToString(value) {
  if (typeof value === "number") {
    const seconds = Math.round(value * 86400);
    const h = Math.floor(seconds / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    return new Date(0, 0, 0, h, m).toLocaleTimeString("en-US", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true
    });
  }
  return value || "";
}

/*********************************************************
 * LOAD EXCEL FROM GITHUB
 *********************************************************/
async function loadMonitoringTemplate() {
  try {
    const response = await fetch("./data/Weekday_Monitoring Template.xlsx");
    if (!response.ok) throw new Error("Fetch failed");

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    headers = rows[0];
    allDataRows = rows.slice(1).map((row, idx) => ({
      __rowId: idx,
      data: row
    }));

    filteredRows = allDataRows;
    currentIndex = 0;
    selectedShift = "All";

    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
    renderCurrent();

  } catch (err) {
    console.error(err);
    document.getElementById("output").innerHTML =
      "<b style='color:red'>Failed to load monitoring template</b>";
  }
}

loadMonitoringTemplate();

/*********************************************************
 * NAVIGATION
 *********************************************************/
function buildNavigation() {
  document.getElementById("navigation").innerHTML = `
    <button onclick="prevCard()">Previous</button>
    <span id="cardCounter"></span>
    <button onclick="nextCard()">Next</button>
  `;
  updateCounter();
}

function updateCounter() {
  const counter = document.getElementById("cardCounter");
  if (!filteredRows.length) {
    counter.innerText = "No data";
    return;
  }
  counter.innerText = `Shift ${currentIndex + 1} of ${filteredRows.length}`;
}

function prevCard() {
  if (currentIndex > 0) {
    currentIndex--;
    renderCurrent();
  }
}

function nextCard() {
  if (currentIndex < filteredRows.length - 1) {
    currentIndex++;
    renderCurrent();
  }
}

/*********************************************************
 * RENDER CARD
 *********************************************************/
function renderCurrent() {
  const rowObj = filteredRows[currentIndex];
  if (!rowObj) return;

  const row = rowObj.data;
  const rowId = rowObj.__rowId;

  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${row[0]} — ${excelTimeToString(row[1])} | ${excelTimeToString(row[2])}
      </div>
  `;

  for (let i = 3; i < headers.length; i++) {
    if (!row[i]) continue;

    const key = `${row[0]}-${headers[i]}-${rowId}`;
    const status = taskStatuses[key] || "";

    html += `
      <div class="task ${status === "escalate" ? "escalated" : ""}">
        <strong>${headers[i]}</strong>
        <span>${row[i]}</span>
        <div class="task-buttons">
          <button class="${status === "good" ? "active-good" : ""}"
            onclick="setStatus('${key}','good')">Good</button>
          <button class="${status === "monitor" ? "active-monitor" : ""}"
            onclick="setStatus('${key}','monitor')">Monitor</button>
          <button class="${status === "escalate" ? "active-escalate" : ""}"
            onclick="setStatus('${key}','escalate',true)">Escalate</button>
        </div>
      </div>
    `;
  }

  document.getElementById("output").innerHTML = html + "</div>";
  updateCounter();
}

/*********************************************************
 * STATUS + MODAL
 *********************************************************/
function setStatus(key, status, openModal = false) {
  taskStatuses[key] = status;

  if (status === "escalate" && openModal && !escalationNotes[key]) {
    currentEscalationKey = key;
    document.getElementById("escalationModal").style.display = "block";
  }
  renderCurrent();
}

function closeModal() {
  document.getElementById("escalationModal").style.display = "none";
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
  let html = `<button onclick="filterByShift('All')">All</button>`;
  shifts.forEach(s => {
    html += `<button onclick="filterByShift('${s}')">${s}</button>`;
  });
  document.getElementById("shiftFilter").innerHTML = html;
}

function filterByShift(shift) {
  selectedShift = shift;
  filteredRows =
    shift === "All"
      ? allDataRows
      : allDataRows.filter(r => r.data[0] === shift);
  currentIndex = 0;
  renderCurrent();
}

/*********************************************************
 * EXPORT + RESET
 *********************************************************/
function buildExportButtons() {
  document.getElementById("exportButtons").innerHTML = `
    <button onclick="exportShiftData(selectedShift)">Export Shift</button>
    <button onclick="exportEscalatedTasks(selectedShift)">Export Escalations</button>
    <button onclick="resetAllStatuses()">Reset Status</button>
  `;
}

function resetAllStatuses() {
  if (!confirm("Reset all task statuses?")) return;
  taskStatuses = {};
  escalationNotes = {};
  renderCurrent();
}

/*********************************************************
 * EXPORT FUNCTIONS
 *********************************************************/
async function exportShiftData(shiftName) {
  const rows =
    shiftName === "All"
      ? allDataRows
      : allDataRows.filter(r => r.data[0] === shiftName);

  if (!rows.length) return alert("No data to export");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Shifts");

  ws.addRow(headers);

  rows.forEach(r => {
    const row = [...r.data];
    row[1] = excelTimeToString(row[1]);
    row[2] = excelTimeToString(row[2]);
    ws.addRow(row);
  });

  const buf = await wb.xlsx.writeBuffer();
  download(buf, "Shift_Report.xlsx");
}

async function exportEscalatedTasks(shiftName) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Escalations");

  ws.addRow(["Shift", "App", "Task", "Issue", "Root Cause", "Remarks"]);

  allDataRows.forEach(r => {
    const shift = r.data[0];
    if (shiftName !== "All" && shift !== shiftName) return;

    headers.slice(3).forEach((app, idx) => {
      const key = `${shift}-${app}-${r.__rowId}`;
      if (taskStatuses[key] === "escalate") {
        const notes = escalationNotes[key] || {};
        ws.addRow([
          shift,
          app,
          r.data[idx + 3],
          notes.issue || "",
          notes.rootCause || "",
          notes.remarks || ""
        ]);
      }
    });
  });

  const buf = await wb.xlsx.writeBuffer();
  download(buf, "Escalation_Report.xlsx");
}

function download(buffer, name) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = name;
  link.click();
}
