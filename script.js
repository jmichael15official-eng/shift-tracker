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
 * DATE (MOUNTAIN TIME)
 *********************************************************/
function getMountainDateString() {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "2-digit"
  }).formatToParts(new Date());

  const get = t => parts.find(p => p.type === t)?.value;
  return `${get("month")}_${get("day")}_${get("year")}`;
}

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
 * LOAD EXCEL (GITHUB)
 *********************************************************/
async function loadMonitoringTemplate() {
  try {
    const res = await fetch("./data/Weekday_Monitoring Template.xlsx");
    if (!res.ok) throw new Error("Fetch failed");

    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    headers = rows[0];
    allDataRows = rows.slice(1).map((row, i) => ({
      __rowId: i,
      data: row
    }));

    filteredRows = allDataRows;
    currentIndex = 0;
    selectedShift = "All";

    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
    renderCurrent();

  } catch (e) {
    console.error(e);
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
  const el = document.getElementById("cardCounter");
  if (!filteredRows.length) {
    el.innerText = "No data";
    return;
  }
  el.innerText = `Shift ${currentIndex + 1} of ${filteredRows.length}`;
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
 * RENDER
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
 * STATUS
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
 * SHIFT FILTER (HIGHLIGHT FIXED)
 *********************************************************/
function buildShiftFilter() {
  const shifts = [...new Set(allDataRows.map(r => r.data[0]))];

  let html = `
    <button
      class="${selectedShift === "All" ? "active-filter" : ""}"
      onclick="filterByShift('All')">All</button>
  `;

  shifts.forEach(s => {
    html += `
      <button
        class="${selectedShift === s ? "active-filter" : ""}"
        onclick="filterByShift('${s}')">${s}</button>
    `;
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
  buildShiftFilter();
  renderCurrent();
}

/*********************************************************
 * EXPORT + RESET
 *********************************************************/
function buildExportButtons() {
  document.getElementById("exportButtons").innerHTML = `
    <button onclick="exportShiftData()">Export Shift</button>
    <button onclick="exportEscalatedTasks()">Export Escalations</button>
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
 * EXPORTS (FILENAME FIXED)
 *********************************************************/
async function exportShiftData() {
  const rows =
    selectedShift === "All"
      ? allDataRows
      : allDataRows.filter(r => r.data[0] === selectedShift);

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

  const date = getMountainDateString();
  const name =
    selectedShift === "All"
      ? `${date}_All_Shifts.xlsx`
      : `${date}_${selectedShift}.xlsx`;

  download(await wb.xlsx.writeBuffer(), name);
}

async function exportEscalatedTasks() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Escalations");

  ws.addRow(["Shift", "App", "Task", "Issue", "Root Cause", "Remarks"]);

  allDataRows.forEach(r => {
    const shift = r.data[0];
    if (selectedShift !== "All" && shift !== selectedShift) return;

    headers.slice(3).forEach((app, idx) => {
      const key = `${shift}-${app}-${r.__rowId}`;
      if (taskStatuses[key] === "escalate") {
        const n = escalationNotes[key] || {};
        ws.addRow([
          shift,
          app,
          r.data[idx + 3],
          n.issue || "",
          n.rootCause || "",
          n.remarks || ""
        ]);
      }
    });
  });

  const date = getMountainDateString();
  const name =
    selectedShift === "All"
      ? `${date}_All_Escalated_Report.xlsx`
      : `${date}_${selectedShift}_Escalated_Report.xlsx`;

  download(await wb.xlsx.writeBuffer(), name);
}

function download(buffer, filename) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}
