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

  const manilaTime = now.toLocaleString("en-US", {
    timeZone: "Asia/Manila",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  });

  const mountainTime = now.toLocaleString("en-US", {
    timeZone: "America/Denver",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  });

  document.getElementById("currentTimes").innerHTML = `
    <strong>Manila:</strong> ${manilaTime} |
    <strong>MT:</strong> ${mountainTime}
  `;
}
setInterval(updateCurrentTimes, 1000);
updateCurrentTimes();

/*********************************************************
 * DATE (MT SAFE)
 *********************************************************/
function getMountainDateString() {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "numeric"
  }).formatToParts(new Date());

  const get = t => parts.find(p => p.type === t)?.value;
  return `${get("month")}_${get("day")}_${get("year")}`;
}

/*********************************************************
 * EXCEL TIME
 *********************************************************/
function excelTimeToString(value) {
  if (typeof value === "number") {
    const totalSeconds = Math.round(value * 86400);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    return new Date(0, 0, 0, hours, minutes).toLocaleTimeString("en-US", {
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
    const res = await fetch("./data/Weekday_Monitoring Template.xlsx");
    if (!res.ok) throw new Error("Fetch failed");

    const buffer = await res.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    headers = rows[0];
    allDataRows = rows.slice(1).map((row, i) => ({
      __rowId: i,
      data: row
    }));

    filteredRows = allDataRows;

    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
    renderCurrentCard();

  } catch (e) {
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
  document.getElementById("cardCounter").innerText =
    filteredRows.length
      ? `Shift ${currentIndex + 1} of ${filteredRows.length}`
      : "No shifts";
}

function prevCard() {
  if (currentIndex > 0) {
    currentIndex--;
    renderCurrentCard();
  }
}

function nextCard() {
  if (currentIndex < filteredRows.length - 1) {
    currentIndex++;
    renderCurrentCard();
  }
}

/*********************************************************
 * RENDER CARD
 *********************************************************/
function renderCurrentCard() {
  const rowObj = filteredRows[currentIndex];
  if (!rowObj) return;

  const row = rowObj.data;
  const rowId = rowObj.__rowId;

  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${row[0]} —
        ${excelTimeToString(row[1])} Manila |
        ${excelTimeToString(row[2])} MT
      </div>
  `;

  for (let i = 3; i < headers.length; i++) {
    if (!row[i]) continue;

    const app = headers[i];
    const key = `${row[0]}-${app}-${rowId}`;
    const status = taskStatuses[key] || "";

    html += `
      <div class="task ${status === "escalate" ? "escalated" : ""}">
        <strong>${app}</strong>
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
 * STATUS + ESCALATION (FIXED)
 *********************************************************/
function setStatus(key, status, openModal = false) {
  taskStatuses[key] = status;

  if (status === "escalate" && openModal) {
    currentEscalationKey = key;

    const notes = escalationNotes[key] || {};
    noteIssue.value = notes.issue || "";
    noteRootCause.value = notes.rootCause || "";
    noteRemarks.value = notes.remarks || "";

    document.getElementById("escalationModal").style.display = "block";
  }

  renderCurrentCard();
}

function closeModal() {
  document.getElementById("escalationModal").style.display = "none";
}

function saveEscalationNotes() {
  escalationNotes[currentEscalationKey] = {
    issue: noteIssue.value.trim(),
    rootCause: noteRootCause.value.trim(),
    remarks: noteRemarks.value.trim()
  };
  closeModal();
}

/*********************************************************
 * SHIFT FILTER
 *********************************************************/
function buildShiftFilter() {
  const shifts = [...new Set(allDataRows.map(r => r.data[0]))];
  let html = `
    <button class="${selectedShift === "All" ? "active-filter" : ""}"
      onclick="filterByShift('All')">All Shift</button>
  `;
  shifts.forEach(s => {
    html += `
      <button class="${selectedShift === s ? "active-filter" : ""}"
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
  renderCurrentCard();
}

/*********************************************************
 * EXPORT + RESET
 *********************************************************/
function buildExportButtons() {
  document.getElementById("exportButtons").innerHTML = `
    <button onclick="exportShiftData()">Export Shift</button>
    <button onclick="exportEscalatedTasks()">Export Escalation</button>
    <button onclick="resetAllStatuses()">Reset Status</button>
  `;
}

function resetAllStatuses() {
  if (!confirm("Reset all task statuses and notes?")) return;
  taskStatuses = {};
  escalationNotes = {};
  renderCurrentCard();
}
