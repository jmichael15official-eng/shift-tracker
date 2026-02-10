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
  document.getElementById("currentTimes").innerHTML =
    `<strong>Manila:</strong> ${manila} | <strong>MT:</strong> ${mountain}`;
}
setInterval(updateCurrentTimes, 1000);
updateCurrentTimes();

/*********************************************************
 * DATE (MT)
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
function excelTimeToString(v) {
  if (typeof v === "number") {
    const s = Math.round(v * 86400);
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    return new Date(0, 0, 0, h, m).toLocaleTimeString("en-US", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true
    });
  }
  return v || "";
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
    allDataRows = rows.slice(1).map((r, i) => ({ __rowId: i, data: r }));
    filteredRows = allDataRows;

    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
    renderCurrent();
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
      : "No data";
}
function prevCard() {
  if (currentIndex > 0) currentIndex--, renderCurrent();
}
function nextCard() {
  if (currentIndex < filteredRows.length - 1) currentIndex++, renderCurrent();
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
 * STATUS + ESCALATION (FIXED)
 *********************************************************/
function setStatus(key, status, openModal = false) {
  taskStatuses[key] = status;

  if (status === "escalate" && openModal) {
    currentEscalationKey = key;

    // load existing notes OR clear
    const notes = escalationNotes[key] || {};
    noteIssue.value = notes.issue || "";
    noteRootCause.value = notes.rootCause || "";
    noteRemarks.value = notes.remarks || "";

    document.getElementById("escalationModal").style.display = "block";
  }

  renderCurrent();
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
  let html =
    `<button class="${selectedShift === "All" ? "active-filter" : ""}"
      onclick="filterByShift('All')">All</button>`;
  shifts.forEach(s => {
    html += `<button class="${selectedShift === s ? "active-filter" : ""}"
      onclick="filterByShift('${s}')">${s}</button>`;
  });
  document.getElementById("shiftFilter").innerHTML = html;
}
function filterByShift(s) {
  selectedShift = s;
  filteredRows = s === "All" ? allDataRows : allDataRows.filter(r => r.data[0] === s);
  currentIndex = 0;
  buildShiftFilter();
  renderCurrent();
}

/*********************************************************
 * EXPORT + RESET (UNCHANGED)
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
