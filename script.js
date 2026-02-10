/* ===================== CONFIG ===================== */

const EXCEL_URL = "./data/monitoring_template.xlsx";

/* ===================== STATE ===================== */

let allDataRows = [];
let headers = [];
let filteredRows = [];
let currentIndex = 0;
let selectedShift = "All";

let taskStatuses = {};
let escalationNotes = {};
let currentEscalationKey = "";

/* ================= CURRENT TIMES ================= */

function updateCurrentTimes() {
  const now = new Date();

  const manila = now.toLocaleString("en-US", {
    timeZone: "Asia/Manila",
    month: "long",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "numeric",
    second: "numeric",
    hour12: true
  });

  const mt = now.toLocaleString("en-US", {
    timeZone: "America/Denver",
    month: "long",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "numeric",
    second: "numeric",
    hour12: true
  });

  document.getElementById("currentTimes").innerHTML = `
    <h2>Current Times</h2>
    <p>Manila: ${manila}</p>
    <p>Mountain Time: ${mt}</p>
  `;
}

updateCurrentTimes();
setInterval(updateCurrentTimes, 1000);

/* ================= EXCEL HELPERS ================= */

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

/* ================= LOAD EXCEL ================= */

async function loadExcelFromGitHub() {
  const res = await fetch(EXCEL_URL);
  if (!res.ok) {
    document.getElementById("output").innerText =
      "Failed to load monitoring template.";
    return;
  }

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
  currentIndex = 0;
  selectedShift = "All";

  buildNavigation();
  buildShiftFilter();
  buildExportButtons();
  renderSingleCard(filteredRows, currentIndex);
}

/* ================= NAVIGATION ================= */

function buildNavigation() {
  document.getElementById("navigation").innerHTML = `
    <button id="prevBtn" onclick="prevCard()">Previous</button>
    <span id="cardCounter"></span>
    <button id="nextBtn" onclick="nextCard()">Next</button>
  `;
  updateCounter();
}

function updateCounter() {
  if (!filteredRows.length) {
    document.getElementById("cardCounter").innerText = "No shifts";
    return;
  }

  document.getElementById("cardCounter").innerText =
    `Shift ${currentIndex + 1} of ${filteredRows.length}`;

  document.getElementById("prevBtn").disabled = currentIndex === 0;
  document.getElementById("nextBtn").disabled =
    currentIndex === filteredRows.length - 1;
}

function prevCard() {
  if (currentIndex > 0) {
    currentIndex--;
    renderSingleCard(filteredRows, currentIndex);
    updateCounter();
  }
}

function nextCard() {
  if (currentIndex < filteredRows.length - 1) {
    currentIndex++;
    renderSingleCard(filteredRows, currentIndex);
    updateCounter();
  }
}

/* ================= CARD RENDER ================= */

function renderSingleCard(rows, index) {
  const obj = rows[index];
  if (!obj) return;

  const row = obj.data;
  const rowId = obj.__rowId;

  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${row[0]} — ${excelTimeToString(row[1])} Manila | ${excelTimeToString(row[2])} MT
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
            onclick="setStatus(this,'good','${key}')">
            <i class="fas fa-check"></i>Good
          </button>
          <button class="${status === "monitor" ? "active-monitor" : ""}"
            onclick="setStatus(this,'monitor','${key}')">
            <i class="fas fa-eye"></i>Monitor
          </button>
          <button class="${status === "escalate" ? "active-escalate" : ""}"
            onclick="setStatus(this,'escalate','${key}',true)">
            <i class="fas fa-exclamation-triangle"></i>Escalate
          </button>
        </div>
      </div>
    `;
  }

  document.getElementById("output").innerHTML = html + "</div>";
}

/* ================= STATUS ================= */

function setStatus(btn, status, key, modal = false) {
  btn.parentElement
    .querySelectorAll("button")
    .forEach(b => b.className = "");

  btn.classList.add(`active-${status}`);
  taskStatuses[key] = status;

  if (status === "escalate" && modal && !escalationNotes[key]) {
    currentEscalationKey = key;
    openModalWindow(key);
  }
}

/* ================= MODAL ================= */

function openModalWindow(key) {
  const n = escalationNotes[key] || {};
  noteIssue.value = n.issue || "";
  noteRootCause.value = n.rootCause || "";
  noteRemarks.value = n.remarks || "";
  escalationModal.style.display = "block";
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

/* ================= FILTER ================= */

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

  shiftFilter.innerHTML = html;
}

function filterByShift(shift) {
  selectedShift = shift;
  filteredRows =
    shift === "All"
      ? allDataRows
      : allDataRows.filter(r => r.data[0] === shift);

  currentIndex = 0;
  buildShiftFilter();
  renderSingleCard(filteredRows, 0);
  updateCounter();
}

/* ================= EXPORT ================= */

function buildExportButtons() {
  exportButtons.innerHTML = `
    <button onclick="exportShiftData(selectedShift)">Export Selected Shift</button>
    <button onclick="exportEscalatedTasks(selectedShift)">
      Export Escalations
    </button>
  `;
}

/* ===== MOUNTAIN DATE ===== */

function getMountainTimeDateString() {
  const now = new Date();
  return new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "numeric"
  }).format(now).replace(/ /g, "_");
}

/* ================= INIT ================= */

window.addEventListener("DOMContentLoaded", loadExcelFromGitHub);
