let allDataRows = [];
let headers = [];
let currentIndex = 0;
let filteredRows = [];
let taskStatuses = {};
let escalationNotes = {};
let currentEscalationKey = "";
let selectedShift = "All";

/* =====================================================
   ================= CURRENT TIMES =====================
   ===================================================== */

function updateCurrentTimes() {
  const now = new Date();

  const manilaTime = now.toLocaleString("en-US", {
    timeZone: "Asia/Manila",
    month: "long",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "numeric",
    second: "numeric",
    hour12: true
  });

  const mountainTime = now.toLocaleString("en-US", {
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
    <p>Manila: ${manilaTime}</p>
    <p>Mountain Time: ${mountainTime}</p>
  `;
}

updateCurrentTimes();
setInterval(updateCurrentTimes, 1000);

/* =====================================================
   ================= FILE UPLOAD =======================
   ===================================================== */

document.getElementById("excelFile").addEventListener("change", handleFile);

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

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = e => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
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

    renderSingleCard(filteredRows, currentIndex);
    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
  };
  reader.readAsArrayBuffer(file);
}

/* =====================================================
   ================= NAVIGATION ========================
   ===================================================== */

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
    document.getElementById("prevBtn").disabled = true;
    document.getElementById("nextBtn").disabled = true;
    return;
  }

  document.getElementById("cardCounter").innerText =
    `Shift ${currentIndex + 1} of ${filteredRows.length}`;

  document.getElementById("prevBtn").disabled = currentIndex === 0;
  document.getElementById("nextBtn").disabled =
    currentIndex >= filteredRows.length - 1;
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

/* =====================================================
   ================= CARD RENDER =======================
   ===================================================== */

function renderSingleCard(rows, index) {
  const rowObj = rows[index];
  if (!rowObj) return;

  const row = rowObj.data;
  const rowId = rowObj.__rowId;

  const shift = row[0];
  const manilaTime = excelTimeToString(row[1]);
  const mtTime = excelTimeToString(row[2]);

  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${shift} — ${manilaTime} Manila | ${mtTime} MT Time
      </div>
  `;

  for (let i = 3; i < headers.length; i++) {
    const app = headers[i];
    const task = row[i];
    if (!task) continue;

    const key = `${shift}-${app}-${rowId}`;
    const status = taskStatuses[key] || "";

    html += `
      <div class="task ${status === "escalate" ? "escalated" : ""}">
        <strong>${app}</strong>
        <span>${task}</span>
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

/* =====================================================
   ================= STATUS HANDLING ===================
   ===================================================== */

function setStatus(button, status, key, openModal = false) {
  const buttons = button.parentElement.querySelectorAll("button");
  buttons.forEach(b =>
    b.classList.remove("active-good", "active-monitor", "active-escalate")
  );

  button.classList.add(`active-${status}`);
  taskStatuses[key] = status;

  if (status === "escalate" && openModal && !escalationNotes[key]) {
    currentEscalationKey = key;
    openModalWindow(key);
  }
}

/* =====================================================
   ================= ESCALATION MODAL ==================
   ===================================================== */

function openModalWindow(key) {
  const notes = escalationNotes[key] || {};
  document.getElementById("noteIssue").value = notes.issue || "";
  document.getElementById("noteRootCause").value = notes.rootCause || "";
  document.getElementById("noteRemarks").value = notes.remarks || "";

  document.getElementById("escalationModal").style.display = "block";
  setTimeout(() => document.getElementById("noteIssue").focus(), 100);
}

function closeModal() {
  document.getElementById("escalationModal").style.display = "none";
}

function saveEscalationNotes() {
  escalationNotes[currentEscalationKey] = {
    issue: document.getElementById("noteIssue").value,
    rootCause: document.getElementById("noteRootCause").value,
    remarks: document.getElementById("noteRemarks").value
  };
  closeModal();
}

/* =====================================================
   ================= SHIFT FILTER ======================
   ===================================================== */

function buildShiftFilter() {
  const shifts = [...new Set(allDataRows.map(r => r.data[0]))];

  let html = `
    <button onclick="filterByShift('All')"
      class="${selectedShift === "All" ? "active-filter" : ""}">
      All Shift
    </button>
  `;

  shifts.forEach(shift => {
    html += `
      <button onclick="filterByShift('${shift}')"
        class="${selectedShift === shift ? "active-filter" : ""}">
        ${shift}
      </button>
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

  filteredRows.length
    ? renderSingleCard(filteredRows, currentIndex)
    : (document.getElementById("output").innerHTML = "No data");

  buildShiftFilter();
  updateCounter();
}

/* =====================================================
   ================= EXPORT BUTTONS ====================
   ===================================================== */

function buildExportButtons() {
  document.getElementById("exportButtons").innerHTML = `
    <button onclick="exportShiftData(selectedShift)">Export Selected Shift</button>
    <button onclick="exportEscalatedTasks(selectedShift)">
      Export Escalation (Selected Shift)
    </button>
  `;
}

/* =====================================================
   ============ MOUNTAIN TIME DATE (SAFE) ===============
   ===================================================== */

function getMountainTimeDateString() {
  const now = new Date();

  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "numeric",
    hour12: false
  }).formatToParts(now);

  const get = t => parts.find(p => p.type === t)?.value;

  let year = get("year");
  let month = get("month");
  let day = get("day");
  let hour = parseInt(get("hour"), 10);

  if (hour >= 15) {
    const temp = new Date(`${month} ${day}, ${year}`);
    temp.setDate(temp.getDate() + 1);
    year = temp.getFullYear();
    month = temp.toLocaleString("en-US", { month: "long" });
    day = temp.getDate();
  }

  return `${month}_${day}_${year}`;
}

/* =====================================================
   ================= EXPORT SHIFT ======================
   ===================================================== */

async function exportShiftData(shiftName) {
  const rows =
    shiftName === "All"
      ? allDataRows
      : allDataRows.filter(r => r.data[0] === shiftName);

  if (!rows.length) {
    alert("No data to export!");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Shifts");

  worksheet.addRow(headers);

  rows.forEach(rowObj => {
    const row = [...rowObj.data];
    row[1] = excelTimeToString(row[1]);
    row[2] = excelTimeToString(row[2]);
    worksheet.addRow(row);
  });

  rows.forEach((rowObj, rIndex) => {
    const shift = rowObj.data[0];
    for (let c = 3; c < headers.length; c++) {
      const app = headers[c];
      const key = `${shift}-${app}-${rowObj.__rowId}`;
      const status = taskStatuses[key];

      if (!status) continue;

      const cell = worksheet.getRow(rIndex + 2).getCell(c + 1);
      if (status === "good")
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } };
      if (status === "monitor")
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      if (status === "escalate") {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
        cell.font = { bold: true };
      }
    }
  });

  const fileName =
    shiftName === "All"
      ? `All_Shifts_${getMountainTimeDateString()}.xlsx`
      : `${getMountainTimeDateString()}_${shiftName}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
}

/* =====================================================
   ============== EXPORT ESCALATED TASKS ================
   ===================================================== */

async function exportEscalatedTasks(shiftName) {
  const escalated = [];

  allDataRows.forEach(rowObj => {
    const row = rowObj.data;
    const shift = row[0];
    if (shiftName !== "All" && shift !== shiftName) return;

    for (let c = 3; c < headers.length; c++) {
      const app = headers[c];
      const key = `${shift}-${app}-${rowObj.__rowId}`;

      if (taskStatuses[key] === "escalate") {
        const notes = escalationNotes[key] || {};
        escalated.push([
          shift,
          app,
          row[c],
          excelTimeToString(row[1]),
          excelTimeToString(row[2]),
          notes.issue || "",
          notes.rootCause || "",
          notes.remarks || ""
        ]);
      }
    }
  });

  if (!escalated.length) {
    alert("No escalated tasks to export!");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Escalated Tasks");

  worksheet.addRow([
    "Shift",
    "App",
    "Task",
    "Manila Time",
    "MT Time",
    "Issue",
    "Root Cause",
    "Remarks"
  ]);

  escalated.forEach(row => worksheet.addRow(row));

  escalated.forEach((_, i) => {
    worksheet.getRow(i + 2).eachCell(cell => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } };
      cell.font = { bold: true };
    });
  });

  const fileName =
    shiftName === "All"
      ? `${getMountainTimeDateString()}_All_Escalated_Report.xlsx`
      : `${getMountainTimeDateString()}_${shiftName}_Escalated_Report.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
}
