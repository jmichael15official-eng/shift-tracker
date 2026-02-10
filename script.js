/* =====================================================
   ================= GLOBAL STATE ======================
   ===================================================== */

let headers = [];
let masterRows = [];     // FULL data from Excel (never mutate)
let viewRows = [];       // Filtered rows for UI
let currentIndex = 0;

let taskStatuses = {};   // key -> good | monitor | escalate
let escalationNotes = {}; // key -> { issue, rootCause, remarks }

let selectedShift = "All";
let currentEscalationKey = "";

/* =====================================================
   ================= CURRENT TIMES =====================
   ===================================================== */

function updateCurrentTimes() {
  const now = new Date();

  const format = tz =>
    now.toLocaleString("en-US", {
      timeZone: tz,
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
    <p>Manila: ${format("Asia/Manila")}</p>
    <p>Mountain Time: ${format("America/Denver")}</p>
  `;
}

updateCurrentTimes();
setInterval(updateCurrentTimes, 1000);

/* =====================================================
   ================= EXCEL HELPERS =====================
   ===================================================== */

function excelTimeToString(value) {
  if (typeof value !== "number") return value || "";
  const totalSeconds = Math.round(value * 86400);
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);

  return new Date(0, 0, 0, h, m).toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true
  });
}

/* =====================================================
   ================= LOAD EXCEL (GITHUB) ================
   ===================================================== */

async function loadExcelFromGitHub() {
  try {
    const res = await fetch("data/Weekday_Monitoring Template.xlsx");
    if (!res.ok) throw new Error("Fetch failed");

    const buffer = await res.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    headers = rows[0];
    masterRows = rows.slice(1).map((row, i) => ({
      __rowId: i,
      data: row
    }));

    selectedShift = "All";
    viewRows = [...masterRows];
    currentIndex = 0;

    buildShiftFilter();
    buildNavigation();
    buildExportButtons();
    renderCurrent();

  } catch (err) {
    console.error(err);
    alert("Failed to load monitoring template");
  }
}

loadExcelFromGitHub();

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
  const total = viewRows.length;
  document.getElementById("cardCounter").innerText =
    total ? `Shift ${currentIndex + 1} of ${total}` : "No shifts";

  document.getElementById("prevBtn").disabled = currentIndex === 0;
  document.getElementById("nextBtn").disabled =
    currentIndex >= total - 1;
}

function prevCard() {
  if (currentIndex > 0) {
    currentIndex--;
    renderCurrent();
  }
}

function nextCard() {
  if (currentIndex < viewRows.length - 1) {
    currentIndex++;
    renderCurrent();
  }
}

/* =====================================================
   ================= CARD RENDER =======================
   ===================================================== */

function renderCurrent() {
  if (!viewRows.length) {
    document.getElementById("output").innerHTML = "No data";
    updateCounter();
    return;
  }

  const rowObj = viewRows[currentIndex];
  const row = rowObj.data;
  const rowId = rowObj.__rowId;

  const shift = row[0];
  const manila = excelTimeToString(row[1]);
  const mt = excelTimeToString(row[2]);

  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${shift} — ${manila} Manila | ${mt} MT
      </div>
  `;

  for (let c = 3; c < headers.length; c++) {
    if (!row[c]) continue;

    const app = headers[c];
    const key = `${shift}-${app}-${rowId}`;
    const status = taskStatuses[key] || "";

    html += `
      <div class="task ${status === "escalate" ? "escalated" : ""}">
        <strong>${app}</strong>
        <span>${row[c]}</span>
        <div class="task-buttons">
          ${statusButton("good", key, status)}
          ${statusButton("monitor", key, status)}
          ${statusButton("escalate", key, status)}
        </div>
      </div>
    `;
  }

  document.getElementById("output").innerHTML = html + "</div>";
  updateCounter();
}

function statusButton(type, key, current) {
  const icons = {
    good: "check",
    monitor: "eye",
    escalate: "exclamation-triangle"
  };

  return `
    <button class="${current === type ? `active-${type}` : ""}"
      onclick="setStatus('${type}','${key}')">
      <i class="fas fa-${icons[type]}"></i>${type}
    </button>
  `;
}

/* =====================================================
   ================= STATUS HANDLING ===================
   ===================================================== */

function setStatus(status, key) {
  taskStatuses[key] = status;

  if (status === "escalate" && !escalationNotes[key]) {
    currentEscalationKey = key;
    openModal(key);
  }

  renderCurrent();
}

/* =====================================================
   ================= ESCALATION MODAL ==================
   ===================================================== */

function openModal(key) {
  const notes = escalationNotes[key] || {};
  noteIssue.value = notes.issue || "";
  noteRootCause.value = notes.rootCause || "";
  noteRemarks.value = notes.remarks || "";
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

/* =====================================================
   ================= SHIFT FILTER ======================
   ===================================================== */

function buildShiftFilter() {
  const shifts = [...new Set(masterRows.map(r => r.data[0]))];

  let html = `<button onclick="filterByShift('All')" class="${selectedShift === "All" ? "active-filter" : ""}">All</button>`;

  shifts.forEach(s => {
    html += `<button onclick="filterByShift('${s}')" class="${selectedShift === s ? "active-filter" : ""}">${s}</button>`;
  });

  document.getElementById("shiftFilter").innerHTML = html;
}

function filterByShift(shift) {
  selectedShift = shift;
  viewRows =
    shift === "All"
      ? [...masterRows]
      : masterRows.filter(r => r.data[0] === shift);

  currentIndex = 0;
  buildShiftFilter();
  renderCurrent();
}

/* =====================================================
   ================= EXPORT HELPERS ====================
   ===================================================== */

function getMountainDate() {
  const now = new Date();
  return now.toLocaleDateString("en-US", {
    timeZone: "America/Denver",
    month: "long",
    day: "numeric",
    year: "numeric"
  }).replace(/ /g, "_");
}

/* =====================================================
   ================= EXPORT SHIFT ======================
   ===================================================== */

async function exportShiftData(shift) {
  const rows =
    shift === "All"
      ? masterRows
      : masterRows.filter(r => r.data[0] === shift);

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

  const buffer = await wb.xlsx.writeBuffer();
  download(buffer, `${getMountainDate()}_${shift}_Shifts.xlsx`);
}

/* =====================================================
   ================= EXPORT ESCALATIONS =================
   ===================================================== */

async function exportEscalatedTasks(shift) {
  const escalated = [];

  masterRows.forEach(r => {
    const row = r.data;
    if (shift !== "All" && row[0] !== shift) return;

    for (let c = 3; c < headers.length; c++) {
      const key = `${row[0]}-${headers[c]}-${r.__rowId}`;
      if (taskStatuses[key] === "escalate") {
        const notes = escalationNotes[key] || {};
        escalated.push([
          row[0],
          headers[c],
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

  if (!escalated.length) return alert("No escalations found");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Escalations");

  ws.addRow([
    "Shift",
    "App",
    "Task",
    "Manila Time",
    "MT Time",
    "Issue",
    "Root Cause",
    "Remarks"
  ]);

  escalated.forEach(r => ws.addRow(r));

  const buffer = await wb.xlsx.writeBuffer();
  download(buffer, `${getMountainDate()}_${shift}_Escalations.xlsx`);
}

/* =====================================================
   ================= DOWNLOAD ==========================
   ===================================================== */

function download(buffer, name) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
}
