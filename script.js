let allDataRows = [];
let headers = [];
let currentIndex = 0;
let filteredRows = [];
let taskStatuses = {};
let escalationNotes = {};
let currentEscalationKey = "";
let selectedShift = "All";

/* ================= CURRENT TIMES ================= */

function updateCurrentTimes() {
  const now = new Date();
  document.getElementById("currentTimes").innerHTML = `
    <p>Manila: ${now.toLocaleString("en-US",{timeZone:"Asia/Manila"})}</p>
    <p>Mountain: ${now.toLocaleString("en-US",{timeZone:"America/Denver"})}</p>
  `;
}
updateCurrentTimes();
setInterval(updateCurrentTimes, 1000);

/* ================= FETCH TEMPLATE ================= */

async function loadTemplate() {
  const res = await fetch("./data/Weekday_Monitoring Template.xlsx");
  const buffer = await res.arrayBuffer();

  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  headers = rows[0];
  allDataRows = rows.slice(1).map((r, i) => ({ __rowId: i, data: r }));

  filteredRows = allDataRows;
  currentIndex = 0;

  renderSingleCard(filteredRows, currentIndex);
  buildNavigation();
  buildShiftFilter();
  buildExportButtons();
}

/* ================= NAVIGATION ================= */

function buildNavigation() {
  document.getElementById("navigation").innerHTML = `
    <button onclick="prevCard()">Prev</button>
    <span id="cardCounter"></span>
    <button onclick="nextCard()">Next</button>
  `;
  updateCounter();
}

function updateCounter() {
  document.getElementById("cardCounter").innerText =
    `Shift ${currentIndex + 1} of ${filteredRows.length}`;
}

function prevCard() {
  if (currentIndex > 0) currentIndex--, renderSingleCard(filteredRows, currentIndex), updateCounter();
}
function nextCard() {
  if (currentIndex < filteredRows.length - 1) currentIndex++, renderSingleCard(filteredRows, currentIndex), updateCounter();
}

/* ================= RENDER ================= */

function renderSingleCard(rows, index) {
  const rowObj = rows[index];
  if (!rowObj) return;

  const row = rowObj.data;
  const shift = row[0];

  let html = `<div class="shift-card"><div class="shift-header">${shift}</div>`;

  for (let i = 3; i < headers.length; i++) {
    const app = headers[i];
    const task = row[i];
    if (!task) continue;

    const key = `${shift}-${app}-${rowObj.__rowId}`;
    const status = taskStatuses[key] || "";

    html += `
      <div class="task ${status === "escalate" ? "escalated" : ""}">
        <strong>${app}</strong> — ${task}
        <div class="task-buttons">
          <button class="${status==="good"?"active-good":""}"
            onclick="setStatus(this,'good','${key}')">Good</button>
          <button class="${status==="monitor"?"active-monitor":""}"
            onclick="setStatus(this,'monitor','${key}')">Monitor</button>
          <button class="${status==="escalate"?"active-escalate":""}"
            onclick="setStatus(this,'escalate','${key}',true)">Escalate</button>
        </div>
      </div>`;
  }

  document.getElementById("output").innerHTML = html + "</div>";
}

/* ================= STATUS ================= */

function setStatus(btn, status, key, openModal=false) {
  btn.parentElement.querySelectorAll("button")
    .forEach(b=>b.className="");

  btn.classList.add(`active-${status}`);
  taskStatuses[key] = status;

  if (status==="escalate" && openModal) {
    currentEscalationKey = key;
    openModalWindow(key);
  }
}

/* ================= RESET ================= */

function resetCurrentShift() {
  if (!confirm("Reset all statuses for this shift?")) return;

  const shift = filteredRows[currentIndex].data[0];
  Object.keys(taskStatuses).forEach(k => {
    if (k.startsWith(shift)) delete taskStatuses[k];
  });
  Object.keys(escalationNotes).forEach(k => {
    if (k.startsWith(shift)) delete escalationNotes[k];
  });

  renderSingleCard(filteredRows, currentIndex);
}

/* ================= EXPORT HELPERS ================= */

function getMountainDate() {
  return new Date().toLocaleDateString("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "numeric"
  });
}

/* ================= EXPORT SHIFT ================= */

async function exportShiftData(shiftName) {
  const rows = shiftName==="All" ? allDataRows : allDataRows.filter(r=>r.data[0]===shiftName);
  if (!rows.length) return alert("No data");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Shift");

  ws.addRow(headers);

  rows.forEach((r,i)=>{
    ws.addRow(r.data);
    for (let c=3;c<headers.length;c++){
      const key = `${r.data[0]}-${headers[c]}-${r.__rowId}`;
      const status = taskStatuses[key];
      if (!status) continue;

      const cell = ws.getRow(i+2).getCell(c+1);
      if (status==="good") cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF00FF00"}};
      if (status==="monitor") cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFFF00"}};
      if (status==="escalate") {
        cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFF0000"}};
        cell.font={bold:true};
      }
    }
  });

  const name = `${getMountainDate()}_${shiftName}.xlsx`;
  const blob = new Blob([await wb.xlsx.writeBuffer()]);
  download(blob, name);
}

/* ================= EXPORT ESCALATED ================= */

async function exportEscalatedTasks(shiftName) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Escalated");

  ws.addRow(["Shift","App","Task","Issue","Root Cause","Remarks"]);

  Object.entries(taskStatuses).forEach(([key,val])=>{
    if (val!=="escalate") return;
    if (shiftName!=="All" && !key.startsWith(shiftName)) return;

    const notes = escalationNotes[key]||{};
    const [shift,app] = key.split("-");
    ws.addRow([shift,app,"",notes.issue||"",notes.rootCause||"",notes.remarks||""]);
  });

  const name = `${getMountainDate()}_Escalated_${shiftName}.xlsx`;
  const blob = new Blob([await wb.xlsx.writeBuffer()]);
  download(blob,name);
}

/* ================= UTIL ================= */

function download(blob,name){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download=name;
  a.click();
}
