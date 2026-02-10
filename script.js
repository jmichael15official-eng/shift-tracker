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
  document.getElementById("currentTimes").innerHTML = `
    <b>Manila:</b> ${now.toLocaleTimeString("en-US", { timeZone: "Asia/Manila" })} |
    <b>MT:</b> ${now.toLocaleTimeString("en-US", { timeZone: "America/Denver" })}
  `;
}
setInterval(updateCurrentTimes, 1000);
updateCurrentTimes();

/*********************************************************
 * DATE (MT SAFE)
 * MT DATE STRING
 *********************************************************/
function getMTFileDate() {
  const p = new Intl.DateTimeFormat("en-US", {
function getMTDateString() {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Denver",
    year: "numeric",
    month: "long",
    day: "numeric"
  }).formatToParts(new Date());

  const g = t => p.find(x => x.type === t)?.value;
  const g = t => parts.find(p => p.type === t)?.value;
  return `${g("month")}_${g("day")}_${g("year")}`;
}

@@ -58,28 +57,32 @@
}

/*********************************************************
 * LOAD TEMPLATE
 * LOAD TEMPLATE (FIXED PATH)
 *********************************************************/
async function loadMonitoringTemplate() {
  try {
    const res = await fetch("./data/Weekday_Monitoring Template.xlsx");
    const res = await fetch("./data/Weekday_Monitoring%20Template.xlsx");
    if (!res.ok) throw new Error();

    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    headers = rows[0];
    allDataRows = rows.slice(1).map((r, i) => ({ __rowId: i, data: r }));
    allDataRows = rows.slice(1).map((row, i) => ({
      __rowId: i,
      data: row
    }));
    filteredRows = allDataRows;

    buildNavigation();
    buildShiftFilter();
    buildExportButtons();
    renderCurrentCard();
  } catch {
    output.innerHTML = "<b style='color:red'>Failed to load monitoring template</b>";
    document.getElementById("output").innerHTML =
      "<b style='color:red'>Failed to load monitoring template</b>";
  }
}
loadMonitoringTemplate();
@@ -114,26 +117,28 @@
  let html = `
    <div class="shift-card">
      <div class="shift-header">
        ${row[0]} — ${excelTimeToString(row[1])} Manila | ${excelTimeToString(row[2])} MT
        ${row[0]} — ${excelTimeToString(row[1])} Manila |
        ${excelTimeToString(row[2])} MT
      </div>
  `;

  for (let i = 3; i < headers.length; i++) {
    if (!row[i]) continue;

    const app = headers[i];
    const key = `${row[0]}-${app}-${id}`;
    const s = taskStatuses[key] || "";
    const status = taskStatuses[key] || "";

    html += `
      <div class="task ${s === "escalate" ? "escalated" : ""}">
      <div class="task ${status === "escalate" ? "escalated" : ""}">
        <strong>${app}</strong>
        <span>${row[i]}</span>
        <div class="task-buttons">
          <button class="${s === "good" ? "active-good" : ""}"
          <button class="${status === "good" ? "active-good" : ""}"
            onclick="setStatus('${key}','good')">Good</button>
          <button class="${s === "monitor" ? "active-monitor" : ""}"
          <button class="${status === "monitor" ? "active-monitor" : ""}"
            onclick="setStatus('${key}','monitor')">Monitor</button>
          <button class="${s === "escalate" ? "active-escalate" : ""}"
          <button class="${status === "escalate" ? "active-escalate" : ""}"
            onclick="setStatus('${key}','escalate',true)">Escalate</button>
        </div>
      </div>
@@ -147,10 +152,10 @@
/*********************************************************
 * STATUS + MODAL
 *********************************************************/
function setStatus(key, status, open = false) {
function setStatus(key, status, openModal = false) {
  taskStatuses[key] = status;

  if (status === "escalate" && open) {
  if (status === "escalate" && openModal) {
    currentEscalationKey = key;
    const n = escalationNotes[key] || {};
    noteIssue.value = n.issue || "";
@@ -186,9 +191,11 @@
        onclick="filterByShift('${s}')">${s}</button>`
    ).join("");
}
function filterByShift(s) {
  selectedShift = s;
  filteredRows = s === "All" ? allDataRows : allDataRows.filter(r => r.data[0] === s);
function filterByShift(shift) {
  selectedShift = shift;
  filteredRows = shift === "All"
    ? allDataRows
    : allDataRows.filter(r => r.data[0] === shift);
  currentIndex = 0;
  buildShiftFilter();
  renderCurrentCard();
@@ -206,82 +213,105 @@
}

/*********************************************************
 * EXPORT SHIFT (FIXED)
 * EXPORT SHIFT (FIXED COLORS)
 *********************************************************/
async function exportShiftData() {
  if (!window.ExcelJS) return alert("ExcelJS not loaded");

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Shift");

  ws.addRow(headers);

  filteredRows.forEach((obj, r) => {
  filteredRows.forEach((obj, rIdx) => {
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
      const status = taskStatuses[key];
      if (!status) continue;

      const cell = ws.getRow(rIdx + 2).getCell(c + 1);
      const colors = {
        good: "FF00FF00",
        monitor: "FFFFFF00",
        escalate: "FFFF0000"
      };

      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: colors[status] }
      };
    }
  });

  const name = `${getMTFileDate()}_${selectedShift}.xlsx`;
  const buf = await wb.xlsx.writeBuffer();
  download(buf, name);
  const fileName = `${getMTDateString()}_${selectedShift}.xlsx`;
  download(await wb.xlsx.writeBuffer(), fileName);
}

/*********************************************************
 * EXPORT ESCALATION
 * EXPORT ESCALATED TASKS (FIXED)
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

  allDataRows.forEach(obj => {
    const row = obj.data;
    for (let c = 3; c < headers.length; c++) {
      const key = `${row[0]}-${headers[c]}-${obj.__rowId}`;
      if (taskStatuses[key] !== "escalate") continue;

      const notes = escalationNotes[key] || {};
      ws.addRow([
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
  });

  const buf = await wb.xlsx.writeBuffer();
  download(buf, `${getMTFileDate()}_Escalations.xlsx`);
  const fileName = `${getMTDateString()}_${selectedShift}_Escalated_Report.xlsx`;
  download(await wb.xlsx.writeBuffer(), fileName);
}

/*********************************************************
 * DOWNLOAD
 *********************************************************/
function download(buf, name) {
  const blob = new Blob([buf], {
function download(buffer, name) {
  const blob = new Blob([buffer], {
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
