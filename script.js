let allDataRows = [];
let headers = [];
let currentIndex = 0;
let filteredRows = [];
let taskStatuses = {};
let escalationNotes = {};
let currentEscalationKey = "";
let selectedShift = "All";

// ---------------- CURRENT TIMES ----------------
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

// ---------------- FILE UPLOAD ----------------
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
    allDataRows = rows.slice(1);
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

// ---------------- NAVIGATION ----------------
function buildNavigation() {
  document.getElementById("navigation").innerHTML = `
    <button id="prevBtn" onclick="prevCard()">Previous</button>
    <span id="cardCounter"></span>
    <button id="nextBtn" onclick="nextCard()">Next</button>
  `;
  updateCounter();
}

function updateCounter() {
  document.getElementById("cardCounter").innerText =
    `Shift ${currentIndex + 1} of ${filteredRows.length}`;
  document.getElementById("prevBtn").disabled = currentIndex === 0;
  document.getElementById("nextBtn").disabled = currentIndex >= filteredRows.length - 1;
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

// ---------------- CARD RENDER ----------------
function renderSingleCard(rows, index) {
  const row = rows[index];
  if (!row) return;

  const shift = row[0];
  const manilaTime = excelTimeToString(row[1]);
  const mtTime = excelTimeToString(row[2]);

  let html = `<div class="shift-card">
      <div class="shift-header">${shift} — ${manilaTime} Manila | ${mtTime} MT Time</div>
  `;

  for (let i = 3; i < headers.length; i++) {
    const app = headers[i];
    const task = row[i];
    if (!task) continue;

    const key = `${shift}-${app}-${index}`;
    const status = taskStatuses[key];

    html += `
      <div class="task">
        <strong>${app}</strong>
        <span>${task}</span>
        <div class="task-buttons">
          <button class="${status==='good'?'active-good':''}" onclick="setStatus(this,'good','${key}')"><i class="fas fa-check"></i>Good</button>
          <button class="${status==='monitor'?'active-monitor':''}" onclick="setStatus(this,'monitor','${key}')"><i class="fas fa-eye"></i>Monitor</button>
          <button class="${status==='escalate'?'active-escalate':''}" onclick="setStatus(this,'escalate','${key}',true)"><i class="fas fa-exclamation-triangle"></i>Escalate</button>
        </div>
      </div>
    `;
  }

  document.getElementById("output").innerHTML = html + "</div>";
}

// ---------------- STATUS HANDLING ----------------
function setStatus(button, status, key, openModal=false){
  const buttons = button.parentElement.querySelectorAll("button");
  buttons.forEach(b => b.classList.remove("active-good","active-monitor","active-escalate"));
  button.classList.add(`active-${status}`);
  taskStatuses[key]=status;

  if(status==="escalate" && openModal){
    currentEscalationKey=key;
    openModalWindow(key);
  }
}

// ---------------- ESCALATION MODAL ----------------
function openModalWindow(key){
  const notes = escalationNotes[key]||{};
  document.getElementById("noteIssue").value=notes.issue||"";
  document.getElementById("noteRootCause").value=notes.rootCause||"";
  document.getElementById("noteRemarks").value=notes.remarks||"";
  document.getElementById("escalationModal").style.display="block";
}

function closeModal(){ document.getElementById("escalationModal").style.display="none"; }

function saveEscalationNotes(){
  escalationNotes[currentEscalationKey]={
    issue: document.getElementById("noteIssue").value,
    rootCause: document.getElementById("noteRootCause").value,
    remarks: document.getElementById("noteRemarks").value
  };
  closeModal();
}

// ---------------- FILTERS ----------------
function buildShiftFilter(){
  const shifts=[...new Set(allDataRows.map(r=>r[0]))];
  let html=`<button onclick="filterByShift('All')" class="${selectedShift==='All'?'active-filter':''}">All Shift</button>`;
  shifts.forEach(s=>{
    html+=`<button onclick="filterByShift('${s}')" class="${selectedShift===s?'active-filter':''}">${s}</button>`;
  });
  document.getElementById("shiftFilter").innerHTML=html;
}

function filterByShift(shift){
  selectedShift=shift;
  filteredRows=shift==='All'?allDataRows:allDataRows.filter(r=>r[0]===shift);
  currentIndex=0;
  filteredRows.length?renderSingleCard(filteredRows,currentIndex):
    (document.getElementById("output").innerHTML="No data");
  buildShiftFilter();
  updateCounter();
}

// ---------------- EXPORT BUTTONS ----------------
function buildExportButtons(){
  document.getElementById("exportButtons").innerHTML=`
    <button onclick="exportShiftData(selectedShift)">Export Selected Shift</button>
    <button onclick="exportEscalatedTasks(selectedShift)">Export Escalation (Selected Shift)</button>
  `;
}

// ---------------- MOUNTAIN TIME DATE WITH ROLLOVER ----------------
function getMountainTimeDateString(){
  const now = new Date();
  const mtString = now.toLocaleString("en-US", {timeZone:"America/Denver"});
  const mtDate = new Date(mtString);

  if(mtDate.getHours() >= 15) mtDate.setDate(mtDate.getDate()+1);

  return mtDate.toLocaleString("en-US",{
    month:"long",
    day:"numeric",
    year:"numeric",
    timeZone:"America/Denver"
  }).replace(/,/g,"").replace(/ /g,"_");
}

// ---------------- EXPORT SHIFT ----------------
async function exportShiftData(shiftName){
  const rowsToExport=shiftName==="All"?allDataRows:allDataRows.filter(r=>r[0]===shiftName);
  if(!rowsToExport.length){ alert("No data to export!"); return; }

  const workbook=new ExcelJS.Workbook();
  const worksheet=workbook.addWorksheet("Shifts");

  worksheet.addRow(headers);
  rowsToExport.forEach(row=>{
    const copy=[...row];
    copy[1]=excelTimeToString(row[1]);
    copy[2]=excelTimeToString(row[2]);
    worksheet.addRow(copy);
  });

  rowsToExport.forEach((row,rIndex)=>{
    const shift=row[0];
    for(let cIndex=3;cIndex<headers.length;cIndex++){
      const appName=headers[cIndex];
      const key=`${shift}-${appName}-${rIndex}`;
      const status=taskStatuses[key];
      if(status){
        const cell=worksheet.getRow(rIndex+2).getCell(cIndex+1);
        if(status==="good") cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF00FF00'}};
        else if(status==="monitor") cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFF00'}};
        else if(status==="escalate"){ cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFF0000'}}; cell.font={bold:true}; }
      }
    }
  });

  const dateStr=getMountainTimeDateString();
  const fileName=shiftName==="All"?`All_Shifts_${dateStr}.xlsx`:`${dateStr}_${shiftName}.xlsx`;

  const buffer=await workbook.xlsx.writeBuffer();
  const blob=new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const link=document.createElement("a");
  link.href=URL.createObjectURL(blob);
  link.download=fileName;
  link.click();
}

// ---------------- EXPORT ESCALATED ----------------
async function exportEscalatedTasks(shiftName){
  const escalatedRows=[];

  allDataRows.forEach((row,rIndex)=>{
    const shift=row[0];
    if(shiftName!=="All" && shift!==shiftName) return;

    const manilaTime=excelTimeToString(row[1]);
    const mtTime=excelTimeToString(row[2]);

    for(let cIndex=3;cIndex<headers.length;cIndex++){
      const appName=headers[cIndex];
      const key=`${shift}-${appName}-${rIndex}`;
      const status=taskStatuses[key];

      if(status==="escalate"){
        const notes=escalationNotes[key]||{};
        escalatedRows.push({
          shift,
          app: appName,
          task: row[cIndex],
          manilaTime,
          mtTime,
          issue: notes.issue||"",
          rootCause: notes.rootCause||"",
          remarks: notes.remarks||""
        });
      }
    }
  });

  if(!escalatedRows.length){ alert("No escalated tasks to export!"); return; }

  const workbook=new ExcelJS.Workbook();
  const worksheet=workbook.addWorksheet("Escalated Tasks");

  worksheet.addRow(["Shift","App","Task","Manila Time","MT Time","Issue","Root Cause","Remarks"]);
  escalatedRows.forEach(row=>{
    worksheet.addRow([row.shift,row.app,row.task,row.manilaTime,row.mtTime,row.issue,row.rootCause,row.remarks]);
  });

  escalatedRows.forEach((row,index)=>{
    const excelRow=worksheet.getRow(index+2);
    excelRow.eachCell(cell=>{
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFF0000'}};
      cell.font={bold:true};
    });
  });

  const dateStr=getMountainTimeDateString();
  const fileName=shiftName==="All"?`${dateStr}_All_Escalated_Report.xlsx`:`${dateStr}_${shiftName}_Escalated_Report.xlsx`;

  const buffer=await workbook.xlsx.writeBuffer();
  const blob=new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const link=document.createElement("a");
  link.href=URL.createObjectURL(blob);
  link.download=fileName;
  link.click();
}
