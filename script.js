let allDataRows = [];
let headers = [];
let filteredRows = [];
let currentIndex = 0;
let selectedShift = "All";
let taskStatuses = {};
let escalationNotes = {};
let currentEscalationKey = "";

/* ================= TIMES ================= */

function updateCurrentTimes() {
  const now = new Date();
  currentTimes.innerHTML = `
    <p>Manila: ${now.toLocaleString("en-US",{timeZone:"Asia/Manila"})}</p>
    <p>Mountain: ${now.toLocaleString("en-US",{timeZone:"America/Denver"})}</p>
  `;
}
updateCurrentTimes();
setInterval(updateCurrentTimes,1000);

/* ================= LOAD TEMPLATE ================= */

window.addEventListener("load", async () => {
  try {
    const res = await fetch("./data/Weekday_Monitoring Template.xlsx");
    if (!res.ok) return;
    loadWorkbook(await res.arrayBuffer());
  } catch {}
});

excelFile.addEventListener("change", e => {
  const r = new FileReader();
  r.onload = ev => loadWorkbook(ev.target.result);
  r.readAsArrayBuffer(e.target.files[0]);
});

function loadWorkbook(buffer) {
  const wb = XLSX.read(buffer,{type:"array"});
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1});

  headers = rows[0];
  allDataRows = rows.slice(1).map((r,i)=>({__rowId:i,data:r}));
  filteredRows = allDataRows;
  currentIndex = 0;

  render();
  buildNav();
  buildFilters();
  buildExportButtons();
}

/* ================= UI ================= */

function buildNav() {
  navigation.innerHTML = `
    <button onclick="prev()">Prev</button>
    <span>${currentIndex+1}/${filteredRows.length}</span>
    <button onclick="next()">Next</button>
  `;
}

function buildFilters() {
  const shifts = [...new Set(allDataRows.map(r=>r.data[0]))];
  shiftFilter.innerHTML = `
    <button onclick="filterShift('All')" class="active-filter">All</button>
    ${shifts.map(s=>`<button onclick="filterShift('${s}')">${s}</button>`).join("")}
  `;
}

function buildExportButtons() {
  exportButtons.innerHTML = `
    <button onclick="exportShift()">Export Shift</button>
    <button onclick="exportEscalated()">Export Escalated</button>
    <button onclick="resetShift()">Reset Shift</button>
  `;
}

function render() {
  const r = filteredRows[currentIndex];
  if (!r) return;

  let html = `<div class="shift-card"><div class="shift-header">${r.data[0]}</div>`;
  for (let i=3;i<headers.length;i++) {
    const key = `${r.data[0]}-${headers[i]}-${r.__rowId}`;
    const status = taskStatuses[key]||"";
    html += `
      <div class="task ${status==="escalate"?"escalated":""}">
        <strong>${headers[i]}</strong> ${r.data[i]}
        <div class="task-buttons">
          <button class="${status==="good"?"active-good":""}" onclick="setStatus('${key}','good')">Good</button>
          <button class="${status==="monitor"?"active-monitor":""}" onclick="setStatus('${key}','monitor')">Monitor</button>
          <button class="${status==="escalate"?"active-escalate":""}" onclick="setStatus('${key}','escalate',true)">Escalate</button>
        </div>
      </div>`;
  }
  output.innerHTML = html+"</div>";
}

/* ================= STATUS ================= */

function setStatus(key,status,modal=false){
  taskStatuses[key]=status;
  if (modal) { currentEscalationKey=key; escalationModal.style.display="block"; }
  render();
}

/* ================= RESET ================= */

function resetShift(){
  if(!confirm("Reset all statuses for this shift?"))return;
  const shift=filteredRows[currentIndex].data[0];
  Object.keys(taskStatuses).forEach(k=>k.startsWith(shift)&&delete taskStatuses[k]);
  Object.keys(escalationNotes).forEach(k=>k.startsWith(shift)&&delete escalationNotes[k]);
  render();
}

/* ================= EXPORT ================= */

function getDate(){
  return new Date().toLocaleDateString("en-US",{timeZone:"America/Denver",month:"long",day:"numeric",year:"numeric"});
}

async function exportShift(){
  const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet("Shift");
  ws.addRow(headers);

  filteredRows.forEach((r,i)=>{
    ws.addRow(r.data);
    for(let c=3;c<headers.length;c++){
      const key=`${r.data[0]}-${headers[c]}-${r.__rowId}`;
      const cell=ws.getRow(i+2).getCell(c+1);
      if(taskStatuses[key]==="good") cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF00FF00"}};
      if(taskStatuses[key]==="monitor") cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFFF00"}};
      if(taskStatuses[key]==="escalate"){cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFF0000"}};cell.font={bold:true};}
    }
  });

  download(await wb.xlsx.writeBuffer(),`${getDate()}_${filteredRows[currentIndex].data[0]}.xlsx`);
}

async function exportEscalated(){
  const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet("Escalated");
  ws.addRow(["Shift","App","Issue","Root Cause","Remarks"]);

  Object.entries(taskStatuses).forEach(([k,v])=>{
    if(v!=="escalate")return;
    const n=escalationNotes[k]||{};
    const [s,a]=k.split("-");
    ws.addRow([s,a,n.issue||"",n.rootCause||"",n.remarks||""]);
  });

  download(await wb.xlsx.writeBuffer(),`${getDate()}_Escalated_${filteredRows[currentIndex].data[0]}.xlsx`);
}

/* ================= HELPERS ================= */

function download(buf,name){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([buf]));
  a.download=name;
  a.click();
}

function closeModal(){ escalationModal.style.display="none"; }
function saveEscalationNotes(){
  escalationNotes[currentEscalationKey]={
    issue:noteIssue.value,
    rootCause:noteRootCause.value,
    remarks:noteRemarks.value
  };
  closeModal();
}
