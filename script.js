/* ============================
   GLOBAL STATE
============================ */
const tasks = [
  { name: "Database Check", status: null },
  { name: "API Health", status: null },
  { name: "Batch Jobs", status: null },
  { name: "Disk Usage", status: null }
];

let selectedTaskIndex = null;

/* ============================
   TIME DISPLAY
============================ */
function updateTime() {
  const now = new Date();
  document.getElementById("currentTimes").innerText =
    now.toLocaleString("en-US", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    });
}
setInterval(updateTime, 1000);
updateTime();

/* ============================
   UI RENDER
============================ */
function renderTasks() {
  const output = document.getElementById("output");
  output.innerHTML = "";

  tasks.forEach((task, index) => {
    const div = document.createElement("div");
    div.className = `task ${task.status === "escalate" ? "escalated" : ""}`;

    div.innerHTML = `
      <strong>${task.name}</strong>

      <div class="task-buttons">
        <button class="${task.status === "good" ? "active-good" : ""}"
                onclick="setStatus(${index}, 'good')">
          <i class="fa-solid fa-check"></i> Good
        </button>

        <button class="${task.status === "monitor" ? "active-monitor" : ""}"
                onclick="setStatus(${index}, 'monitor')">
          <i class="fa-solid fa-eye"></i> Monitor
        </button>

        <button class="${task.status === "escalate" ? "active-escalate" : ""}"
                onclick="openEscalation(${index})">
          <i class="fa-solid fa-triangle-exclamation"></i> Escalate
        </button>
      </div>
    `;
    output.appendChild(div);
  });
}

renderTasks();

/* ============================
   STATUS HANDLING
============================ */
function setStatus(index, status) {
  tasks[index].status = status;
  renderTasks();
}

function openEscalation(index) {
  selectedTaskIndex = index;
  document.getElementById("escalationModal").style.display = "block";
}

function closeModal() {
  document.getElementById("escalationModal").style.display = "none";
}

function saveEscalationNotes() {
  tasks[selectedTaskIndex].status = "escalate";
  closeModal();
  renderTasks();
}

/* ============================
   RESET WITH CONFIRMATION
============================ */
function resetAllStatuses() {
  if (!confirm("Are you sure you want to reset all statuses?")) return;

  tasks.forEach(t => t.status = null);
  renderTasks();
}

/* ============================
   FETCH EXCEL TEMPLATE (GITHUB)
============================ */
async function fetchTemplate() {
  const url =
    "https://raw.githubusercontent.com/YOUR_USERNAME/shift-tracker/main/data/Weekday_Monitoring%20Template.xlsx";

  const response = await fetch(url);
  const arrayBuffer = await response.arrayBuffer();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

/* ============================
   EXPORT TO EXCEL
============================ */
async function exportExcel(isEscalated) {
  const workbook = await fetchTemplate();
  const sheet = workbook.worksheets[0];

  let row = 2;

  tasks.forEach(task => {
    const excelRow = sheet.getRow(row);
    excelRow.getCell(1).value = task.name;
    excelRow.getCell(2).value = task.status || "N/A";

    if (task.status === "good") {
      excelRow.getCell(2).fill = fillColor("00C853");
    }
    if (task.status === "monitor") {
      excelRow.getCell(2).fill = fillColor("FFD600");
    }
    if (task.status === "escalate") {
      excelRow.getCell(2).fill = fillColor("D50000");
    }

    excelRow.commit();
    row++;
  });

  const date = new Date().toLocaleDateString("en-US", {
    month: "long",
    day: "2-digit",
    year: "numeric"
  });

  const shift = "1st Shift";
  const fileName = isEscalated
    ? `${date}_Escalated_${shift}.xlsx`
    : `${date}_${shift}.xlsx`;

  const buffer = await workbook.xlsx.writeBuffer();
  saveAsExcel(buffer, fileName);
}

/* ============================
   HELPERS
============================ */
function fillColor(hex) {
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: hex }
  };
}

function saveAsExcel(buffer, fileName) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
}
