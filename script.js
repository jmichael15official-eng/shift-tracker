/* ==========================
   DATA
========================== */
const allTasks = [
  { name: "Database Check", shift: "1st", status: null },
  { name: "API Health", shift: "1st", status: null },
  { name: "Batch Jobs", shift: "2nd", status: null },
  { name: "Disk Usage", shift: "3rd", status: null }
];

let currentShift = "all";
let selectedTaskIndex = null;

/* ==========================
   TIME
========================== */
function updateTime() {
  document.getElementById("currentTimes").innerText =
    new Date().toLocaleString();
}
setInterval(updateTime, 1000);
updateTime();

/* ==========================
   RENDER
========================== */
function renderTasks() {
  const output = document.getElementById("output");
  output.innerHTML = "";

  allTasks
    .filter(t => currentShift === "all" || t.shift === currentShift)
    .forEach((task, index) => {
      const div = document.createElement("div");
      div.className = "task";

      div.innerHTML = `
        <strong>${task.name} (${task.shift} Shift)</strong>

        <div class="task-buttons">
          <button class="${task.status === 'good' ? 'active-good' : ''}"
            onclick="setStatus(${index}, 'good')">
            <i class="fa-solid fa-check"></i> Good
          </button>

          <button class="${task.status === 'monitor' ? 'active-monitor' : ''}"
            onclick="setStatus(${index}, 'monitor')">
            <i class="fa-solid fa-eye"></i> Monitor
          </button>

          <button class="${task.status === 'escalate' ? 'active-escalate' : ''}"
            onclick="openEscalation(${index})">
            <i class="fa-solid fa-triangle-exclamation"></i> Escalate
          </button>
        </div>
      `;
      output.appendChild(div);
    });
}

renderTasks();

/* ==========================
   STATUS
========================== */
function setStatus(index, status) {
  allTasks[index].status = status;
  renderTasks();
}

/* ==========================
   SHIFT FILTER
========================== */
function setShift(shift) {
  currentShift = shift;
  document.querySelectorAll("#shiftFilter button")
    .forEach(btn => btn.classList.remove("active-filter"));

  event.target.classList.add("active-filter");
  renderTasks();
}

/* ==========================
   MODAL
========================== */
function openEscalation(index) {
  selectedTaskIndex = index;
  document.getElementById("escalationModal").style.display = "block";
}

function closeModal() {
  document.getElementById("escalationModal").style.display = "none";
}

function saveEscalationNotes() {
  allTasks[selectedTaskIndex].status = "escalate";
  closeModal();
  renderTasks();
}

/* ==========================
   RESET
========================== */
function resetAllStatuses() {
  if (!confirm("Reset all statuses?")) return;
  allTasks.forEach(t => t.status = null);
  renderTasks();
}

/* ==========================
   EXPORT (FIXED)
========================== */
async function exportExcel(isEscalated) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Shift Report");

  ws.addRow(["Task", "Shift", "Status"]);

  allTasks.forEach(t => {
    const row = ws.addRow([t.name, t.shift, t.status || "N/A"]);

    const cell = row.getCell(3);
    if (t.status === "good") cell.fill = fill("00C853");
    if (t.status === "monitor") cell.fill = fill("FFD600");
    if (t.status === "escalate") cell.fill = fill("D50000");
  });

  const date = new Date().toLocaleDateString("en-US", {
    month: "long", day: "2-digit", year: "numeric"
  });

  const shiftName = currentShift === "all" ? "All_Shifts" : `${currentShift}_Shift`;
  const fileName = isEscalated
    ? `${date}_Escalated_${shiftName}.xlsx`
    : `${date}_${shiftName}.xlsx`;

  const buffer = await wb.xlsx.writeBuffer();
  download(buffer, fileName);
}

function fill(color) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: color } };
}

function download(buffer, name) {
  const blob = new Blob([buffer]);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
}
