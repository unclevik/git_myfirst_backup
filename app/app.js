const headers = [
  "Идентификатор курса",
  "Длинное наименование курса в системе",
  "Наименование на английском языке",
  "Короткое наименование учебного курса в системе",
  "Фамилия",
  "Имя и отчество (имена)",
  "Фамилия на английском языке",
  "Имя (имена)  на английском языке",
  "Логин",
  "Роль в курсе",
  "Дата регистрации",
];

const tableBody = document.getElementById("tableBody");
const addRowBtn = document.getElementById("addRowBtn");
const deleteSelectedBtn = document.getElementById("deleteSelectedBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const exportBtn = document.getElementById("exportBtn");
const selectAll = document.getElementById("selectAll");
const fileNameInput = document.getElementById("fileNameInput");
const rowCount = document.getElementById("rowCount");

function loadRowsFromStorage() {
  try {
    const raw = localStorage.getItem("course_rows");
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveRowsToStorage(rows) {
  try {
    localStorage.setItem("course_rows", JSON.stringify(rows));
  } catch (e) {
    // ignore
  }
}

function getRows() {
  const rows = [];
  for (const tr of tableBody.querySelectorAll("tr")) {
    const inputs = tr.querySelectorAll("input.cell-input");
    const row = {};
    headers.forEach((h, idx) => {
      row[h] = inputs[idx]?.value ?? "";
    });
    rows.push(row);
  }
  return rows;
}

function updateRowCount() {
  const count = tableBody.querySelectorAll("tr").length;
  rowCount.textContent = `${count} ${decl(count, ["запись", "записи", "записей"])}`;
}

function decl(n, forms) {
  const num = Math.abs(n) % 100;
  const n1 = num % 10;
  if (num > 10 && num < 20) return forms[2];
  if (n1 > 1 && n1 < 5) return forms[1];
  if (n1 === 1) return forms[0];
  return forms[2];
}

function addRow(prefill = {}) {
  const tr = document.createElement("tr");

  const selTd = document.createElement("td");
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.className = "row-select";
  selTd.appendChild(cb);
  tr.appendChild(selTd);

  headers.forEach((h) => {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.className = "cell-input";
    input.placeholder = h;
    input.value = prefill[h] || "";
    td.appendChild(input);
    tr.appendChild(td);
  });

  tableBody.appendChild(tr);
  updateRowCount();
}

function deleteSelected() {
  const selected = tableBody.querySelectorAll(".row-select:checked");
  selected.forEach((cb) => cb.closest("tr").remove());
  updateRowCount();
  persist();
}

function clearAll() {
  tableBody.innerHTML = "";
  updateRowCount();
  persist();
}

function persist() {
  saveRowsToStorage(getRows());
}

function exportToExcel() {
  const rows = getRows();
  const sheetData = [headers, ...rows.map((r) => headers.map((h) => r[h] || ""))];

  const ws = XLSX.utils.aoa_to_sheet(sheetData);

  // Try to coerce "Дата регистрации" to date format if looks like YYYY-MM-DD or DD.MM.YYYY
  const dateColIndex = headers.indexOf("Дата регистрации");
  if (dateColIndex >= 0) {
    for (let i = 1; i < sheetData.length; i++) {
      const value = sheetData[i][dateColIndex];
      const parsed = parseDate(value);
      if (parsed) {
        const cellRef = XLSX.utils.encode_cell({ r: i, c: dateColIndex });
        ws[cellRef] = { t: "d", v: parsed };
      }
    }
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; ++R) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: dateColIndex });
      if (ws[cellRef]) ws[cellRef].z = "yyyy-mm-dd";
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Курсы");
  const fileName = (fileNameInput.value || "courses").replace(/\.+$/, "");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

function parseDate(text) {
  if (!text) return null;
  const t = String(text).trim();
  // ISO or YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
    const d = new Date(t + "T00:00:00");
    return isNaN(d) ? null : d;
  }
  // DD.MM.YYYY
  const m = t.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (m) {
    const d = new Date(`${m[3]}-${m[2]}-${m[1]}T00:00:00`);
    return isNaN(d) ? null : d;
  }
  return null;
}

addRowBtn.addEventListener("click", () => {
  addRow();
});

deleteSelectedBtn.addEventListener("click", deleteSelected);
clearAllBtn.addEventListener("click", clearAll);
exportBtn.addEventListener("click", exportToExcel);

selectAll.addEventListener("change", (e) => {
  const checked = e.target.checked;
  tableBody.querySelectorAll(".row-select").forEach((cb) => (cb.checked = checked));
});

// Persist on input changes
tableBody.addEventListener("input", () => persist());

// Restore saved rows or add one empty row
const saved = loadRowsFromStorage();
if (saved.length > 0) {
  saved.forEach((r) => addRow(r));
} else {
  addRow();
}

updateRowCount();