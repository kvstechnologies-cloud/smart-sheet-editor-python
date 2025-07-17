let workbookData = [];
let selectedSheetName = '';

const fileInput = document.getElementById('fileInput');
const sheetSelect = document.getElementById('sheetSelect');
const tableContainer = document.getElementById('tableContainer');
const downloadBtn = document.getElementById('downloadBtn');
const uploadBtn = document.getElementById('uploadBtn');
const controls = document.getElementById('controls');
const dropZone = document.getElementById('dropZone');

// ✅ Browse button triggers file input
document.getElementById('browseBtn').addEventListener('click', (e) => {
  e.preventDefault();
  fileInput.click();
});

// ✅ Upload button
uploadBtn.addEventListener('click', upload);

// ✅ Drag & drop support
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  fileInput.files = e.dataTransfer.files;
});

// ✅ Upload file to backend
async function upload() {
  if (!fileInput.files.length) return;
  const formData = new FormData();
  formData.append('file', fileInput.files[0]);

  const res = await fetch('/upload', {
    method: 'POST',
    body: formData
  });

  const json = await res.json();
  workbookData = json.sheets || [];
  sheetSelect.innerHTML = '';

  if (workbookData.length) {
    workbookData.forEach(sheet => {
      const opt = document.createElement('option');
      opt.value = sheet.name;
      opt.textContent = sheet.name;
      sheetSelect.appendChild(opt);
    });

    controls.style.display = 'block';
    selectedSheetName = workbookData[0].name;
    sheetSelect.value = selectedSheetName;
    displaySheet(workbookData[0]);
  }
}

sheetSelect.addEventListener('change', () => {
  selectedSheetName = sheetSelect.value;
  const sheet = workbookData.find(s => s.name === selectedSheetName);
  if (sheet) displaySheet(sheet);
});

function displaySheet(sheet) {
  tableContainer.innerHTML = '';
  if (!sheet.data.length) {
    tableContainer.textContent = 'No data in sheet';
    return;
  }

  const cols = Object.keys(sheet.data[0]);
  const table = document.createElement('table');

  const headerRow = document.createElement('tr');
  headerRow.appendChild(document.createElement('th')); // For row numbers

  cols.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    headerRow.appendChild(th);
  });

  table.appendChild(headerRow);

  sheet.data.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    const rowNum = document.createElement('td');
    rowNum.textContent = rowIndex + 1;
    tr.appendChild(rowNum);

    cols.forEach(col => {
      const td = document.createElement('td');
      const value = row[col];
      if (typeof value === 'number') {
        const input = document.createElement('input');
        input.type = 'number';
        input.value = value;
        input.dataset.row = rowIndex;
        input.dataset.col = col;
        input.addEventListener('change', (e) => {
          const r = parseInt(e.target.dataset.row, 10);
          const c = e.target.dataset.col;
          sheet.data[r][c] = parseFloat(e.target.value);
        });
        td.appendChild(input);
      } else {
        td.textContent = value ?? '';
      }
      tr.appendChild(td);
    });

    table.appendChild(tr);
  });

  tableContainer.appendChild(table);
}

downloadBtn.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  workbookData.forEach(sheet => {
    const ws = XLSX.utils.json_to_sheet(sheet.data);
    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });
  XLSX.writeFile(wb, 'edited.xlsx');
});
