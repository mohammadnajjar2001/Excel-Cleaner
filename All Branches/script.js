const fileInput = document.getElementById('fileInput');
const dateInput = document.getElementById('dateInput');
const cleanButton = document.getElementById('cleanButton');
const downloadButton = document.getElementById('downloadButton');
const statusText = document.getElementById('status');
const resultTableBody = document.querySelector('#resultTable tbody');
const resultDateHeader = document.querySelector('#resultTable thead th:nth-child(2)');

let currentRows = [];
let filteredRows = [];
let previewRows = [];
let workbookSheets = [];
let exportHeaders = ['Task', 'Date'];
let originalFileName = 'Excel-Cleaner-Filtered.xlsx';
let lastSelectedDate = '';

function showStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? '#ba1a1a' : '#445066';
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y) {
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
    }
  }

  const normalized = String(value).trim().replace(/\./g, '-').replace(/\u200f/g, '');
  const date = new Date(normalized);
  if (!Number.isNaN(date.getTime())) {
    return date;
  }

  return null;
}

function normalizeDate(value) {
  const date = parseExcelDate(value);
  if (!date) return null;
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function isDateLike(value) {
  return Boolean(parseExcelDate(value));
}

function findColumns(headers) {
  const taskHeader = headers.find((header) => {
    if (!header) return false;
    const normalized = String(header).trim().toLowerCase();
    return normalized === 'task' || normalized === 'المهمة' || normalized.includes('task') || normalized.includes('المهمة');
  });

  let dateHeader = headers.find((header) => {
    if (!header) return false;
    const normalized = String(header).trim().toLowerCase();
    return normalized === 'date' || normalized === 'تاريخ' || normalized.includes('date') || normalized.includes('تاريخ');
  });

  if (!dateHeader) {
    dateHeader = headers.find((header) => {
      if (!header) return false;
      return /date|تاريخ/i.test(String(header));
    });
  }

  return { taskHeader, dateHeader };
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = event.target.result;
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '',
          raw: false,
        });
        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('فشل قراءة الملف.'));
    reader.readAsArrayBuffer(file);
  });
}

function findTaskColumn(rows) {
  const maxRows = Math.min(rows.length, 4);

  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const value = String(row[colIndex] || '').trim().toLowerCase();
      if (value === 'task' || value === 'المهمة' || value.includes('task') || value.includes('المهمة')) {
        return colIndex;
      }
    }
  }

  return 0;
}

function findDateHeaderRow(rows, taskColIndex) {
  const maxRows = Math.min(rows.length, 4);

  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    let dateCount = 0;

    for (let colIndex = taskColIndex + 1; colIndex < row.length; colIndex += 1) {
      if (isDateLike(row[colIndex])) {
        dateCount += 1;
      }
    }

    if (dateCount >= 2) {
      return rowIndex;
    }
  }

  return null;
}

function prepareWideRows(rows) {
  const taskColIndex = findTaskColumn(rows);
  const headerRowIndex = findDateHeaderRow(rows, taskColIndex);

  if (headerRowIndex === null) {
    return null;
  }

  const headerRow = rows[headerRowIndex] || [];
  const dateHeaders = [];

  for (let colIndex = taskColIndex + 1; colIndex < headerRow.length; colIndex += 1) {
    if (isDateLike(headerRow[colIndex])) {
      const label = normalizeDate(headerRow[colIndex]) || String(headerRow[colIndex] || '');
      dateHeaders.push({ index: colIndex, label });
    }
  }

  if (!dateHeaders.length) {
    return null;
  }

  const parsedRows = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const task = row[taskColIndex];
    if (task === '' || task === undefined) {
      continue;
    }

    const dateValues = {};
    dateHeaders.forEach((dateHeader) => {
      dateValues[dateHeader.label] = row[dateHeader.index] || '';
    });

    parsedRows.push({ task, dateValues });
  }

  return {
    mode: 'wide',
    rows: parsedRows,
    dateHeaders: dateHeaders.map((header) => header.label),
  };
}

function findHeaderRowForLongRows(rows) {
  const maxRows = Math.min(rows.length, 4);

  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const headers = row.map((cell) => String(cell || '').trim());
    const { taskHeader, dateHeader } = findColumns(headers);
    if (taskHeader && dateHeader) {
      return { rowIndex, headers };
    }
  }

  return null;
}

function prepareLongRows(rows) {
  const headerInfo = findHeaderRowForLongRows(rows);
  if (!headerInfo) {
    return null;
  }

  const { rowIndex, headers } = headerInfo;
  const taskColIndex = headers.findIndex((header) => header === findColumns(headers).taskHeader);
  const dateColIndex = headers.findIndex((header) => header === findColumns(headers).dateHeader);

  if (taskColIndex === -1 || dateColIndex === -1) {
    return null;
  }

  const parsedRows = [];

  for (let index = rowIndex + 1; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const task = row[taskColIndex];
    const date = row[dateColIndex];
    if ((task === '' || task === undefined) && (date === '' || date === undefined)) {
      continue;
    }

    parsedRows.push({ task, date });
  }

  return {
    mode: 'long',
    rows: parsedRows,
  };
}

function prepareRows(data) {
  if (!Array.isArray(data) || data.length === 0) {
    return [];
  }

  const wideResult = prepareWideRows(data);
  if (wideResult) {
    return wideResult;
  }

  const longResult = prepareLongRows(data);
  return longResult;
}

function isValidValueCell(value) {
  if (value === null || value === undefined) {
    return false;
  }

  const trimmed = String(value).trim();
  if (!trimmed) {
    return false;
  }

  if (typeof value === 'number') {
    return value > 0;
  }

  const percentMatch = trimmed.match(/^(-?\d+(?:\.\d+)?)%$/);
  if (percentMatch) {
    return parseFloat(percentMatch[1]) > 0;
  }

  if (!Number.isNaN(Number(trimmed))) {
    return Number(trimmed) > 0;
  }

  return true;
}

function filterLongRows(rows, dateValue) {
  const normalizedTarget = normalizeDate(dateValue);
  if (!normalizedTarget) return [];

  return rows
    .filter((row) => normalizeDate(row.date) === normalizedTarget && isValidValueCell(row.date))
    .map((row) => ({
      task: row.task,
      date: row.date,
    }));
}

function filterWideRows(rows, dateValue, dateHeaders) {
  const normalizedTarget = normalizeDate(dateValue);
  if (!normalizedTarget) return { filtered: [], selectedHeader: null };

  const selectedHeader = dateHeaders.find((header) => normalizeDate(header) === normalizedTarget);
  if (!selectedHeader) {
    return { filtered: [], selectedHeader: null };
  }

  const filtered = rows
    .filter((row) => isValidValueCell(row.dateValues[selectedHeader]))
    .map((row) => ({
      task: row.task,
      date: row.dateValues[selectedHeader],
    }));

  return { filtered, selectedHeader };
}

function updateTableHeader(label) {
  resultDateHeader.textContent = label || 'التاريخ';
}

function renderTable(rows) {
  resultTableBody.innerHTML = '';

  if (!rows.length) {
    const emptyRow = document.createElement('tr');
    emptyRow.innerHTML = '<td colspan="2">لا توجد نتائج مطابقة.</td>';
    resultTableBody.appendChild(emptyRow);
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    const taskCell = document.createElement('td');
    const dateCell = document.createElement('td');

    taskCell.textContent = row.task || '';
    dateCell.textContent = row.date === undefined ? '' : String(row.date);

    tr.appendChild(taskCell);
    tr.appendChild(dateCell);
    resultTableBody.appendChild(tr);
  });
}

function getSheetNameFromFile(fileName) {
  const baseName = fileName.replace(/\.xlsx$/i, '').replace(/[^\w\u0600-\u06FF \-]/g, '');
  const trimmed = baseName.trim();
  if (!trimmed) {
    return 'Sheet1';
  }
  return trimmed.substring(0, 31);
}

function buildWorkbookFromSheets(sheets) {
  const workbook = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    let sheetName = sheet.name;
    let suffix = 1;
    while (workbook.SheetNames.includes(sheetName)) {
      sheetName = `${sheet.name}-${suffix}`;
      suffix += 1;
    }

    const sortedRows = sortRowsDescending([...sheet.rows]);

    const sheetData = [
      sheet.headers,
      ...sortedRows.map((row) => [row.task || '', String(row.date || '')])
    ]; const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    worksheet['!sheetViews'] = [{ RTL: true }];
    worksheet['!cols'] = [{ wch: 60 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = workbook.Workbook.Views || [{ RTL: true }];
  return workbook;
}

function processFile(file, selectedDate) {
  return parseWorkbook(file).then((workbookRows) => {
    const prepared = prepareRows(workbookRows);
    const sheetName = getSheetNameFromFile(file.name);
    const sheet = {
      name: sheetName,
      headers: ['Task', 'Date'],
      rows: [],
    };

    if (!prepared) {
      return sheet;
    }

    if (prepared.mode === 'wide') {
      const { filtered, selectedHeader } = filterWideRows(prepared.rows, selectedDate, prepared.dateHeaders);
      sheet.headers = ['Task', selectedHeader || 'Date'];
      sheet.rows = filtered;
      return sheet;
    }

    if (prepared.mode === 'long') {
      sheet.rows = filterLongRows(prepared.rows, selectedDate);
      return sheet;
    }

    return sheet;
  });
}

function processFiles(files, selectedDate) {
  return Promise.all(files.map((file) => processFile(file, selectedDate)));
}

function buildWorkbook(rows, sheetName) {
  const sheetData = [exportHeaders, ...rows.map((row) => [row.task || '', String(row.date || '')])];
  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
  worksheet['!sheetViews'] = [{ RTL: true }];
  worksheet['!cols'] = [{ wch: 60 }, { wch: 20 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName || 'Cleaned');
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = workbook.Workbook.Views || [];
  workbook.Workbook.Views.push({ RTL: true });

  return workbook;
}
function sortRowsDescending(rows) {
  return rows.sort((a, b) => {
    const parseValue = (val) => {
      if (val === null || val === undefined) return { type: 3, value: 0 };

      const str = String(val).trim();

      // نسبة مئوية
      if (str.endsWith('%')) {
        return {
          type: 2,
          value: parseFloat(str.replace('%', '')) || 0,
        };
      }

      // رقم
      if (!isNaN(str)) {
        return {
          type: 1,
          value: parseFloat(str),
        };
      }

      // نص أو شيء آخر
      return {
        type: 3,
        value: 0,
      };
    };

    const aVal = parseValue(a.date);
    const bVal = parseValue(b.date);

    // أولاً حسب النوع
    if (aVal.type !== bVal.type) {
      return aVal.type - bVal.type; // 1 (أرقام) ثم 2 (%) ثم 3 (نص)
    }

    // ثم حسب القيمة داخل نفس النوع
    return bVal.value - aVal.value;
  });
}
function downloadWorkbook(workbook, fileName) {
  XLSX.writeFile(workbook, fileName);
}

cleanButton.addEventListener('click', async () => {
  const files = Array.from(fileInput.files);
  // const selectedDate = dateInput.value;
  const filterType = document.getElementById('filterType').value;

  let selectedDate = dateInput.value;

  // إذا اختار أسبوع
  if (filterType === 'week1') selectedDate = "1";
  if (filterType === 'week2') selectedDate = "2";
  if (filterType === 'week3') selectedDate = "3";
  if (filterType === 'week4') selectedDate = "4";

  if (!files.length) {
    showStatus('الرجاء اختيار ملف Excel على الأقل.', true);
    return;
  }

  if (!selectedDate) {
    showStatus('الرجاء اختيار تاريخ قبل المتابعة.', true);
    return;
  }

  lastSelectedDate = normalizeDate(selectedDate) || selectedDate;
  originalFileName = `Excel-Cleaner-${lastSelectedDate}.xlsx`;

  showStatus('جارٍ معالجة الملفات...');
  downloadButton.disabled = true;
  resultTableBody.innerHTML = '';
  updateTableHeader('التاريخ');

  try {
    workbookSheets = await processFiles(files, lastSelectedDate);
    filteredRows = workbookSheets.flatMap((sheet) => sheet.rows);
    filteredRows = sortRowsDescending(filteredRows);
    previewRows = [...filteredRows];

    if (!filteredRows.length) {
      renderTable(filteredRows);
      showStatus('لم يتم العثور على أي بيانات مطابقة في الملفات المحددة.');
      return;
    }

    const firstHeader = workbookSheets.find((sheet) => sheet.rows.length > 0)?.headers[1] || 'التاريخ';
    updateTableHeader(firstHeader);
    renderTable(filteredRows);

    showStatus(`تم تنظيف البيانات بنجاح من ${files.length} ملف${files.length === 1 ? '' : 'ات'}.`);
    downloadButton.disabled = false;
  } catch (error) {
    console.error(error);
    showStatus('حدث خطأ أثناء قراءة الملفات. تأكد من أن الملفات صالحة.', true);
  }
});

downloadButton.addEventListener('click', () => {
  if (!workbookSheets.length) {
    showStatus('لا توجد بيانات جاهزة للتنزيل.', true);
    return;
  }

  const workbook = buildWorkbookFromSheets(workbookSheets);
  const normalizedDate = normalizeDate(lastSelectedDate) || lastSelectedDate;
  const fileName = `الأفرع-كاملةً-${normalizedDate}.xlsx`;
  downloadWorkbook(workbook, fileName);
});

const filterType = document.getElementById('filterType');
const dateField = document.getElementById('dateField');

function toggleDateInput() {
  if (filterType.value === 'date') {
    dateField.style.display = 'block';
  } else {
    dateField.style.display = 'none';
  }
}

// تشغيل أول مرة عند تحميل الصفحة
toggleDateInput();

// عند تغيير الاختيار
filterType.addEventListener('change', toggleDateInput);