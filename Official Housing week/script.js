const fileInput = document.getElementById('fileInput');
const selectedFileContainer = document.getElementById('selectedFile');
const filterTypeInput = document.getElementById('filterType');
const dateField = document.getElementById('dateField');
const dateInput = document.getElementById('dateInput');
const processButton = document.getElementById('processButton');
const downloadButton = document.getElementById('downloadButton');
const statusText = document.getElementById('status');
const resultTableBody = document.querySelector('#resultTable tbody');
const valueHeader = document.getElementById('valueHeader');

let selectedFile = null;
let processedSheets = [];
let outputWorkbook = null;
let outputLabel = '';

function showStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? '#ba1a1a' : '#445066';
}

function renderSelectedFile() {
  selectedFileContainer.innerHTML = '<div class="selected-file-title">الملف المختار</div>';

  if (!selectedFile) {
    const empty = document.createElement('div');
    empty.className = 'empty-file';
    empty.textContent = 'لم يتم اختيار أي ملف بعد.';
    selectedFileContainer.appendChild(empty);
    return;
  }

  const item = document.createElement('div');
  item.className = 'file-item';

  const info = document.createElement('span');
  info.className = 'file-info';
  info.textContent = `${selectedFile.name} (${(selectedFile.size / 1024).toFixed(1)} KB)`;

  const removeButton = document.createElement('button');
  removeButton.className = 'remove-file-btn';
  removeButton.type = 'button';
  removeButton.textContent = 'حذف';
  removeButton.addEventListener('click', () => {
    selectedFile = null;
    fileInput.value = '';
    resetResults();
    renderSelectedFile();
    showStatus('تم حذف الملف المختار.');
  });

  item.appendChild(info);
  item.appendChild(removeButton);
  selectedFileContainer.appendChild(item);
}

function resetResults() {
  processedSheets = [];
  outputWorkbook = null;
  downloadButton.disabled = true;
  valueHeader.textContent = 'القيمة';
  resultTableBody.innerHTML = '';
}

function toggleDateInput() {
  dateField.style.display = filterTypeInput.value === 'date' ? 'flex' : 'none';
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  if (!file) return;

  selectedFile = file;
  resetResults();
  renderSelectedFile();
  showStatus('تم اختيار الملف. يمكنك الآن تحديد نوع التصفية.');
});

filterTypeInput.addEventListener('change', () => {
  toggleDateInput();
  resetResults();
});

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number' && value > 31) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y) {
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
    }
  }

  const text = String(value ?? '')
    .trim()
    .replace(/\u061c|\u200e|\u200f/g, '')
    .replace(/[٠-٩]/g, (digit) => '٠١٢٣٤٥٦٧٨٩'.indexOf(digit))
    .replace(/[۰-۹]/g, (digit) => '۰۱۲۳۴۵۶۷۸۹'.indexOf(digit))
    .replace(/[／⁄]/g, '/')
    .replace(/\./g, '-');
  if (!text || /^[1-4]$/.test(text)) return null;

  const partsMatch = text.match(/^(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})$/);
  if (partsMatch) {
    const first = Number(partsMatch[1]);
    const second = Number(partsMatch[2]);
    const third = Number(partsMatch[3]);
    const fullYear = third < 100 ? 2000 + third : third;

    if (partsMatch[1].length === 4) {
      return new Date(first, second - 1, third);
    }

    if (second > 12) {
      return new Date(fullYear, first - 1, second);
    }

    if (first > 12 || partsMatch[3].length === 4) {
      return new Date(fullYear, second - 1, first);
    }
  }

  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeDate(value) {
  const date = parseExcelDate(value);
  if (!date) return null;

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getMonthNumber(value) {
  const text = normalizeText(value)
    .replace(/[٠-٩]/g, (digit) => '٠١٢٣٤٥٦٧٨٩'.indexOf(digit))
    .replace(/[۰-۹]/g, (digit) => '۰۱۲۳۴۵۶۷۸۹'.indexOf(digit))
    .replace(/[اأإآ]ل/g, 'ال');

  if (text === '44') return '4';
  if (!text.includes('شهر') && !text.includes('month')) return null;
  if (text.includes('الرابع') || text.includes('fourth') || text.includes('four')) return '4';
  if (/(^|[^\d])4([^\d]|$)/.test(text)) return '4';

  return null;
}

function isDateLike(value) {
  return Boolean(parseExcelDate(value));
}

function getDateInputValue() {
  return dateInput.value || dateInput.valueAsDate || dateInput.getAttribute('value') || '';
}

function normalizeText(value) {
  return String(value ?? '')
    .trim()
    .replace(/\u200f/g, '')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function getWeekNumber(value) {
  const text = normalizeText(value)
    .replace(/[٠-٩]/g, (digit) => '٠١٢٣٤٥٦٧٨٩'.indexOf(digit))
    .replace(/[۰-۹]/g, (digit) => '۰۱۲۳۴۵۶۷۸۹'.indexOf(digit))
    .replace(/[／⁄]/g, '/')
    .replace(/[اأإآ]ل/g, 'ال')
    .replace(/first|one|الاول|الأول/g, '1')
    .replace(/second|two|الثاني/g, '2')
    .replace(/third|three|الثالث/g, '3')
    .replace(/fourth|four|الرابع/g, '4');

  if (/^\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}$/.test(text)) {
    return null;
  }

  const match = text.match(/[1-4]/);
  return match ? match[0] : null;
}

function getComparableValue(value) {
  const date = normalizeDate(value);
  if (date) return date;

  const week = getWeekNumber(value);
  if (week) return week;

  return normalizeText(value);
}

function periodHeaderMatchesFilter(header, filter) {
  if (filter.type === 'date') {
    const normalizedTarget = normalizeDate(filter.key);
    if (!normalizedTarget) return false;
    return normalizeDate(header.label) === normalizedTarget || normalizeDate(header.key) === normalizedTarget;
  }

  if (filter.type === 'month') {
    return getMonthNumber(header.label) === filter.key || getMonthNumber(header.key) === filter.key;
  }

  return header.key === filter.key;
}

function periodValueMatchesFilter(value, filter) {
  if (filter.type === 'date') {
    const normalizedTarget = normalizeDate(filter.key);
    if (!normalizedTarget) return false;
    return normalizeDate(value) === normalizedTarget;
  }

  if (filter.type === 'month') {
    return getMonthNumber(value) === filter.key;
  }

  return getComparableValue(value) === filter.key;
}

function isPeriodHeader(value) {
  return Boolean(normalizeDate(value) || getWeekNumber(value));
}

function isMatchingHeaderForFilter(value, filterType) {
  if (filterType === 'date') {
    return isDateLike(value);
  }

  if (filterType === 'month') {
    return getMonthNumber(value) === '4';
  }

  if (filterType === 'week') {
    return Boolean(getWeekNumber(value));
  }

  return isPeriodHeader(value);
}

function isValidValueCell(value) {
  if (value === null || value === undefined) return false;

  const trimmed = String(value).trim();
  if (!trimmed) return false;

  if (typeof value === 'number') return value > 0;

  const percentMatch = trimmed.match(/^(-?\d+(?:\.\d+)?)%$/);
  if (percentMatch) return parseFloat(percentMatch[1]) > 0;

  if (!Number.isNaN(Number(trimmed))) return Number(trimmed) > 0;

  return true;
}

function findColumns(headers) {
  const taskHeader = headers.find((header) => {
    const value = normalizeText(header);
    return value === 'task' || value === 'المهمة' || value.includes('task') || value.includes('المهمة');
  });

  const dateHeader = headers.find((header) => {
    const value = normalizeText(header);
    return value === 'date' || value === 'week' || value.includes('date') || value.includes('تاريخ') || value.includes('اسبوع') || value.includes('أسبوع');
  });

  return { taskHeader, dateHeader };
}

function findTaskColumn(rows) {
  const maxRows = Math.min(rows.length, 5);

  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const value = normalizeText(row[colIndex]);
      if (value === 'task' || value === 'المهمة' || value.includes('task') || value.includes('المهمة')) {
        return colIndex;
      }
    }
  }

  return 0;
}

function findPeriodHeaderRow(rows, taskColIndex, filterType = 'all') {
  const maxRows = Math.min(rows.length, 5);

  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    let periodCount = 0;

    for (let colIndex = taskColIndex + 1; colIndex < row.length; colIndex += 1) {
      if (isMatchingHeaderForFilter(row[colIndex], filterType)) periodCount += 1;
    }

    if (periodCount >= (filterType === 'month' ? 1 : 2)) return rowIndex;
  }

  return null;
}

function prepareWideRows(rows, filterType = 'all') {
  const taskColIndex = findTaskColumn(rows);
  const headerRowIndex = findPeriodHeaderRow(rows, taskColIndex, filterType);

  if (headerRowIndex === null) return null;

  const headerRow = rows[headerRowIndex] || [];
  const periodHeaders = [];

  for (let colIndex = taskColIndex + 1; colIndex < headerRow.length; colIndex += 1) {
    if (isMatchingHeaderForFilter(headerRow[colIndex], filterType)) {
      periodHeaders.push({
        index: colIndex,
        label: String(headerRow[colIndex] ?? '').trim(),
        key: getComparableValue(headerRow[colIndex]),
      });
    }
  }

  if (!periodHeaders.length) return null;

  const parsedRows = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const task = row[taskColIndex];
    if (task === '' || task === undefined) continue;

    const values = {};
    periodHeaders.forEach((header) => {
      values[header.key] = row[header.index] || '';
    });

    parsedRows.push({ task, values });
  }

  return { mode: 'wide', rows: parsedRows, periodHeaders };
}

function prepareLongRows(rows) {
  const maxRows = Math.min(rows.length, 5);

  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const headers = row.map((cell) => String(cell ?? '').trim());
    const { taskHeader, dateHeader } = findColumns(headers);

    if (!taskHeader || !dateHeader) continue;

    const taskColIndex = headers.findIndex((header) => header === taskHeader);
    const dateColIndex = headers.findIndex((header) => header === dateHeader);
    const valueColIndex = headers.findIndex((header) => {
      const value = normalizeText(header);
      return value.includes('value') || value.includes('number') || value.includes('count') || value.includes('القيمة') || value.includes('العدد');
    });

    const parsedRows = [];

    for (let index = rowIndex + 1; index < rows.length; index += 1) {
      const sourceRow = rows[index] || [];
      const task = sourceRow[taskColIndex];
      const period = sourceRow[dateColIndex];
      const value = valueColIndex === -1 ? period : sourceRow[valueColIndex];

      if ((task === '' || task === undefined) && (period === '' || period === undefined)) continue;
      parsedRows.push({ task, period, value });
    }

    return { mode: 'long', rows: parsedRows, valueHeader: valueColIndex === -1 ? dateHeader : headers[valueColIndex] };
  }

  return null;
}

function prepareRows(rows, filterType = 'all') {
  if (!Array.isArray(rows) || rows.length === 0) return null;
  return prepareWideRows(rows, filterType) || prepareLongRows(rows);
}

function getSelectedFilter() {
  const filterType = filterTypeInput.value;

  if (filterType === 'date') {
    const selectedDate = getDateInputValue();
    return {
      type: 'date',
      key: normalizeDate(selectedDate),
      label: normalizeDate(selectedDate) || String(selectedDate || ''),
      header: normalizeDate(selectedDate) || String(selectedDate || 'التاريخ'),
    };
  }

  if (filterType === 'all') {
    return { type: 'all', key: 'all', label: 'كل البيانات', header: 'القيمة' };
  }

  if (filterType === '44') {
    return { type: 'month', key: '4', label: 'الشهر الرابع كامل', header: 'الشهر الرابع كامل' };
  }

  const week = filterType.replace('week', '');
  return { type: 'week', key: week, label: `الأسبوع ${week}`, header: `الأسبوع ${week}` };
}

function filterPreparedRows(prepared, filter) {
  if (!prepared) return { rows: [], header: filter.header };

  if (prepared.mode === 'wide') {
    if (filter.type === 'all') {
      const allRows = [];
      prepared.rows.forEach((row) => {
        prepared.periodHeaders.forEach((header) => {
          const value = row.values[header.key];
          if (isValidValueCell(value)) {
            const proofread = proofreadTask(row.task);
            allRows.push({
              task: proofread.task,
              originalTask: proofread.originalTask,
              taskChanged: proofread.taskChanged,
              value,
              period: header.label,
            });
          }
        });
      });
      return { rows: allRows, header: 'القيمة' };
    }

    if (filter.type === 'month') {
      const monthHeaders = prepared.periodHeaders.filter((header) => periodHeaderMatchesFilter(header, filter));
      const monthRows = [];

      prepared.rows.forEach((row) => {
        monthHeaders.forEach((header) => {
          const value = row.values[header.key];
          if (isValidValueCell(value)) {
            const proofread = proofreadTask(row.task);
            monthRows.push({
              task: proofread.task,
              originalTask: proofread.originalTask,
              taskChanged: proofread.taskChanged,
              value,
              period: header.label,
            });
          }
        });
      });

      return { rows: monthRows, header: filter.header };
    }

    const selectedHeader = prepared.periodHeaders.find((header) => periodHeaderMatchesFilter(header, filter));
    if (!selectedHeader) return { rows: [], header: filter.header };

    const rows = prepared.rows
      .filter((row) => isValidValueCell(row.values[selectedHeader.key]))
      .map((row) => {
        const proofread = proofreadTask(row.task);
        return {
          task: proofread.task,
          originalTask: proofread.originalTask,
          taskChanged: proofread.taskChanged,
          value: row.values[selectedHeader.key],
          period: selectedHeader.label,
        };
      });

    return { rows, header: selectedHeader.label || filter.header };
  }

  if (filter.type === 'all') {
    return {
      rows: prepared.rows
        .filter((row) => isValidValueCell(row.value))
        .map((row) => {
          const proofread = proofreadTask(row.task);
          return {
            ...row,
            task: proofread.task,
            originalTask: proofread.originalTask,
            taskChanged: proofread.taskChanged,
          };
        }),
      header: prepared.valueHeader || 'القيمة',
    };
  }

  const rows = prepared.rows
    .filter((row) => periodValueMatchesFilter(row.period, filter) && isValidValueCell(row.value))
    .map((row) => {
      const proofread = proofreadTask(row.task);
      return {
        task: proofread.task,
        originalTask: proofread.originalTask,
        taskChanged: proofread.taskChanged,
        value: row.value,
        period: row.period,
      };
    });

  return { rows, header: prepared.valueHeader || filter.header };
}

function normalizeOriginalTaskForCompare(value) {
  return String(value ?? '')
    .replace(/\u0640/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeTaskSpacing(value) {
  return String(value ?? '')
    .replace(/\u0640/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\s+([،؛:,.!?])/g, '$1')
    .replace(/([،؛:,.!?])(?=\S)/g, '$1 ')
    .replace(/\s*-\s*/g, ' - ')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/[.。]+$/g, '')
    .trim();
}

function proofreadTask(task) {
  const originalTask = String(task ?? '');
  const correctedTask = normalizeTaskSpacing(originalTask);

  return {
    task: correctedTask,
    originalTask,
    taskChanged: correctedTask !== normalizeOriginalTaskForCompare(originalTask),
  };
}

function sortRowsDescending(rows) {
  return rows.sort((a, b) => {
    const left = getValueCategory(a.value);
    const right = getValueCategory(b.value);

    if (left.type !== right.type) return left.type - right.type;
    return right.value - left.value;
  });
}

function getValueCategory(value) {
  const text = String(value ?? '').trim();
  if (text.endsWith('%')) return { type: 2, value: parseFloat(text.replace('%', '')) || 0 };
  if (!Number.isNaN(Number(text)) && text !== '') return { type: 1, value: Number(text) };
  return { type: 3, value: 0 };
}

function addNumericPercentSeparatorRows(rows) {
  const output = [];

  rows.forEach((row, index) => {
    if (index > 0) {
      const previousType = getValueCategory(rows[index - 1].value).type;
      const currentType = getValueCategory(row.value).type;

      if (previousType === 1 && currentType === 2) {
        output.push({ task: '', period: '', value: '' }, { task: '', period: '', value: '' });
      }
    }

    output.push(row);
  });

  return output;
}

function normalizeArabicSheetName(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[\u064B-\u065F\u0670]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\u0640/g, '')
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function isHousingBranchSheet(sheetName) {
  return normalizeArabicSheetName(sheetName).includes(normalizeArabicSheetName('فرع السكن'));
}

function getSafeSheetName(name, usedNames) {
  const cleaned = String(name || 'Sheet')
    .replace(/[:\\/?*\[\]]/g, '')
    .trim()
    .substring(0, 31) || 'Sheet';

  let sheetName = cleaned;
  let suffix = 1;

  while (usedNames.has(sheetName)) {
    const tail = `-${suffix}`;
    sheetName = `${cleaned.substring(0, 31 - tail.length)}${tail}`;
    suffix += 1;
  }

  usedNames.add(sheetName);
  return sheetName;
}

function buildWorkbook(sheets, filter) {
  const workbook = XLSX.utils.book_new();
  const usedNames = new Set();

  sheets.forEach((sheet) => {
    const sortedRows = sortRowsDescending([...sheet.rows]);
    const exportRows = addNumericPercentSeparatorRows(sortedRows);
    const headerRow = filter.type === 'all' ? ['المهمة', 'الفترة', sheet.header] : ['المهمة', sheet.header];
    const rowWidth = headerRow.length;
    const cardRows = isHousingBranchSheet(sheet.name)
      ? [
          ['مهام ال card', ...Array(rowWidth - 1).fill('')],
          Array(rowWidth).fill(''),
          Array(rowWidth).fill(''),
        ]
      : [];
    const sheetData = [
      ...cardRows,
      headerRow,
      ...exportRows.map((row) => (
        filter.type === 'all'
          ? [row.task || '', row.period || '', row.value ?? '']
          : [row.task || '', row.value ?? '']
      )),
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    if (cardRows.length && worksheet['A1']) {
      worksheet['A1'].s = {
        fill: { fgColor: { rgb: 'D9EAD3' } },
        font: { bold: true },
        alignment: { horizontal: 'center' },
      };
    }
    exportRows.forEach((row, index) => {
      if (!row.taskChanged) return;

      const cellAddress = XLSX.utils.encode_cell({ r: cardRows.length + 1 + index, c: 0 });
      worksheet[cellAddress] = worksheet[cellAddress] || { t: 's', v: row.task || '' };
      worksheet[cellAddress].s = {
        fill: {
          patternType: 'solid',
          fgColor: { rgb: 'FFFF00' },
          bgColor: { rgb: 'FFFF00' },
        },
        font: { color: { rgb: '000000' } },
      };
    });
    worksheet['!sheetViews'] = [{ RTL: true }];
    worksheet['!cols'] = filter.type === 'all'
      ? [{ wch: 55 }, { wch: 18 }, { wch: 18 }]
      : [{ wch: 55 }, { wch: 18 }];

    XLSX.utils.book_append_sheet(workbook, worksheet, getSafeSheetName(sheet.name, usedNames));
  });

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = [{ RTL: true }];
  return workbook;
}

function renderPreview(sheets, header) {
  resultTableBody.innerHTML = '';
  valueHeader.textContent = header || 'القيمة';

  const rows = sheets.flatMap((sheet) => sheet.rows.slice(0, 20).map((row) => ({ ...row, sheet: sheet.name })));

  if (!rows.length) {
    const emptyRow = document.createElement('tr');
    emptyRow.innerHTML = '<td colspan="3">لا توجد نتائج مطابقة.</td>';
    resultTableBody.appendChild(emptyRow);
    return;
  }

  rows.slice(0, 80).forEach((row) => {
    const tr = document.createElement('tr');
    [row.sheet, row.task, row.value].forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value ?? '';
      tr.appendChild(td);
    });
    resultTableBody.appendChild(tr);
  });
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('فشل قراءة الملف.'));
    reader.readAsArrayBuffer(file);
  });
}

processButton.addEventListener('click', async () => {
  if (!selectedFile) {
    showStatus('الرجاء اختيار ملف Excel أولاً.', true);
    return;
  }

  const filter = getSelectedFilter();
  if (filter.type === 'date' && !filter.key) {
    showStatus('الرجاء اختيار تاريخ صالح قبل المتابعة.', true);
    return;
  }

  resetResults();
  showStatus('جارٍ معالجة الشيتات...');

  try {
    const workbook = await parseWorkbook(selectedFile);
    const sheets = workbook.SheetNames.map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
      const prepared = prepareRows(rows, filter.type);
      const filtered = filterPreparedRows(prepared, filter);
      return {
        name: sheetName,
        rows: filtered.rows,
        header: filtered.header,
      };
    });

    processedSheets = sheets;
    outputWorkbook = buildWorkbook(processedSheets, filter);
    outputLabel = filter.label;

    const totalRows = processedSheets.reduce((sum, sheet) => sum + sheet.rows.length, 0);
    const firstHeader = processedSheets.find((sheet) => sheet.rows.length)?.header || filter.header;
    renderPreview(processedSheets, firstHeader);

    if (!totalRows) {
      showStatus('لم يتم العثور على بيانات مطابقة، لكن يمكن تنزيل ملف بنفس الشيتات فارغ النتائج.');
    } else {
      showStatus(`تمت معالجة ${processedSheets.length} شيت، وعدد النتائج المطابقة ${totalRows}.`);
    }

    downloadButton.disabled = false;
  } catch (error) {
    console.error(error);
    showStatus('حدث خطأ أثناء قراءة الملف. تأكد من أن ملف Excel صالح.', true);
  }
});

downloadButton.addEventListener('click', () => {
  if (!outputWorkbook) {
    showStatus('لا توجد نتيجة جاهزة للتنزيل.', true);
    return;
  }

  const baseName = selectedFile.name.replace(/\.(xlsx|xls)$/i, '') || 'السكن-الأسبوعي';
  const safeLabel = outputLabel.replace(/[^\w\u0600-\u06FF-]+/g, '-');
  XLSX.writeFile(outputWorkbook, `${baseName}-${safeLabel}.xlsx`);
});

toggleDateInput();
renderSelectedFile();
