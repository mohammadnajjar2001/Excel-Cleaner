const sharedDateInput = document.getElementById('sharedDate');
const affiliationDateInput = document.getElementById('affiliationDate');
const runAllButton = document.getElementById('runAllButton');
const statusBox = document.getElementById('status');

const tools = [
  {
    name: 'فرع الانتساب والتوظيف',
    frameId: 'affiliationFrame',
    dateInput: affiliationDateInput,
    processButtonId: 'cleanButton',
    keepsFilesInList: true,
  },
  {
    name: 'جميع الفروع',
    frameId: 'allBranchesFrame',
    dateInput: sharedDateInput,
    processButtonId: 'cleanButton',
    keepsFilesInList: true,
  },
  {
    name: 'السكن الوظيفي',
    frameId: 'housingFrame',
    dateInput: sharedDateInput,
    processButtonId: 'processButton',
  },
  {
    name: 'الدعم والرعاية',
    frameId: 'supportFrame',
    dateInput: sharedDateInput,
    processButtonId: 'processButton',
  },
];

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.style.background = isError ? '#fff1f2' : '#eef6f6';
  statusBox.style.color = isError ? '#9f1239' : '#385b5a';
}

function todayValue() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getFrameDocument(tool) {
  const frame = document.getElementById(tool.frameId);
  return frame?.contentDocument || frame?.contentWindow?.document || null;
}

function dispatchNativeChange(element) {
  element.dispatchEvent(new Event('input', { bubbles: true }));
  element.dispatchEvent(new Event('change', { bubbles: true }));
}

function setInputValue(frameWindow, element, value) {
  const prototype = frameWindow.HTMLInputElement.prototype;
  const descriptor = Object.getOwnPropertyDescriptor(prototype, 'value');

  if (descriptor?.set) {
    descriptor.set.call(element, value);
  } else {
    element.value = value;
  }

  element.setAttribute('value', value);
  dispatchNativeChange(element);
}

function setSelectValue(frameWindow, element, value) {
  const prototype = frameWindow.HTMLSelectElement.prototype;
  const descriptor = Object.getOwnPropertyDescriptor(prototype, 'value');

  if (descriptor?.set) {
    descriptor.set.call(element, value);
  } else {
    element.value = value;
  }

  dispatchNativeChange(element);
}

function wait(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function prepareTool(tool) {
  const frame = document.getElementById(tool.frameId);
  const doc = getFrameDocument(tool);
  if (!doc) {
    throw new Error(`لم يتم تحميل أداة ${tool.name} بعد.`);
  }

  const fileInput = doc.getElementById('fileInput');
  if (!tool.keepsFilesInList && !fileInput?.files?.length) {
    return { ready: false, reason: `لم يتم اختيار ملف في ${tool.name}.` };
  }

  const filterType = doc.getElementById('filterType');
  if (filterType) {
    setSelectValue(frame.contentWindow, filterType, 'date');
  }

  const dateInput = doc.getElementById('dateInput');
  const selectedDate = tool.dateInput.value;
  if (dateInput && selectedDate) {
    setInputValue(frame.contentWindow, dateInput, selectedDate);
  }

  const processButton = doc.getElementById(tool.processButtonId);
  const downloadButton = doc.getElementById('downloadButton');
  if (!processButton || !downloadButton) {
    throw new Error(`تعذر الوصول إلى أزرار ${tool.name}.`);
  }

  return { ready: true, processButton, downloadButton, dateInput, selectedDate };
}

function syncToolDate(tool) {
  const frame = document.getElementById(tool.frameId);
  const doc = getFrameDocument(tool);
  if (!frame?.contentWindow || !doc) return;

  const filterType = doc.getElementById('filterType');
  const dateInput = doc.getElementById('dateInput');
  const selectedDate = tool.dateInput.value;

  if (filterType) {
    setSelectValue(frame.contentWindow, filterType, 'date');
  }

  if (dateInput && selectedDate) {
    setInputValue(frame.contentWindow, dateInput, selectedDate);
  }
}

function syncAllDates() {
  tools.forEach(syncToolDate);
}

function waitForDownload(downloadButton, timeoutMs = 30000) {
  const startedAt = Date.now();

  return new Promise((resolve) => {
    const timer = window.setInterval(() => {
      if (!downloadButton.disabled) {
        window.clearInterval(timer);
        resolve(true);
        return;
      }

      if (Date.now() - startedAt > timeoutMs) {
        window.clearInterval(timer);
        resolve(false);
      }
    }, 250);
  });
}

function getXlsxApi() {
  for (const tool of tools) {
    const frameWindow = document.getElementById(tool.frameId)?.contentWindow;
    if (frameWindow?.XLSX) {
      return frameWindow.XLSX;
    }
  }

  return null;
}

function createCellStyle(XLSX, { fillColor, bold = false, fontSize = 14 } = {}) {
  const style = {
    font: {
      name: 'Arial',
      bold,
      sz: fontSize,
      color: { rgb: '000000' },
    },
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
      readingOrder: 2,
    },
    border: {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } },
    },
  };

  if (fillColor) {
    style.fill = {
      patternType: 'solid',
      fgColor: { rgb: fillColor },
      bgColor: { rgb: fillColor },
    };
  }

  return style;
}

function collectPrintSections() {
  const frameOrder = [
    'allBranchesFrame',
    'affiliationFrame',
    'housingFrame',
    'supportFrame',
  ];

  return frameOrder.flatMap((frameId) => {
    const frameWindow = document.getElementById(frameId)?.contentWindow;
    if (typeof frameWindow?.getUnifiedPrintSections !== 'function') {
      return [];
    }

    return frameWindow.getUnifiedPrintSections();
  });
}

function downloadPrintWorkbook() {
  const XLSX = getXlsxApi();
  const sections = collectPrintSections();

  if (!XLSX || !sections.length) {
    return false;
  }

  const rows = [];
  const merges = [];
  const rowTypes = [];

  sections.forEach((section, sectionIndex) => {
    const titleRowIndex = rows.length;
    rows.push([section.title, '']);
    rowTypes.push('title');
    merges.push({ s: { r: titleRowIndex, c: 0 }, e: { r: titleRowIndex, c: 1 } });

    rows.push(['المهمة', 'العدد/النسبة']);
    rowTypes.push('header');

    section.rows.forEach((row) => {
      rows.push([row.task || '', row.value ?? '']);
      rowTypes.push(row.taskChanged ? 'changed' : 'body');
    });

    if (sectionIndex < sections.length - 1) {
      rows.push(['', ''], ['', '']);
      rowTypes.push('blank', 'blank');
    }
  });

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const titleStyle = createCellStyle(XLSX, { fillColor: '9A9558', bold: true, fontSize: 16 });
  const headerStyle = createCellStyle(XLSX, { fillColor: '00B050', bold: true, fontSize: 14 });
  const bodyStyle = createCellStyle(XLSX, { fontSize: 14 });
  const changedStyle = createCellStyle(XLSX, { fillColor: 'FFFF00', fontSize: 14 });

  rows.forEach((row, rowIndex) => {
    for (let colIndex = 0; colIndex < 2; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      worksheet[address] = worksheet[address] || { t: 's', v: '' };

      if (rowTypes[rowIndex] === 'title') {
        worksheet[address].s = titleStyle;
      } else if (rowTypes[rowIndex] === 'header') {
        worksheet[address].s = headerStyle;
      } else if (rowTypes[rowIndex] === 'changed' && colIndex === 0) {
        worksheet[address].s = changedStyle;
      } else if (rowTypes[rowIndex] !== 'blank') {
        worksheet[address].s = bodyStyle;
      }
    }
  });

  worksheet['!merges'] = merges;
  worksheet['!sheetViews'] = [{ RTL: true }];
  worksheet['!cols'] = [{ wch: 72 }, { wch: 18 }];
  worksheet['!rows'] = rowTypes.map((type) => {
    if (type === 'title') return { hpt: 42 };
    if (type === 'header') return { hpt: 30 };
    if (type === 'blank') return { hpt: 14 };
    return { hpt: 27 };
  });

  XLSX.utils.book_append_sheet(workbook, worksheet, 'للطباعة');
  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = [{ RTL: true }];

  const selectedDateLabel = sharedDateInput.value.replace(/[\\/:*?"<>|]+/g, '-') || 'بدون تاريخ';
  XLSX.writeFile(workbook, `للطباعة (${selectedDateLabel}).xlsx`);
  return true;
}

async function runTool(tool) {
  const prepared = prepareTool(tool);
  if (!prepared.ready) {
    return { name: tool.name, status: 'skipped', message: prepared.reason };
  }

  if (prepared.dateInput && prepared.selectedDate && prepared.dateInput.value !== prepared.selectedDate) {
    return { name: tool.name, status: 'failed', message: `لم يتم تثبيت التاريخ في ${tool.name}.` };
  }

  await wait(150);
  prepared.processButton.click();
  const readyToDownload = await waitForDownload(prepared.downloadButton);
  if (!readyToDownload) {
    return { name: tool.name, status: 'failed', message: `لم تجهز نتيجة ${tool.name} خلال الوقت المحدد.` };
  }

  prepared.downloadButton.click();
  return { name: tool.name, status: 'done', message: `تم تنزيل ${tool.name}.` };
}

async function runAllTools() {
  if (!sharedDateInput.value) {
    setStatus('اختر التاريخ الموحد أولاً.', true);
    return;
  }

  if (!affiliationDateInput.value) {
    setStatus('اختر تاريخ الانتساب أولاً.', true);
    return;
  }

  runAllButton.disabled = true;
  setStatus('جارٍ تشغيل الأدوات الجاهزة...');

  const results = [];
  for (const tool of tools) {
    setStatus(`جارٍ معالجة ${tool.name}...`);
    try {
      results.push(await runTool(tool));
    } catch (error) {
      results.push({ name: tool.name, status: 'failed', message: error.message });
    }
  }

  const doneCount = results.filter((result) => result.status === 'done').length;
  const skipped = results.filter((result) => result.status === 'skipped').map((result) => result.name);
  const failed = results.filter((result) => result.status === 'failed').map((result) => result.name);
  const printDownloaded = downloadPrintWorkbook();

  const parts = [`تم تنزيل ${doneCount} ملف.`];
  if (printDownloaded) parts.push('وتم تنزيل ملف الطباعة.');
  if (skipped.length) parts.push(`لم يتم اختيار ملفات في: ${skipped.join('، ')}.`);
  if (failed.length) parts.push(`تعذر إكمال: ${failed.join('، ')}.`);

  setStatus(parts.join(' '), failed.length > 0);
  runAllButton.disabled = false;
}

const today = todayValue();
sharedDateInput.value = today;
affiliationDateInput.value = today;
tools.forEach((tool) => {
  document.getElementById(tool.frameId)?.addEventListener('load', () => syncToolDate(tool));
});
sharedDateInput.addEventListener('change', syncAllDates);
affiliationDateInput.addEventListener('change', syncAllDates);
runAllButton.addEventListener('click', runAllTools);
