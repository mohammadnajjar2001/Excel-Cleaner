// عناصر الواجهة
const fileInput = document.getElementById('fileInput');
const processBtn = document.getElementById('processBtn');
const statusText = document.getElementById('status');
const selectedFileContainer = document.getElementById('selectedFile');

let selectedFile = null;

// تحديث حالة الواجهة
function showStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? 'red' : '#2c3e50';
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
  removeButton.setAttribute('aria-label', `حذف ${selectedFile.name}`);
  removeButton.addEventListener('click', () => {
    selectedFile = null;
    fileInput.value = '';
    renderSelectedFile();
    showStatus('تم حذف الملف المختار. بانتظار رفع ملف جديد.');
  });

  item.appendChild(info);
  item.appendChild(removeButton);
  selectedFileContainer.appendChild(item);
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  if (!file) return;

  selectedFile = file;
  renderSelectedFile();
  showStatus('تم اختيار الملف. يمكنك استبداله باختيار ملف آخر أو حذفه.');
});

// قراءة ملف Excel وتحويله إلى صفوف
function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
          raw: false
        });

        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };

    reader.readAsArrayBuffer(file);
  });
}

// استخراج المحافظات من الجدول
// استخراج المحافظات
function extractGovernorates(rows) {
  const result = [];

  if (!rows.length) return result;

  const headerRow = rows[0];

  for (let col = 0; col < headerRow.length; col += 3) {
    const govName = headerRow[col];
    if (!govName) continue;

    const taskCol = col;
    const valueCol = col + 1;

    const govRows = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];

      const task = row[taskCol];
      const value = row[valueCol];

      // 🔥 التصفية الجديدة:
      // تجاهل المهمة إذا كانت القيمة فارغة أو صفر
      if (!task) continue;
      if (value === "" || value === null) continue;
      if (parseFloat(value) === 0) continue;

      govRows.push({
        task: task,
        value: value
      });
    }

    result.push({
      name: govName,
      rows: govRows
    });
  }

  return result;
}

// ترتيب القيم تنازليًا
function sortByValueDesc(rows) {
  return [...rows].sort((a, b) => {
    const valA = parseFloat(a.value) || 0;
    const valB = parseFloat(b.value) || 0;
    return valB - valA;
  });
}

// بناء ملف Excel جديد
function buildWorkbook(governorates) {
  const workbook = XLSX.utils.book_new();

  governorates.forEach((gov) => {
    const numericTasks = gov.rows.filter(r => !String(r.value).includes('%'));
    const percentageTasks = gov.rows.filter(r => String(r.value).includes('%'));

    const sortedNumeric = sortByValueDesc(numericTasks);
    const sortedPercentage = sortByValueDesc(
      percentageTasks.map(r => ({
        ...r,
        value: parseFloat(r.value) || 0
      }))
    );

    const sheetData = [
      ['المهمة', 'العدد'],
      ...sortedNumeric.map(r => [r.task, r.value]),
      [],
      [],
      [],
      ...sortedPercentage.map(r => [r.task, r.value + '%'])
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    worksheet['!sheetViews'] = [{ RTL: true }];
    worksheet['!cols'] = [{ wch: 50 }, { wch: 20 }];

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      gov.name.substring(0, 31)
    );
  });

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = [{ RTL: true }];

  return workbook;
}

// زر المعالجة
processBtn.addEventListener('click', async () => {
  const file = selectedFile;

  if (!file) {
    showStatus('❌ الرجاء اختيار ملف أولاً', true);
    return;
  }

  showStatus('⏳ جاري معالجة الملف...');

  try {
    const rows = await parseWorkbook(file);
    const governorates = extractGovernorates(rows);

    if (!governorates.length) {
      showStatus('❌ لم يتم التعرف على البيانات', true);
      return;
    }

    const workbook = buildWorkbook(governorates);

    // إضافة التاريخ بصيغة YYYY-MM-DD
    const today = new Date().toISOString().split('T')[0];

    XLSX.writeFile(workbook, `السكن-الوظيفي-${today}.xlsx`);

    showStatus('✅ تم تقسيم المحافظات بنجاح');

  } catch (error) {
    console.error(error);
    showStatus('❌ حدث خطأ أثناء المعالجة', true);
  }
});
