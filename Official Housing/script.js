const fileInput = document.getElementById('fileInput');
const processBtn = document.getElementById('processBtn');
const statusText = document.getElementById('status');

function showStatus(msg, error = false) {
  statusText.textContent = msg;
  statusText.style.color = error ? 'red' : '#2c3e50';
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
          raw: false
        });

        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };

    reader.readAsArrayBuffer(file);
  });
}

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

      if (!task && !value) continue;

      govRows.push({
        task: task || '',
        value: value || ''
      });
    }

    result.push({
      name: govName,
      rows: govRows
    });
  }

  return result;
}

function sortByValueDesc(rows) {
  return [...rows].sort((a, b) => {
    const valA = parseFloat(a.value) || 0;
    const valB = parseFloat(b.value) || 0;
    return valB - valA;
  });
}

function buildWorkbook(governorates) {
  const workbook = XLSX.utils.book_new();

  governorates.forEach((gov) => {

    // نفترض أن القيمة النسبية تحتوي %
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
      // ['📊 المهام العددية (من الأكبر للأصغر)'],
      ['المهمة', 'العدد'],
      ...sortedNumeric.map(r => [r.task, r.value]),

      // [],
      // ['📈 المهام النسبية (من الأكبر للأصغر)'],
      // ['المهمة', 'النسبة'],
      ...sortedPercentage.map(r => [r.task, r.value + '%'])
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    worksheet['!sheetViews'] = [{ RTL: true }];
    worksheet['!cols'] = [
      { wch: 50 },
      { wch: 20 }
    ];

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
  const file = fileInput.files[0];

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

    XLSX.writeFile(workbook, 'Governorates.xlsx');

    showStatus('✅ تم تقسيم المحافظات بنجاح');

  } catch (err) {
    console.error(err);
    showStatus('❌ حدث خطأ أثناء المعالجة', true);
  }
});