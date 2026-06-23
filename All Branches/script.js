const fileInput = document.getElementById('fileInput');
const dateInput = document.getElementById('dateInput');
const cleanButton = document.getElementById('cleanButton');
const downloadButton = document.getElementById('downloadButton');
const statusText = document.getElementById('status');
const resultTableBody = document.querySelector('#resultTable tbody');
const resultDateHeader = document.querySelector('#resultTable thead th:nth-child(2)');
const selectedFilesContainer = document.getElementById('selectedFiles');
const ruleSourceInput = document.getElementById('ruleSourceInput');
const ruleReplacementInput = document.getElementById('ruleReplacementInput');
const saveRuleButton = document.getElementById('saveRuleButton');
const cancelRuleEditButton = document.getElementById('cancelRuleEditButton');
const resetRulesButton = document.getElementById('resetRulesButton');
const rulesSearchInput = document.getElementById('rulesSearchInput');
const rulesStatus = document.getElementById('rulesStatus');
const rulesList = document.getElementById('rulesList');

let currentRows = [];
let filteredRows = [];
let previewRows = [];
let workbookSheets = [];
let selectedFiles = [];
let exportHeaders = ['Task', 'Date'];
let originalFileName = 'Excel-Cleaner-Filtered.xlsx';
let lastSelectedDate = '';
let proofreadingRuleEntries = [];
let editingRuleId = null;

const allowedSheetNames = [
  'المتابعة',
  'شؤون الضباط',
  'صف الضباط',
  'تنمية القدرات',
  'البيانات',
  'المدنيين',
  'التخطيط',
  'الأفراد',
  'المالية',
  'التوظيف',
  'المعلوماتية',
  'الخدمات',
  'الديوان',
  'المركبات',
  'القوى البشرية',
];

const sheetNameAliases = [
  { sheet: 'شؤون الضباط', aliases: ['شؤون الضباط', 'شئون الضباط'] },
  { sheet: 'صف الضباط', aliases: ['صف الضباط', 'صف ضباط'] },
  { sheet: 'القوى البشرية', aliases: ['قسم القوى البشرية'] },
  { sheet: 'الأفراد', aliases: ['الأفراد', 'الافراد'] },
  { sheet: 'المدنيين', aliases: ['المدنيين', 'العاملين المدنيين', 'مدنيين','العاملين'] },
  { sheet: 'المتابعة', aliases: ['المتابعة', 'المتابعه'] },
  { sheet: 'المعلوماتية', aliases: ['المعلوماتية', 'المعلوماتيه'] },
  { sheet: 'التوظيف', aliases: ['التوظيف'] },
  { sheet: 'التخطيط', aliases: ['التخطيط','تخطيط'] },
  { sheet: 'تنمية القدرات', aliases: ['تنمية', 'القدرات'] },
  { sheet: 'البيانات', aliases: ['البيانات'] },
  { sheet: 'المالية', aliases: ['المالية', 'الماليه', 'المالي'] },
  { sheet: 'الخدمات', aliases: ['الخدمات', 'خدمات'] },
  { sheet: 'الديوان', aliases: ['الديوان'] },
  { sheet: 'المركبات', aliases: ['المركبات'] },
];

const branchSheetTitles = {
  المتابعة: 'فرع المتابعة والتقييم',
  'شؤون الضباط': 'فرع شؤون الضباط',
  'صف الضباط': 'فرع صف الضباط',
  'تنمية القدرات': 'فرع تنمية القدرات',
  البيانات: 'فرع البيانات',
  المدنيين: 'فرع شؤون العاملين المدنيين والمتعاقدين',
  التخطيط: 'فرع تخطيط القوى البشرية',
  الأفراد: 'فرع الأفراد',
  المالية: 'الفرع المالي',
  التوظيف: 'قسم التوظيف المركزي',
  المعلوماتية: 'قسم المعلوماتية',
  الخدمات: 'مكتب الخدمات اللوجستية',
  الديوان: 'قسم الديوان المركزي',
  المركبات: 'قسم المركبات',
  'القوى البشرية': 'قسم القوى البشرية',
};

const taskProofreadingRules = [
  { pattern: /احالة مزكرات/g, replacement: 'إحالة مذكرات' },
  { pattern: /البريد الوارد و الضادر/g, replacement: 'البريد الوارد والصادر' },
  { pattern: /البريد الوارد والضادر/g, replacement: 'البريد الوارد والصادر' },
  { pattern: /البريد الضادر/g, replacement: 'البريد الصادر' },
  { pattern: /البريد الضادره/g, replacement: 'البريد الصادر' },
  { pattern: /الضادر/g, replacement: 'الصادر' },
  { pattern: /الضادره/g, replacement: 'الصادرة' },
  { pattern: /ضادر/g, replacement: 'صادر' },
  { pattern: /ضادره/g, replacement: 'صادرة' },
  { pattern: /احالات/g, replacement: 'إحالات' },
  { pattern: /احالة/g, replacement: 'إحالة' },
  { pattern: /الغاءات/g, replacement: 'إلغاءات' },
  { pattern: /الغاء/g, replacement: 'إلغاء' },
  { pattern: /المنع سفر/g, replacement: 'منع سفر' },
  { pattern: /الكف يد/g, replacement: 'كف يد' },
  { pattern: /مزكرات/g, replacement: 'مذكرات' },
  { pattern: /مزكرة/g, replacement: 'مذكرة' },
  { pattern: /مخاطبه/g, replacement: 'مخاطبة' },
  { pattern: /المخاطبه/g, replacement: 'المخاطبة' },
  { pattern: /مراسله/g, replacement: 'مراسلة' },
  { pattern: /المراسله/g, replacement: 'المراسلة' },
  { pattern: /الصادره/g, replacement: 'الصادرة' },
  { pattern: /الواردات/g, replacement: 'الواردات' },
  { pattern: /وارده/g, replacement: 'واردة' },
  { pattern: /الصادر/g, replacement: 'الصادر' },
  { pattern: /الاجازات/g, replacement: 'الإجازات' },
  { pattern: /اجازات/g, replacement: 'إجازات' },
  { pattern: /اجازة/g, replacement: 'إجازة' },
  { pattern: /اجازه/g, replacement: 'إجازة' },
  { pattern: /الاجازه/g, replacement: 'الإجازة' },
  { pattern: /اذونات/g, replacement: 'إذونات' },
  { pattern: /اذن/g, replacement: 'إذن' },
  { pattern: /الاذونات/g, replacement: 'الإذونات' },
  { pattern: /الاذن/g, replacement: 'الإذن' },
  { pattern: /ادخال/g, replacement: 'إدخال' },
  { pattern: /الادخال/g, replacement: 'الإدخال' },
  { pattern: /ادارة/g, replacement: 'إدارة' },
  { pattern: /الادارات/g, replacement: 'الإدارات' },
  { pattern: /الادارة/g, replacement: 'الإدارة' },
  { pattern: /الاجراءات/g, replacement: 'الإجراءات' },
  { pattern: /اجراءات/g, replacement: 'إجراءات' },
  { pattern: /اجراء/g, replacement: 'إجراء' },
  { pattern: /الاحالات/g, replacement: 'الإحالات' },
  { pattern: /الاحالة/g, replacement: 'الإحالة' },
  { pattern: /الايجارات/g, replacement: 'الإيجارات' },
  { pattern: /ايجارات/g, replacement: 'إيجارات' },
  { pattern: /ايجار/g, replacement: 'إيجار' },
  { pattern: /الاخلاءات/g, replacement: 'الإخلاءات' },
  { pattern: /اخلاءات/g, replacement: 'إخلاءات' },
  { pattern: /اخلاء/g, replacement: 'إخلاء' },
  { pattern: /الاعفاءات/g, replacement: 'الإعفاءات' },
  { pattern: /اعفاءات/g, replacement: 'إعفاءات' },
  { pattern: /اعفاء/g, replacement: 'إعفاء' },
  { pattern: /الاقامات/g, replacement: 'الإقامات' },
  { pattern: /اقامات/g, replacement: 'إقامات' },
  { pattern: /اقامة/g, replacement: 'إقامة' },
  { pattern: /اقامه/g, replacement: 'إقامة' },
  { pattern: /الاقامة/g, replacement: 'الإقامة' },
  { pattern: /الاقامه/g, replacement: 'الإقامة' },
  { pattern: /الاستمارات/g, replacement: 'الاستمارات' },
  { pattern: /استماره/g, replacement: 'استمارة' },
  { pattern: /الاستماره/g, replacement: 'الاستمارة' },
  { pattern: /المعامله/g, replacement: 'المعاملة' },
  { pattern: /معامله/g, replacement: 'معاملة' },
  { pattern: /المعاملات/g, replacement: 'المعاملات' },
  { pattern: /الاستبيانات/g, replacement: 'الاستبيانات' },
  { pattern: /استبيان/g, replacement: 'استبيان' },
  { pattern: /الالكتروني/g, replacement: 'الإلكتروني' },
  { pattern: /الالكترونية/g, replacement: 'الإلكترونية' },
  { pattern: /الالكترونيه/g, replacement: 'الإلكترونية' },
  { pattern: /الكترونيا/g, replacement: 'إلكترونياً' },
  { pattern: /الكتروني/g, replacement: 'إلكتروني' },
  { pattern: /الكترونية/g, replacement: 'إلكترونية' },
  { pattern: /الكترونيه/g, replacement: 'إلكترونية' },
  { pattern: /الإلكترونيه/g, replacement: 'الإلكترونية' },
  { pattern: /إلكترونيه/g, replacement: 'إلكترونية' },
  { pattern: /الارشفة/g, replacement: 'الأرشفة' },
  { pattern: /ارشفة/g, replacement: 'أرشفة' },
  { pattern: /الارشفه/g, replacement: 'الأرشفة' },
  { pattern: /ارشفه/g, replacement: 'أرشفة' },
  { pattern: /الاحصاءات/g, replacement: 'الإحصاءات' },
  { pattern: /احصاءات/g, replacement: 'إحصاءات' },
  { pattern: /احصاء/g, replacement: 'إحصاء' },
  { pattern: /الانشاء/g, replacement: 'الإنشاء' },
  { pattern: /انشاء/g, replacement: 'إنشاء' },
  { pattern: /الاصدار/g, replacement: 'الإصدار' },
  { pattern: /اصدار/g, replacement: 'إصدار' },
  { pattern: /الانجازات/g, replacement: 'الإنجازات' },
  { pattern: /انجازات/g, replacement: 'إنجازات' },
  { pattern: /انجاز/g, replacement: 'إنجاز' },
  { pattern: /الانجاز/g, replacement: 'الإنجاز' },
  { pattern: /الاعمال/g, replacement: 'الأعمال' },
  { pattern: /اعمال/g, replacement: 'أعمال' },
  { pattern: /الاوامر/g, replacement: 'الأوامر' },
  { pattern: /اوامر/g, replacement: 'أوامر' },
  { pattern: /امر/g, replacement: 'أمر' },
  { pattern: /الاولوية/g, replacement: 'الأولوية' },
  { pattern: /اولويه/g, replacement: 'أولوية' },
  { pattern: /اولوية/g, replacement: 'أولوية' },
  { pattern: /اسبوعية/g, replacement: 'أسبوعية' },
  { pattern: /اسبوعيه/g, replacement: 'أسبوعية' },
  { pattern: /اسبوعي/g, replacement: 'أسبوعي' },
  { pattern: /اسبوع/g, replacement: 'أسبوع' },
  { pattern: /الاسبوع/g, replacement: 'الأسبوع' },
  { pattern: /الاسبوعية/g, replacement: 'الأسبوعية' },
  { pattern: /الاسبوعيه/g, replacement: 'الأسبوعية' },
  { pattern: /اوامر/g, replacement: 'أوامر' },
  { pattern: /ايصالات/g, replacement: 'إيصالات' },
  { pattern: /ايصال/g, replacement: 'إيصال' },
  { pattern: /الملاحظات/g, replacement: 'الملاحظات' },
  { pattern: /تامين/g, replacement: 'تأمين' },
  { pattern: /تدقيق/g, replacement: 'تدقيق' },
  { pattern: /تحديث/g, replacement: 'تحديث' },
  { pattern: /تعديل/g, replacement: 'تعديل' },
  { pattern: /تصنيف/g, replacement: 'تصنيف' },
  { pattern: /تعميم/g, replacement: 'تعميم' },
  { pattern: /تنسيق/g, replacement: 'تنسيق' },
  { pattern: /تقرير/g, replacement: 'تقرير' },
  { pattern: /تقارير/g, replacement: 'تقارير' },
  { pattern: /متابعة/g, replacement: 'متابعة' },
  { pattern: /متابعه/g, replacement: 'متابعة' },
  { pattern: /المتابعه/g, replacement: 'المتابعة' },
  { pattern: /معالجة/g, replacement: 'معالجة' },
  { pattern: /معالجه/g, replacement: 'معالجة' },
  { pattern: /المعالجه/g, replacement: 'المعالجة' },
  { pattern: /المراجعه/g, replacement: 'المراجعة' },
  { pattern: /مراجعه/g, replacement: 'مراجعة' },
  { pattern: /مطابقه/g, replacement: 'مطابقة' },
  { pattern: /المطابقه/g, replacement: 'المطابقة' },
  { pattern: /مخالفه/g, replacement: 'مخالفة' },
  { pattern: /المخالفه/g, replacement: 'المخالفة' },
  { pattern: /مخالفات/g, replacement: 'مخالفات' },
  { pattern: /مرفقات/g, replacement: 'مرفقات' },
  { pattern: /مرفق/g, replacement: 'مرفق' },
  { pattern: /مستندات/g, replacement: 'مستندات' },
  { pattern: /مستند/g, replacement: 'مستند' },
  { pattern: /نموذج/g, replacement: 'نموذج' },
  { pattern: /نماذج/g, replacement: 'نماذج' },
  { pattern: /لائحه/g, replacement: 'لائحة' },
  { pattern: /اللائحه/g, replacement: 'اللائحة' },
  { pattern: /شهاده/g, replacement: 'شهادة' },
  { pattern: /الشهاده/g, replacement: 'الشهادة' },
  { pattern: /شهادات/g, replacement: 'شهادات' },
  { pattern: /بيانات/g, replacement: 'بيانات' },
  { pattern: /بيان/g, replacement: 'بيان' },
  { pattern: /احضار/g, replacement: 'إحضار' },
  { pattern: /الحاق/g, replacement: 'إلحاق' },
  { pattern: /الحاقي/g, replacement: 'إلحاقي' },
  { pattern: /المركزيه/g, replacement: 'المركزية' },
  { pattern: /المركزية/g, replacement: 'المركزية' },
  { pattern: /الفروع/g, replacement: 'الفروع' },
  { pattern: /الاقسام/g, replacement: 'الأقسام' },
  { pattern: /اقسام/g, replacement: 'أقسام' },
  { pattern: /سكنيه/g, replacement: 'سكنية' },
  { pattern: /السكنيه/g, replacement: 'السكنية' },
  { pattern: /اليوميه/g, replacement: 'اليومية' },
  { pattern: /يوميه/g, replacement: 'يومية' },
  { pattern: /ربعيه/g, replacement: 'ربعية' },
  { pattern: /سنويه/g, replacement: 'سنوية' },
  { pattern: /المدنيه/g, replacement: 'المدنية' },
  { pattern: /مدنيه/g, replacement: 'مدنية' },
  { pattern: /الشكاوي/g, replacement: 'الشكاوى' },
  { pattern: /شكوي/g, replacement: 'شكوى' },
  { pattern: /دعوي/g, replacement: 'دعوى' },
  { pattern: /الدعاوي/g, replacement: 'الدعاوى' },
  { pattern: /الملاحضات/g, replacement: 'الملاحظات' },
  { pattern: /ملاحضات/g, replacement: 'ملاحظات' },
  { pattern: /ملاحضه/g, replacement: 'ملاحظة' },
  { pattern: /الخاصه/g, replacement: 'الخاصة' },
  { pattern: /خاصه/g, replacement: 'خاصة' },
  { pattern: /الشهريه/g, replacement: 'الشهرية' },
  { pattern: /شهريه/g, replacement: 'شهرية' },
  { pattern: /السنويه/g, replacement: 'السنوية' },
  { pattern: /سنويه/g, replacement: 'سنوية' },
  { pattern: /المهام/g, replacement: 'المهام' },
  { pattern: /المهمه/g, replacement: 'المهمة' },
  { pattern: /المهنية/g, replacement: 'المهنية' },
  { pattern: /وظيفه/g, replacement: 'وظيفة' },
  { pattern: /الوظيفه/g, replacement: 'الوظيفة' },
  { pattern: /وظائف/g, replacement: 'وظائف' },
  { pattern: /توظيف/g, replacement: 'توظيف' },
  { pattern: /المعلوماتيه/g, replacement: 'المعلوماتية' },
  { pattern: /الماليه/g, replacement: 'المالية' },
  { pattern: /الخدميه/g, replacement: 'الخدمية' },
  { pattern: /البشريه/g, replacement: 'البشرية' },
  { pattern: /القوه/g, replacement: 'القوة' },
  { pattern: /القوي/g, replacement: 'القوى' },
  { pattern: /الافراد/g, replacement: 'الأفراد' },
  { pattern: /الضباط/g, replacement: 'الضباط' },
  { pattern: /شئون/g, replacement: 'شؤون' },
  { pattern: /مسؤلية/g, replacement: 'مسؤولية' },
  { pattern: /مسؤوليه/g, replacement: 'مسؤولية' },
  { pattern: /مسؤل/g, replacement: 'مسؤول' },
  { pattern: /مسوول/g, replacement: 'مسؤول' },
];

const RULES_STORAGE_KEY = 'allBranchesTaskProofreadingRules';
const defaultTaskProofreadingEntries = taskProofreadingRules.map((rule, index) => ({
  id: `default-${index}`,
  source: rule.pattern.source,
  replacement: rule.replacement,
}));

function createRuleId() {
  return `rule-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function replaceExactProofreadingMatch(value, source, replacement) {
  const wordChar = 'A-Za-z0-9\\u0600-\\u06FF\\u0750-\\u077F\\u08A0-\\u08FF\\uFB50-\\uFDFF\\uFE70-\\uFEFF';
  const pattern = new RegExp(`(^|[^${wordChar}])(${escapeRegExp(source)})(?=$|[^${wordChar}])`, 'g');
  return String(value).replace(pattern, (match, prefix) => `${prefix}${replacement}`);
}

function sanitizeRuleEntry(entry, fallbackId = createRuleId()) {
  const source = String(entry?.source ?? '').trim();
  const replacement = String(entry?.replacement ?? '').trim();

  if (!source || !replacement) {
    return null;
  }

  return {
    id: String(entry?.id || fallbackId),
    source,
    replacement,
  };
}

function saveProofreadingRules() {
  try {
    localStorage.setItem(RULES_STORAGE_KEY, JSON.stringify(proofreadingRuleEntries));
  } catch (error) {
    console.warn('Unable to save proofreading rules.', error);
  }
}

function loadProofreadingRules() {
  let hasStoredRules = false;

  try {
    const storedRules = JSON.parse(localStorage.getItem(RULES_STORAGE_KEY) || 'null');
    if (Array.isArray(storedRules)) {
      hasStoredRules = true;
      proofreadingRuleEntries = storedRules
        .map((entry) => sanitizeRuleEntry(entry))
        .filter(Boolean);
    }
  } catch (error) {
    proofreadingRuleEntries = [];
  }

  if (!hasStoredRules && !proofreadingRuleEntries.length) {
    proofreadingRuleEntries = defaultTaskProofreadingEntries.map((entry) => ({ ...entry }));
    saveProofreadingRules();
  }
}

function getActiveTaskProofreadingRules() {
  return proofreadingRuleEntries.filter((entry) => entry.source && entry.replacement);
}

function resetRuleForm() {
  editingRuleId = null;
  if (ruleSourceInput) ruleSourceInput.value = '';
  if (ruleReplacementInput) ruleReplacementInput.value = '';
  if (saveRuleButton) saveRuleButton.textContent = 'إضافة للسجل';
  if (cancelRuleEditButton) cancelRuleEditButton.hidden = true;
}

function updateRulesStatus(visibleCount = proofreadingRuleEntries.length) {
  if (!rulesStatus) return;
  rulesStatus.textContent = `السجل يحتوي على ${proofreadingRuleEntries.length} قاعدة تصحيح. المعروض الآن: ${visibleCount}.`;
}

function renderProofreadingRules() {
  if (!rulesList) return;

  const query = String(rulesSearchInput?.value || '').trim().toLowerCase();
  const visibleRules = proofreadingRuleEntries.filter((entry) => {
    const source = entry.source.toLowerCase();
    const replacement = entry.replacement.toLowerCase();
    return !query || source.includes(query) || replacement.includes(query);
  });

  rulesList.innerHTML = '';

  if (!visibleRules.length) {
    rulesList.innerHTML = '<div class="empty-rules">لا توجد كلمات مطابقة للبحث الحالي.</div>';
    updateRulesStatus(0);
    return;
  }

  visibleRules.forEach((entry) => {
    const item = document.createElement('div');
    item.className = 'rule-item';
    item.dataset.id = entry.id;

    const text = document.createElement('div');
    text.className = 'rule-text';

    const source = document.createElement('span');
    source.className = 'rule-source';
    source.textContent = entry.source;

    const arrow = document.createElement('span');
    arrow.className = 'rule-arrow';
    arrow.textContent = '←';

    const replacement = document.createElement('span');
    replacement.className = 'rule-replacement';
    replacement.textContent = entry.replacement;

    text.append(source, arrow, replacement);

    const actions = document.createElement('div');
    actions.className = 'rule-buttons';

    const editButton = document.createElement('button');
    editButton.type = 'button';
    editButton.className = 'rule-edit-btn';
    editButton.textContent = 'تعديل';
    editButton.dataset.action = 'edit-rule';

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'rule-delete-btn';
    deleteButton.textContent = 'حذف';
    deleteButton.dataset.action = 'delete-rule';

    actions.append(editButton, deleteButton);
    item.append(text, actions);
    rulesList.appendChild(item);
  });

  updateRulesStatus(visibleRules.length);
}

function saveRuleFromForm() {
  const source = ruleSourceInput?.value.trim() || '';
  const replacement = ruleReplacementInput?.value.trim() || '';

  if (!source || !replacement) {
    showStatus('اكتب الكلمة الحالية والتصحيح قبل الحفظ.', true);
    return;
  }

  const duplicate = proofreadingRuleEntries.find((entry) => entry.source === source && entry.id !== editingRuleId);
  if (duplicate) {
    showStatus('هذه الكلمة موجودة مسبقًا في سجل التدقيق.', true);
    return;
  }

  if (editingRuleId) {
    proofreadingRuleEntries = proofreadingRuleEntries.map((entry) =>
      entry.id === editingRuleId ? { ...entry, source, replacement } : entry
    );
    showStatus('تم تعديل كلمة التدقيق بنجاح.');
  } else {
    proofreadingRuleEntries.push({ id: createRuleId(), source, replacement });
    showStatus('تمت إضافة كلمة جديدة إلى سجل التدقيق.');
  }

  saveProofreadingRules();
  resetRuleForm();
  renderProofreadingRules();
  resetProcessedResults();
}

function startRuleEdit(ruleId) {
  const entry = proofreadingRuleEntries.find((rule) => rule.id === ruleId);
  if (!entry) return;

  editingRuleId = entry.id;
  if (ruleSourceInput) ruleSourceInput.value = entry.source;
  if (ruleReplacementInput) ruleReplacementInput.value = entry.replacement;
  if (saveRuleButton) saveRuleButton.textContent = 'حفظ التعديل';
  if (cancelRuleEditButton) cancelRuleEditButton.hidden = false;
  ruleSourceInput?.focus();
}

function deleteRule(ruleId) {
  const entry = proofreadingRuleEntries.find((rule) => rule.id === ruleId);
  if (!entry) return;

  const confirmed = window.confirm(`هل تريد حذف "${entry.source}" من سجل التدقيق؟`);
  if (!confirmed) return;

  proofreadingRuleEntries = proofreadingRuleEntries.filter((rule) => rule.id !== ruleId);
  saveProofreadingRules();
  resetRuleForm();
  renderProofreadingRules();
  resetProcessedResults();
  showStatus('تم حذف الكلمة من سجل التدقيق.');
}

function resetProofreadingRulesToDefault() {
  const confirmed = window.confirm('سيتم حذف التعديلات الحالية واستعادة الكلمات الأصلية. هل تريد المتابعة؟');
  if (!confirmed) return;

  proofreadingRuleEntries = defaultTaskProofreadingEntries.map((entry) => ({ ...entry }));
  saveProofreadingRules();
  resetRuleForm();
  renderProofreadingRules();
  resetProcessedResults();
  showStatus('تمت استعادة سجل التدقيق الأصلي.');
}

function initializeProofreadingRulesManager() {
  loadProofreadingRules();
  renderProofreadingRules();

  saveRuleButton?.addEventListener('click', saveRuleFromForm);
  cancelRuleEditButton?.addEventListener('click', () => {
    resetRuleForm();
    showStatus('تم إلغاء تعديل كلمة التدقيق.');
  });
  resetRulesButton?.addEventListener('click', resetProofreadingRulesToDefault);
  rulesSearchInput?.addEventListener('input', renderProofreadingRules);
  rulesList?.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;

    const ruleId = button.closest('.rule-item')?.dataset.id;
    if (!ruleId) return;

    if (button.dataset.action === 'edit-rule') {
      startRuleEdit(ruleId);
    } else if (button.dataset.action === 'delete-rule') {
      deleteRule(ruleId);
    }
  });
}

function normalizeTaskSpacing(value, removeTrailingDot = true) {
  let normalized = String(value ?? '')
    .replace(/\u0640/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\s+([،؛:,.!?])/g, '$1')
    .replace(/([،؛:,.!?])(?=\S)/g, '$1 ')
    .replace(/\s*-\s*/g, ' - ')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .trim();

  if (removeTrailingDot) {
    normalized = normalized.replace(/[.。]+$/g, '').trim();
  }

  return normalized;
}

function proofreadTask(task) {
  const originalTask = String(task ?? '');
  const comparableOriginalTask = normalizeTaskSpacing(originalTask, false);
  let correctedTask = normalizeTaskSpacing(originalTask);

  getActiveTaskProofreadingRules().forEach((rule) => {
    correctedTask = replaceExactProofreadingMatch(correctedTask, rule.source, rule.replacement);
  });

  correctedTask = normalizeTaskSpacing(correctedTask);

  return {
    task: correctedTask,
    originalTask,
    taskChanged: correctedTask !== comparableOriginalTask,
  };
}

function showStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.style.color = isError ? '#ba1a1a' : '#445066';
}

function getFileKey(file) {
  return `${file.name}-${file.size}-${file.lastModified}`;
}

function resetProcessedResults() {
  workbookSheets = [];
  filteredRows = [];
  previewRows = [];
  downloadButton.disabled = true;
  resultTableBody.innerHTML = '';
  updateTableHeader('التاريخ');
}

function renderSelectedFiles() {
  selectedFilesContainer.innerHTML = '<div class="selected-files-title">الملفات المختارة</div>';

  if (!selectedFiles.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-files';
    empty.textContent = 'لم يتم اختيار أي ملفات بعد.';
    selectedFilesContainer.appendChild(empty);
    return;
  }

  const list = document.createElement('ul');
  list.className = 'files-list';

  selectedFiles.forEach((file, index) => {
    const item = document.createElement('li');
    item.className = 'file-item';

    const info = document.createElement('span');
    info.className = 'file-info';
    info.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

    const removeButton = document.createElement('button');
    removeButton.className = 'remove-file-btn';
    removeButton.type = 'button';
    removeButton.textContent = 'حذف';
    removeButton.setAttribute('aria-label', `حذف ${file.name}`);
    removeButton.addEventListener('click', () => {
      selectedFiles.splice(index, 1);
      renderSelectedFiles();
      resetProcessedResults();
      showStatus(selectedFiles.length ? `تم حذف الملف. المتبقي ${selectedFiles.length}.` : 'تم حذف جميع الملفات المختارة.');
    });

    item.appendChild(info);
    item.appendChild(removeButton);
    list.appendChild(item);
  });

  selectedFilesContainer.appendChild(list);
}

fileInput.addEventListener('change', () => {
  const newFiles = Array.from(fileInput.files);
  let addedCount = 0;

  newFiles.forEach((file) => {
    const exists = selectedFiles.some((selectedFile) => getFileKey(selectedFile) === getFileKey(file));
    if (!exists) {
      selectedFiles.push(file);
      addedCount += 1;
    }
  });

  fileInput.value = '';
  renderSelectedFiles();
  resetProcessedResults();

  if (addedCount > 0) {
    showStatus(`تمت إضافة ${addedCount} ملف. إجمالي الملفات المختارة: ${selectedFiles.length}.`);
  } else if (newFiles.length) {
    showStatus('هذه الملفات موجودة مسبقًا في القائمة.');
  }
});

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
    .map((row) => {
      const proofread = proofreadTask(row.task);
      return {
        task: proofread.task,
        originalTask: proofread.originalTask,
        taskChanged: proofread.taskChanged,
        date: row.date,
      };
    });
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
    .map((row) => {
      const proofread = proofreadTask(row.task);
      return {
        task: proofread.task,
        originalTask: proofread.originalTask,
        taskChanged: proofread.taskChanged,
        date: row.dateValues[selectedHeader],
      };
    });

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

function normalizeArabicName(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[\u064B-\u065F\u0670]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    .replace(/\u0640/g, '')
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getSheetNameFromFile(fileName) {
  const baseName = fileName.replace(/\.(xlsx|xls|csv)$/i, '');
  const normalizedBaseName = normalizeArabicName(baseName);

  const matched = sheetNameAliases.find((entry) =>
    entry.aliases.some((alias) => normalizedBaseName.includes(normalizeArabicName(alias)))
  );

  if (matched) {
    return matched.sheet;
  }

  const directMatch = allowedSheetNames.find((sheetName) =>
    normalizedBaseName.includes(normalizeArabicName(sheetName))
  );

  if (directMatch) {
    return directMatch;
  }

  const cleanedBaseName = baseName.replace(/[^\w\u0600-\u06FF \-]/g, '');
  const trimmed = cleanedBaseName.trim();
  return trimmed ? trimmed.substring(0, 31) : 'Sheet1';
}

function mergeSheetsByName(sheets) {
  const merged = new Map();

  sheets.forEach((sheet) => {
    const sheetName = allowedSheetNames.includes(sheet.name) ? sheet.name : getSheetNameFromFile(sheet.name);

    if (!merged.has(sheetName)) {
      merged.set(sheetName, {
        name: sheetName,
        headers: sheet.headers,
        rows: [],
      });
    }

    const current = merged.get(sheetName);
    if (sheet.rows.length && sheet.headers?.[1]) {
      current.headers = sheet.headers;
    }
    current.rows.push(...sheet.rows);
  });

  return Array.from(merged.values()).sort((a, b) => {
    const aIndex = allowedSheetNames.indexOf(a.name);
    const bIndex = allowedSheetNames.indexOf(b.name);

    if (aIndex === -1 && bIndex === -1) {
      return a.name.localeCompare(b.name, 'ar');
    }

    if (aIndex === -1) return 1;
    if (bIndex === -1) return -1;
    return aIndex - bIndex;
  });
}

function makeUniqueSheetName(workbook, baseName) {
  const safeBaseName = String(baseName || 'Sheet1').substring(0, 31);
  let sheetName = safeBaseName;
  let suffix = 1;

  while (workbook.SheetNames.includes(sheetName)) {
    const suffixText = `-${suffix}`;
    sheetName = `${safeBaseName.substring(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }

  return sheetName;
}

function getBranchSheetTitle(sheetName) {
  return branchSheetTitles[sheetName] || sheetName || 'فرع';
}

function createCellStyle({ fillColor, fontColor = '000000', bold = false, fontSize = 14, horizontal = 'center' } = {}) {
  const style = {
    font: {
      name: 'Arial',
      bold,
      sz: fontSize,
      color: { rgb: fontColor },
    },
    alignment: {
      horizontal,
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

function buildWorkbookFromSheets(sheets) {
  const workbook = XLSX.utils.book_new();

  mergeSheetsByName(sheets).forEach((sheet) => {
    const sheetName = makeUniqueSheetName(workbook, sheet.name);

    const sortedRows = sortRowsDescending([...sheet.rows]);
    const exportRows = addNumericPercentSeparatorRows(sortedRows);

    const sheetData = [
      [getBranchSheetTitle(sheet.name), ''],
      ['المهمة', 'العدد/النسبة'],
      ...exportRows.map((row) => [row.task || '', String(row.date || '')]),
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    const titleStyle = createCellStyle({ fillColor: '9A9558', bold: true, fontSize: 16 });
    const headerStyle = createCellStyle({ fillColor: '00B050', bold: true, fontSize: 14 });
    const taskStyle = createCellStyle({ fontSize: 14 });
    const valueStyle = createCellStyle({ fontSize: 14 });
    const changedTaskStyle = createCellStyle({ fillColor: 'FFFF00', fontSize: 14 });

    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    worksheet['A1'].s = titleStyle;
    worksheet['B1'] = worksheet['B1'] || { t: 's', v: '' };
    worksheet['B1'].s = titleStyle;
    worksheet['A2'].s = headerStyle;
    worksheet['B2'].s = headerStyle;

    exportRows.forEach((row, index) => {
      const rowIndex = index + 2;
      const taskCellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
      const valueCellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 1 });

      worksheet[taskCellAddress] = worksheet[taskCellAddress] || { t: 's', v: row.task || '' };
      worksheet[valueCellAddress] = worksheet[valueCellAddress] || { t: 's', v: String(row.date || '') };
      worksheet[taskCellAddress].s = row.taskChanged ? changedTaskStyle : taskStyle;
      worksheet[valueCellAddress].s = valueStyle;
    });

    worksheet['!sheetViews'] = [{ RTL: true }];
    worksheet['!cols'] = [{ wch: 72 }, { wch: 18 }];
    worksheet['!rows'] = [{ hpt: 42 }, { hpt: 30 }, ...exportRows.map(() => ({ hpt: 27 }))];
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  workbook.Workbook = workbook.Workbook || {};
  workbook.Workbook.Views = workbook.Workbook.Views || [{ RTL: true }];
  return workbook;
}

function getUnifiedPrintSections() {
  return mergeSheetsByName(workbookSheets)
    .map((sheet) => {
      const rows = addNumericPercentSeparatorRows(sortRowsDescending([...sheet.rows]));
      return {
        title: getBranchSheetTitle(sheet.name),
        rows: rows.map((row) => ({
          task: row.task || '',
          value: String(row.date || ''),
          taskChanged: Boolean(row.taskChanged),
        })),
      };
    })
    .filter((section) => section.rows.length > 0);
}

window.getUnifiedPrintSections = getUnifiedPrintSections;

function getValueCategory(value) {
  if (value === null || value === undefined) {
    return { type: 3, value: 0 };
  }

  const str = String(value).trim();

  if (str.endsWith('%')) {
    return {
      type: 2,
      value: parseFloat(str.replace('%', '')) || 0,
    };
  }

  if (!isNaN(str)) {
    return {
      type: 1,
      value: parseFloat(str),
    };
  }

  return {
    type: 3,
    value: 0,
  };
}

function addNumericPercentSeparatorRows(rows) {
  const output = [];

  rows.forEach((row, index) => {
    //يتم التفعيل عند الضرورة فقططططط
    // if (index > 0) {
    //   const previousType = getValueCategory(rows[index - 1].date).type;
    //   const currentType = getValueCategory(row.date).type;

    //   if (previousType === 1 && currentType === 2) {
    //     output.push({ task: '', date: '' }, { task: '', date: '' });
    //   }
    // }

    output.push(row);
  });

  return output;
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
    const aVal = getValueCategory(a.date);
    const bVal = getValueCategory(b.date);

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
  const files = selectedFiles;
  // const selectedDate = dateInput.value;
  const filterType = document.getElementById('filterType').value;

  let selectedDate = dateInput.value;

  // إذا اختار أسبوع
  if (filterType === 'week1') selectedDate = "1";
  if (filterType === 'week2') selectedDate = "2";
  if (filterType === 'week3') selectedDate = "3";
  if (filterType === 'week4') selectedDate = "4";
  if (filterType === 'all') selectedDate = "55";

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

    const correctedTasksCount = filteredRows.filter((row) => row.taskChanged).length;
    const proofreadingMessage = correctedTasksCount
      ? ` وتم تدقيق ${correctedTasksCount} مهمة وتلوينها بالأصفر عند التصدير.`
      : ' ولم يتم العثور على مهام تحتاج تعديلاً.';

    showStatus(`تم تنظيف البيانات بنجاح من ${files.length} ملف${files.length === 1 ? '' : 'ات'}.${proofreadingMessage}`);
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
initializeProofreadingRulesManager();

// عند تغيير الاختيار
filterType.addEventListener('change', toggleDateInput);
