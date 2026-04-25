const state = {
  excelFile: null,
  pdfFile: null,
  rows: [],
  results: [],
  activeFilter: "all",
};

const els = {
  excelFile: document.querySelector("#excelFile"),
  pdfFile: document.querySelector("#pdfFile"),
  excelName: document.querySelector("#excelName"),
  pdfName: document.querySelector("#pdfName"),
  compareBtn: document.querySelector("#compareBtn"),
  clearBtn: document.querySelector("#clearBtn"),
  exportBtn: document.querySelector("#exportBtn"),
  threshold: document.querySelector("#threshold"),
  thresholdValue: document.querySelector("#thresholdValue"),
  strictBranch: document.querySelector("#strictBranch"),
  status: document.querySelector("#status"),
  resultsList: document.querySelector("#resultsList"),
  searchBox: document.querySelector("#searchBox"),
  branchCount: document.querySelector("#branchCount"),
  taskCount: document.querySelector("#taskCount"),
  okCount: document.querySelector("#okCount"),
  errorCount: document.querySelector("#errorCount"),
};

const headerAliases = {
  branch: ["branch", "branches", "فرع", "الفرع", "اسم الفرع", "الأفرع", "الافرع"],
  task: ["task", "tasks", "المهمة", "مهام", "المهام", "اسم المهمة", "الوصف", "العمل"],
  number: ["number", "no", "id", "#", "رقم", "الرقم", "رقم المهمة", "م", "تسلسل"],
  value: [
    "value",
    "count",
    "qty",
    "quantity",
    "percent",
    "percentage",
    "rate",
    "progress",
    "عدد",
    "العدد",
    "نسبة",
    "النسبة",
    "نسبة الانجاز",
    "نسبة الإنجاز",
    "الانجاز",
    "الإنجاز",
    "القيمة",
    "معدل",
  ],
};

function setStatus(message, type = "") {
  els.status.textContent = message;
  els.status.className = `status ${type}`.trim();
}

function normalizeDigitsAndPercent(value) {
  const arabicDigits = "٠١٢٣٤٥٦٧٨٩";
  const persianDigits = "۰۱۲۳۴۵۶۷۸۹";
  return String(value ?? "")
    .replace(/[٠-٩]/g, (digit) => arabicDigits.indexOf(digit))
    .replace(/[۰-۹]/g, (digit) => persianDigits.indexOf(digit))
    .replace(/[٪﹪％]/g, "%");
}

function normalizeArabic(value) {
  return normalizeDigitsAndPercent(value)
    .normalize("NFKC")
    .replace(/\uFFFD/g, " ")
    .replace(/[\u064B-\u065F\u0670]/g, "")
    .replace(/[إأآا]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/ؤ/g, "و")
    .replace(/ئ/g, "ي")
    .replace(/\u0640/g, "")
    .replace(/[^\p{L}\p{N}\s%]/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function tokenSet(value) {
  return new Set(
    normalizeArabic(value)
      .split(" ")
      .filter((token) => token.length > 1),
  );
}

function removeValuesFromText(value) {
  return normalizeDigitsAndPercent(value)
    .replace(/%\s*-?\d+(?:[.,]\d+)?/g, " ")
    .replace(/-?\d+(?:[.,]\d+)?\s*%/g, " ")
    .replace(/(^|\s)-?\d+(?:[.,]\d+)?(?=\s|$)/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function similarity(a, b) {
  const left = tokenSet(removeValuesFromText(a));
  const right = tokenSet(removeValuesFromText(b));
  if (!left.size || !right.size) return 0;
  let hits = 0;
  left.forEach((token) => {
    const partialHit = [...right].some((candidate) => {
      if (candidate === token) return true;
      const shorter = candidate.length < token.length ? candidate : token;
      const longer = candidate.length >= token.length ? candidate : token;
      return shorter.length >= 4 && longer.includes(shorter);
    });
    if (partialHit) hits += 1;
  });
  return Math.round((2 * hits * 100) / (left.size + right.size));
}

function findHeaderIndex(headers, kind) {
  const aliases = headerAliases[kind].map(normalizeArabic);
  return headers.findIndex((header) => aliases.includes(normalizeArabic(header)));
}

function looksLikeValue(value) {
  const text = normalizeDigitsAndPercent(value).trim();
  return /^%?\s*-?\d+(?:[.,]\d+)?\s*%?$/.test(text);
}

function normalizeValue(value) {
  const raw = normalizeDigitsAndPercent(value).trim();
  if (!raw) return "";
  const hasPercent = raw.includes("%");
  const number = Number(raw.replace(/[^\d.,-]/g, "").replace(",", "."));
  if (!Number.isFinite(number)) return normalizeArabic(raw);
  const compact = Number.isInteger(number) ? String(number) : String(number).replace(/0+$/, "").replace(/\.$/, "");
  return hasPercent ? `${compact}%` : compact;
}

function valueExistsInText(value, text) {
  const normalized = normalizeValue(value);
  if (!normalized) return true;
  const normalizedText = normalizeDigitsAndPercent(text).replace(/,/g, ".");
  const numberText = normalized.replace("%", "");
  const escapedNumber = numberText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const patterns = normalized.includes("%")
    ? [
        new RegExp(`(^|[^\\d])(?:${escapedNumber}\\s*%|%\\s*${escapedNumber})([^\\d]|$)`),
      ]
    : [new RegExp(`(^|[^\\d])${escapedNumber}\\s*%?([^\\d]|$)`)];

  if (normalized.includes("%")) {
    const percentNumber = Number(numberText);
    if (Number.isFinite(percentNumber)) {
      const decimalEquivalent = percentNumber / 100;
      const decimalText = Number.isInteger(decimalEquivalent)
        ? String(decimalEquivalent)
        : String(decimalEquivalent).replace(/0+$/, "").replace(/\.$/, "");
      const escapedDecimal = decimalText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      patterns.push(new RegExp(`(^|[^\\d])${escapedDecimal}([^\\d]|$)`));
    }
  }

  return patterns.some((pattern) => pattern.test(normalizedText));
}

function taskExistsInText(task, text) {
  const taskOnly = normalizeArabic(removeValuesFromText(task));
  const textOnly = normalizeArabic(removeValuesFromText(text));
  return Boolean(taskOnly && textOnly.includes(taskOnly));
}

function nearbyTextForValue(bestLine, sectionText) {
  if (!bestLine) return sectionText;
  const lines = String(sectionText || "").split("\n").filter(Boolean);
  const bestIndex = lines.findIndex((line) => normalizeArabic(line) === normalizeArabic(bestLine));
  if (bestIndex < 0) return `${bestLine}\n${sectionText}`;
  return lines.slice(Math.max(0, bestIndex - 1), bestIndex + 2).join("\n");
}

function splitRowIntoVisualLines(row, viewportWidth) {
  const sorted = row.items.sort((a, b) => b.x - a.x);
  const parts = [];
  const gapThreshold = Math.max(36, viewportWidth * 0.055);

  sorted.forEach((item) => {
    const current = parts[parts.length - 1];
    const itemRight = item.x + item.width;
    const gapFromCurrent = current ? current.minX - itemRight : 0;

    if (!current || gapFromCurrent > gapThreshold) {
      parts.push({
        y: row.y,
        minX: item.x,
        maxX: item.x + item.width,
        texts: [item.text],
      });
      return;
    }

    current.minX = Math.min(current.minX, item.x);
    current.maxX = Math.max(current.maxX, item.x + item.width);
    current.texts.push(item.text);
  });

  return parts
    .map((part) => ({
      page: part.page,
      y: part.y,
      x: (part.minX + part.maxX) / 2,
      minX: part.minX,
      maxX: part.maxX,
      line: part.texts.join(" ").replace(/\s+/g, " ").trim(),
    }))
    .filter((part) => part.line);
}

function orderPdfItemsByRightThenLeftGroups(textItems, viewportWidth) {
  const pageMiddle = viewportWidth / 2;
  const groups = [
    { side: "right", items: [] },
    { side: "left", items: [] },
  ];

  textItems.forEach((item) => {
    const center = item.x + item.width / 2;
    groups[center >= pageMiddle ? 0 : 1].items.push(item);
  });

  const yTolerance = 4;
  return groups.flatMap((group) => {
    if (!group.items.length) return [];
    const rows = [];

    group.items
      .sort((a, b) => b.y - a.y)
      .forEach((item) => {
        let row = rows.find((existing) => Math.abs(existing.y - item.y) <= yTolerance);
        if (!row) {
          row = { y: item.y, items: [] };
          rows.push(row);
        }
        row.items.push(item);
        row.y = (row.y * (row.items.length - 1) + item.y) / row.items.length;
      });

    return rows
      .sort((a, b) => b.y - a.y)
      .map((row) =>
        row.items
          .sort((a, b) => b.x - a.x)
          .map((item) => item.text)
          .join(" ")
          .replace(/\s+/g, " ")
          .trim(),
      )
      .filter(Boolean);
  });
}

function orderPdfLinesRightColumnFirst(lineParts, viewportWidth) {
  if (!lineParts.length) return [];
  const columns = [];
  const columnTolerance = Math.max(70, viewportWidth * 0.12);

  lineParts
    .sort((a, b) => b.x - a.x)
    .forEach((part) => {
      let column = columns.find((existing) => {
        const overlaps =
          Math.min(existing.maxX, part.maxX) - Math.max(existing.minX, part.minX);
        return Math.abs(existing.x - part.x) <= columnTolerance || overlaps > 0;
      });

      if (!column) {
        column = {
          x: part.x,
          minX: part.minX,
          maxX: part.maxX,
          parts: [],
        };
        columns.push(column);
      }

      column.parts.push(part);
      column.minX = Math.min(column.minX, part.minX);
      column.maxX = Math.max(column.maxX, part.maxX);
      column.x = (column.minX + column.maxX) / 2;
    });

  return columns
    .sort((a, b) => b.x - a.x)
    .flatMap((column) => column.parts.sort((a, b) => b.y - a.y))
    .map((part) => part.line);
}

function firstTextIndex(row) {
  return row.findIndex((cell) => String(cell ?? "").trim().length > 0);
}

function isHeaderRow(row) {
  const normalized = row.map(normalizeArabic);
  return Object.values(headerAliases).some((aliases) => {
    const normalizedAliases = aliases.map(normalizeArabic);
    return normalized.some((cell) => normalizedAliases.includes(cell));
  });
}

function getCell(rows, rowIndex, colIndex) {
  return rows[rowIndex]?.[colIndex] ?? "";
}

function sheetToDisplayRows(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const rows = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = sheet[address];
      row.push(cell ? String(cell.w ?? cell.v ?? "").trim() : "");
    }
    rows.push(row);
  }

  return rows;
}

function columnStats(rows, colIndex, startRow) {
  let filled = 0;
  let taskLike = 0;
  let valueLike = 0;

  for (let rowIndex = startRow; rowIndex < rows.length; rowIndex += 1) {
    const value = String(getCell(rows, rowIndex, colIndex)).trim();
    if (!value) continue;
    filled += 1;
    if (looksLikeValue(value)) valueLike += 1;
    if (!looksLikeValue(value) && normalizeArabic(value).length > 2) taskLike += 1;
  }

  return { filled, taskLike, valueLike };
}

function parseSheetByColumns(sheetName, rows) {
  const tasks = [];
  const cleanRows = rows.filter((row) => row.some((cell) => String(cell).trim()));
  if (!cleanRows.length) return tasks;

  const headerRow = cleanRows[0].map((cell) => String(cell).trim());
  const hasHeader = isHeaderRow(headerRow);
  const startRow = hasHeader ? 1 : 0;
  const branchCol = hasHeader ? findHeaderIndex(headerRow, "branch") : -1;
  const numberCol = hasHeader ? findHeaderIndex(headerRow, "number") : -1;
  const explicitValueCol = hasHeader ? findHeaderIndex(headerRow, "value") : -1;
  const explicitTaskCol = hasHeader ? findHeaderIndex(headerRow, "task") : -1;
  const maxCols = Math.max(...cleanRows.map((row) => row.length));
  const used = new Set();

  for (let colIndex = 0; colIndex < maxCols; colIndex += 1) {
    if (colIndex === branchCol || colIndex === numberCol || colIndex === explicitValueCol) continue;
    const stats = columnStats(cleanRows, colIndex, startRow);
    const headerLooksTask = explicitTaskCol === colIndex || headerAliases.task.map(normalizeArabic).includes(normalizeArabic(headerRow[colIndex]));
    const isTaskColumn = headerLooksTask || (stats.taskLike > 0 && stats.taskLike >= stats.valueLike);
    if (!isTaskColumn) continue;

    let pairedValueCol = explicitValueCol;
    const nextStats = columnStats(cleanRows, colIndex + 1, startRow);
    if (pairedValueCol < 0 && nextStats.valueLike > 0 && nextStats.valueLike >= nextStats.taskLike) {
      pairedValueCol = colIndex + 1;
      used.add(pairedValueCol);
    }

    for (let rowIndex = startRow; rowIndex < cleanRows.length; rowIndex += 1) {
      const task = String(getCell(cleanRows, rowIndex, colIndex)).trim();
      if (!task || looksLikeValue(task) || normalizeArabic(task) === normalizeArabic(headerRow[colIndex])) continue;

      const branch = String(getCell(cleanRows, rowIndex, branchCol)).trim() || sheetName.trim() || "بدون فرع";
      const explicitNumber = numberCol >= 0 && String(getCell(cleanRows, rowIndex, numberCol)).trim();
      const pairedValue = pairedValueCol >= 0 ? String(getCell(cleanRows, rowIndex, pairedValueCol)).trim() : "";

      tasks.push({
        branch,
        task,
        number: explicitNumber ? String(getCell(cleanRows, rowIndex, numberCol)).trim() : String(tasks.length + 1),
        hasExplicitNumber: Boolean(explicitNumber),
        value: looksLikeValue(pairedValue) ? pairedValue : "",
        sheet: sheetName,
      });
    }

    used.add(colIndex);
  }

  return tasks.filter((task, index, list) => {
    const key = `${normalizeArabic(task.branch)}|${normalizeArabic(task.task)}|${normalizeValue(task.value)}`;
    return list.findIndex((item) => `${normalizeArabic(item.branch)}|${normalizeArabic(item.task)}|${normalizeValue(item.value)}` === key) === index;
  });
}

function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("تعذر قراءة ملف Excel."));
    reader.onload = () => {
      try {
        const workbook = XLSX.read(reader.result, { type: "array", cellDates: false });
        const tasks = [];

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const rows = sheetToDisplayRows(sheet);
          tasks.push(...parseSheetByColumns(sheetName, rows));
        });

        resolve(tasks);
      } catch (error) {
        reject(new Error("ملف Excel غير مدعوم أو يحتوي على تنسيق غير متوقع."));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

async function waitForPdfJs() {
  for (let i = 0; i < 80; i += 1) {
    if (window.pdfjsLib) return window.pdfjsLib;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error("تعذر تحميل قارئ PDF. تأكد من الاتصال بالإنترنت عند فتح الصفحة.");
}

async function waitForTesseract() {
  for (let i = 0; i < 80; i += 1) {
    if (window.Tesseract) return window.Tesseract;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error("تعذر تحميل قارئ الصور OCR.");
}

function orderOcrLinesRightThenLeft(lines, pageWidth) {
  const cleaned = lines
    .map((line) => {
      const bbox = line.bbox || {};
      const text = String(line.text || "").replace(/\s+/g, " ").trim();
      return {
        line: text,
        x: ((bbox.x0 || 0) + (bbox.x1 || 0)) / 2,
        y: bbox.y0 || 0,
      };
    })
    .filter((item) => item.line);

  const groups = [
    cleaned.filter((item) => item.x >= pageWidth / 2),
    cleaned.filter((item) => item.x < pageWidth / 2),
  ];

  return groups.flatMap((group) => group.sort((a, b) => a.y - b.y));
}

async function parsePdfWithOcr(pdf) {
  const Tesseract = await waitForTesseract();
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    setStatus(`جاري قراءة PDF كصورة - صفحة ${pageNumber} من ${pdf.numPages}...`, "busy");
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { willReadFrequently: true });
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: context, viewport }).promise;

    const result = await Tesseract.recognize(canvas, "ara+eng", {
      logger(message) {
        if (message.status === "recognizing text") {
          const percent = Math.round((message.progress || 0) * 100);
          setStatus(`جاري قراءة PDF كصورة - صفحة ${pageNumber}: ${percent}%`, "busy");
        }
      },
    });

    const ordered = orderOcrLinesRightThenLeft(result.data.lines || [], canvas.width);
    const lines = ordered.map((item) => item.line);
    pages.push({ pageNumber, lines, text: lines.join("\n") });
  }

  return {
    source: "ocr",
    pages,
    lines: pages.flatMap((page) => page.lines.map((line) => ({ page: page.pageNumber, line }))),
    text: pages.map((page) => page.text).join("\n"),
  };
}

async function parsePdf(file) {
  const pdfjsLib = await waitForPdfJs();
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const buffer = await file.arrayBuffer();
  let pdf;
  try {
    pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  } catch (error) {
    pdfjsLib.disableWorker = true;
    pdf = await pdfjsLib.getDocument({ data: buffer, disableWorker: true }).promise;
  }

  try {
    const ocrData = await parsePdfWithOcr(pdf);
    if (ocrData.text.trim()) return ocrData;
  } catch (error) {
    setStatus("تعذر OCR للـPDF، جاري تجربة القراءة النصية...", "busy");
  }

  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    const textItems = content.items
      .map((item) => ({
        text: String(item.str || "").replace(/\s+/g, " ").trim(),
        x: item.transform[4],
        y: item.transform[5],
        width: item.width || 0,
      }))
      .filter((item) => item.text);

    const rows = [];
    const yTolerance = 4;
    textItems
      .sort((a, b) => b.y - a.y)
      .forEach((item) => {
        let row = rows.find((existing) => Math.abs(existing.y - item.y) <= yTolerance);
        if (!row) {
          row = { y: item.y, items: [] };
          rows.push(row);
        }
        row.items.push(item);
        row.y = (row.y * (row.items.length - 1) + item.y) / row.items.length;
      });

    let lines = orderPdfItemsByRightThenLeftGroups(textItems, viewport.width);
    if (!lines.length) {
      const lineParts = rows.flatMap((row) =>
        splitRowIntoVisualLines({ ...row, items: row.items, page: pageNumber }, viewport.width).map((part) => ({
          ...part,
          page: pageNumber,
        })),
      );
      lines = orderPdfLinesRightColumnFirst(lineParts, viewport.width);
    }
    if (!lines.length) {
      const fallbackRows = new Map();
      textItems.forEach((item) => {
        const y = Math.round(item.y);
        const existing = fallbackRows.get(y) || [];
        existing.push(item);
        fallbackRows.set(y, existing);
      });
      lines = [...fallbackRows.entries()]
        .sort((a, b) => b[0] - a[0])
        .map(([, items]) =>
          items
            .sort((a, b) => b.x - a.x)
            .map((item) => item.text)
            .join(" ")
            .replace(/\s+/g, " ")
            .trim(),
        )
        .filter(Boolean);
    }

    pages.push({ pageNumber, lines, text: lines.join("\n") });
  }

  return {
    source: "text",
    pages,
    lines: pages.flatMap((page) => page.lines.map((line) => ({ page: page.pageNumber, line }))),
    text: pages.map((page) => page.text).join("\n"),
  };
}

function buildBranchSections(tasks, pdfData) {
  const branches = [...new Set(tasks.map((task) => task.branch))];
  const normalizedPdf = normalizeArabic(pdfData.text);
  const branchPositions = branches
    .map((branch) => ({ branch, position: normalizedPdf.indexOf(normalizeArabic(branch)) }))
    .filter((item) => item.position >= 0)
    .sort((a, b) => a.position - b.position);

  const sections = new Map();
  branchPositions.forEach((item, index) => {
    const next = branchPositions[index + 1]?.position ?? normalizedPdf.length;
    sections.set(item.branch, normalizedPdf.slice(item.position, next));
  });

  return { sections, foundBranches: new Set(branchPositions.map((item) => item.branch)) };
}

function bestPdfLine(task, pdfLines, sectionText, strictBranch) {
  const candidates = strictBranch
    ? pdfLines.filter((item) => sectionText.includes(normalizeArabic(item.line)))
    : pdfLines;
  const usable = candidates.length ? candidates : pdfLines;
  let best = { score: 0, line: "", page: "" };

  usable.forEach((item) => {
    const score = similarity(task, item.line);
    if (score > best.score) best = { score, line: item.line, page: item.page };
  });

  return best;
}

function compare(tasks, pdfData) {
  const threshold = Number(els.threshold.value);
  const strictBranch = els.strictBranch.checked;
  const { sections, foundBranches } = buildBranchSections(tasks, pdfData);
  const normalizedFullPdf = normalizeArabic(pdfData.text);
  const checkBranchInPdf = pdfData.source !== "ocr";

  return tasks.map((task) => {
    const branchFound = !checkBranchInPdf || foundBranches.has(task.branch);
    const sectionText = checkBranchInPdf ? sections.get(task.branch) || normalizedFullPdf : normalizedFullPdf;
    const numberText = normalizeArabic(task.number);
    const directTaskMatch = taskExistsInText(task.task, sectionText);
    const best = bestPdfLine(task.task, pdfData.lines, sectionText, checkBranchInPdf && strictBranch && branchFound);
    const numberOk = !task.hasExplicitNumber || sectionText.includes(numberText);
    const taskOk = directTaskMatch || best.score >= threshold;
    const valueOk = !task.value || (taskOk && valueExistsInText(task.value, nearbyTextForValue(best.line, sectionText)));
    const errors = [];

    if (checkBranchInPdf && !branchFound) errors.push("الفرع غير موجود في PDF");
    if (!taskOk) errors.push("المهمة غير مطابقة أو غير موجودة");
    if (!numberOk) errors.push("رقم المهمة غير موجود داخل الفرع");
    if (!valueOk) errors.push("العدد أو النسبة غير مطابقة للمهمة");

    return {
      ...task,
      ok: errors.length === 0,
      errors,
      score: directTaskMatch ? 100 : best.score,
      matchedLine: best.line,
      page: best.page,
    };
  });
}

function updateMetrics() {
  const branchCount = new Set(state.rows.map((row) => row.branch)).size;
  const okCount = state.results.filter((result) => result.ok).length;
  els.branchCount.textContent = branchCount;
  els.taskCount.textContent = state.rows.length;
  els.okCount.textContent = okCount;
  els.errorCount.textContent = Math.max(state.results.length - okCount, 0);
}

function renderResults() {
  const query = normalizeArabic(els.searchBox.value);
  const filtered = state.results.filter((result) => {
    const filterOk =
      state.activeFilter === "all" ||
      (state.activeFilter === "ok" && result.ok) ||
      (state.activeFilter === "error" && !result.ok);
    const searchOk =
      !query ||
      normalizeArabic(`${result.branch} ${result.number} ${result.task} ${result.errors.join(" ")}`).includes(query);
    return filterOk && searchOk;
  });

  els.resultsList.innerHTML = "";
  if (!filtered.length && state.results.length) {
    els.resultsList.innerHTML = '<div class="status">لا توجد نتائج ضمن هذا التصفية.</div>';
    return;
  }

  filtered.forEach((result) => {
    const card = document.createElement("article");
    card.className = `result-card ${result.ok ? "ok" : "error"}`;

    const details = document.createElement("div");
    details.innerHTML = `
      <h2 class="task-title">${escapeHtml(result.task)}</h2>
      <div class="meta">
        <span class="pill">الفرع: ${escapeHtml(result.branch)}</span>
        <span class="pill">الرقم: ${result.hasExplicitNumber ? escapeHtml(result.number) : "تسلسل " + escapeHtml(result.number)}</span>
        ${result.value ? `<span class="pill">العدد/النسبة: ${escapeHtml(result.value)}</span>` : ""}
        ${
          result.ok
            ? '<span class="pill ok">مطابق</span>'
            : result.errors.map((error) => `<span class="pill danger">${escapeHtml(error)}</span>`).join("")
        }
      </div>
    `;

    const score = document.createElement("div");
    score.className = `score ${result.ok ? "ok" : "error"}`;
    score.textContent = `${result.score}%`;

    card.append(details, score);
    if (result.matchedLine) {
      const matched = document.createElement("p");
      matched.className = "matched-line";
      matched.textContent = `أقرب سطر${result.page ? ` - صفحة ${result.page}` : ""}: ${result.matchedLine}`;
      card.append(matched);
    }

    els.resultsList.append(card);
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function updateReadyState() {
  els.compareBtn.disabled = !(state.excelFile && state.pdfFile);
}

async function runComparison() {
  try {
    setStatus("جاري قراءة الملفات والمطابقة...", "busy");
    els.compareBtn.disabled = true;
    els.exportBtn.disabled = true;
    state.rows = await parseWorkbook(state.excelFile);
    if (!state.rows.length) throw new Error("لم يتم العثور على مهام داخل ملف Excel.");
    const pdfData = await parsePdf(state.pdfFile);
    if (!pdfData.text.trim()) throw new Error("لم يتم استخراج نص من PDF. قد يكون الملف صورًا فقط.");

    state.results = compare(state.rows, pdfData);
    updateMetrics();
    renderResults();
    els.exportBtn.disabled = false;
    const errors = state.results.filter((result) => !result.ok).length;
    setStatus(errors ? `اكتملت المطابقة مع ${errors} خطأ.` : "اكتملت المطابقة بدون أخطاء.", errors ? "error" : "");
  } catch (error) {
    setStatus(error.message || "حدث خطأ غير متوقع.", "error");
  } finally {
    updateReadyState();
  }
}

function exportCsv() {
  const headers = ["الفرع", "رقم المهمة", "المهمة", "العدد/النسبة", "الحالة", "الأخطاء", "نسبة المطابقة", "صفحة PDF", "أقرب سطر"];
  const rows = state.results.map((result) => [
    result.branch,
    result.hasExplicitNumber ? result.number : `تسلسل ${result.number}`,
    result.task,
    result.value || "",
    result.ok ? "مطابق" : "خطأ",
    result.errors.join(" | "),
    `${result.score}%`,
    result.page || "",
    result.matchedLine || "",
  ]);

  const csv = [headers, ...rows]
    .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "نتائج-مطابقة-المهام.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function bindFileInput(input, nameEl, key) {
  input.addEventListener("change", () => {
    const file = input.files?.[0] || null;
    state[key] = file;
    nameEl.textContent = file ? file.name : "لم يتم اختيار ملف";
    updateReadyState();
  });
}

function bindDropZone(zoneId, input) {
  const zone = document.querySelector(zoneId);
  ["dragenter", "dragover"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.add("dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.remove("dragging");
    });
  });
  zone.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    input.dispatchEvent(new Event("change"));
  });
}

bindFileInput(els.excelFile, els.excelName, "excelFile");
bindFileInput(els.pdfFile, els.pdfName, "pdfFile");
bindDropZone("#excelDrop", els.excelFile);
bindDropZone("#pdfDrop", els.pdfFile);

els.compareBtn.addEventListener("click", runComparison);
els.clearBtn.addEventListener("click", () => {
  state.excelFile = null;
  state.pdfFile = null;
  state.rows = [];
  state.results = [];
  els.excelFile.value = "";
  els.pdfFile.value = "";
  els.excelName.textContent = "لم يتم اختيار ملف";
  els.pdfName.textContent = "لم يتم اختيار ملف";
  els.resultsList.innerHTML = "";
  els.searchBox.value = "";
  els.exportBtn.disabled = true;
  updateReadyState();
  updateMetrics();
  setStatus("اختر الملفين للبدء.");
});
els.exportBtn.addEventListener("click", exportCsv);
els.searchBox.addEventListener("input", renderResults);
els.threshold.addEventListener("input", () => {
  els.thresholdValue.textContent = `${els.threshold.value}%`;
});

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
    tab.classList.add("active");
    state.activeFilter = tab.dataset.filter;
    renderResults();
  });
});
