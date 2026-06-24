const $ = (id) => document.getElementById(id);
const state = { files: [], running: false, cancelled: false, downloadUrl: null, records: [] };
const DEFAULT_MATCH_THRESHOLD = 0.52;
const SIGNATURE_SIZE = 48;
let lastAutoNamingTemplate = "";

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  updateTemplateMode();
  initializeNamingTemplate();
  const required = ["pdfjsLib", "PDFLib", "JSZip", "XLSX"];
  const missing = required.filter((name) => !window[name]);
  if (missing.length) showToast(`网页组件加载失败：${missing.join("、")}，请检查网络后刷新`);
});

function bindEvents() {
  const fileInput = $("fileInput");
  const folderInput = $("folderInput");
  $("chooseFiles").onclick = (event) => { event.stopPropagation(); fileInput.click(); };
  $("chooseFolder").onclick = (event) => { event.stopPropagation(); folderInput.click(); };
  $("dropZone").onclick = () => fileInput.click();
  $("dropZone").onkeydown = (event) => {
    if (event.key === "Enter" || event.key === " ") { event.preventDefault(); fileInput.click(); }
  };
  [fileInput, folderInput].forEach((input) => input.onchange = () => addFiles(input.files));
  ["dragenter", "dragover"].forEach((name) => $("dropZone").addEventListener(name, (event) => {
    event.preventDefault(); $("dropZone").classList.add("is-dragging");
  }));
  ["dragleave", "drop"].forEach((name) => $("dropZone").addEventListener(name, (event) => {
    event.preventDefault(); $("dropZone").classList.remove("is-dragging");
  }));
  $("dropZone").ondrop = (event) => addFiles(event.dataTransfer.files);
  $("processMode").onchange = updateProcessingMode;
  $("templateMode").onchange = updateTemplateMode;
  $("fieldNames").oninput = () => syncNamingTemplate();
  $("namingSeparator").onchange = applyNamingSeparator;
  $("namingTemplate").oninput = () => {
    $("namingTemplate").dataset.auto = "false";
    $("namingSeparator").value = "custom";
  };
  $("startButton").onclick = startProcessing;
  $("stopButton").onclick = stopProcessing;
  $("clearButton").onclick = clearAll;
}

function addFiles(fileList) {
  const known = new Set(state.files.map((file) => `${file.name}:${file.size}:${file.lastModified}`));
  for (const file of Array.from(fileList || [])) {
    if (!file.name.toLowerCase().endsWith(".pdf")) continue;
    const key = `${file.name}:${file.size}:${file.lastModified}`;
    if (!known.has(key)) { state.files.push(file); known.add(key); }
  }
  renderFileSummary();
  updateProcessingMode();
}

function renderFileSummary() {
  const summary = $("fileSummary");
  summary.hidden = !state.files.length;
  summary.textContent = state.files.length
    ? `已选择 ${state.files.length} 个 PDF，共 ${formatBytes(state.files.reduce((sum, file) => sum + file.size, 0))}：${state.files.map((file) => file.name).join("、")}`
    : "";
  $("startButton").disabled = !state.files.length || state.running;
}

function updateTemplateMode() {
  const selected = $("templateMode").value === "selected";
  $("templatePageField").hidden = !selected;
  $("templatePage").disabled = !selected;
}

function resolveProcessMode() {
  const selected = $("processMode").value;
  if (selected !== "auto") return selected;
  return state.files.length > 1 ? "rename" : "split";
}

function updateProcessingMode() {
  const renameOnly = resolveProcessMode() === "rename";
  $("templateSettings").hidden = renameOnly;
  updateTemplateMode();
}

function parseFieldNames() {
  return Array.from(new Set($("fieldNames").value.split(/[、,，;；\n]+/).map((item) => item.trim()).filter(Boolean)));
}

function selectedNamingSeparator() {
  return { underscore: "_", hyphen: "-", space: " ", none: "" }[$("namingSeparator").value] ?? "_";
}

function buildDefaultNamingTemplate(fields, separator = selectedNamingSeparator()) {
  return fields.map((field) => `{${field}}`).join(separator);
}

function initializeNamingTemplate() {
  const generated = buildDefaultNamingTemplate(parseFieldNames());
  const input = $("namingTemplate");
  if (!input.value.trim() || input.value.trim() === generated) input.value = generated;
  input.dataset.auto = "true";
  lastAutoNamingTemplate = generated;
}

function syncNamingTemplate() {
  const input = $("namingTemplate");
  const generated = buildDefaultNamingTemplate(parseFieldNames());
  if (input.dataset.auto === "true" || !input.value.trim() || input.value.trim() === lastAutoNamingTemplate) {
    input.value = generated;
    input.dataset.auto = "true";
  }
  lastAutoNamingTemplate = generated;
}

function applyNamingSeparator() {
  if ($("namingSeparator").value === "custom") return;
  const generated = buildDefaultNamingTemplate(parseFieldNames());
  $("namingTemplate").value = generated;
  $("namingTemplate").dataset.auto = "true";
  lastAutoNamingTemplate = generated;
}

async function startProcessing() {
  if (state.running || !state.files.length) return;
  const fields = parseFieldNames();
  if (!fields.length) return showToast("请至少输入一个提取字段名称");
  const processMode = resolveProcessMode();
  const templatePage = $("templateMode").value === "selected" ? Number($("templatePage").value) : 1;
  if (processMode === "split" && (!Number.isInteger(templatePage) || templatePage < 1)) return showToast("模板页码必须是大于 0 的整数");
  const expectedCountRaw = $("expectedCount").value.trim();
  const expectedCount = processMode === "split" && expectedCountRaw ? Number(expectedCountRaw) : null;
  if (expectedCount !== null && (!Number.isInteger(expectedCount) || expectedCount < 1)) return showToast("预期拆分份数必须是大于 0 的整数");
  const namingTemplate = $("namingTemplate").value.trim() || buildDefaultNamingTemplate(fields);
  const placeholders = Array.from(namingTemplate.matchAll(/\{([^{}]+)\}/g), (match) => match[1].trim());
  const unknownFields = Array.from(new Set(placeholders.filter((field) => !fields.includes(field))));
  if (unknownFields.length) return showToast(`文件名规则包含未配置字段：${unknownFields.join("、")}`);

  state.running = true;
  state.cancelled = false;
  state.records = [];
  setRunningUI(true);
  resetResults();
  $("progressCard").hidden = false;
  const zip = new JSZip();
  const usedNames = new Set();

  try {
    let fileIndex = 0;
    for (const file of state.files) {
      checkCancelled();
      fileIndex += 1;
      log(`[${fileIndex}/${state.files.length}] ${processMode === "rename" ? "识别并重命名" : "分析并拆分"} ${file.name}`);
      await processOnePdf(file, templatePage, fields, zip, usedNames, fileIndex, processMode, namingTemplate, expectedCount);
    }
    if (!state.records.length) throw new Error("没有生成可下载的 PDF");
    addExcelCatalog(zip, state.records, fields);
    setProgress(94, "正在打包下载文件");
    const blob = await zip.generateAsync(
      { type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } },
      (meta) => setProgress(94 + meta.percent * 0.06, "正在打包下载文件"),
    );
    if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
    state.downloadUrl = URL.createObjectURL(blob);
    $("downloadButton").href = state.downloadUrl;
    $("downloadButton").download = "PDF处理结果.zip";
    $("downloadButton").classList.remove("is-disabled");
    $("downloadButton").setAttribute("aria-disabled", "false");
    setProgress(100, `处理完成，共生成 ${state.records.length} 个 PDF`);
    renderResults();
    showToast("处理完成，请点击“下载结果”");
  } catch (error) {
    if (state.cancelled) log("任务已停止");
    else { log(`处理失败：${error.message}`); showToast(error.message); }
  } finally {
    state.running = false;
    setRunningUI(false);
  }
}

async function processOnePdf(file, templatePage, fields, zip, usedNames, fileIndex, processMode, namingTemplate, expectedCount) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const pdfjs = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
  if (processMode === "split" && templatePage > pdfjs.numPages) throw new Error(`${file.name} 只有 ${pdfjs.numPages} 页，模板页 ${templatePage} 不存在`);
  const sourcePdf = await PDFLib.PDFDocument.load(bytes, { ignoreEncryption: false });
  let uniqueStarts = [1];

  if (processMode === "split") {
    const templateSignature = await pageSignature(await pdfjs.getPage(templatePage));
    const pageScores = [];
    for (let pageNumber = 1; pageNumber <= pdfjs.numPages; pageNumber += 1) {
      checkCancelled();
      const page = await pdfjs.getPage(pageNumber);
      const signature = await pageSignature(page);
      const similarity = signatureSimilarity(templateSignature, signature);
      pageScores.push({ page: pageNumber, score: similarity });
      const base = ((fileIndex - 1) / state.files.length) * 55;
      const share = (pageNumber / pdfjs.numPages) * (55 / state.files.length);
      setProgress(base + share, `${file.name}：快速扫描 ${pageNumber}/${pdfjs.numPages} 页`);
    }
    const thresholdInfo = chooseAdaptiveThreshold(pageScores, templatePage, expectedCount);
    const starts = [1, ...pageScores.filter((item) => item.page !== 1 && item.score >= thresholdInfo.threshold).map((item) => item.page)];
    uniqueStarts = Array.from(new Set(starts)).sort((a, b) => a - b);
    log(`${file.name}：自适应分界阈值 ${thresholdInfo.threshold.toFixed(2)}，封面组与普通页间隔 ${thresholdInfo.gap.toFixed(2)}`);
    for (const start of uniqueStarts.filter((page) => page !== 1)) {
      const score = pageScores.find((item) => item.page === start)?.score || 0;
      log(`${file.name} 第 ${start} 页：模板相似度 ${score.toFixed(2)}，判定为分界页`);
    }
    log(`${file.name}：找到分界页 ${uniqueStarts.join("、")}`);
    if (expectedCount !== null && uniqueStarts.length !== expectedCount) {
      throw new Error(`${file.name} 检测到 ${uniqueStarts.length} 份，与预期 ${expectedCount} 份不一致，已停止输出。分界页：${uniqueStarts.join("、")}`);
    }
  } else {
    log(`${file.name}：独立文件模式，仅识别第一页，不拆分原 PDF`);
    const base = ((fileIndex - 1) / state.files.length) * 55;
    setProgress(base + (55 / state.files.length), `${file.name}：准备识别第一页`);
  }

  for (let index = 0; index < uniqueStarts.length; index += 1) {
    checkCancelled();
    const start = uniqueStarts[index];
    const end = index + 1 < uniqueStarts.length ? uniqueStarts[index + 1] - 1 : pdfjs.numPages;
    const page = await pdfjs.getPage(start);
    const text = await recognizeBoundaryPage(page, file.name, start, false, fields);
    const title = firstUsefulLine(text);
    let extracted = extractFields(text, fields, title);
    if (start < end && shouldRecognizeFallback(extracted, fields)) {
      const fallbackPageNumber = start + 1;
      log(`${file.name} 第 ${start} 页字段不完整或编号易受印章干扰，补充识别第 ${fallbackPageNumber} 页`);
      const fallbackText = await recognizeBoundaryPage(await pdfjs.getPage(fallbackPageNumber), file.name, fallbackPageNumber, true, fields);
      extracted = mergeFieldValues(extracted, extractFields(fallbackText, fields, firstUsefulLine(fallbackText)), fields);
    }
    const rawStem = renderNamingTemplate(namingTemplate, extracted, fields) || `${stripExtension(file.name)}_${start}-${end}`;
    const filename = uniqueFilename(cleanFilename(rawStem), usedNames);
    const outPdf = await PDFLib.PDFDocument.create();
    const copied = await outPdf.copyPages(sourcePdf, Array.from({ length: end - start + 1 }, (_, offset) => start - 1 + offset));
    copied.forEach((copiedPage) => outPdf.addPage(copiedPage));
    zip.file(filename, await outPdf.save());
    state.records.push({
      index: state.records.length + 1,
      filename,
      source: file.name,
      start,
      end,
      pages: end - start + 1,
      title,
      fields: extracted,
      processedAt: new Date().toLocaleString("zh-CN"),
    });
    const done = index + 1;
    setProgress(55 + (done / uniqueStarts.length) * 38, `${file.name}：生成 ${done}/${uniqueStarts.length} 份`);
  }
  await pdfjs.destroy();
}

async function pageSignature(page) {
  const viewport = page.getViewport({ scale: 0.30 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(viewport.width));
  canvas.height = Math.max(1, Math.round(viewport.height));
  await page.render({ canvasContext: canvas.getContext("2d", { alpha: false }), viewport }).promise;
  const small = document.createElement("canvas");
  small.width = SIGNATURE_SIZE; small.height = SIGNATURE_SIZE;
  const context = small.getContext("2d", { willReadFrequently: true });
  context.drawImage(canvas, 0, 0, SIGNATURE_SIZE, SIGNATURE_SIZE);
  const data = context.getImageData(0, 0, SIGNATURE_SIZE, SIGNATURE_SIZE).data;
  const pixelCount = SIGNATURE_SIZE * SIGNATURE_SIZE;
  const gray = new Uint8Array(pixelCount);
  for (let index = 0; index < pixelCount; index += 1) {
    const offset = index * 4;
    gray[index] = Math.round((data[offset] * 0.299) + (data[offset + 1] * 0.587) + (data[offset + 2] * 0.114));
  }
  const sorted = Array.from(gray).sort((a, b) => a - b);
  const background = sorted[Math.floor(sorted.length * 0.90)];
  const threshold = Math.max(150, Math.min(235, background - 18));
  const mask = Array.from(gray, (value) => value < threshold);
  const rows = Array.from({ length: SIGNATURE_SIZE }, (_, y) => mask.slice(y * SIGNATURE_SIZE, (y + 1) * SIGNATURE_SIZE).filter(Boolean).length / SIGNATURE_SIZE);
  const columns = Array.from({ length: SIGNATURE_SIZE }, (_, x) => {
    let ink = 0;
    for (let y = 0; y < SIGNATURE_SIZE; y += 1) if (mask[(y * SIGNATURE_SIZE) + x]) ink += 1;
    return ink / SIGNATURE_SIZE;
  });
  return { mask, rows, columns, inkRatio: mask.filter(Boolean).length / mask.length };
}

function signatureSimilarity(left, right) {
  let dice = 0;
  for (let shiftY = -1; shiftY <= 1; shiftY += 1) {
    for (let shiftX = -1; shiftX <= 1; shiftX += 1) {
      dice = Math.max(dice, shiftedDice(left.mask, right.mask, shiftX, shiftY));
    }
  }
  const projection = (cosineSimilarity(left.rows, right.rows) + cosineSimilarity(left.columns, right.columns)) / 2;
  const density = Math.min(left.inkRatio, right.inkRatio) / Math.max(left.inkRatio, right.inkRatio, 0.0001);
  return (dice * 0.65) + (projection * 0.20) + (density * 0.15);
}

function chooseAdaptiveThreshold(pageScores, templatePage, expectedCount = null) {
  const sorted = pageScores
    .filter((item) => item.page !== 1 && item.page !== templatePage && item.score >= 0.25 && item.score < 0.96)
    .map((item) => item.score)
    .sort((left, right) => right - left);
  const forcedStarts = templatePage === 1 ? 1 : 2;
  const remainingExpected = expectedCount === null ? null : expectedCount - forcedStarts;
  if (remainingExpected !== null && remainingExpected >= 0 && remainingExpected < sorted.length) {
    if (remainingExpected === 0) return { threshold: 0.97, gap: 1, source: "expected-count" };
    const upper = sorted[remainingExpected - 1];
    const lower = sorted[remainingExpected];
    return { threshold: (upper + lower) / 2, gap: upper - lower, source: "expected-count" };
  }

  let bestGap = 0;
  let threshold = DEFAULT_MATCH_THRESHOLD;
  let lowestReliableThreshold = null;
  let lowestReliableGap = 0;
  for (let index = 0; index < sorted.length - 1; index += 1) {
    const upper = sorted[index];
    const lower = sorted[index + 1];
    const gap = upper - lower;
    // 同类封面可能因扫描深浅形成多个分组。优先取最下层仍有明显断层的分界，
    // 避免只识别颜色较深的一组封面。
    const candidateThreshold = (upper + lower) / 2;
    if (gap >= 0.045 && upper >= 0.36 && lower <= 0.42 && candidateThreshold >= 0.32) {
      if (lowestReliableThreshold === null || candidateThreshold < lowestReliableThreshold) {
        lowestReliableThreshold = candidateThreshold;
        lowestReliableGap = gap;
      }
    }
    if (upper < 0.43 || lower > 0.60) continue;
    if (gap > bestGap) {
      bestGap = gap;
      threshold = candidateThreshold;
    }
  }
  if (lowestReliableThreshold !== null) {
    return { threshold: lowestReliableThreshold, gap: lowestReliableGap, source: "lowest-cover-cluster" };
  }
  if (bestGap < 0.07) return { threshold: DEFAULT_MATCH_THRESHOLD, gap: bestGap };
  return { threshold: Math.max(0.40, Math.min(0.58, threshold)), gap: bestGap };
}

function renderNamingTemplate(template, extracted, fields) {
  const values = Object.fromEntries(fields.map((field) => [field, extracted[field] || ""]));
  const rendered = String(template || "").replace(/\{([^{}]+)\}/g, (_, field) => values[field.trim()] || "");
  return rendered
    .replace(/_{2,}/g, "_")
    .replace(/-{2,}/g, "-")
    .replace(/[\s_-]+$/g, "")
    .replace(/^[\s_-]+/g, "")
    .trim();
}

function shiftedDice(left, right, shiftX, shiftY) {
  let intersection = 0;
  let leftInk = 0;
  let rightInk = 0;
  for (let y = 0; y < SIGNATURE_SIZE; y += 1) {
    for (let x = 0; x < SIGNATURE_SIZE; x += 1) {
      const leftValue = left[(y * SIGNATURE_SIZE) + x];
      const shiftedX = x + shiftX;
      const shiftedY = y + shiftY;
      const rightValue = shiftedX >= 0 && shiftedX < SIGNATURE_SIZE && shiftedY >= 0 && shiftedY < SIGNATURE_SIZE
        ? right[(shiftedY * SIGNATURE_SIZE) + shiftedX]
        : false;
      if (leftValue) leftInk += 1;
      if (rightValue) rightInk += 1;
      if (leftValue && rightValue) intersection += 1;
    }
  }
  return leftInk + rightInk ? (2 * intersection) / (leftInk + rightInk) : 1;
}

function cosineSimilarity(left, right) {
  let product = 0;
  let leftLength = 0;
  let rightLength = 0;
  for (let index = 0; index < left.length; index += 1) {
    product += left[index] * right[index];
    leftLength += left[index] ** 2;
    rightLength += right[index] ** 2;
  }
  return leftLength && rightLength ? product / Math.sqrt(leftLength * rightLength) : 0;
}

async function recognizeBoundaryPage(page, filename, pageNumber, fieldFallback = false, fields = []) {
  const nativeText = await extractNativeText(page);
  if (nativeText.replace(/\s/g, "").length >= 12) {
    log(`${filename} 第 ${pageNumber} 页：直接读取文字层`);
    return { text: nativeText, lines: [] };
  }
  setProgress(null, `${filename} 第 ${pageNumber} 页：上传 PP-OCRv6 高精度识别`);
  // 快速分界扫描只使用 48×48 特征图；真正 OCR 只发生在分界页，
  // 因此这里提高渲染分辨率，优先保证编号和中文名称的识别准确率。
  const viewport = page.getViewport({ scale: fieldFallback ? 2.3 : 2.1 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(viewport.width);
  canvas.height = Math.round(viewport.height);
  await page.render({ canvasContext: canvas.getContext("2d", { alpha: false }), viewport }).promise;
  let recognitionImage = canvas;
  if (fieldFallback) {
    const cropTop = Math.round(canvas.height * 0.16);
    const cropBottom = Math.round(canvas.height * 0.48);
    const crop = document.createElement("canvas");
    crop.width = canvas.width;
    crop.height = cropBottom - cropTop;
    crop.getContext("2d", { alpha: false }).drawImage(canvas, 0, cropTop, canvas.width, crop.height, 0, 0, crop.width, crop.height);
    recognitionImage = crop;
  }

  const blob = await canvasToBlob(recognitionImage, "image/jpeg", 0.92);
  const form = new FormData();
  form.append("image", blob, `${stripExtension(filename)}-${pageNumber}.jpg`);
  let response;
  try {
    response = await fetch("./api/ocr", { method: "POST", body: form });
  } catch (error) {
    throw new Error("高精度 OCR 服务无法连接，已停止处理，未生成错误文件");
  }
  let result = {};
  try { result = await response.json(); } catch {}
  if (!response.ok) {
    throw new Error(result.detail || `高精度 OCR 服务异常（${response.status}）`);
  }
  log(`${filename} 第 ${pageNumber} 页：${result.engine} 平均置信度 ${Number(result.averageConfidence || 0).toFixed(1)}%`);
  return { text: result.text || "", lines: result.lines || [], refined: {}, engine: result.engine };
}

function canvasToBlob(canvas, type, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("无法生成 OCR 页面图片")), type, quality);
  });
}

async function extractNativeText(page) {
  const content = await page.getTextContent();
  let text = "";
  for (const item of content.items) text += `${item.str || ""}${item.hasEOL ? "\n" : " "}`;
  return text.trim();
}

function extractFields(recognized, fields, title) {
  const lines = normalizeOcrText(recognized).split("\n").filter(Boolean);
  const positionedLines = recognized && typeof recognized === "object" ? recognized.lines || [] : [];
  const refined = recognized && typeof recognized === "object" ? recognized.refined || {} : {};
  const values = {};
  for (const name of fields) {
    if (name === "标题") { values[name] = cleanValue(title); continue; }
    const positioned = findSpatialFieldValue(positionedLines, name, fields) || findFieldValue(lines, name, fields);
    const refinedValue = sanitizeFieldValue(refined[name] || "", name);
    values[name] = reconcileRefinedValue(positioned, refinedValue, name);
  }
  return values;
}

function reconcileRefinedValue(positioned, refined, fieldName) {
  if (!refined) return positioned;
  if (!positioned) return refined;
  if (isIdentifierField(fieldName)) {
    return fieldValueScore(refined, fieldName) >= fieldValueScore(positioned, fieldName) - 0.05 ? refined : positioned;
  }
  const positionedCanonical = canonicalLabel(positioned);
  const refinedCanonical = canonicalLabel(refined);
  for (let offset = 1; offset <= 3; offset += 1) {
    const probe = refinedCanonical.slice(offset, offset + 6);
    if (probe.length >= 4 && positionedCanonical.startsWith(probe)) {
      return sanitizeFieldValue(`${refined.slice(0, offset)}${positioned}`, fieldName);
    }
  }
  const similar = textSimilarity(positionedCanonical, refinedCanonical) >= 0.55;
  return similar && refined.length >= positioned.length && fieldValueScore(refined, fieldName) >= fieldValueScore(positioned, fieldName) - 0.05
    ? refined
    : positioned;
}

function selectRefinedFieldValue(text, fieldName, fields) {
  const lines = normalizeOcrText(text)
    .split("\n")
    .map((line) => truncateAtFieldBoundary(line, fieldName, fields))
    .map((line) => line.replace(/第\s*[0-9Iil]+\s*页.*$/i, ""))
    .map((line) => sanitizeFieldValue(line, fieldName))
    .filter((line) => line && !looksLikeAnyFieldLabel(line, fields));
  if (!lines.length) return "";
  if (isIdentifierField(fieldName)) {
    return lines.sort((left, right) => fieldValueScore(right, fieldName) - fieldValueScore(left, fieldName))[0];
  }
  return sanitizeFieldValue(lines.slice(0, 2).join(""), fieldName);
}

function findSpatialFieldValue(lines, fieldName, allFields) {
  let bestValue = "";
  let bestScore = 0;
  for (const labelLine of lines) {
    const labelMatch = fieldLabelMatch(labelLine.text, fieldName);
    if (labelMatch.score < 0.72) continue;
    let value = labelMatch.value;
    if (!value) {
      const labelBox = labelLine.bbox;
      const labelHeight = Math.max(1, labelBox.y1 - labelBox.y0);
      const labelCenterY = (labelBox.y0 + labelBox.y1) / 2;
      const sameRow = lines
        .filter((candidate) => candidate !== labelLine && !looksLikeAnyFieldLabel(candidate.text, allFields))
        .filter((candidate) => {
          const box = candidate.bbox;
          const height = Math.max(1, box.y1 - box.y0);
          const centerY = (box.y0 + box.y1) / 2;
          return box.x0 >= labelBox.x1 - (labelHeight * 0.5) && Math.abs(centerY - labelCenterY) <= Math.max(labelHeight, height) * 1.15;
        })
        .sort((left, right) => left.bbox.x0 - right.bbox.x0);
      const adjacent = [];
      for (const candidate of sameRow) {
        if (!adjacent.length) {
          adjacent.push(candidate);
          continue;
        }
        const previous = adjacent[adjacent.length - 1];
        const gap = candidate.bbox.x0 - previous.bbox.x1;
        if (gap <= Math.max(42, labelHeight * 2.5)) adjacent.push(candidate);
        else break;
      }
      value = adjacent.map((candidate) => candidate.text).join(" ");
      if (adjacent.length) {
        const lastValueLine = adjacent[adjacent.length - 1];
        const continuation = lines
          .filter((candidate) => candidate !== labelLine && !adjacent.includes(candidate) && !looksLikeAnyFieldLabel(candidate.text, allFields))
          .filter((candidate) => {
            const box = candidate.bbox;
            return box.y0 > lastValueLine.bbox.y0
              && box.y0 - lastValueLine.bbox.y1 <= labelHeight * 1.6
              && box.x0 >= labelBox.x0 - labelHeight
              && box.x0 <= lastValueLine.bbox.x1 + (labelHeight * 4)
              && !/^第\s*[0-9Iil]+\s*页/i.test(candidate.text);
          })
          .sort((left, right) => (left.bbox.y0 - right.bbox.y0) || (left.bbox.x0 - right.bbox.x0))[0];
        if (continuation) value = `${value}${needsWordSeparator(value, continuation.text) ? " " : ""}${continuation.text}`;
      }
      if (!value) {
        const below = lines
          .filter((candidate) => candidate !== labelLine && !looksLikeAnyFieldLabel(candidate.text, allFields))
          .filter((candidate) => candidate.bbox.y0 >= labelBox.y0 && candidate.bbox.y0 - labelBox.y1 <= labelHeight * 3)
          .sort((left, right) => (left.bbox.y0 - right.bbox.y0) || (left.bbox.x0 - right.bbox.x0));
        value = below[0]?.text || "";
      }
    }
    value = sanitizeFieldValue(truncateAtFieldBoundary(value, fieldName, allFields), fieldName);
    const score = (labelMatch.score * 2) + fieldValueScore(value, fieldName);
    if (value && score > bestScore) {
      bestScore = score;
      bestValue = value;
    }
  }
  return bestValue;
}

function fieldLabelMatch(line, fieldName) {
  const exactIndex = line.indexOf(fieldName);
  if (exactIndex >= 0) {
    return { score: 1, value: line.slice(exactIndex + fieldName.length).replace(/^\s*[:：;,，]?\s*/, "") };
  }
  const delimiterIndex = line.search(/[:：;,，]/);
  const prefix = delimiterIndex >= 0 ? line.slice(0, delimiterIndex) : line.slice(0, Math.max(fieldName.length + 2, fieldName.length * 2));
  const target = canonicalLabel(fieldName);
  const candidate = canonicalLabel(prefix);
  const contains = candidate.includes(target)
    || (target.includes(candidate) && target.length - candidate.length <= 1);
  return { score: contains ? 0.96 : textSimilarity(target, candidate), value: delimiterIndex >= 0 ? line.slice(delimiterIndex + 1) : "" };
}

function looksLikeAnyFieldLabel(line, fields) {
  if (fields.some((field) => fieldLabelMatch(line, field).score >= 0.72)) return true;
  return /^[\u3400-\u9fff]{2,12}\s*[:：]/.test(line);
}

function findFieldValue(lines, fieldName, allFields) {
  const target = canonicalLabel(fieldName);
  let bestValue = "";
  let bestScore = 0;
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const exactIndex = line.indexOf(fieldName);
    let value = "";
    let labelScore = 0;

    if (exactIndex >= 0) {
      labelScore = 1;
      value = line.slice(exactIndex + fieldName.length).replace(/^\s*[:：;,，]?\s*/, "");
    } else {
      const delimiterIndex = line.search(/[:：;,，]/);
      const prefixLength = Math.min(line.length, Math.max(fieldName.length + 2, fieldName.length * 2));
      const prefix = delimiterIndex >= 0 ? line.slice(0, delimiterIndex) : line.slice(0, prefixLength);
      const canonicalPrefix = canonicalLabel(prefix);
      const containsTarget = canonicalPrefix.includes(target)
        || (target.includes(canonicalPrefix) && target.length - canonicalPrefix.length <= 1);
      labelScore = containsTarget ? 0.96 : textSimilarity(target, canonicalPrefix);
      value = delimiterIndex >= 0 ? line.slice(delimiterIndex + 1) : line.slice(prefix.length);
    }
    if (labelScore < 0.72) continue;

    value = truncateAtFieldBoundary(cleanValue(value), fieldName, allFields);
    const nextLine = lines[index + 1] || "";
    if ((!value || needsContinuation(value)) && nextLine && !looksLikeNewField(nextLine, allFields)) {
      value = cleanValue(`${value}${needsWordSeparator(value, nextLine) ? " " : ""}${nextLine}`);
    }
    value = sanitizeFieldValue(value, fieldName);
    const score = (labelScore * 2) + fieldValueScore(value, fieldName);
    if (value && score > bestScore) {
      bestScore = score;
      bestValue = value;
    }
  }
  return bestValue;
}

function truncateAtFieldBoundary(value, currentField, allFields) {
  let end = value.length;
  for (const field of allFields) {
    if (field === currentField) continue;
    const index = value.indexOf(field);
    if (index >= 0) end = Math.min(end, index);
  }
  const genericBoundary = value.search(/\s+(?:第\s*\d+\s*页|承包单位|清单编号|细目名称|计算单编号|序号|单位|数量)\s*[:：]?/);
  if (genericBoundary >= 0) end = Math.min(end, genericBoundary);
  return cleanValue(value.slice(0, end));
}

function normalizeOcrText(text) {
  const rawText = text && typeof text === "object" ? text.text || "" : text;
  return String(rawText || "")
    .replace(/\f/g, "\n")
    .split(/\r?\n/)
    .map(normalizeOcrLine)
    .filter(Boolean)
    .join("\n");
}

function normalizeOcrLine(value) {
  let line = String(value || "").normalize("NFKC").replace(/[\t\u00a0]+/g, " ").trim();
  let previous = "";
  while (line !== previous) {
    previous = line;
    line = line
      .replace(/([\u3400-\u9fff])\s+(?=[\u3400-\u9fffA-Za-z0-9])/g, "$1")
      .replace(/([A-Za-z0-9])\s+(?=[\u3400-\u9fff])/g, "$1");
  }
  return line.replace(/\s*([:：;,，~～])\s*/g, "$1").replace(/[ ]{2,}/g, " ").trim();
}

function canonicalLabel(value) {
  const replacements = {
    "項": "项", "編": "编", "號": "号", "稱": "称", "現": "现",
    "頁": "页", "圖": "图", "單": "单", "節": "节", "業": "业",
    "處": "处", "類": "类", "結": "结", "測": "测", "設": "设",
    "說": "说", "層": "层",
  };
  return normalizeOcrLine(value)
    .split("")
    .map((char) => replacements[char] || char)
    .join("")
    .replace(/[^\p{L}\p{N}]/gu, "")
    .toLowerCase();
}

function textSimilarity(left, right) {
  if (!left || !right) return 0;
  const distance = levenshteinDistance(left, right);
  return 1 - (distance / Math.max(left.length, right.length));
}

function levenshteinDistance(left, right) {
  const previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      const cost = left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1;
      current[rightIndex] = Math.min(current[rightIndex - 1] + 1, previous[rightIndex] + 1, previous[rightIndex - 1] + cost);
    }
    for (let index = 0; index < current.length; index += 1) previous[index] = current[index];
  }
  return previous[right.length];
}

function needsContinuation(value) {
  if (!value || value.length < 3) return true;
  if (/[~～(（\-—]$/.test(value)) return true;
  const opens = (value.match(/[(（]/g) || []).length;
  const closes = (value.match(/[)）]/g) || []).length;
  return opens > closes;
}

function needsWordSeparator(left, right) {
  return /[A-Za-z0-9]$/.test(left) && /^[A-Za-z0-9]/.test(right);
}

function looksLikeNewField(line, fields) {
  if (/^(第?\s*\d+\s*页|序号|清单|细目|单位|数量|日期|承包单位|工程名称)/.test(line)) return true;
  const prefix = line.split(/[:：;,，]/, 1)[0];
  return fields.some((field) => textSimilarity(canonicalLabel(field), canonicalLabel(prefix)) >= 0.72);
}

function sanitizeFieldValue(value, fieldName) {
  let result = cleanValue(value)
    .replace(/^[\s:：;,，|]+/, "")
    .replace(/[\s:：;,，|]+$/, "")
    .replace(/\s*([~～()（）./+\-—])\s*/g, "$1")
    .replace(/[～—]/g, (char) => char === "～" ? "~" : "-");
  result = result
    .replace(/(\d)\s*一\s*([A-Za-z]{0,3}\d)/g, "$1~$2")
    .replace(/\([^0-9A-Za-z()]{0,6}(\d+\+\d+(?:\.\d+)?)~([A-Za-z]{1,3})(\d+\+\d+(?:\.\d+)?)/g, "($2$1~$2$3")
    .replace(/联道/g, "隧道")
    .replace(/错杆/g, "锚杆")
    .replace(/(回填|C20)[雁奏丰]/g, "$1砼")
    .replace(/\bI[Tl1]\s*型/g, "I型")
    .replace(/S-Va[lI1](?=[)）\-]|$)/g, "S-Va")
    .replace(/砼[A-Za-z]{1,3}$/g, "砼")
    .replace(/锚$/g, "锚杆")
    .replace(/暗杆|漏杆|扭杆/g, "锚杆")
    .replace(/涌道|疏道/g, "隧道")
    .replace(/(回填|C\d{2})[雁奏丰幸]/g, "$1砼")
    .replace(/混凝[土士]/g, "混凝土")
    .replace(/\s*第\s*\d+\s*[页頁]\s*$/i, "")
    .replace(/\s*[Pp][Aa][Gg][Ee]\s*\d+\s*$/i, "");
  if (isIdentifierField(fieldName)) {
    result = result
      .replace(/\s+/g, "")
      .replace(/(\d)[Oo](?=\d)/g, (_, digit) => `${digit}0`)
      .replace(/(\d)[lI](?=\d)/g, (_, digit) => `${digit}1`)
      .replace(/^[Oo](?=\d{3,})/g, "0");
  }
  return result.trim().slice(0, 150);
}

function isIdentifierField(name) {
  return /(编号|编码|代码|合同号|证号|单号|流水号|id)$/i.test(canonicalLabel(name));
}

function fieldValueScore(value, fieldName) {
  const compact = String(value || "").replace(/\s+/g, "");
  if (!compact) return 0;
  const lengthScore = Math.min(1, compact.length / 12);
  if (isIdentifierField(fieldName)) {
    const allowed = (compact.match(/[A-Za-z0-9._/#()-]/g) || []).length / compact.length;
    const digit = /\d/.test(compact) ? 0.15 : 0;
    const separator = /[-_/]/.test(compact) ? 0.10 : 0;
    const chinesePenalty = (compact.match(/[\u3400-\u9fff]/g) || []).length / compact.length;
    return (allowed * 0.65) + digit + separator + (lengthScore * 0.10) - (chinesePenalty * 0.35);
  }
  const meaningful = (compact.match(/[\p{L}\p{N}]/gu) || []).length / compact.length;
  return (meaningful * 0.65) + (lengthScore * 0.35);
}

function shouldRecognizeFallback(extracted, fields) {
  return fields.some((field) => fieldValueScore(extracted[field], field) < (isIdentifierField(field) ? 0.72 : 0.55));
}

function mergeFieldValues(primary, fallback, fields) {
  const merged = { ...primary };
  for (const field of fields) {
    const primaryScore = fieldValueScore(primary[field], field);
    const fallbackScore = fieldValueScore(fallback[field], field);
    const valuesAreSimilar = primary[field] && fallback[field]
      ? textSimilarity(canonicalLabel(primary[field]), canonicalLabel(fallback[field])) >= 0.55
      : false;
    const fallbackIsMoreComplete = String(fallback[field] || "").length >= String(primary[field] || "").length;
    const shouldUseFallback = isIdentifierField(field)
      ? fallbackScore > primaryScore
      : ((!primary[field] || primaryScore < 0.70) && fallbackScore > primaryScore)
        || (valuesAreSimilar && fallbackIsMoreComplete && fallbackScore >= primaryScore - 0.05);
    if (fallback[field] && shouldUseFallback) merged[field] = fallback[field];
  }
  return merged;
}

function addExcelCatalog(zip, records, fields) {
  const rows = records.map((record) => ({
    序号: record.index,
    文件名: record.filename,
    原PDF: record.source,
    开始页: record.start,
    结束页: record.end,
    页数: record.pages,
    识别标题: record.title,
    ...Object.fromEntries(fields.map((name) => [name, record.fields[name] || ""])),
    处理时间: record.processedAt,
  }));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "PDF目录");
  zip.file("PDF目录.xlsx", XLSX.write(workbook, { type: "array", bookType: "xlsx" }));
}

function renderResults() {
  $("resultCount").textContent = `${state.records.length} 项`;
  $("emptyState").hidden = Boolean(state.records.length);
  $("resultTableWrap").hidden = !state.records.length;
  $("resultBody").innerHTML = state.records.map((record) => `
    <tr><td>${record.index}</td><td>${escapeHtml(record.filename)}</td><td>${escapeHtml(record.source)}</td>
    <td>${record.start}–${record.end}</td><td>${record.pages}</td><td>${escapeHtml(record.title)}</td>
    <td>${escapeHtml(Object.entries(record.fields).filter(([, value]) => value).map(([key, value]) => `${key}：${value}`).join("；"))}</td></tr>`).join("");
}

async function stopProcessing() {
  state.cancelled = true;
  $("progressLabel").textContent = "正在停止";
}

function clearAll() {
  if (state.running) return showToast("请先停止当前任务");
  state.files = [];
  state.records = [];
  $("fileInput").value = "";
  $("folderInput").value = "";
  $("progressCard").hidden = true;
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = null;
  $("downloadButton").removeAttribute("href");
  $("downloadButton").classList.add("is-disabled");
  renderFileSummary();
  updateProcessingMode();
  resetResults();
}

function setRunningUI(running) {
  $("startButton").disabled = running || !state.files.length;
  $("stopButton").hidden = !running;
  $("clearButton").disabled = running;
}

function setProgress(value, label) {
  if (value !== null) {
    const safe = Math.max(0, Math.min(100, value));
    $("progressBar").style.width = `${safe}%`;
    $("progressValue").textContent = `${Math.round(safe)}%`;
  }
  $("progressLabel").textContent = label;
}

function log(message) {
  const output = $("logOutput");
  output.textContent += `${message}\n`;
  output.scrollTop = output.scrollHeight;
}

function resetResults() {
  $("resultCount").textContent = "0 项";
  $("emptyState").hidden = false;
  $("resultTableWrap").hidden = true;
  $("resultBody").innerHTML = "";
  $("logOutput").textContent = "";
  setProgress(0, "准备处理");
}

function checkCancelled() { if (state.cancelled) throw new Error("任务已停止"); }
function stripExtension(name) { return name.replace(/\.pdf$/i, ""); }
function firstUsefulLine(text) { return normalizeOcrText(text).split("\n").find((line) => line.length >= 2 && line.length <= 100) || ""; }
function cleanValue(value) { return normalizeOcrLine(value).replace(/\s+/g, " ").trim(); }
function cleanFilename(value) { return `${String(value || "未命名").replace(/[\\/:*?"<>|\r\n\t]+/g, "_").replace(/\s+/g, " ").replace(/_+/g, "_").trim().slice(0, 180) || "未命名"}.pdf`; }
function uniqueFilename(filename, used) { const stem = filename.replace(/\.pdf$/i, ""); let result = filename; let index = 1; while (used.has(result)) result = `${stem}_${String(index++).padStart(2, "0")}.pdf`; used.add(result); return result; }
function escapeRegex(value) { return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]); }
function formatBytes(bytes) { return bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(1)} MB`; }

let toastTimer;
function showToast(message) { const toast = $("toast"); toast.textContent = message; toast.classList.add("is-visible"); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 3600); }
