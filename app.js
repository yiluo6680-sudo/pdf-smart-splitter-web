const $ = (id) => document.getElementById(id);
const state = { files: [], running: false, cancelled: false, worker: null, downloadUrl: null, records: [] };
const MATCH_THRESHOLD = 0.82;

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  updateTemplateMode();
  const required = ["pdfjsLib", "PDFLib", "JSZip", "XLSX", "Tesseract"];
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
  $("templateMode").onchange = updateTemplateMode;
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

function parseFieldNames() {
  return Array.from(new Set($("fieldNames").value.split(/[、,，;；\n]+/).map((item) => item.trim()).filter(Boolean)));
}

async function startProcessing() {
  if (state.running || !state.files.length) return;
  const fields = parseFieldNames();
  if (!fields.length) return showToast("请至少输入一个提取字段名称");
  const templatePage = $("templateMode").value === "selected" ? Number($("templatePage").value) : 1;
  if (!Number.isInteger(templatePage) || templatePage < 1) return showToast("模板页码必须是大于 0 的整数");

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
      log(`[${fileIndex}/${state.files.length}] 分析 ${file.name}`);
      await processOnePdf(file, templatePage, fields, zip, usedNames, fileIndex);
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
    if (state.worker) { await state.worker.terminate(); state.worker = null; }
    state.running = false;
    setRunningUI(false);
  }
}

async function processOnePdf(file, templatePage, fields, zip, usedNames, fileIndex) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const pdfjs = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
  if (templatePage > pdfjs.numPages) throw new Error(`${file.name} 只有 ${pdfjs.numPages} 页，模板页 ${templatePage} 不存在`);
  const sourcePdf = await PDFLib.PDFDocument.load(bytes, { ignoreEncryption: false });
  const templateSignature = await pageSignature(await pdfjs.getPage(templatePage));
  const starts = [1];

  for (let pageNumber = 1; pageNumber <= pdfjs.numPages; pageNumber += 1) {
    checkCancelled();
    const page = await pdfjs.getPage(pageNumber);
    const signature = await pageSignature(page);
    if (pageNumber !== 1 && signatureSimilarity(templateSignature, signature) >= MATCH_THRESHOLD) starts.push(pageNumber);
    const base = ((fileIndex - 1) / state.files.length) * 55;
    const share = (pageNumber / pdfjs.numPages) * (55 / state.files.length);
    setProgress(base + share, `${file.name}：快速扫描 ${pageNumber}/${pdfjs.numPages} 页`);
  }
  const uniqueStarts = Array.from(new Set(starts)).sort((a, b) => a - b);
  log(`${file.name}：找到分界页 ${uniqueStarts.join("、")}`);

  for (let index = 0; index < uniqueStarts.length; index += 1) {
    checkCancelled();
    const start = uniqueStarts[index];
    const end = index + 1 < uniqueStarts.length ? uniqueStarts[index + 1] - 1 : pdfjs.numPages;
    const page = await pdfjs.getPage(start);
    const text = await recognizeBoundaryPage(page, file.name, start);
    const title = firstUsefulLine(text);
    const extracted = extractFields(text, fields, title);
    const rawStem = fields.map((name) => extracted[name]).filter(Boolean).join("_") || `${stripExtension(file.name)}_${start}-${end}`;
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
  const viewport = page.getViewport({ scale: 0.32 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(viewport.width));
  canvas.height = Math.max(1, Math.round(viewport.height));
  await page.render({ canvasContext: canvas.getContext("2d", { alpha: false }), viewport }).promise;
  const small = document.createElement("canvas");
  small.width = 32; small.height = 32;
  const context = small.getContext("2d", { willReadFrequently: true });
  context.drawImage(canvas, 0, 0, 32, 32);
  const data = context.getImageData(0, 0, 32, 32).data;
  let mean = 0;
  const gray = new Uint8Array(1024);
  for (let index = 0; index < 1024; index += 1) {
    const offset = index * 4;
    gray[index] = Math.round((data[offset] * 0.299) + (data[offset + 1] * 0.587) + (data[offset + 2] * 0.114));
    mean += gray[index];
  }
  mean /= 1024;
  const threshold = Math.min(245, mean - 7);
  return Array.from(gray, (value) => value < threshold);
}

function signatureSimilarity(left, right) {
  let same = 0;
  for (let index = 0; index < left.length; index += 1) if (left[index] === right[index]) same += 1;
  return same / left.length;
}

async function recognizeBoundaryPage(page, filename, pageNumber) {
  const nativeText = await extractNativeText(page);
  if (nativeText.replace(/\s/g, "").length >= 12) {
    log(`${filename} 第 ${pageNumber} 页：直接读取文字层`);
    return nativeText;
  }
  setProgress(null, `${filename} 第 ${pageNumber} 页：OCR 识别中`);
  if (!state.worker) {
    log("首次 OCR 正在下载中文识别模型，请稍候");
    state.worker = await Tesseract.createWorker("chi_sim+eng", 1, {
      logger: (message) => {
        if (message.status === "recognizing text") {
          $("progressLabel").textContent = `OCR 识别中 ${Math.round((message.progress || 0) * 100)}%`;
        }
      },
    });
  }
  const viewport = page.getViewport({ scale: 2 });
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(viewport.width);
  canvas.height = Math.round(viewport.height);
  await page.render({ canvasContext: canvas.getContext("2d", { alpha: false }), viewport }).promise;
  const result = await state.worker.recognize(canvas);
  return result.data.text || "";
}

async function extractNativeText(page) {
  const content = await page.getTextContent();
  let text = "";
  for (const item of content.items) text += `${item.str || ""}${item.hasEOL ? "\n" : " "}`;
  return text.trim();
}

function extractFields(text, fields, title) {
  const values = {};
  for (const name of fields) {
    if (name === "标题") { values[name] = title; continue; }
    const escaped = escapeRegex(name);
    const sameLine = text.match(new RegExp(`${escaped}[ \\t]*[:：]?[ \\t]*([^\\r\\n]+)`, "i"));
    const nextLine = text.match(new RegExp(`${escaped}[ \\t]*[:：]?[ \\t]*\\r?\\n[ \\t]*([^\\r\\n]+)`, "i"));
    values[name] = cleanValue((sameLine || nextLine || ["", ""])[1]);
  }
  return values;
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
  if (state.worker) { await state.worker.terminate(); state.worker = null; }
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
function firstUsefulLine(text) { return text.split(/\r?\n/).map((line) => line.trim()).find((line) => line.length >= 2 && line.length <= 100) || ""; }
function cleanValue(value) { return String(value || "").replace(/\s+/g, " ").trim(); }
function cleanFilename(value) { return `${String(value || "未命名").replace(/[\\/:*?"<>|\r\n\t]+/g, "_").replace(/\s+/g, " ").replace(/_+/g, "_").trim().slice(0, 180) || "未命名"}.pdf`; }
function uniqueFilename(filename, used) { const stem = filename.replace(/\.pdf$/i, ""); let result = filename; let index = 1; while (used.has(result)) result = `${stem}_${String(index++).padStart(2, "0")}.pdf`; used.add(result); return result; }
function escapeRegex(value) { return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function escapeHtml(value) { return String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[char]); }
function formatBytes(bytes) { return bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(1)} MB`; }

let toastTimer;
function showToast(message) { const toast = $("toast"); toast.textContent = message; toast.classList.add("is-visible"); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 3600); }
