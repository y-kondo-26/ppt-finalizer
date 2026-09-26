// UI wiring: DOM state, event handlers, and orchestration. The PPTX
// analysis/mutation engine lives in pptx-core.js (loaded before this file)
// so its pure logic can be unit tested without a DOM.

const CUSTOM_VALUE = "__custom__";
const PPTX_TYPES = [{
  description: "PowerPoint",
  accept: { "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"] },
}];
const PRESET_FONTS = [
  "Noto Sans JP",
  "Meiryo",
  "Yu Gothic UI",
  "Segoe UI",
  "Arial",
  "Calibri",
];
const PRESET_SET = new Set(PRESET_FONTS);
const SETTINGS_STORAGE_KEY = "ppt-finalizer-settings";

const supportsFsAccess = "showOpenFilePicker" in window;

let selectedFile = null;
let sourceFileHandle = null;
let parentDirHandle = null;
let pptxFonts = [];
let mediaAnalysis = null;
let cleanupPlan = null;
let currentFileSize = 0;
let selectedTitleFont = DEFAULT_FONT;
let selectedBodyFont = DEFAULT_FONT;
let activeSideTab = "analysis";
let hasResultContent = false;
let pptxZipCache = null;
const mediaThumbUrlCache = new Map();

const fontOptionsPanel = document.getElementById("fontOptionsPanel");
const optReplaceFonts = document.getElementById("optReplaceFonts");
const titleFontSelect = document.getElementById("titleFontSelect");
const titleCustomFontInput = document.getElementById("titleCustomFontInput");
const bodyFontSelect = document.getElementById("bodyFontSelect");
const bodyCustomFontInput = document.getElementById("bodyCustomFontInput");
const scanningMsg = document.getElementById("scanningMsg");
const dropzone = document.getElementById("dropzone");
const openBtn = document.getElementById("openBtn");
const pickBtn = document.getElementById("pickBtn");
const fileInput = document.getElementById("fileInput");
const fileNameEl = document.getElementById("fileName");
const fileModeEl = document.getElementById("fileMode");
const overwriteBtn = document.getElementById("overwriteBtn");
const downloadBtn = document.getElementById("downloadBtn");
const clearBtn = document.getElementById("clearBtn");
const errorMsg = document.getElementById("errorMsg");
const resultPanel = document.getElementById("resultPanel");
const statsEl = document.getElementById("stats");
const fsNote = document.getElementById("fsNote");
const analysisStack = document.getElementById("analysisStack");
const sidePlaceholder = document.getElementById("sidePlaceholder");
const sideTabs = document.getElementById("sideTabs");
const sideTabAnalysis = document.getElementById("sideTabAnalysis");
const sideTabResult = document.getElementById("sideTabResult");
const sideTabAnalysisBtn = document.getElementById("sideTabAnalysisBtn");
const sideTabResultBtn = document.getElementById("sideTabResultBtn");
const resultTabBadge = document.getElementById("resultTabBadge");
const resultEmpty = document.getElementById("resultEmpty");
const fileAnalysisSize = document.getElementById("fileAnalysisSize");
const fileAnalysisSlides = document.getElementById("fileAnalysisSlides");
const fileAnalysisAfterSize = document.getElementById("fileAnalysisAfterSize");
const fileSizeWarning = document.getElementById("fileSizeWarning");
const reductionBlock = document.getElementById("reductionBlock");
const reductionBar = document.getElementById("reductionBar");
const fileAnalysisEstimateBar = document.getElementById("fileAnalysisEstimateBar");
const titleFontPreview = document.getElementById("titleFontPreview");
const bodyFontPreview = document.getElementById("bodyFontPreview");
const fontAnalysisBadge = document.getElementById("fontAnalysisBadge");
const fontSummary = document.getElementById("fontSummary");
const fontTableBody = document.getElementById("fontTableBody");
const fontEmpty = document.getElementById("fontEmpty");
const mediaAnalysisBadge = document.getElementById("mediaAnalysisBadge");
const mediaSummary = document.getElementById("mediaSummary");
const mediaTableBody = document.getElementById("mediaTableBody");
const mediaEmpty = document.getElementById("mediaEmpty");
const cleanupPanel = document.getElementById("cleanupPanel");
const optRemoveOrphanMedia = document.getElementById("optRemoveOrphanMedia");
const optCompressImages = document.getElementById("optCompressImages");
const imagePpiSelect = document.getElementById("imagePpiSelect");
const jpegQualitySelect = document.getElementById("jpegQualitySelect");
const compressImageOptions = document.getElementById("compressImageOptions");
const optRemoveUnusedStructure = document.getElementById("optRemoveUnusedStructure");
const optRemoveNotes = document.getElementById("optRemoveNotes");
const optRemoveProperties = document.getElementById("optRemoveProperties");
const uncheckAllOptionsBtn = document.getElementById("uncheckAllOptionsBtn");
const cleanupOptionCheckboxes = [
  optRemoveOrphanMedia,
  optCompressImages,
  optRemoveUnusedStructure,
  optRemoveNotes,
  optRemoveProperties,
];
const orphanMediaPreview = document.getElementById("orphanMediaPreview");
const compressImagesPreview = document.getElementById("compressImagesPreview");
const structurePreview = document.getElementById("structurePreview");
const structureOptionDesc = document.getElementById("structureOptionDesc");
const notesPreview = document.getElementById("notesPreview");
const propertiesPreview = document.getElementById("propertiesPreview");
const mediaThumbTooltip = document.getElementById("mediaThumbTooltip");
const mediaThumbImg = mediaThumbTooltip.querySelector("img");
const mediaThumbFallback = mediaThumbTooltip.querySelector(".media-thumb-fallback");

if (!supportsFsAccess) {
  openBtn.hidden = true;
  pickBtn.hidden = false;
  overwriteBtn.hidden = true;
  fsNote.textContent = "このブラウザでは別名ダウンロードのみ利用できます。上書き保存は Chrome / Edge をご利用ください。";
} else {
  overwriteBtn.hidden = false;
}

function revokeMediaThumbUrls() {
  for (const url of mediaThumbUrlCache.values()) {
    URL.revokeObjectURL(url);
  }
  mediaThumbUrlCache.clear();
  mediaThumbTooltip.hidden = true;
  mediaThumbImg.hidden = false;
  mediaThumbImg.removeAttribute("src");
  mediaThumbFallback.hidden = true;
  mediaThumbFallback.textContent = "";
}

async function getMediaBlobUrl(path) {
  if (mediaThumbUrlCache.has(path)) return mediaThumbUrlCache.get(path);
  if (!pptxZipCache?.files[path]) return null;
  const data = await pptxZipCache.files[path].async("uint8array");
  const url = URL.createObjectURL(new Blob([data], { type: mimeForMediaPath(path) }));
  mediaThumbUrlCache.set(path, url);
  return url;
}

function positionMediaThumbTooltip(clientX, clientY) {
  const rect = mediaThumbTooltip.getBoundingClientRect();
  const pad = 14;
  let x = clientX + pad;
  let y = clientY + pad;
  if (x + rect.width > window.innerWidth - 8) {
    x = Math.max(8, clientX - rect.width - pad);
  }
  if (y + rect.height > window.innerHeight - 8) {
    y = Math.max(8, clientY - rect.height - pad);
  }
  mediaThumbTooltip.style.left = `${x}px`;
  mediaThumbTooltip.style.top = `${y}px`;
}

let mediaThumbLoadToken = 0;

async function showMediaThumb(path, clientX, clientY) {
  if (!isPreviewableMediaPath(path)) return;
  const token = ++mediaThumbLoadToken;
  const url = await getMediaBlobUrl(path);
  if (!url || token !== mediaThumbLoadToken) return;
  mediaThumbImg.src = url;
  mediaThumbImg.hidden = false;
  mediaThumbFallback.hidden = true;
  mediaThumbTooltip.hidden = false;
  mediaThumbTooltip.setAttribute("aria-hidden", "false");
  positionMediaThumbTooltip(clientX, clientY);
  requestAnimationFrame(() => positionMediaThumbTooltip(clientX, clientY));
}

function hideMediaThumb() {
  mediaThumbLoadToken += 1;
  mediaThumbTooltip.hidden = true;
  mediaThumbTooltip.setAttribute("aria-hidden", "true");
  mediaThumbImg.removeAttribute("src");
}

function getMediaPreviewCell(target) {
  return target?.closest?.(".media-name[data-path], .media-inline-thumb[data-path]");
}

async function populateInlineMediaThumbs() {
  const thumbs = mediaTableBody.querySelectorAll(".media-inline-thumb[data-path]");
  await Promise.all([...thumbs].map(async (span) => {
    const path = span.dataset.path;
    span.classList.add("is-loading");
    const url = await getMediaBlobUrl(path);
    span.classList.remove("is-loading");
    if (!url) return;
    const img = document.createElement("img");
    img.src = url;
    img.alt = "";
    span.appendChild(img);
  }));
}

function addOption(group, value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  group.appendChild(option);
}

function populateFontSelect(selectEl, customInput, previous) {
  selectEl.innerHTML = "";

  const presetGroup = document.createElement("optgroup");
  presetGroup.label = "おすすめ";
  for (const name of PRESET_FONTS) {
    addOption(presetGroup, name, name);
  }
  selectEl.appendChild(presetGroup);

  const pptxOnly = pptxFonts.filter(({ name }) => !PRESET_SET.has(name));
  if (pptxOnly.length > 0) {
    const pptxGroup = document.createElement("optgroup");
    pptxGroup.label = "この PPTX 内";
    for (const { name, count } of pptxOnly) {
      addOption(pptxGroup, name, `${name}（${count} 箇所）`);
    }
    selectEl.appendChild(pptxGroup);
  }

  addOption(selectEl, CUSTOM_VALUE, "その他（自由入力）");

  const knownValues = new Set([
    ...PRESET_FONTS,
    ...pptxOnly.map(({ name }) => name),
  ]);

  if (knownValues.has(previous)) {
    selectEl.value = previous;
    customInput.hidden = true;
    return previous;
  }
  if (previous && previous !== DEFAULT_FONT) {
    selectEl.value = CUSTOM_VALUE;
    customInput.hidden = false;
    customInput.value = previous;
    return previous;
  }
  selectEl.value = DEFAULT_FONT;
  customInput.hidden = true;
  customInput.value = "";
  return DEFAULT_FONT;
}

function rebuildFontDropdowns() {
  selectedTitleFont = populateFontSelect(titleFontSelect, titleCustomFontInput, selectedTitleFont);
  selectedBodyFont = populateFontSelect(bodyFontSelect, bodyCustomFontInput, selectedBodyFont);
  updateFontPreviews();
}

function quoteFontFamily(name) {
  return `"${String(name).replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

async function isFontAvailable(fontName) {
  const trimmed = String(fontName || "").trim();
  if (!trimmed || !document.fonts?.check) return true;
  try {
    await document.fonts.load(`16px ${quoteFontFamily(trimmed)}`);
    return document.fonts.check(`16px ${quoteFontFamily(trimmed)}`);
  } catch {
    return true;
  }
}

function updateFontPreviewElement(previewEl, fontName, sampleText, enabled) {
  if (!previewEl) return;
  previewEl.classList.remove("is-missing");
  if (!enabled) {
    previewEl.style.fontFamily = "";
    previewEl.textContent = "フォント統一オフ";
    return;
  }
  previewEl.style.fontFamily = `${quoteFontFamily(fontName)}, sans-serif`;
  previewEl.textContent = sampleText;
  isFontAvailable(fontName).then((available) => {
    if (!available && optReplaceFonts.checked) {
      previewEl.classList.add("is-missing");
      previewEl.textContent = `「${fontName}」がこの PC に見つかりません`;
    }
  });
}

function updateFontPreviews() {
  const enabled = optReplaceFonts.checked;
  updateFontPreviewElement(
    titleFontPreview,
    getFontFromPicker(titleFontSelect, titleCustomFontInput),
    "見出し Sample あいう ABC 123",
    enabled
  );
  updateFontPreviewElement(
    bodyFontPreview,
    getFontFromPicker(bodyFontSelect, bodyCustomFontInput),
    "本文 Sample あいう ABC 123",
    enabled
  );
}

function getFontFromPicker(selectEl, customInput) {
  if (selectEl.value === CUSTOM_VALUE) {
    return customInput.value.trim() || DEFAULT_FONT;
  }
  return selectEl.value.trim() || DEFAULT_FONT;
}

function getTargetFonts() {
  return {
    title: getFontFromPicker(titleFontSelect, titleCustomFontInput),
    body: getFontFromPicker(bodyFontSelect, bodyCustomFontInput),
  };
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) return;
    const settings = JSON.parse(raw);
    if (typeof settings.titleFont === "string" && settings.titleFont.trim()) {
      selectedTitleFont = settings.titleFont.trim();
    }
    if (typeof settings.bodyFont === "string" && settings.bodyFont.trim()) {
      selectedBodyFont = settings.bodyFont.trim();
    }
    if (typeof settings.replaceFonts === "boolean") {
      optReplaceFonts.checked = settings.replaceFonts;
    }
    if (typeof settings.removeOrphanMedia === "boolean") {
      optRemoveOrphanMedia.checked = settings.removeOrphanMedia;
    }
    if (typeof settings.compressImages === "boolean") {
      optCompressImages.checked = settings.compressImages;
    }
    if (settings.imagePpi != null && imagePpiSelect) {
      const ppi = String(settings.imagePpi);
      if ([...imagePpiSelect.options].some((opt) => opt.value === ppi)) {
        imagePpiSelect.value = ppi;
      }
    }
    if (settings.jpegQuality != null && jpegQualitySelect) {
      const quality = String(settings.jpegQuality);
      if ([...jpegQualitySelect.options].some((opt) => opt.value === quality)) {
        jpegQualitySelect.value = quality;
      }
    }
    if (typeof settings.removeUnusedStructure === "boolean") {
      optRemoveUnusedStructure.checked = settings.removeUnusedStructure;
    }
    if (typeof settings.removeNotes === "boolean") {
      optRemoveNotes.checked = settings.removeNotes;
    }
    if (typeof settings.removeProperties === "boolean") {
      optRemoveProperties.checked = settings.removeProperties;
    }
  } catch {
    /* ignore corrupt settings */
  }
}

function saveSettings() {
  try {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify({
      titleFont: getFontFromPicker(titleFontSelect, titleCustomFontInput),
      bodyFont: getFontFromPicker(bodyFontSelect, bodyCustomFontInput),
      replaceFonts: optReplaceFonts.checked,
      removeOrphanMedia: optRemoveOrphanMedia.checked,
      compressImages: optCompressImages.checked,
      imagePpi: Number(imagePpiSelect.value) || DEFAULT_IMAGE_PPI,
      jpegQuality: Number(jpegQualitySelect.value) || DEFAULT_JPEG_QUALITY,
      removeUnusedStructure: optRemoveUnusedStructure.checked,
      removeNotes: optRemoveNotes.checked,
      removeProperties: optRemoveProperties.checked,
    }));
  } catch {
    /* ignore quota / private mode */
  }
}

function bindFontPicker(selectEl, customInput, getSelected, setSelected) {
  selectEl.addEventListener("change", () => {
    if (selectEl.value === CUSTOM_VALUE) {
      customInput.hidden = false;
      customInput.focus();
      if (!customInput.value.trim()) {
        customInput.value = getSelected() === CUSTOM_VALUE ? "" : getSelected();
      }
      setSelected(customInput.value.trim() || DEFAULT_FONT);
    } else {
      customInput.hidden = true;
      setSelected(selectEl.value);
    }
    saveSettings();
    updateFontPreviews();
  });
  customInput.addEventListener("input", () => {
    setSelected(customInput.value.trim() || DEFAULT_FONT);
    saveSettings();
    updateFontPreviews();
  });
}

function updateFileMode() {
  if (!selectedFile) {
    fileModeEl.hidden = true;
    return;
  }
  fileModeEl.hidden = false;
  fileModeEl.textContent = sourceFileHandle
    ? "上書き保存可"
    : "ダウンロードのみ";
}

function setActionState() {
  const hasFile = Boolean(selectedFile);
  downloadBtn.disabled = !hasFile;
  clearBtn.disabled = !hasFile;
  overwriteBtn.disabled = !hasFile || !sourceFileHandle;
  updateFileMode();
}

async function verifyPermission(handle, readWrite = true) {
  const opts = { mode: readWrite ? "readwrite" : "read" };
  if ((await handle.queryPermission(opts)) === "granted") return true;
  if ((await handle.requestPermission(opts)) === "granted") return true;
  return false;
}

async function writeBlobToHandle(handle, blob) {
  const writable = await handle.createWritable({ keepExistingData: false });
  await writable.write(blob);
  await writable.close();
}

function friendlyFsError(err) {
  const msg = err?.message || String(err);
  if (
    err?.name === "InvalidStateError" &&
    msg.includes("state cached in an interface object")
  ) {
    return "ファイルまたはフォルダの状態がディスク上で変わりました。PowerPoint で開いていないか確認し、OneDrive 等の同期が終わってから「ファイルを開く」で選び直してください。";
  }
  return msg;
}

async function runFsOp(fn) {
  try {
    return await fn();
  } catch (err) {
    if (err?.name === "InvalidStateError") {
      parentDirHandle = null;
    }
    throw new Error(friendlyFsError(err));
  }
}

function backupTimestamp() {
  return String(Math.floor(Date.now() / 1000));
}

function backupBaseName(originalName) {
  return originalName.replace(/\.pptx$/i, "");
}

async function resolveBackupName(dirHandle, originalName) {
  const stem = backupBaseName(originalName);
  let backupName = `${stem}_backup_${backupTimestamp()}.pptx`;
  try {
    await dirHandle.getFileHandle(backupName, { create: false });
    // Same-second collision (rare): fall back to millisecond precision.
    backupName = `${stem}_backup_${Date.now()}.pptx`;
  } catch (err) {
    if (err.name !== "NotFoundError") throw err;
  }
  return backupName;
}

async function ensureParentDirectory(fileHandle) {
  if (parentDirHandle) return parentDirHandle;

  if (typeof fileHandle.getParent === "function") {
    try {
      parentDirHandle = await fileHandle.getParent();
      return parentDirHandle;
    } catch {
      /* fall through */
    }
  }

  parentDirHandle = await window.showDirectoryPicker({
    mode: "readwrite",
    startIn: fileHandle,
  });
  return parentDirHandle;
}

async function createBackupInDirectory(dirHandle, originalFile) {
  const backupName = await resolveBackupName(dirHandle, originalFile.name);
  const backupHandle = await dirHandle.getFileHandle(backupName, { create: true });
  await writeBlobToHandle(backupHandle, originalFile);
  return backupName;
}

function getFinalizeOptions() {
  return {
    replaceFonts: optReplaceFonts.checked,
    removeOrphanMedia: optRemoveOrphanMedia.checked,
    compressImages: optCompressImages.checked,
    imagePpi: Number(imagePpiSelect.value) || DEFAULT_IMAGE_PPI,
    jpegQuality: Number(jpegQualitySelect.value) || DEFAULT_JPEG_QUALITY,
    removeUnusedStructure: optRemoveUnusedStructure.checked,
    removeNotes: optRemoveNotes.checked,
    removeProperties: optRemoveProperties.checked,
  };
}

function updateCompressOptionsState() {
  const enabled = optCompressImages.checked;
  if (imagePpiSelect) imagePpiSelect.disabled = !enabled;
  if (jpegQualitySelect) jpegQualitySelect.disabled = !enabled;
  if (compressImageOptions) compressImageOptions.hidden = !enabled;
}

function applySideTabView() {
  const tab = activeSideTab;
  if (sideTabAnalysisBtn) {
    sideTabAnalysisBtn.classList.toggle("is-active", tab === "analysis");
    sideTabAnalysisBtn.setAttribute("aria-selected", tab === "analysis" ? "true" : "false");
  }
  if (sideTabResultBtn) {
    sideTabResultBtn.classList.toggle("is-active", tab === "result");
    sideTabResultBtn.setAttribute("aria-selected", tab === "result" ? "true" : "false");
  }
  if (sideTabAnalysis) sideTabAnalysis.hidden = tab !== "analysis";
  if (sideTabResult) sideTabResult.hidden = tab !== "result";

  if (resultEmpty) {
    resultEmpty.hidden = tab !== "result" || hasResultContent;
  }
  if (resultPanel && hasResultContent) {
    resultPanel.hidden = tab !== "result";
  } else if (resultPanel) {
    resultPanel.hidden = true;
  }
  if (resultTabBadge) {
    resultTabBadge.hidden = !hasResultContent || tab === "result";
  }
}

function setSideTab(tab) {
  activeSideTab = tab;
  applySideTabView();
}

function updateSideChrome() {
  const hasAnalysis = analysisStack && !analysisStack.hidden;
  const hasContent = hasAnalysis || hasResultContent;

  if (sidePlaceholder) sidePlaceholder.hidden = hasContent;
  if (sideTabs) sideTabs.hidden = !hasContent;

  if (!hasContent) {
    if (sideTabAnalysis) sideTabAnalysis.hidden = true;
    if (sideTabResult) sideTabResult.hidden = true;
    return;
  }

  applySideTabView();
}

function updateReductionEstimate() {
  if (!fileAnalysisAfterSize) return;
  if (!cleanupPlan || !pptxZipCache || !currentFileSize) {
    fileAnalysisAfterSize.textContent = "—";
    if (reductionBlock) reductionBlock.hidden = true;
    if (reductionBar) reductionBar.setAttribute("aria-valuenow", "0");
    if (fileAnalysisEstimateBar) fileAnalysisEstimateBar.style.width = "0%";
    return;
  }

  const bytes = computeReductionEstimate(cleanupPlan, getFinalizeOptions(), pptxZipCache);
  if (bytes <= 0) {
    fileAnalysisAfterSize.textContent = `${formatBytes(currentFileSize)}（削減見込みなし）`;
    if (reductionBlock) reductionBlock.hidden = true;
    if (reductionBar) reductionBar.setAttribute("aria-valuenow", "0");
    if (fileAnalysisEstimateBar) fileAnalysisEstimateBar.style.width = "0%";
    return;
  }

  const after = Math.max(0, currentFileSize - bytes);
  const pctNum = currentFileSize > 0 ? (bytes / currentFileSize) * 100 : 0;
  const pct = pctNum.toFixed(1);
  const barPct = Math.min(100, Math.max(0.5, pctNum));

  fileAnalysisAfterSize.textContent =
    `${formatBytes(after)}（${pct}% 削減）`;
  if (reductionBlock) reductionBlock.hidden = false;
  if (fileAnalysisEstimateBar) fileAnalysisEstimateBar.style.width = `${barPct}%`;
  if (reductionBar) reductionBar.setAttribute("aria-valuenow", String(Math.round(barPct)));
}

function renderCleanupPreview() {
  if (!cleanupPlan) {
    cleanupPanel.hidden = true;
    return;
  }
  cleanupPanel.hidden = false;
  const slideCount = cleanupPlan.structure.slidePathToNum.size;
  if (structureOptionDesc) {
    structureOptionDesc.textContent =
      `${slideCount} 枚のスライドから参照されていない型・マスター XML および関連するメディアを削除します`;
  }
  orphanMediaPreview.textContent = cleanupPlan.slideOrphanMedia.items.length
    ? `${formatBytes(cleanupPlan.slideOrphanMedia.totalSize)}（${cleanupPlan.slideOrphanMedia.items.length} 件` +
      `${cleanupPlan.slideOrphanMedia.missingCount ? `、参照のみ ${cleanupPlan.slideOrphanMedia.missingCount} 件` : ""}）`
    : "0 B（0 件）";
  const compressJobs = buildImageCompressJobs(
    cleanupPlan.imageCompressUsages,
    getFinalizeOptions(),
    mediaPathsRemovedByOptions(cleanupPlan, getFinalizeOptions())
  );
  const compressBytes = compressJobs.reduce(
    (sum, job) => sum + Math.max(0, (job.compressedSize || 0) - job.estimatedCompressed),
    0
  );
  compressImagesPreview.textContent = compressJobs.length
    ? `${compressJobs.length} 件（見込み ${formatBytes(compressBytes)}）`
    : "0 件";
  updateCompressOptionsState();
  const structureMedia = cleanupPlan.structureFreedMedia;
  const structureMediaPart = structureMedia.items.length
    ? ` + 関連メディア ${structureMedia.items.length} 件（${formatBytes(structureMedia.totalCompressedSize)}）`
    : "";
  structurePreview.textContent =
    `${cleanupPlan.unusedLayoutCount} レイアウト + ${cleanupPlan.unusedMasterCount} マスター${structureMediaPart}`;
  notesPreview.textContent = cleanupPlan.notes.count
    ? `${cleanupPlan.notes.count} 件（${formatBytes(cleanupPlan.notes.bytes)}）`
    : "0 件";
  if (cleanupPlan.properties.fieldCount > 0) {
    const sample = cleanupPlan.properties.fields
      .slice(0, 3)
      .map((field) => field.label)
      .join("、");
    const suffix = cleanupPlan.properties.fieldCount > 3 ? " 等" : "";
    propertiesPreview.textContent = `${cleanupPlan.properties.fieldCount} 項目（${sample}${suffix}）`;
  } else {
    propertiesPreview.textContent = "0 項目";
  }
  updateReductionEstimate();
}

function updateFileSizeWarning(totalFileSize) {
  if (!fileSizeWarning) return;
  const check = assessInputFileSize(totalFileSize);
  if (check.level === "warn") {
    fileSizeWarning.hidden = false;
    fileSizeWarning.textContent = check.message;
    return;
  }
  fileSizeWarning.hidden = true;
  fileSizeWarning.textContent = "";
}

function renderFileAnalysis(totalFileSize, slideCount) {
  currentFileSize = totalFileSize;
  analysisStack.hidden = false;
  fileAnalysisSize.textContent = formatBytes(totalFileSize);
  fileAnalysisSlides.textContent = `${slideCount} 枚`;
  updateFileSizeWarning(totalFileSize);
  updateReductionEstimate();
  setSideTab("analysis");
  updateSideChrome();
}

function renderFontAnalysis(fonts) {
  if (!fonts || fonts.length === 0) {
    analysisStack.hidden = false;
    fontAnalysisBadge.textContent = "フォントなし";
    fontSummary.textContent = "";
    fontTableBody.innerHTML = "";
    fontEmpty.hidden = false;
    updateSideChrome();
    return;
  }

  analysisStack.hidden = false;
  fontEmpty.hidden = true;
  const totalUses = fonts.reduce((sum, font) => sum + font.count, 0);
  const embeddedKinds = fonts.filter((font) => font.fileBytes > 0).length;
  let totalEmbedded = 0;
  if (pptxZipCache) {
    for (const path of Object.keys(pptxZipCache.files)) {
      if (!path.startsWith("ppt/fonts/") || pptxZipCache.files[path].dir) continue;
      totalEmbedded += getZipEntrySize(pptxZipCache.files[path]);
    }
  }
  fontAnalysisBadge.textContent = `${fonts.length} 種類 / 埋込 ${formatBytes(totalEmbedded)}`;
  fontSummary.textContent =
    `検出フォント: ${fonts.length} 種類 / typeface 合計: ${totalUses} 箇所 / ` +
    `埋込フォント: ${embeddedKinds} 種類（${formatBytes(totalEmbedded)}）`;

  fontTableBody.innerHTML = fonts
    .map((font) => {
      const fileLabel = font.fileBytes > 0 ? formatBytes(font.fileBytes) : "—";
      const fileTitle = font.fileDetail
        ? ` title="${escapeHtml(font.fileDetail)}"`
        : "";
      return `<tr>
      <td>${escapeHtml(font.name)}</td>
      <td class="file-size"${fileTitle}>${escapeHtml(fileLabel)}</td>
      <td class="count">${font.count}</td>
    </tr>`;
    })
    .join("");
}

function releasePptxZipCache() {
  revokeMediaThumbUrls();
  pptxZipCache = null;
}

function assertInputFileSizeAllowed(file) {
  const check = assessInputFileSize(file?.size || 0);
  if (check.level === "reject") {
    throw new Error(check.message);
  }
  return check;
}

async function analyzePptxFile(file) {
  assertInputFileSizeAllowed(file);
  revokeMediaThumbUrls();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  pptxZipCache = zip;
  const slidePathToNum = await getPresentationSlideOrder(zip);
  const [fonts, media, plan] = await Promise.all([
    extractFontsFromZipAsync(zip),
    analyzeMediaUsage(zip, slidePathToNum),
    computeCleanupPlan(zip),
  ]);
  return { fonts, media, cleanupPlan: plan, totalSize: file.size };
}

function renderMediaAnalysis(analysis) {
  if (!analysis || analysis.items.length === 0) {
    mediaAnalysisBadge.textContent = "メディアなし";
    mediaSummary.textContent = "ppt/media/ 内のメディアはありません。";
    mediaTableBody.innerHTML = "";
    mediaEmpty.hidden = false;
    return;
  }

  mediaEmpty.hidden = true;
  mediaAnalysisBadge.textContent =
    `${formatBytes(analysis.totalMediaCompressedSize)} / 未使用メディア ${formatBytes(analysis.orphanMediaCompressedSize)}`;

  const previewableCount = analysis.items.filter((item) => item.isPreviewable).length;
  const previewHint = previewableCount > 0
    ? "プレビュー列／ホバーで拡大。"
    : "";
  if (analysis.items.length > 30) {
    mediaSummary.textContent = `サイズ上位 30 件。${previewHint}`.trim();
  } else {
    mediaSummary.textContent = `${analysis.items.length} 件。${previewHint}`.trim();
  }

  const jobsByPath = new Map();
  if (cleanupPlan?.imageCompressUsages) {
    for (const job of buildImageCompressJobs(cleanupPlan.imageCompressUsages, getFinalizeOptions())) {
      jobsByPath.set(job.path, job);
    }
  }
  const dimsByPath = new Map();
  for (const item of cleanupPlan?.imageCompressUsages || []) {
    if (item.width && item.height) dimsByPath.set(item.path, item);
  }

  mediaTableBody.innerHTML = analysis.items
    .slice(0, 30)
    .map((item) => {
      const slideText = item.slides.length
        ? escapeHtml(item.slides.join("、"))
        : '<span style="color:var(--muted)">（参照元不明）</span>';
      const rowStyle = item.isOrphan ? ' style="opacity:0.92"' : "";
      const thumbCell = item.isPreviewable
        ? `<td class="media-thumb-cell"><span class="media-inline-thumb" data-path="${escapeHtml(item.path)}" title="ホバーで拡大"></span></td>`
        : `<td class="media-thumb-cell"></td>`;
      const job = jobsByPath.get(item.path);
      const dims = dimsByPath.get(item.path);
      let pxLine = "";
      if (job) {
        pxLine = `<span class="media-px">${job.width}×${job.height} → ${job.targetWidth}×${job.targetHeight}</span>`;
      } else if (dims) {
        pxLine = `<span class="media-px">${dims.width}×${dims.height}</span>`;
      }
      const nameCell = item.isPreviewable
        ? `<td class="media-name media-name--preview" data-path="${escapeHtml(item.path)}" title="ホバーでプレビュー">${escapeHtml(item.name)}${pxLine}</td>`
        : `<td class="media-name">${escapeHtml(item.name)}${pxLine}</td>`;
      return `<tr${rowStyle}>
        ${thumbCell}
        ${nameCell}
        <td class="size-compressed">${formatBytes(item.compressedSize)}</td>
        <td class="slides">${slideText}</td>
      </tr>`;
    })
    .join("");
  populateInlineMediaThumbs();
}

async function encodeImageBytes(bytes, spec) {
  const mime = spec.mime || "image/png";
  const blob = new Blob([bytes], { type: mime });
  let bitmap;
  try {
    const resizeOpts = spec.needsResize
      ? {
        resizeWidth: spec.targetWidth,
        resizeHeight: spec.targetHeight,
        resizeQuality: "high",
      }
      : {};
    bitmap = await createImageBitmap(blob, resizeOpts);
  } catch {
    return null;
  }
  try {
    const width = bitmap.width;
    const height = bitmap.height;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(bitmap, 0, 0, width, height);
    const outMime = mime === "image/jpeg" ? "image/jpeg" : mime;
    const quality = outMime === "image/jpeg" ? spec.jpegQuality : undefined;
    const outBlob = await new Promise((resolve) => {
      canvas.toBlob((result) => resolve(result), outMime, quality);
    });
    if (!outBlob) return null;
    return { bytes: new Uint8Array(await outBlob.arrayBuffer()) };
  } finally {
    if (typeof bitmap.close === "function") bitmap.close();
  }
}

async function extractFontsFromPptx(file) {
  const { fonts } = await analyzePptxFile(file);
  return fonts;
}

async function finalizePptx(file, fonts, options) {
  const titleFont = fonts.title.trim() || DEFAULT_FONT;
  const bodyFont = fonts.body.trim() || DEFAULT_FONT;
  if (!titleFont || !bodyFont) {
    throw new Error("フォント名を入力してください。");
  }
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error(".pptx ファイルを選択してください。");
  }
  assertInputFileSizeAllowed(file);

  releasePptxZipCache();
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const plan = await computeCleanupPlan(zip);
  const embeddedMap = await mapEmbeddedFontFileSizes(zip);
  const embeddedFontKinds = embeddedMap.size;
  const embeddedFontBytes = [...embeddedMap.values()].reduce((sum, item) => sum + item.bytes, 0);

  let fontStats = { replacements: 0, filesChanged: 0, embeddedFontsRemoved: 0 };
  if (options.replaceFonts) {
    fontStats = await applyFontReplaceToZip(zip, fonts);
  }

  let structureStats = { layoutsRemoved: 0, mastersRemoved: 0, mediaRemoved: 0, mediaBytes: 0 };
  if (options.removeUnusedStructure) {
    structureStats = await removeUnusedStructure(zip, plan);
    const structureMediaStats = await removePackageOrphanMedia(zip);
    structureStats.mediaRemoved = structureMediaStats.count;
    structureStats.mediaBytes = structureMediaStats.bytes;
  }

  let mediaStats = { count: 0, bytes: 0 };
  let notesStats = { count: 0, bytes: 0 };
  let propertiesStats = { cleared: 0 };

  if (options.removeNotes) {
    notesStats = await removeNotesFromZip(zip);
  }

  if (options.removeProperties) {
    propertiesStats = await clearDocumentProperties(zip);
  }

  if (options.removeOrphanMedia) {
    mediaStats = await removePackageOrphanMedia(zip);
  }

  let imageStats = { count: 0, bytes: 0, skipped: 0 };
  if (options.compressImages) {
    imageStats = await compressImagesInZip(
      zip,
      options,
      { encode: encodeImageBytes },
      plan.imageCompressUsages
    );
  }

  const blob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return {
    blob,
    replacements: fontStats.replacements,
    filesChanged: fontStats.filesChanged,
    embeddedFontsRemoved: fontStats.embeddedFontsRemoved,
    embeddedFontKinds,
    embeddedFontBytes,
    layoutsRemoved: structureStats.layoutsRemoved,
    mastersRemoved: structureStats.mastersRemoved,
    structureMediaRemoved: structureStats.mediaRemoved,
    structureMediaBytes: structureStats.mediaBytes,
    orphanMediaRemoved: mediaStats.count,
    orphanMediaBytes: mediaStats.bytes,
    imagesCompressed: imageStats.count,
    imageCompressBytes: imageStats.bytes,
    imagesSkipped: imageStats.skipped,
    notesRemoved: notesStats.count,
    notesBytes: notesStats.bytes,
    propertiesCleared: propertiesStats.cleared,
    titleFont,
    bodyFont,
    fontsReplaced: options.replaceFonts,
    targetFont: bodyFont,
    sourceFonts: pptxFonts.map((f) => f.name),
    originalSize: file.size,
    outputSize: blob.size,
  };
}

async function loadFromFile(file, handle = null) {
  selectedFile = file;
  sourceFileHandle = handle;
  if (!handle) parentDirHandle = null;

  fileNameEl.textContent = file ? file.name : "";
  hideError();
  hasResultContent = false;
  resultPanel.hidden = true;
  setActionState();

  if (!file) {
    pptxFonts = [];
    mediaAnalysis = null;
    cleanupPlan = null;
    pptxZipCache = null;
    revokeMediaThumbUrls();
    analysisStack.hidden = true;
    cleanupPanel.hidden = true;
    updateFileSizeWarning(0);
    rebuildFontDropdowns();
    updateSideChrome();
    return;
  }

  scanningMsg.hidden = false;
  scanningMsg.textContent = "分析中…";

  try {
    const analysis = await analyzePptxFile(file);
    pptxFonts = analysis.fonts;
    mediaAnalysis = analysis.media;
    cleanupPlan = analysis.cleanupPlan;
    rebuildFontDropdowns();
    renderFileAnalysis(analysis.totalSize, mediaAnalysis?.slideCount ?? 0);
    renderFontAnalysis(analysis.fonts);
    renderMediaAnalysis(mediaAnalysis);
    renderCleanupPreview();
    setSideTab("analysis");
    updateSideChrome();
  } catch (err) {
    showError(`分析に失敗しました: ${err.message || err}`);
    pptxFonts = [];
    mediaAnalysis = null;
    cleanupPlan = null;
    pptxZipCache = null;
    revokeMediaThumbUrls();
    analysisStack.hidden = true;
    cleanupPanel.hidden = true;
    updateFileSizeWarning(0);
    rebuildFontDropdowns();
    updateSideChrome();
  } finally {
    scanningMsg.hidden = true;
    scanningMsg.textContent = "分析中…";
  }
}

async function openFileWithHandle() {
  if (!supportsFsAccess) {
    fileInput.click();
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      mode: "readwrite",
      types: PPTX_TYPES,
    });
    if (!(await verifyPermission(handle, true))) {
      throw new Error("ファイルへの書き込み権限が必要です。");
    }
    const file = await handle.getFile();
    parentDirHandle = null;
    await loadFromFile(file, handle);
  } catch (err) {
    if (err.name === "AbortError") return;
    showError(err.message || String(err));
  }
}

async function loadFromDrop(dataTransfer) {
  const item = dataTransfer.items?.[0];
  if (item && typeof item.getAsFileSystemHandle === "function") {
    try {
      const handle = await item.getAsFileSystemHandle();
      if (handle?.kind === "file") {
        if (await verifyPermission(handle, true)) {
          const file = await handle.getFile();
          parentDirHandle = null;
          await loadFromFile(file, handle);
          return;
        }
      }
    } catch {
      /* fall through to File */
    }
  }

  const file = dataTransfer.files?.[0];
  if (file) {
    await loadFromFile(file, null);
  }
}

async function overwriteOriginal() {
  if (!selectedFile || !sourceFileHandle) {
    throw new Error("上書き保存するには「ファイルを開く（上書き可）」で PPTX を選んでください。");
  }

  if (!(await verifyPermission(sourceFileHandle, true))) {
    throw new Error("ファイルへの書き込み権限が必要です。");
  }

  // ディレクトリハンドルのキャッシュは InvalidStateError の原因になりやすい
  parentDirHandle = null;

  const originalFile = await runFsOp(() => sourceFileHandle.getFile());
  const dirHandle = await runFsOp(() => ensureParentDirectory(sourceFileHandle));
  if (!(await verifyPermission(dirHandle, true))) {
    throw new Error("バックアップ先フォルダへの書き込み権限が必要です。");
  }

  const backupName = await runFsOp(() =>
    createBackupInDirectory(dirHandle, originalFile)
  );

  const fonts = getTargetFonts();
  const options = getFinalizeOptions();
  const result = await finalizePptx(originalFile, fonts, options);
  await runFsOp(() => writeBlobToHandle(sourceFileHandle, result.blob));

  selectedFile = new File([result.blob], originalFile.name, {
    type: originalFile.type || "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  scanningMsg.hidden = false;
  scanningMsg.textContent = "分析中…";
  try {
    const analysis = await analyzePptxFile(selectedFile);
    pptxFonts = analysis.fonts;
    mediaAnalysis = analysis.media;
    cleanupPlan = analysis.cleanupPlan;
    rebuildFontDropdowns();
    renderFileAnalysis(analysis.totalSize, mediaAnalysis?.slideCount ?? 0);
    renderFontAnalysis(analysis.fonts);
    renderMediaAnalysis(mediaAnalysis);
    renderCleanupPreview();
  } finally {
    scanningMsg.hidden = true;
  }
  setActionState();

  return {
    ...result,
    outputName: originalFile.name,
    backupName,
    saveMode: "overwrite",
  };
}

function showError(message) {
  errorMsg.textContent = message;
  errorMsg.hidden = false;
}

function hideError() {
  errorMsg.hidden = true;
  errorMsg.textContent = "";
}

function renderStats(stats) {
  const convertedFontCount = stats.sourceFonts.length;
  const embeddedLine = stats.embeddedFontKinds > 0
    ? `${stats.embeddedFontKinds} 種類（${formatBytes(stats.embeddedFontBytes)}）`
    : "0 種類";
  const modeLine = stats.saveMode === "overwrite"
    ? `<dt>保存方法</dt><dd>元ファイルに上書き</dd>`
    : `<dt>保存方法</dt><dd>別名ダウンロード</dd>`;
  const sizeLine = stats.originalSize != null && stats.outputSize != null
    ? `<dt>ファイルサイズ</dt><dd>${formatBytes(stats.originalSize)} → ${formatBytes(stats.outputSize)}</dd>
       <dt>削減</dt><dd>${formatSizeChange(stats.originalSize, stats.outputSize)}</dd>`
    : "";
  const outputLine = `<dt>出力ファイル</dt><dd>${escapeHtml(stats.outputName)}</dd>`;
  const summaryBlock = `<dl class="stats">${sizeLine}${modeLine}${outputLine}</dl>`;

  const backupLine = stats.backupName
    ? `<dt>バックアップ</dt><dd>${escapeHtml(stats.backupName)}（元ファイルと同じフォルダ）</dd>`
    : "";

  const cleanupLines = [];
  if (stats.orphanMediaRemoved > 0) {
    cleanupLines.push(
      `<dt>孤立メディア削除</dt><dd>${stats.orphanMediaRemoved} 件（${formatBytes(stats.orphanMediaBytes)}）</dd>`
    );
  }
  if (stats.imagesCompressed > 0 || stats.imagesSkipped > 0) {
    let detail = `${stats.imagesCompressed || 0} 枚`;
    if (stats.imageCompressBytes > 0) {
      detail += `（${formatBytes(stats.imageCompressBytes)}）`;
    }
    if (stats.imagesSkipped > 0) {
      detail += `、スキップ ${stats.imagesSkipped} 枚`;
    }
    cleanupLines.push(`<dt>画像圧縮</dt><dd>${detail}</dd>`);
  }
  if (stats.layoutsRemoved > 0 || stats.mastersRemoved > 0) {
    let structureDetail =
      `レイアウト ${stats.layoutsRemoved}、マスター ${stats.mastersRemoved}`;
    if (stats.structureMediaRemoved > 0) {
      structureDetail +=
        `、関連メディア ${stats.structureMediaRemoved} 件（${formatBytes(stats.structureMediaBytes)}）`;
    }
    cleanupLines.push(
      `<dt>未使用レイアウト／マスター削除</dt><dd>${structureDetail}</dd>`
    );
  }
  if (stats.notesRemoved > 0) {
    cleanupLines.push(
      `<dt>ノート削除</dt><dd>${stats.notesRemoved} 件（${formatBytes(stats.notesBytes)}）</dd>`
    );
  }
  if (stats.propertiesCleared > 0) {
    cleanupLines.push(
      `<dt>プロパティ削除</dt><dd>${stats.propertiesCleared} 項目</dd>`
    );
  }

  const fontLines = stats.fontsReplaced
    ? `<dt>タイトルフォント</dt><dd>${escapeHtml(stats.titleFont)}</dd>
       <dt>本文フォント</dt><dd>${escapeHtml(stats.bodyFont)}</dd>
       <dt>置換されたフォント</dt><dd>${convertedFontCount} フォント</dd>
       <dt>typeface 置換数</dt><dd>${stats.replacements}</dd>
       <dt>埋め込みフォント削除</dt><dd>${embeddedLine}</dd>`
    : `<dt>フォント統一</dt><dd>未実施（元のフォントのまま）</dd>`;

  const changeLines = [fontLines, ...cleanupLines, backupLine].filter(Boolean);

  statsEl.innerHTML = `
    <div class="stats-block">${summaryBlock}</div>
    <h3 class="stats-section-title">変更内容</h3>
    <div class="stats-block"><dl class="stats">${changeLines.join("")}</dl></div>
  `;
  hasResultContent = true;
  resultPanel.hidden = false;
  setSideTab("result");
  updateSideChrome();
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function outputFilename(originalName) {
  return `${originalName.replace(/\.pptx$/i, "")}_finalized.pptx`;
}

function clearAll() {
  selectedFile = null;
  sourceFileHandle = null;
  parentDirHandle = null;
  currentFileSize = 0;
  fileInput.value = "";
  pptxFonts = [];
  mediaAnalysis = null;
  cleanupPlan = null;
  pptxZipCache = null;
  revokeMediaThumbUrls();
  analysisStack.hidden = true;
  cleanupPanel.hidden = true;
  hasResultContent = false;
  resultPanel.hidden = true;
  activeSideTab = "analysis";
  loadFromFile(null, null);
}

async function runDownload() {
  if (!selectedFile) return;
  hideError();
  downloadBtn.disabled = true;
  downloadBtn.textContent = "仕上げ中…";

  try {
    const fonts = getTargetFonts();
    const options = getFinalizeOptions();
    const result = await finalizePptx(selectedFile, fonts, options);
    if (selectedFile && !pptxZipCache) {
      try {
        pptxZipCache = await JSZip.loadAsync(await selectedFile.arrayBuffer());
      } catch {
        /* サムネイル用キャッシュの復旧失敗は無視 */
      }
    }
    const outputName = outputFilename(selectedFile.name);
    downloadBlob(result.blob, outputName);
    renderStats({ ...result, outputName, saveMode: "download" });
  } catch (err) {
    showError(err.message || String(err));
  } finally {
    downloadBtn.disabled = !selectedFile;
    downloadBtn.textContent = "仕上げてダウンロード";
  }
}

async function runOverwrite() {
  if (!selectedFile || !sourceFileHandle) return;
  hideError();
  overwriteBtn.disabled = true;
  overwriteBtn.textContent = "保存中…";

  try {
    const result = await overwriteOriginal();
    renderStats(result);
  } catch (err) {
    if (err.name === "AbortError") return;
    showError(err.message || String(err));
  } finally {
    overwriteBtn.disabled = !selectedFile || !sourceFileHandle;
    overwriteBtn.textContent = "仕上げて上書き（バックアップ付き）";
  }
}

function updateFontOptionsState() {
  const enabled = optReplaceFonts.checked;
  titleFontSelect.disabled = !enabled;
  titleCustomFontInput.disabled = !enabled;
  bodyFontSelect.disabled = !enabled;
  bodyCustomFontInput.disabled = !enabled;
  fontOptionsPanel.classList.toggle("is-disabled", !enabled);
  updateFontPreviews();
}

function onFontReplaceToggle() {
  updateFontOptionsState();
  saveSettings();
  updateReductionEstimate();
}

function onSideTabClick(event) {
  const tab = event.currentTarget?.dataset?.sideTab;
  if (!tab) return;
  setSideTab(tab);
}

function onCleanupOptionChange() {
  saveSettings();
  renderCleanupPreview();
  if (mediaAnalysis) renderMediaAnalysis(mediaAnalysis);
}

function uncheckAllCleanupOptions() {
  for (const checkbox of cleanupOptionCheckboxes) {
    checkbox.checked = false;
  }
  onCleanupOptionChange();
}

optRemoveOrphanMedia.addEventListener("change", onCleanupOptionChange);
optCompressImages.addEventListener("change", onCleanupOptionChange);
imagePpiSelect.addEventListener("change", onCleanupOptionChange);
jpegQualitySelect.addEventListener("change", onCleanupOptionChange);
optRemoveUnusedStructure.addEventListener("change", onCleanupOptionChange);
optRemoveNotes.addEventListener("change", onCleanupOptionChange);
optRemoveProperties.addEventListener("change", onCleanupOptionChange);
uncheckAllOptionsBtn.addEventListener("click", uncheckAllCleanupOptions);
optReplaceFonts.addEventListener("change", onFontReplaceToggle);
if (sideTabAnalysisBtn) sideTabAnalysisBtn.addEventListener("click", onSideTabClick);
if (sideTabResultBtn) sideTabResultBtn.addEventListener("click", onSideTabClick);

mediaTableBody.addEventListener("mouseover", (event) => {
  const cell = getMediaPreviewCell(event.target);
  if (!cell) return;
  showMediaThumb(cell.dataset.path, event.clientX, event.clientY);
});

mediaTableBody.addEventListener("mousemove", (event) => {
  if (mediaThumbTooltip.hidden) return;
  const cell = getMediaPreviewCell(event.target);
  if (!cell) return;
  positionMediaThumbTooltip(event.clientX, event.clientY);
});

mediaTableBody.addEventListener("mouseout", (event) => {
  const fromCell = getMediaPreviewCell(event.target);
  if (!fromCell) return;
  const toCell = getMediaPreviewCell(event.relatedTarget);
  if (toCell?.dataset.path === fromCell.dataset.path) return;
  hideMediaThumb();
});

openBtn.addEventListener("click", openFileWithHandle);
pickBtn.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (file) loadFromFile(file, null);
});

["dragenter", "dragover"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (e) => {
  loadFromDrop(e.dataTransfer);
});

clearBtn.addEventListener("click", clearAll);
bindFontPicker(
  titleFontSelect,
  titleCustomFontInput,
  () => selectedTitleFont,
  (value) => { selectedTitleFont = value; }
);
bindFontPicker(
  bodyFontSelect,
  bodyCustomFontInput,
  () => selectedBodyFont,
  (value) => { selectedBodyFont = value; }
);
downloadBtn.addEventListener("click", runDownload);
overwriteBtn.addEventListener("click", runOverwrite);

[titleCustomFontInput, bodyCustomFontInput].forEach((input) => {
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && selectedFile && !downloadBtn.disabled) {
      if (sourceFileHandle && !overwriteBtn.disabled) {
        runOverwrite();
      } else {
        runDownload();
      }
    }
  });
});

loadSettings();
rebuildFontDropdowns();
updateFontOptionsState();
updateCompressOptionsState();
setActionState();
updateSideChrome();
showAppVersion();

async function showAppVersion() {
  const line = document.getElementById("appVersionLine");
  const el = document.getElementById("appVersion");
  if (!line || !el) return;
  try {
    const res = await fetch("version.json", { cache: "no-store" });
    if (!res.ok) throw new Error("version fetch failed");
    const data = await res.json();
    const version = String(data.version || "").trim();
    if (!version) throw new Error("empty version");
    el.textContent = `v${version}`;
    line.hidden = false;
  } catch {
    line.hidden = true;
  }
}
