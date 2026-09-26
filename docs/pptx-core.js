// PPTX manipulation engine: pure string/XML helpers plus JSZip-based
// analysis and mutation functions. No DOM access here so this file can be
// loaded both as a classic <script> (browser) and via require() (tests).

const DEFAULT_FONT = "Noto Sans JP";
const TYPEFACE_CAPTURE_RE = /typeface="([^"]*)"/g;
const TYPEFACE_RE = /typeface="[^"]*"/g;
const TITLE_PH_RE = /<p:ph[^>]*\btype="(?:title|ctrTitle|subTitle)"/;
const BODY_PH_RE = /<p:ph[^>]*\btype="body"/;
const STYLE_BLOCK_RES = [
  { re: /<(?:p|a):titleStyle\b[\s\S]*?<\/(?:p|a):titleStyle>/g, role: "title" },
  { re: /<(?:p|a):bodyStyle\b[\s\S]*?<\/(?:p|a):bodyStyle>/g, role: "body" },
  { re: /<(?:p|a):otherStyle\b[\s\S]*?<\/(?:p|a):otherStyle>/g, role: "body" },
];
const EMBEDDED_FONT_BLOCK_RE = /<p:embeddedFont>([\s\S]*?)<\/p:embeddedFont>/g;
const PREVIEWABLE_IMAGE_RE = /\.(png|jpe?g|gif|webp|bmp|svg)$/i;
const EMBEDDED_FONT_LST_RE = /<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/g;
const FONT_REL_RE = /<Relationship[^>]+Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font"[^>]*\/>\s*/g;
const REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const MEDIA_REL_MARKERS = ["/image", "/video", "/audio", "/media"];
const EMU_PER_INCH = 914400;
const SRC_RECT_BASE = 100000;
const DEFAULT_IMAGE_PPI = 150;
const DEFAULT_JPEG_QUALITY = 0.75;
const COMPRESSIBLE_IMAGE_RE = /\.(png|jpe?g|webp)$/i;
const DEFAULT_SLIDE_CX_EMU = 12192000;
const DEFAULT_SLIDE_CY_EMU = 6858000;
const IMAGE_ENCODE_CONCURRENCY = 3;
const IMAGE_DIM_PREFIX_BYTES = 256 * 1024;
const WARN_INPUT_FILE_BYTES = 50 * 1024 * 1024;
const MAX_INPUT_FILE_BYTES = 200 * 1024 * 1024;

function decodeXmlEntities(value) {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function escapeXmlAttr(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isPreviewableMediaPath(path) {
  return PREVIEWABLE_IMAGE_RE.test(path);
}

function isCompressibleImagePath(path) {
  return COMPRESSIBLE_IMAGE_RE.test(path);
}

function mimeForMediaPath(path) {
  const ext = (path.split(".").pop() || "").toLowerCase();
  const map = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
    bmp: "image/bmp",
    svg: "image/svg+xml",
  };
  return map[ext] || "application/octet-stream";
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatSizeChange(before, after) {
  const delta = before - after;
  const pct = before > 0 ? (Math.abs(delta) / before) * 100 : 0;
  if (delta > 0) {
    return `${formatBytes(delta)} 削減（−${pct.toFixed(1)}%）`;
  }
  if (delta < 0) {
    return `${formatBytes(-delta)} 増加（+${pct.toFixed(1)}%）`;
  }
  return "変化なし（0%）";
}

function assessInputFileSize(bytes) {
  const size = Number(bytes) || 0;
  if (size > MAX_INPUT_FILE_BYTES) {
    return {
      level: "reject",
      bytes: size,
      message:
        `ファイルが大きすぎます（${formatBytes(size)}）。` +
        `${formatBytes(MAX_INPUT_FILE_BYTES)} 以下の PPTX を指定してください。`,
    };
  }
  if (size >= WARN_INPUT_FILE_BYTES) {
    return {
      level: "warn",
      bytes: size,
      message:
        `ファイルサイズが ${formatBytes(size)} あります。` +
        "分析・仕上げに時間がかかるか、メモリ不足になることがあります。",
    };
  }
  return { level: "ok", bytes: size, message: "" };
}

function countReplace(str, re, replacement) {
  const matches = str.match(re);
  const count = matches ? matches.length : 0;
  const data = str.replace(re, replacement);
  return { data, count };
}

function replaceTypefacesInString(text, font) {
  const safeFont = escapeXmlAttr(font.trim() || DEFAULT_FONT);
  return text.replace(TYPEFACE_RE, `typeface="${safeFont}"`);
}

function countTypefaces(text) {
  TYPEFACE_RE.lastIndex = 0;
  const matches = text.match(TYPEFACE_RE);
  return matches ? matches.length : 0;
}

function findNextSpShapeStart(xml, fromIndex) {
  let i = fromIndex;
  while (i < xml.length) {
    const idx = xml.indexOf("<p:sp", i);
    if (idx === -1) return -1;
    const next = xml.charAt(idx + 5);
    if (next === ">" || next === " " || next === "/") return idx;
    i = idx + 5;
  }
  return -1;
}

function findBalancedSpBlockEnd(xml, startIndex) {
  const closeTag = "</p:sp>";
  let depth = 1;
  let i = startIndex + 5;
  while (i < xml.length && depth > 0) {
    const nextOpen = findNextSpShapeStart(xml, i);
    const nextClose = xml.indexOf(closeTag, i);
    if (nextClose === -1) return -1;
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth += 1;
      i = nextOpen + 5;
    } else {
      depth -= 1;
      if (depth === 0) return nextClose + closeTag.length;
      i = nextClose + closeTag.length;
    }
  }
  return -1;
}

function replaceFontsInStyleBlocks(xml, titleFont, bodyFont) {
  let result = xml;
  let replacements = 0;
  for (const { re, role } of STYLE_BLOCK_RES) {
    re.lastIndex = 0;
    result = result.replace(re, (block) => {
      const font = role === "title" ? titleFont : bodyFont;
      replacements += countTypefaces(block);
      return replaceTypefacesInString(block, font);
    });
  }
  return { data: result, count: replacements };
}

function replaceFontsContextual(data, titleFont, bodyFont) {
  const safeTitle = (titleFont || DEFAULT_FONT).trim() || DEFAULT_FONT;
  const safeBody = (bodyFont || DEFAULT_FONT).trim() || DEFAULT_FONT;
  let replacements = 0;
  const spStore = [];
  let masked = "";
  let cursor = 0;
  while (cursor < data.length) {
    const start = findNextSpShapeStart(data, cursor);
    if (start === -1) {
      masked += data.slice(cursor);
      break;
    }
    masked += data.slice(cursor, start);
    const end = findBalancedSpBlockEnd(data, start);
    if (end === -1) {
      masked += data.slice(start);
      break;
    }
    spStore.push(data.slice(start, end));
    masked += `\x00SP${spStore.length - 1}\x00`;
    cursor = end;
  }

  // Apply body font to unmasked XML first; title/body style blocks are updated afterward
  // so their typeface attributes are not overwritten by the global body pass.
  const remaining = countReplace(masked, TYPEFACE_RE, `typeface="${escapeXmlAttr(safeBody)}"`);
  masked = remaining.data;
  replacements += remaining.count;

  const styles = replaceFontsInStyleBlocks(masked, safeTitle, safeBody);
  masked = styles.data;
  replacements += styles.count;

  const restored = masked.replace(/\x00SP(\d+)\x00/g, (_, indexText) => {
    const block = spStore[Number(indexText)];
    let font = safeBody;
    if (TITLE_PH_RE.test(block)) font = safeTitle;
    else if (BODY_PH_RE.test(block)) font = safeBody;
    replacements += countTypefaces(block);
    return replaceTypefacesInString(block, font);
  });

  return { data: restored, count: replacements };
}

function getZipEntrySize(entry) {
  if (!entry) return 0;
  const data = entry._data;
  if (!data) return 0;
  if (typeof data.uncompressedSize === "number" && data.uncompressedSize > 0) {
    return data.uncompressedSize;
  }
  return data.length || data.compressedSize || 0;
}

function getZipEntryCompressedSize(entry) {
  if (!entry) return 0;
  const data = entry._data;
  if (!data) return 0;
  if (typeof data.compressedSize === "number" && data.compressedSize >= 0) {
    return data.compressedSize;
  }
  return data.length || 0;
}

function zipEntryExists(zip, path) {
  return Boolean(path && zip.files[path] && !zip.files[path].dir);
}

function parseRelationships(xml) {
  const rels = [];
  const re = /<Relationship\s+([^>]+?)\/?>/g;
  let match;
  while ((match = re.exec(xml)) !== null) {
    const attrs = match[1];
    const id = /(?:\bId|Id)="([^"]+)"/.exec(attrs)?.[1]
      || /Id="([^"]+)"/.exec(attrs)?.[1];
    const type = /Type="([^"]+)"/.exec(attrs)?.[1];
    const target = /Target="([^"]+)"/.exec(attrs)?.[1];
    if (type && target) rels.push({ id, type, target });
  }
  return rels;
}

function resolveZipPath(basePath, target) {
  const normalized = target.replace(/\\/g, "/");
  if (normalized.startsWith("/")) {
    return normalized.replace(/^\//, "");
  }
  const baseDir = basePath.includes("/")
    ? basePath.slice(0, basePath.lastIndexOf("/") + 1)
    : "";
  const parts = (baseDir + normalized).split("/");
  const stack = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") stack.pop();
    else stack.push(part);
  }
  return stack.join("/");
}

function isMediaRelationship(type, target = "") {
  const normalizedTarget = String(target || "").replace(/\\/g, "/");
  if (/(?:^|\/)media\//.test(normalizedTarget)) {
    return true;
  }
  return MEDIA_REL_MARKERS.some((marker) => type.includes(marker));
}

function ownerPathFromRelsPath(relsPath) {
  const match = relsPath.match(/^(ppt\/[^/]+)\/_rels\/(.+)\.rels$/);
  if (match) return `${match[1]}/${match[2]}`;
  return null;
}

async function getPresentationSlideOrder(zip) {
  const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const presFile = zip.file("ppt/presentation.xml");
  if (!presRelsFile || !presFile) return new Map();

  const presRels = parseRelationships(await presRelsFile.async("string"));
  const rIdToSlide = new Map();
  for (const rel of presRels) {
    if (rel.type === REL_SLIDE) {
      rIdToSlide.set(rel.id, resolveZipPath("ppt/presentation.xml", rel.target));
    }
  }

  const presXml = await presFile.async("string");
  const slidePathToNum = new Map();
  const sldIdRe = /<p:sldId\b[^>]*\br:id="([^"]+)"/g;
  let match;
  let index = 0;
  while ((match = sldIdRe.exec(presXml)) !== null) {
    index += 1;
    const slidePath = rIdToSlide.get(match[1]);
    if (slidePath) slidePathToNum.set(slidePath, index);
  }
  return slidePathToNum;
}

function formatSlideLabel(slideNum, via) {
  if (!slideNum) return via;
  return via ? `スライド ${slideNum}（${via}）` : `スライド ${slideNum}`;
}

function addMediaUsage(mediaUsage, mediaPath, label) {
  if (!mediaPath.startsWith("ppt/media/")) return;
  if (!mediaUsage.has(mediaPath)) {
    mediaUsage.set(mediaPath, new Set());
  }
  mediaUsage.get(mediaPath).add(label);
}

function isOrphanUsageLabel(label) {
  return label.startsWith("未使用");
}

function sortUsageLabels(labels) {
  return labels.sort((a, b) => {
    const aOrphan = isOrphanUsageLabel(a);
    const bOrphan = isOrphanUsageLabel(b);
    if (aOrphan !== bOrphan) return aOrphan ? 1 : -1;
    return a.localeCompare(b, "ja");
  });
}

async function analyzeMediaUsage(zip, slidePathToNum) {
  const mediaSizes = new Map();
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("ppt/media/") || zip.files[path].dir) continue;
    const entry = zip.files[path];
    mediaSizes.set(path, {
      size: getZipEntrySize(entry),
      compressedSize: getZipEntryCompressedSize(entry),
    });
  }

  const slideToLayout = new Map();
  const layoutToMaster = new Map();
  const notesToSlide = new Map();
  const chartToSlides = new Map();
  const mediaUsage = new Map();

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const ownerPath = ownerPathFromRelsPath(path);
    if (!ownerPath) continue;

    const rels = parseRelationships(await zip.files[path].async("string"));

    if (ownerPath.startsWith("ppt/slides/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_LAYOUT) {
          slideToLayout.set(ownerPath, resolveZipPath(ownerPath, rel.target));
        }
        if (rel.type === REL_NOTES_SLIDE) {
          notesToSlide.set(resolveZipPath(ownerPath, rel.target), ownerPath);
        }
        if (rel.type.includes("/chart")) {
          const chartPath = resolveZipPath(ownerPath, rel.target);
          if (!chartToSlides.has(chartPath)) chartToSlides.set(chartPath, []);
          chartToSlides.get(chartPath).push(ownerPath);
        }
      }
    }

    if (ownerPath.startsWith("ppt/slideLayouts/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_MASTER) {
          layoutToMaster.set(ownerPath, resolveZipPath(ownerPath, rel.target));
        }
      }
    }
  }

  const layoutToSlides = new Map();
  for (const [slidePath, layoutPath] of slideToLayout) {
    if (!layoutToSlides.has(layoutPath)) layoutToSlides.set(layoutPath, []);
    layoutToSlides.get(layoutPath).push(slidePath);
  }

  const masterToSlides = new Map();
  for (const [slidePath, layoutPath] of slideToLayout) {
    const masterPath = layoutToMaster.get(layoutPath);
    if (!masterPath) continue;
    if (!masterToSlides.has(masterPath)) masterToSlides.set(masterPath, []);
    masterToSlides.get(masterPath).push(slidePath);
  }

  function slideLabelsForPaths(slidePaths, via) {
    const labels = new Set();
    for (const slidePath of slidePaths) {
      const num = slidePathToNum.get(slidePath);
      if (num) labels.add(formatSlideLabel(num, via));
    }
    return labels;
  }

  function ownerShortName(ownerPath) {
    return ownerPath.split("/").pop() || ownerPath;
  }

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const ownerPath = ownerPathFromRelsPath(path);
    if (!ownerPath) continue;

    const rels = parseRelationships(await zip.files[path].async("string"));

    for (const rel of rels) {
      if (!isMediaRelationship(rel.type, rel.target)) continue;
      const mediaPath = resolveZipPath(ownerPath, rel.target);

      if (ownerPath.startsWith("ppt/slides/")) {
        const num = slidePathToNum.get(ownerPath);
        if (num) {
          addMediaUsage(mediaUsage, mediaPath, formatSlideLabel(num, "直接"));
        } else {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `${ownerShortName(ownerPath)}（表示順外）`
          );
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/slideLayouts/")) {
        const slidePaths = layoutToSlides.get(ownerPath) || [];
        if (slidePaths.length === 0) {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `未使用レイアウト: ${ownerShortName(ownerPath)}`
          );
        } else {
          for (const label of slideLabelsForPaths(slidePaths, "レイアウト")) {
            addMediaUsage(mediaUsage, mediaPath, label);
          }
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/slideMasters/")) {
        const slidePaths = masterToSlides.get(ownerPath) || [];
        if (slidePaths.length === 0) {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `未使用マスター: ${ownerShortName(ownerPath)}`
          );
        } else {
          for (const label of slideLabelsForPaths(slidePaths, "マスター")) {
            addMediaUsage(mediaUsage, mediaPath, label);
          }
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/charts/")) {
        const slidePaths = chartToSlides.get(ownerPath) || [];
        for (const label of slideLabelsForPaths(slidePaths, "チャート")) {
          addMediaUsage(mediaUsage, mediaPath, label);
        }
        if (slidePaths.length === 0) {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `未参照チャート: ${ownerShortName(ownerPath)}`
          );
        }
        continue;
      }

      if (ownerPath.startsWith("ppt/notesSlides/")) {
        const slidePath = notesToSlide.get(ownerPath);
        const num = slidePath ? slidePathToNum.get(slidePath) : null;
        if (num) {
          addMediaUsage(mediaUsage, mediaPath, formatSlideLabel(num, "ノート"));
        } else {
          addMediaUsage(
            mediaUsage,
            mediaPath,
            `ノート: ${ownerShortName(ownerPath)}（対応スライド不明）`
          );
        }
      }
    }
  }

  const items = [...mediaSizes.entries()]
    .map(([path, sizes]) => {
      const usage = mediaUsage.get(path);
      const slides = usage ? sortUsageLabels([...usage]) : [];
      const isOrphan = slides.length > 0 && slides.every(isOrphanUsageLabel);
      return {
        path,
        name: path.split("/").pop(),
        size: sizes.size,
        compressedSize: sizes.compressedSize,
        slides,
        isOrphan,
        isPreviewable: isPreviewableMediaPath(path),
      };
    })
    .sort((a, b) =>
      b.size - a.size ||
      b.compressedSize - a.compressedSize ||
      a.name.localeCompare(b.name, "ja")
    );

  const totalMediaSize = items.reduce((sum, item) => sum + item.size, 0);
  const totalMediaCompressedSize = items.reduce((sum, item) => sum + item.compressedSize, 0);
  const orphanMediaSize = items
    .filter((item) => item.isOrphan)
    .reduce((sum, item) => sum + item.size, 0);
  const orphanMediaCompressedSize = items
    .filter((item) => item.isOrphan)
    .reduce((sum, item) => sum + item.compressedSize, 0);
  const orphanCount = items.filter((item) => item.isOrphan).length;

  return {
    items,
    totalMediaSize,
    totalMediaCompressedSize,
    orphanMediaSize,
    orphanMediaCompressedSize,
    orphanCount,
    slideCount: slidePathToNum.size,
  };
}

function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function partFileName(partPath) {
  return partPath.split("/").pop();
}

function relsPathForPart(partPath) {
  const idx = partPath.lastIndexOf("/");
  if (idx === -1) return `_rels/${partFileName(partPath)}.rels`;
  return `${partPath.slice(0, idx)}/_rels/${partFileName(partPath)}.rels`;
}

async function buildDeckStructure(zip) {
  const slidePathToNum = await getPresentationSlideOrder(zip);
  const slideToLayout = new Map();
  const layoutToMaster = new Map();
  const allLayouts = new Set();
  const allMasters = new Set();

  for (const path of Object.keys(zip.files)) {
    if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)) allLayouts.add(path);
    if (/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(path)) allMasters.add(path);
  }

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const owner = ownerPathFromRelsPath(path);
    if (!owner) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));

    if (owner.startsWith("ppt/slides/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_LAYOUT) {
          slideToLayout.set(owner, resolveZipPath(owner, rel.target));
        }
      }
    }
    if (owner.startsWith("ppt/slideLayouts/")) {
      for (const rel of rels) {
        if (rel.type === REL_SLIDE_MASTER) {
          layoutToMaster.set(owner, resolveZipPath(owner, rel.target));
        }
      }
    }
  }

  const usedLayouts = new Set(slideToLayout.values());
  const unusedLayouts = [...allLayouts].filter((layout) => !usedLayouts.has(layout));
  const usedMasters = new Set();
  for (const layout of usedLayouts) {
    const master = layoutToMaster.get(layout);
    if (master) usedMasters.add(master);
  }
  const unusedMasters = [...allMasters].filter((master) => !usedMasters.has(master));

  return {
    slidePathToNum,
    slideToLayout,
    layoutToMaster,
    allLayouts,
    allMasters,
    usedLayouts,
    unusedLayouts,
    usedMasters,
    unusedMasters,
  };
}

async function collectAllReferencedMedia(zip) {
  const referenced = new Set();
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const owner = ownerPathFromRelsPath(path);
    if (!owner) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));
    for (const rel of rels) {
      if (!isMediaRelationship(rel.type, rel.target)) continue;
      const mediaPath = resolveZipPath(owner, rel.target);
      if (mediaPath.startsWith("ppt/media/")) referenced.add(mediaPath);
    }
  }
  return referenced;
}

async function computePackageOrphanMedia(zip) {
  const referenced = await collectAllReferencedMedia(zip);
  const orphanPaths = [];

  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("ppt/media/") || zip.files[path].dir) continue;
    if (!referenced.has(path)) orphanPaths.push(path);
  }

  const items = orphanPaths
    .map((path) => ({
      path,
      name: partFileName(path),
      size: zipEntryExists(zip, path) ? getZipEntrySize(zip.files[path]) : 0,
      compressedSize: zipEntryExists(zip, path)
        ? getZipEntryCompressedSize(zip.files[path])
        : 0,
      missing: !zipEntryExists(zip, path),
    }))
    .sort((a, b) => b.size - a.size || Number(a.missing) - Number(b.missing));

  return {
    items,
    totalSize: items.reduce((sum, item) => sum + item.size, 0),
    totalCompressedSize: items.reduce((sum, item) => sum + item.compressedSize, 0),
    missingCount: items.filter((item) => item.missing).length,
    paths: new Set(items.filter((item) => zipEntryExists(zip, item.path)).map((item) => item.path)),
  };
}

async function collectMediaOwners(zip) {
  const mediaOwners = new Map();

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const owner = ownerPathFromRelsPath(path);
    if (!owner) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));
    for (const rel of rels) {
      if (!isMediaRelationship(rel.type, rel.target)) continue;
      const mediaPath = resolveZipPath(owner, rel.target);
      if (!mediaPath.startsWith("ppt/media/")) continue;
      if (!mediaOwners.has(mediaPath)) mediaOwners.set(mediaPath, new Set());
      mediaOwners.get(mediaPath).add(owner);
    }
  }

  return mediaOwners;
}

async function computeStructureFreedMedia(zip, layoutsToRemove, mastersToRemove) {
  const doomed = new Set([...layoutsToRemove, ...mastersToRemove]);
  const mediaOwners = await collectMediaOwners(zip);
  const freedPaths = [];

  for (const [mediaPath, owners] of mediaOwners) {
    if (owners.size > 0 && [...owners].every((owner) => doomed.has(owner))) {
      freedPaths.push(mediaPath);
    }
  }

  const items = freedPaths
    .map((path) => ({
      path,
      name: partFileName(path),
      size: zipEntryExists(zip, path) ? getZipEntrySize(zip.files[path]) : 0,
      compressedSize: zipEntryExists(zip, path)
        ? getZipEntryCompressedSize(zip.files[path])
        : 0,
      missing: !zipEntryExists(zip, path),
    }))
    .sort((a, b) => b.size - a.size || Number(a.missing) - Number(b.missing));

  return {
    items,
    totalSize: items.reduce((sum, item) => sum + item.size, 0),
    totalCompressedSize: items.reduce((sum, item) => sum + item.compressedSize, 0),
    missingCount: items.filter((item) => item.missing).length,
    paths: new Set(items.filter((item) => zipEntryExists(zip, item.path)).map((item) => item.path)),
  };
}

async function computeNotesInfo(zip) {
  const paths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(path)
  );
  const masterPaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)
  );
  let bytes = 0;
  let compressedBytes = 0;
  for (const path of [...paths, ...masterPaths]) {
    bytes += getZipEntrySize(zip.files[path]);
    compressedBytes += getZipEntryCompressedSize(zip.files[path]);
    const relsPath = relsPathForPart(path);
    if (zip.files[relsPath]) {
      bytes += getZipEntrySize(zip.files[relsPath]);
      compressedBytes += getZipEntryCompressedSize(zip.files[relsPath]);
    }
  }
  return { count: paths.length, bytes, compressedBytes, paths };
}

function readXmlElementText(xml, fullTag) {
  const escapedTag = fullTag.replace(":", "\\:");
  const re = new RegExp(`<${escapedTag}[^>]*>([\\s\\S]*?)<\\/${escapedTag}>`, "i");
  const match = xml.match(re);
  if (!match) return "";
  return decodeXmlEntities(match[1].replace(/<[^>]+>/g, "")).trim();
}

function setXmlElementText(xml, fullTag, value) {
  const escapedTag = fullTag.replace(":", "\\:");
  const re = new RegExp(`(<${escapedTag}[^>]*>)([\\s\\S]*?)(<\\/${escapedTag}>)`, "i");
  if (!re.test(xml)) return xml;
  return xml.replace(re, `$1${escapeXmlAttr(value)}$3`);
}

async function scanDocumentProperties(zip) {
  const fields = [];
  const coreFile = zip.file("docProps/core.xml");
  if (coreFile) {
    const xml = await coreFile.async("string");
    const coreTags = [
      ["dc:title", "タイトル"],
      ["dc:subject", "件名"],
      ["dc:creator", "作成者"],
      ["dc:description", "説明"],
      ["cp:lastModifiedBy", "最終更新者"],
      ["cp:keywords", "キーワード"],
      ["cp:category", "カテゴリ"],
      ["cp:contentStatus", "コンテンツ状態"],
    ];
    for (const [tag, label] of coreTags) {
      const value = readXmlElementText(xml, tag);
      if (value) fields.push({ label, value });
    }
    const revision = readXmlElementText(xml, "cp:revision");
    if (revision && revision !== "0" && revision !== "1") {
      fields.push({ label: "リビジョン", value: revision });
    }
  }

  const appFile = zip.file("docProps/app.xml");
  if (appFile) {
    const xml = await appFile.async("string");
    const appTags = [
      ["Company", "会社"],
      ["Manager", "管理者"],
      ["HyperlinkBase", "ハイパーリンク基準"],
    ];
    for (const [tag, label] of appTags) {
      const value = readXmlElementText(xml, tag);
      if (value) fields.push({ label, value });
    }
  }

  const customFile = zip.file("docProps/custom.xml");
  if (customFile) {
    const xml = await customFile.async("string");
    const customCount = (xml.match(/<property\b/gi) || []).length;
    if (customCount > 0) {
      fields.push({ label: "カスタムプロパティ", value: `${customCount} 件` });
    }
  }

  if (zip.file("docProps/thumbnail.jpeg")) {
    fields.push({ label: "サムネイル", value: "あり" });
  }

  let removableBytes = 0;
  const thumbFile = zip.file("docProps/thumbnail.jpeg");
  if (thumbFile) {
    removableBytes += getZipEntryCompressedSize(thumbFile);
  }
  if (customFile) {
    const customXml = await customFile.async("string");
    if ((customXml.match(/<property\b/gi) || []).length > 0) {
      removableBytes += Math.max(0, getZipEntryCompressedSize(customFile) - 180);
    }
  }

  return { fields, fieldCount: fields.length, removableBytes };
}

function deleteZipEntry(zip, path) {
  if (zip.files[path]) delete zip.files[path];
}

async function clearDocumentProperties(zip) {
  let cleared = 0;
  const corePath = "docProps/core.xml";
  const coreFile = zip.file(corePath);
  if (coreFile) {
    let xml = await coreFile.async("string");
    const clearTags = [
      "dc:title", "dc:subject", "dc:creator", "dc:description",
      "cp:lastModifiedBy", "cp:keywords", "cp:category", "cp:contentStatus",
    ];
    for (const tag of clearTags) {
      const before = readXmlElementText(xml, tag);
      if (before) cleared += 1;
      xml = setXmlElementText(xml, tag, "");
    }
    const revision = readXmlElementText(xml, "cp:revision");
    if (revision && revision !== "0") {
      cleared += 1;
      xml = setXmlElementText(xml, "cp:revision", "0");
    }
    zip.file(corePath, xml);
  }

  const appPath = "docProps/app.xml";
  const appFile = zip.file(appPath);
  if (appFile) {
    let xml = await appFile.async("string");
    for (const tag of ["Company", "Manager", "HyperlinkBase"]) {
      const before = readXmlElementText(xml, tag);
      if (before) cleared += 1;
      xml = setXmlElementText(xml, tag, "");
    }
    zip.file(appPath, xml);
  }

  const customPath = "docProps/custom.xml";
  if (zip.file(customPath)) {
    const customCount = ((await zip.file(customPath).async("string")).match(/<property\b/gi) || []).length;
    if (customCount > 0) cleared += customCount;
    zip.file(
      customPath,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>'
    );
  }

  if (zip.file("docProps/thumbnail.jpeg")) {
    cleared += 1;
    deleteZipEntry(zip, "docProps/thumbnail.jpeg");
    const ctPath = "[Content_Types].xml";
    const ctFile = zip.file(ctPath);
    if (ctFile) {
      let ctXml = await ctFile.async("string");
      ctXml = ctXml.replace(
        /<Override[^>]+PartName="\/docProps\/thumbnail\.jpeg"[^>]*\/>\s*/g,
        ""
      );
      zip.file(ctPath, ctXml);
    }
  }

  return { cleared };
}

function removeRelationshipById(relsXml, rId) {
  if (!rId) return relsXml;
  return relsXml.replace(
    new RegExp(`<Relationship\\s+[^>]*\\bId="${escapeRegex(rId)}"[^>]*/>\\s*`, "g"),
    ""
  );
}

function removeXmlElementsByRId(xml, rId) {
  if (!rId) return xml;
  return xml.replace(
    new RegExp(`<[^>]+\\br:id="${escapeRegex(rId)}"[^>]*/>\\s*`, "g"),
    ""
  );
}

async function removeContentTypeOverrides(zip, deletedPaths) {
  const ctPath = "[Content_Types].xml";
  const ctFile = zip.file(ctPath);
  if (!ctFile) return;
  let xml = await ctFile.async("string");
  for (const partPath of deletedPaths) {
    if (!partPath.endsWith(".xml")) continue;
    const partName = `/${partPath}`;
    xml = xml.replace(
      new RegExp(`<Override[^>]+PartName="${escapeRegex(partName)}"[^>]*/>\\s*`, "g"),
      ""
    );
  }
  zip.file(ctPath, xml);
}

async function removeNotesFromZip(zip) {
  const deleted = new Set();
  const notesPaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(path)
  );
  let bytes = 0;

  for (const path of Object.keys(zip.files)) {
    if (!/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(path)) continue;
    let relsXml = await zip.files[path].async("string");
    const rels = parseRelationships(relsXml);
    let changed = false;
    for (const rel of rels) {
      if (rel.type === REL_NOTES_SLIDE) {
        relsXml = removeRelationshipById(relsXml, rel.id);
        changed = true;
      }
    }
    if (changed) zip.file(path, relsXml);
  }

  for (const notesPath of notesPaths) {
    bytes += getZipEntrySize(zip.files[notesPath]);
    const relsPath = relsPathForPart(notesPath);
    if (zip.files[relsPath]) bytes += getZipEntrySize(zip.files[relsPath]);
    deleteZipEntry(zip, notesPath);
    deleteZipEntry(zip, relsPath);
    deleted.add(notesPath);
    deleted.add(relsPath);
  }

  const presRelsPath = "ppt/_rels/presentation.xml.rels";
  const presPath = "ppt/presentation.xml";
  const presRelsFile = zip.file(presRelsPath);
  const presFile = zip.file(presPath);
  const notesMasterPaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(path)
  );

  if (presRelsFile && presFile) {
    let relsXml = await presRelsFile.async("string");
    let presXml = await presFile.async("string");
    const rels = parseRelationships(relsXml);
    const masterRIds = new Set();

    for (const rel of rels) {
      if (rel.type === REL_NOTES_MASTER) {
        masterRIds.add(rel.id);
        relsXml = removeRelationshipById(relsXml, rel.id);
      }
    }
    if (masterRIds.size > 0) {
      zip.file(presRelsPath, relsXml);
      for (const rId of masterRIds) {
        presXml = removeXmlElementsByRId(presXml, rId);
      }
      presXml = presXml.replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>\s*/g, "");
      zip.file(presPath, presXml);
    }
  }

  for (const masterPath of notesMasterPaths) {
    bytes += getZipEntrySize(zip.files[masterPath]);
    const relsPath = relsPathForPart(masterPath);
    if (zip.files[relsPath]) bytes += getZipEntrySize(zip.files[relsPath]);
    deleteZipEntry(zip, masterPath);
    deleteZipEntry(zip, relsPath);
    deleted.add(masterPath);
    deleted.add(relsPath);
  }

  await removeContentTypeOverrides(zip, deleted);
  return { count: notesPaths.length, bytes };
}

async function computeCleanupPlan(zip) {
  const structure = await buildDeckStructure(zip);
  const slideOrphans = await computePackageOrphanMedia(zip);
  const notes = await computeNotesInfo(zip);
  const properties = await scanDocumentProperties(zip);

  const layoutsToRemove = new Set(structure.unusedLayouts);
  for (const master of structure.unusedMasters) {
    for (const layout of structure.allLayouts) {
      if (structure.layoutToMaster.get(layout) === master) layoutsToRemove.add(layout);
    }
  }

  const layoutsOnUsedMasters = structure.unusedLayouts.filter((layout) => {
    const master = structure.layoutToMaster.get(layout);
    return !master || structure.usedMasters.has(master);
  });

  const structureFreedMedia = await computeStructureFreedMedia(
    zip,
    [...layoutsToRemove],
    [...structure.unusedMasters]
  );
  const imageCompressUsages = await collectImageCompressUsages(zip);

  return {
    structure,
    slideOrphanMedia: slideOrphans,
    structureFreedMedia,
    imageCompressUsages,
    layoutsToRemove: [...layoutsToRemove],
    mastersToRemove: [...structure.unusedMasters],
    unusedLayoutCount: layoutsOnUsedMasters.length,
    unusedMasterCount: structure.unusedMasters.length,
    notes,
    properties,
  };
}

async function removeLayoutFromMaster(zip, layoutPath, layoutToMaster) {
  const masterPath = layoutToMaster.get(layoutPath);
  if (!masterPath) return;
  const masterRelsPath = relsPathForPart(masterPath);
  const masterRelsFile = zip.file(masterRelsPath);
  const masterFile = zip.file(masterPath);
  if (!masterRelsFile || !masterFile) return;

  let relsXml = await masterRelsFile.async("string");
  const rels = parseRelationships(relsXml);
  let layoutRId = null;
  for (const rel of rels) {
    if (rel.type === REL_SLIDE_LAYOUT && resolveZipPath(masterPath, rel.target) === layoutPath) {
      layoutRId = rel.id;
      break;
    }
  }
  if (!layoutRId) return;

  relsXml = removeRelationshipById(relsXml, layoutRId);
  zip.file(masterRelsPath, relsXml);

  let masterXml = await masterFile.async("string");
  masterXml = removeXmlElementsByRId(masterXml, layoutRId);
  zip.file(masterPath, masterXml);
}

async function removeMasterFromPresentation(zip, masterPath) {
  const presRelsPath = "ppt/_rels/presentation.xml.rels";
  const presPath = "ppt/presentation.xml";
  const presRelsFile = zip.file(presRelsPath);
  const presFile = zip.file(presPath);
  if (!presRelsFile || !presFile) return;

  let relsXml = await presRelsFile.async("string");
  const rels = parseRelationships(relsXml);
  let masterRId = null;
  for (const rel of rels) {
    if (rel.type === REL_SLIDE_MASTER && resolveZipPath(presPath, rel.target) === masterPath) {
      masterRId = rel.id;
      break;
    }
  }
  if (!masterRId) return;

  relsXml = removeRelationshipById(relsXml, masterRId);
  zip.file(presRelsPath, relsXml);

  let presXml = await presFile.async("string");
  presXml = removeXmlElementsByRId(presXml, masterRId);
  zip.file(presPath, presXml);
}

async function removeUnusedStructure(zip, plan) {
  const deleted = new Set();
  const layoutToMaster = plan.structure.layoutToMaster;

  for (const layoutPath of plan.layoutsToRemove) {
    if (plan.structure.usedMasters.has(layoutToMaster.get(layoutPath))) {
      await removeLayoutFromMaster(zip, layoutPath, layoutToMaster);
    }
    deleteZipEntry(zip, layoutPath);
    deleteZipEntry(zip, relsPathForPart(layoutPath));
    deleted.add(layoutPath);
    deleted.add(relsPathForPart(layoutPath));
  }

  for (const masterPath of plan.mastersToRemove) {
    await removeMasterFromPresentation(zip, masterPath);
    deleteZipEntry(zip, masterPath);
    deleteZipEntry(zip, relsPathForPart(masterPath));
    deleted.add(masterPath);
    deleted.add(relsPathForPart(masterPath));
  }

  await removeContentTypeOverrides(zip, deleted);
  return {
    layoutsRemoved: plan.layoutsToRemove.length,
    mastersRemoved: plan.mastersToRemove.length,
  };
}

async function removePackageOrphanMedia(zip) {
  const referenced = await collectAllReferencedMedia(zip);
  let bytes = 0;
  let count = 0;
  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("ppt/media/") || zip.files[path].dir) continue;
    if (referenced.has(path)) continue;
    bytes += getZipEntrySize(zip.files[path]);
    deleteZipEntry(zip, path);
    count += 1;
  }
  return { count, bytes };
}

async function removeSlideOrphanMedia(zip, orphanPaths) {
  let bytes = 0;
  let count = 0;
  for (const path of orphanPaths) {
    if (!zip.files[path]) continue;
    bytes += getZipEntrySize(zip.files[path]);
    deleteZipEntry(zip, path);
    count += 1;
  }
  return { count, bytes };
}

async function applyFontReplaceToZip(zip, fonts) {
  let replacements = 0;
  let filesChanged = 0;
  let embeddedFontsRemoved = 0;

  for (const path of Object.keys(zip.files)) {
    const entry = zip.files[path];
    if (entry.dir) continue;

    if (path.startsWith("ppt/fonts/")) {
      embeddedFontsRemoved += 1;
      deleteZipEntry(zip, path);
      continue;
    }

    if (path.endsWith(".xml") || path.endsWith(".rels")) {
      const text = await entry.async("string");
      const { data, count } = replaceFontsInXml(path, text, fonts);
      if (data !== text) filesChanged += 1;
      replacements += count;
      zip.file(path, data);
    }
  }

  return { replacements, filesChanged, embeddedFontsRemoved };
}

function asUint8Array(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (typeof Buffer !== "undefined" && Buffer.isBuffer && Buffer.isBuffer(bytes)) {
    return new Uint8Array(bytes);
  }
  return new Uint8Array(bytes);
}

function readU16BE(buf, offset) {
  return (buf[offset] << 8) | buf[offset + 1];
}

function readU32BE(buf, offset) {
  return (
    ((buf[offset] << 24) | (buf[offset + 1] << 16) | (buf[offset + 2] << 8) | buf[offset + 3]) >>> 0
  );
}

function asciiAt(buf, offset, length) {
  let out = "";
  for (let i = 0; i < length; i++) out += String.fromCharCode(buf[offset + i]);
  return out;
}

function readPngDimensions(buf) {
  if (buf.length < 24) return null;
  if (buf[0] !== 0x89 || asciiAt(buf, 1, 3) !== "PNG") return null;
  if (asciiAt(buf, 12, 4) !== "IHDR") return null;
  const width = readU32BE(buf, 16);
  const height = readU32BE(buf, 20);
  if (!width || !height) return null;
  return { width, height };
}

function readJpegDimensions(buf) {
  if (buf.length < 10 || buf[0] !== 0xFF || buf[1] !== 0xD8) return null;
  let i = 2;
  while (i + 8 < buf.length) {
    if (buf[i] !== 0xFF) {
      i += 1;
      continue;
    }
    const marker = buf[i + 1];
    if (marker === 0xFF) {
      i += 1;
      continue;
    }
    if (marker === 0xD8 || marker === 0xD9 || marker === 0x01 || (marker >= 0xD0 && marker <= 0xD7)) {
      i += 2;
      continue;
    }
    if (i + 3 >= buf.length) break;
    const length = readU16BE(buf, i + 2);
    const isSof = marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC;
    if (isSof) {
      if (i + 8 >= buf.length) return null;
      const height = readU16BE(buf, i + 5);
      const width = readU16BE(buf, i + 7);
      if (!width || !height) return null;
      return { width, height };
    }
    if (length < 2) break;
    i += 2 + length;
  }
  return null;
}

function readWebpDimensions(buf) {
  if (buf.length < 30) return null;
  if (asciiAt(buf, 0, 4) !== "RIFF" || asciiAt(buf, 8, 4) !== "WEBP") return null;
  const fourcc = asciiAt(buf, 12, 4);
  if (fourcc === "VP8X") {
    if (buf.length < 30) return null;
    const width = 1 + (buf[24] | (buf[25] << 8) | (buf[26] << 16));
    const height = 1 + (buf[27] | (buf[28] << 8) | (buf[29] << 16));
    return { width, height };
  }
  if (fourcc === "VP8L") {
    if (buf.length < 25 || buf[20] !== 0x2F) return null;
    const bits = buf[21] | (buf[22] << 8) | (buf[23] << 16) | (buf[24] << 24);
    const width = (bits & 0x3FFF) + 1;
    const height = ((bits >> 14) & 0x3FFF) + 1;
    return { width, height };
  }
  if (fourcc === "VP8 ") {
    for (let i = 20; i + 9 < buf.length; i++) {
      if (buf[i] === 0x9D && buf[i + 1] === 0x01 && buf[i + 2] === 0x2A) {
        const width = (buf[i + 3] | (buf[i + 4] << 8)) & 0x3FFF;
        const height = (buf[i + 5] | (buf[i + 6] << 8)) & 0x3FFF;
        if (!width || !height) return null;
        return { width, height };
      }
    }
  }
  return null;
}

function readImageDimensions(bytes) {
  const buf = asUint8Array(bytes);
  return readPngDimensions(buf) || readJpegDimensions(buf) || readWebpDimensions(buf);
}

function copyUint8Prefix(bytes, maxBytes) {
  const src = asUint8Array(bytes);
  const n = Math.min(src.length, Math.max(0, maxBytes));
  return new Uint8Array(src.subarray(0, n));
}

function isStoredZipCompression(compression) {
  if (!compression) return false;
  return compression.magic === "\x00\x00" || compression.name === "STORE";
}

async function readZipEntryPrefix(entry, maxBytes = IMAGE_DIM_PREFIX_BYTES) {
  if (!entry) return new Uint8Array(0);
  const limit = Number.isFinite(maxBytes) && maxBytes > 0 ? maxBytes : IMAGE_DIM_PREFIX_BYTES;
  const data = entry._data;
  if (data && isStoredZipCompression(data.compression) && data.compressedContent) {
    const stored = data.compressedContent;
    if (typeof stored.then === "function") {
      return copyUint8Prefix(await stored, limit);
    }
    return copyUint8Prefix(stored, limit);
  }
  return copyUint8Prefix(await entry.async("uint8array"), limit);
}

async function readImageDimensionsFromEntry(entry) {
  return readImageDimensions(await readZipEntryPrefix(entry, IMAGE_DIM_PREFIX_BYTES));
}

function parseSrcRect(xml) {
  const match = /<a:srcRect\b([^>]*)\/?>/.exec(xml);
  if (!match) {
    return { l: 0, t: 0, r: 0, b: 0, visibleRatioW: 1, visibleRatioH: 1 };
  }
  const attrs = match[1];
  const read = (name) => {
    const found = new RegExp(`\\b${name}="(-?\\d+)"`).exec(attrs);
    return found ? Number(found[1]) : 0;
  };
  const l = read("l");
  const t = read("t");
  const r = read("r");
  const b = read("b");
  const visibleRatioW = Math.max(0, 1 - (l + r) / SRC_RECT_BASE);
  const visibleRatioH = Math.max(0, 1 - (t + b) / SRC_RECT_BASE);
  return { l, t, r, b, visibleRatioW, visibleRatioH };
}

function parseExtentAttrs(tag) {
  if (!tag) return null;
  const cx = Number(/cx="(-?\d+)"/.exec(tag)?.[1]);
  const cy = Number(/cy="(-?\d+)"/.exec(tag)?.[1]);
  if (!Number.isFinite(cx) || !Number.isFinite(cy)) return null;
  return { cx: Math.abs(cx), cy: Math.abs(cy) };
}

function parseXfrmExtents(xml) {
  const xfrmMatch = /<(?:a|p|xdr|cdr):xfrm\b[^>]*>[\s\S]*?<\/(?:a|p|xdr|cdr):xfrm>/.exec(xml);
  if (!xfrmMatch) return null;
  const block = xfrmMatch[0];
  const ext = parseExtentAttrs(/<a:ext\b[^>]*>/.exec(block)?.[0]);
  const chExt = parseExtentAttrs(/<a:chExt\b[^>]*>/.exec(block)?.[0]);
  if (!ext) return null;
  return {
    cx: ext.cx,
    cy: ext.cy,
    chCx: chExt ? chExt.cx : 0,
    chCy: chExt ? chExt.cy : 0,
  };
}

function parseGroupScale(grpBlock) {
  const prMatch = /<(?:p|a|xdr|cdr):grpSpPr\b[\s\S]*?<\/(?:p|a|xdr|cdr):grpSpPr>/.exec(grpBlock);
  if (!prMatch) return null;
  const xfrm = parseXfrmExtents(prMatch[0]);
  if (!xfrm || !xfrm.cx || !xfrm.cy || !xfrm.chCx || !xfrm.chCy) return null;
  return {
    scaleX: xfrm.cx / xfrm.chCx,
    scaleY: xfrm.cy / xfrm.chCy,
  };
}

function parseBlipEmbedId(xml) {
  const blip = /<a:blip\b[^>]*>/.exec(xml);
  if (!blip) return null;
  return /\br:embed="([^"]+)"/.exec(blip[0])?.[1] || null;
}

function findBalancedTaggedEnd(xml, startIndex, localName) {
  const openRe = new RegExp(`<[A-Za-z0-9]+:${localName}(?=[\\s>/])`, "g");
  const closeRe = new RegExp(`</[A-Za-z0-9]+:${localName}>`, "g");
  const gt = xml.indexOf(">", startIndex);
  if (gt === -1) return -1;
  if (xml[gt - 1] === "/") return gt + 1;
  let depth = 1;
  let i = gt + 1;
  while (i < xml.length && depth > 0) {
    openRe.lastIndex = i;
    closeRe.lastIndex = i;
    const open = openRe.exec(xml);
    const close = closeRe.exec(xml);
    if (!close) return -1;
    if (open && open.index < close.index) {
      depth += 1;
      i = open.index + open[0].length;
    } else {
      depth -= 1;
      i = close.index + close[0].length;
    }
  }
  return depth === 0 ? i : -1;
}

function neededPixelsForUse(displayCxEmu, displayCyEmu, visibleRatioW, visibleRatioH, ppi) {
  const visW = visibleRatioW > 0 ? visibleRatioW : 1;
  const visH = visibleRatioH > 0 ? visibleRatioH : 1;
  const width = Math.max(1, Math.ceil((displayCxEmu / EMU_PER_INCH) * ppi / visW));
  const height = Math.max(1, Math.ceil((displayCyEmu / EMU_PER_INCH) * ppi / visH));
  return { width, height };
}

function collectPicUse(block, rIdToMedia, scaleX, scaleY, onUse) {
  const rId = parseBlipEmbedId(block);
  if (!rId) return;
  const mediaPath = rIdToMedia.get(rId);
  if (!mediaPath || !isCompressibleImagePath(mediaPath)) return;
  const xfrm = parseXfrmExtents(block);
  if (!xfrm || !xfrm.cx || !xfrm.cy) return;
  const srcRect = parseSrcRect(block);
  onUse(mediaPath, xfrm.cx * scaleX, xfrm.cy * scaleY, srcRect);
}

function readXmlTagNumber(tag, name) {
  if (!tag) return 0;
  const found = new RegExp(`\\b${name}="(-?\\d+)"`).exec(tag);
  return found ? Math.abs(Number(found[1])) : 0;
}

function collectTableBlipUses(frameBlock, rIdToMedia, scaleX, scaleY, onUse) {
  const tblOpen = /<a:tbl\b/.exec(frameBlock);
  if (!tblOpen) return;
  const tblEnd = findBalancedTaggedEnd(frameBlock, tblOpen.index, "tbl");
  if (tblEnd === -1) return;
  const tblXml = frameBlock.slice(tblOpen.index, tblEnd);

  const colWidths = [];
  const colRe = /<a:gridCol\b[^>]*>/g;
  let colMatch;
  while ((colMatch = colRe.exec(tblXml)) !== null) {
    colWidths.push(readXmlTagNumber(colMatch[0], "w"));
  }

  const rowHeights = [];
  const rows = [];
  let rowCursor = 0;
  while (rowCursor < tblXml.length) {
    const sliced = tblXml.slice(rowCursor);
    const trOpen = /<a:tr\b/.exec(sliced);
    if (!trOpen) break;
    const start = rowCursor + trOpen.index;
    const end = findBalancedTaggedEnd(tblXml, start, "tr");
    if (end === -1) break;
    const openEnd = tblXml.indexOf(">", start);
    const height = readXmlTagNumber(
      openEnd === -1 ? "" : tblXml.slice(start, openEnd + 1),
      "h"
    );
    rowHeights.push(height);
    rows.push(tblXml.slice(start, end));
    rowCursor = end;
  }

  const frameXfrm = parseXfrmExtents(frameBlock);
  const tableW = colWidths.reduce((sum, width) => sum + width, 0);
  const tableH = rowHeights.reduce((sum, height) => sum + height, 0);
  const frameScaleX = frameXfrm && frameXfrm.cx && tableW ? frameXfrm.cx / tableW : 1;
  const frameScaleY = frameXfrm && frameXfrm.cy && tableH ? frameXfrm.cy / tableH : 1;

  rows.forEach((rowXml, rowIndex) => {
    let colIndex = 0;
    let cellCursor = 0;
    while (cellCursor < rowXml.length && colIndex < Math.max(colWidths.length, 1)) {
      const sliced = rowXml.slice(cellCursor);
      const tcOpen = /<a:tc\b/.exec(sliced);
      if (!tcOpen) break;
      const start = cellCursor + tcOpen.index;
      const end = findBalancedTaggedEnd(rowXml, start, "tc");
      if (end === -1) break;
      const cellXml = rowXml.slice(start, end);
      const openEnd = rowXml.indexOf(">", start);
      const openTag = openEnd === -1 ? "" : rowXml.slice(start, openEnd + 1);
      const gridSpan = Math.max(1, readXmlTagNumber(openTag, "gridSpan") || 1);
      const rowSpan = Math.max(1, readXmlTagNumber(openTag, "rowSpan") || 1);

      let cellW = 0;
      for (let i = 0; i < gridSpan && colIndex + i < colWidths.length; i++) {
        cellW += colWidths[colIndex + i];
      }
      let cellH = 0;
      for (let i = 0; i < rowSpan && rowIndex + i < rowHeights.length; i++) {
        cellH += rowHeights[rowIndex + i];
      }
      if (!cellW && frameXfrm) cellW = frameXfrm.cx;
      if (!cellH && frameXfrm) cellH = frameXfrm.cy;

      const rId = parseBlipEmbedId(cellXml);
      if (rId) {
        const mediaPath = rIdToMedia.get(rId);
        if (mediaPath && isCompressibleImagePath(mediaPath) && cellW && cellH) {
          onUse(
            mediaPath,
            cellW * frameScaleX * scaleX,
            cellH * frameScaleY * scaleY,
            parseSrcRect(cellXml)
          );
        }
      }

      colIndex += 1;
      cellCursor = end;
    }
  });
}

function walkDrawingXml(xml, rIdToMedia, scaleX, scaleY, onUse) {
  const re = /<(?:p|a|pic|xdr|cdr):(graphicFrame|grpSp|cxnSp|pic|sp)\b/g;
  let cursor = 0;
  while (cursor < xml.length) {
    re.lastIndex = cursor;
    const match = re.exec(xml);
    if (!match) break;
    const kind = match[1];
    const start = match.index;
    const end = findBalancedTaggedEnd(xml, start, kind);
    if (end === -1) {
      cursor = start + match[0].length;
      continue;
    }
    const block = xml.slice(start, end);
    if (kind === "grpSp") {
      const innerStart = xml.indexOf(">", start);
      if (innerStart !== -1) {
        const scale = parseGroupScale(block) || { scaleX: 1, scaleY: 1 };
        walkDrawingXml(
          xml.slice(innerStart + 1, end),
          rIdToMedia,
          scaleX * scale.scaleX,
          scaleY * scale.scaleY,
          onUse
        );
      }
    } else if (kind === "graphicFrame") {
      collectTableBlipUses(block, rIdToMedia, scaleX, scaleY, onUse);
    } else {
      collectPicUse(block, rIdToMedia, scaleX, scaleY, onUse);
    }
    cursor = end;
  }
}

function collectBackgroundImageUses(xml, rIdToMedia, slideCx, slideCy, onUse) {
  if (!slideCx || !slideCy) return;
  const bgRe = /<(?:p|a):bg\b[\s\S]*?<\/(?:p|a):bg>/g;
  let match;
  while ((match = bgRe.exec(xml)) !== null) {
    const rId = parseBlipEmbedId(match[0]);
    if (!rId) continue;
    const mediaPath = rIdToMedia.get(rId);
    if (!mediaPath || !isCompressibleImagePath(mediaPath)) continue;
    onUse(mediaPath, slideCx, slideCy, parseSrcRect(match[0]));
  }
}

async function readSlideSizeEmu(zip) {
  const presFile = zip.file("ppt/presentation.xml");
  if (!presFile) return { cx: DEFAULT_SLIDE_CX_EMU, cy: DEFAULT_SLIDE_CY_EMU };
  const xml = await presFile.async("string");
  const tag = /<p:sldSz\b[^>]*>/.exec(xml)?.[0];
  const parsed = parseExtentAttrs(tag || "");
  if (!parsed || !parsed.cx || !parsed.cy) {
    return { cx: DEFAULT_SLIDE_CX_EMU, cy: DEFAULT_SLIDE_CY_EMU };
  }
  return parsed;
}

function addImageCompressUse(byPath, mediaPath, displayCxEmu, displayCyEmu, srcRect) {
  if (!byPath.has(mediaPath)) {
    byPath.set(mediaPath, { path: mediaPath, uses: [] });
  }
  byPath.get(mediaPath).uses.push({
    displayCxEmu,
    displayCyEmu,
    visibleRatioW: srcRect.visibleRatioW,
    visibleRatioH: srcRect.visibleRatioH,
  });
}

async function collectImageCompressUsages(zip) {
  const byPath = new Map();
  const slideSize = await readSlideSizeEmu(zip);
  const onUse = (mediaPath, displayCxEmu, displayCyEmu, srcRect) => {
    addImageCompressUse(byPath, mediaPath, displayCxEmu, displayCyEmu, srcRect);
  };

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".rels")) continue;
    const ownerPath = ownerPathFromRelsPath(path);
    if (!ownerPath || !zip.files[ownerPath]) continue;
    const rels = parseRelationships(await zip.files[path].async("string"));
    const rIdToMedia = new Map();
    for (const rel of rels) {
      if (!rel.id || !rel.type.includes("/image")) continue;
      const mediaPath = resolveZipPath(ownerPath, rel.target);
      if (mediaPath.startsWith("ppt/media/") && isCompressibleImagePath(mediaPath)) {
        rIdToMedia.set(rel.id, mediaPath);
      }
    }
    if (rIdToMedia.size === 0) continue;
    const xml = await zip.files[ownerPath].async("string");
    walkDrawingXml(xml, rIdToMedia, 1, 1, onUse);
    collectBackgroundImageUses(xml, rIdToMedia, slideSize.cx, slideSize.cy, onUse);
  }

  const items = [];
  for (const item of byPath.values()) {
    if (!zipEntryExists(zip, item.path)) continue;
    const entry = zip.files[item.path];
    const dims = await readImageDimensionsFromEntry(entry);
    items.push({
      path: item.path,
      uses: item.uses,
      width: dims?.width || 0,
      height: dims?.height || 0,
      mime: mimeForMediaPath(item.path),
      size: getZipEntrySize(entry),
      compressedSize: getZipEntryCompressedSize(entry),
    });
  }
  return items;
}

function normalizeImageCompressOptions(options = {}) {
  const ppi = Number(options.imagePpi);
  const quality = Number(options.jpegQuality);
  return {
    imagePpi: Number.isFinite(ppi) && ppi > 0 ? ppi : DEFAULT_IMAGE_PPI,
    jpegQuality: Number.isFinite(quality) && quality > 0 && quality <= 1
      ? quality
      : DEFAULT_JPEG_QUALITY,
  };
}

function buildImageCompressJobs(usages, options = {}, excludePaths = null) {
  const { imagePpi, jpegQuality } = normalizeImageCompressOptions(options);
  const jobs = [];
  for (const item of usages || []) {
    if (excludePaths && excludePaths.has(item.path)) continue;
    if (!item.uses || item.uses.length === 0) continue;
    if (!item.width || !item.height) continue;
    if (!isCompressibleImagePath(item.path)) continue;

    let maxScale = 0;
    for (const use of item.uses) {
      const needed = neededPixelsForUse(
        use.displayCxEmu,
        use.displayCyEmu,
        use.visibleRatioW,
        use.visibleRatioH,
        imagePpi
      );
      const scale = Math.max(needed.width / item.width, needed.height / item.height);
      if (scale > maxScale) maxScale = scale;
    }
    const scale = Math.min(1, maxScale);
    const targetWidth = Math.max(1, Math.round(item.width * scale));
    const targetHeight = Math.max(1, Math.round(item.height * scale));
    const needsResize = targetWidth < item.width || targetHeight < item.height;
    const isJpeg = /jpe?g$/i.test(item.path);
    if (!needsResize && !isJpeg) continue;

    const areaRatio = (targetWidth * targetHeight) / (item.width * item.height);
    let factor = areaRatio;
    if (isJpeg) factor *= jpegQuality / 0.92;
    factor = Math.min(Math.max(factor, 0.02), 0.98);
    const estimatedCompressed = Math.max(1, Math.round((item.compressedSize || item.size || 0) * factor));

    jobs.push({
      path: item.path,
      mime: item.mime || mimeForMediaPath(item.path),
      width: item.width,
      height: item.height,
      targetWidth,
      targetHeight,
      needsResize,
      jpegQuality: isJpeg ? jpegQuality : undefined,
      size: item.size,
      compressedSize: item.compressedSize,
      estimatedCompressed,
    });
  }
  return jobs;
}

function mediaPathsRemovedByOptions(plan, options) {
  const paths = new Set();
  if (!plan) return paths;
  if (options.removeOrphanMedia) {
    for (const item of plan.slideOrphanMedia?.items || []) paths.add(item.path);
  }
  if (options.removeUnusedStructure) {
    for (const item of plan.structureFreedMedia?.items || []) paths.add(item.path);
  }
  return paths;
}

function estimateImageCompressBytes(usages, options, excludePaths) {
  return buildImageCompressJobs(usages, options, excludePaths)
    .reduce((sum, job) => sum + Math.max(0, (job.compressedSize || 0) - job.estimatedCompressed), 0);
}

async function collectImageCompressJobs(zip, options = {}) {
  const usages = await collectImageCompressUsages(zip);
  return buildImageCompressJobs(usages, options);
}

function codecResultBytes(result) {
  if (!result) return null;
  if (result instanceof Uint8Array) return result;
  if (result.bytes instanceof Uint8Array) return result.bytes;
  if (result.bytes) return asUint8Array(result.bytes);
  return null;
}

async function mapPool(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const index = next;
      next += 1;
      results[index] = await fn(items[index], index);
    }
  }
  const workers = Math.max(1, Math.min(limit, items.length));
  await Promise.all(Array.from({ length: workers }, worker));
  return results;
}

async function compressImagesInZip(zip, options, codec, usages) {
  const empty = { count: 0, bytes: 0, skipped: 0, jobs: [] };
  if (!options?.compressImages) return empty;
  if (!codec || typeof codec.encode !== "function") return empty;

  const jobs = usages
    ? buildImageCompressJobs(usages, options)
    : await collectImageCompressJobs(zip, options);
  let count = 0;
  let bytes = 0;
  let skipped = 0;

  await mapPool(jobs, IMAGE_ENCODE_CONCURRENCY, async (job) => {
    try {
      if (!zip.files[job.path]) return;
      const input = await zip.files[job.path].async("uint8array");
      const encoded = codecResultBytes(await codec.encode(input, job));
      if (!encoded) {
        skipped += 1;
        return;
      }
      if (encoded.byteLength >= input.byteLength) return;
      zip.file(job.path, encoded);
      count += 1;
      bytes += input.byteLength - encoded.byteLength;
    } catch {
      skipped += 1;
    }
  });

  return { count, bytes, skipped, jobs };
}

function zipEntryCompressedBytes(zip, path) {
  if (!zipEntryExists(zip, path)) return 0;
  return getZipEntryCompressedSize(zip.files[path]);
}

function computeReductionEstimate(plan, options, zip) {
  if (!plan || !zip) return 0;

  let bytes = 0;

  if (options.replaceFonts) {
    for (const path of Object.keys(zip.files)) {
      if (path.startsWith("ppt/fonts/") && !zip.files[path].dir) {
        bytes += getZipEntryCompressedSize(zip.files[path]);
      }
    }
  }

  if (options.removeOrphanMedia) {
    for (const item of plan.slideOrphanMedia.items) {
      if (!item.missing) {
        bytes += item.compressedSize || zipEntryCompressedBytes(zip, item.path);
      }
    }
  }

  if (options.removeUnusedStructure) {
    for (const path of plan.layoutsToRemove) {
      bytes += zipEntryCompressedBytes(zip, path);
      bytes += zipEntryCompressedBytes(zip, relsPathForPart(path));
    }
    for (const path of plan.mastersToRemove) {
      bytes += zipEntryCompressedBytes(zip, path);
      bytes += zipEntryCompressedBytes(zip, relsPathForPart(path));
    }
    bytes += plan.structureFreedMedia.totalCompressedSize;
  }

  if (options.removeNotes && plan.notes.count > 0) {
    bytes += plan.notes.compressedBytes || plan.notes.bytes;
  }

  if (options.removeProperties && plan.properties.removableBytes > 0) {
    bytes += plan.properties.removableBytes;
  }

  if (options.compressImages) {
    bytes += estimateImageCompressBytes(
      plan.imageCompressUsages,
      options,
      mediaPathsRemovedByOptions(plan, options)
    );
  }

  return bytes;
}

function replaceFontsInXml(name, data, fonts) {
  const titleFont = fonts.title.trim() || DEFAULT_FONT;
  const bodyFont = fonts.body.trim() || DEFAULT_FONT;
  let total = 0;

  if (name === "ppt/presentation.xml") {
    let result = countReplace(data, EMBEDDED_FONT_LST_RE, "");
    data = result.data;
    total += result.count;
    data = data.replace(/embedTrueTypeFonts="1"/g, 'embedTrueTypeFonts="0"');
    const contextual = replaceFontsContextual(data, titleFont, bodyFont);
    return { data: contextual.data, count: total + contextual.count };
  }

  if (name === "ppt/_rels/presentation.xml.rels") {
    const result = countReplace(data, FONT_REL_RE, "");
    return { data: result.data, count: result.count };
  }

  if (!name.endsWith(".xml")) {
    return { data, count: 0 };
  }

  const contextual = replaceFontsContextual(data, titleFont, bodyFont);
  return { data: contextual.data, count: contextual.count };
}

function buildEmbeddedFontDetail(files) {
  return files
    .sort((a, b) => b.size - a.size)
    .map((file) => `${file.name} (${formatBytes(file.size)})`)
    .join(" / ");
}

async function mapEmbeddedFontFileSizes(zip) {
  const typefaceFiles = new Map();

  const presFile = zip.file("ppt/presentation.xml");
  const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  if (!presFile || !presRelsFile) {
    return typefaceFiles;
  }

  const presXml = await presFile.async("string");
  const rels = parseRelationships(await presRelsFile.async("string"));
  const rIdToFontPath = new Map();
  for (const rel of rels) {
    if (rel.type === REL_FONT) {
      rIdToFontPath.set(rel.id, resolveZipPath("ppt/presentation.xml", rel.target));
    }
  }

  EMBEDDED_FONT_BLOCK_RE.lastIndex = 0;
  let block;
  while ((block = EMBEDDED_FONT_BLOCK_RE.exec(presXml)) !== null) {
    const inner = block[1];
    const tfMatch = inner.match(/typeface="([^"]*)"/);
    if (!tfMatch) continue;
    const typeface = decodeXmlEntities(tfMatch[1]).trim();
    if (!typeface) continue;

    const rIds = new Set();
    const ridRe = /\br:id="([^"]+)"/g;
    let ridMatch;
    while ((ridMatch = ridRe.exec(inner)) !== null) {
      rIds.add(ridMatch[1]);
    }

    const files = [];
    let bytes = 0;
    for (const rId of rIds) {
      const fontPath = rIdToFontPath.get(rId);
      if (!fontPath || !zip.files[fontPath] || files.includes(fontPath)) continue;
      files.push(fontPath);
      bytes += getZipEntrySize(zip.files[fontPath]);
    }

    if (files.length === 0) continue;
    typefaceFiles.set(typeface, {
      bytes,
      files: files.map((path) => ({
        path,
        name: partFileName(path),
        size: getZipEntrySize(zip.files[path]),
      })),
    });
  }

  return typefaceFiles;
}

async function extractFontsFromZipAsync(zip) {
  const fonts = new Map();
  const embeddedMap = await mapEmbeddedFontFileSizes(zip);

  function ensureFont(name) {
    if (!fonts.has(name)) {
      fonts.set(name, { name, count: 0 });
    }
    return fonts.get(name);
  }

  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".xml")) continue;
    const text = await zip.files[path].async("string");

    TYPEFACE_CAPTURE_RE.lastIndex = 0;
    let match;
    while ((match = TYPEFACE_CAPTURE_RE.exec(text)) !== null) {
      const name = decodeXmlEntities(match[1]).trim();
      if (name) ensureFont(name).count += 1;
    }
  }

  return [...fonts.values()]
    .map((font) => {
      const embedded = embeddedMap.get(font.name);
      return {
        name: font.name,
        count: font.count,
        fileBytes: embedded?.bytes || 0,
        fileDetail: embedded?.files?.length
          ? buildEmbeddedFontDetail(embedded.files)
          : "",
      };
    })
    .sort((a, b) =>
      b.fileBytes - a.fileBytes ||
      b.count - a.count ||
      a.name.localeCompare(b.name, "ja")
    );
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DEFAULT_FONT,
    decodeXmlEntities,
    escapeXmlAttr,
    escapeHtml,
    isPreviewableMediaPath,
    isCompressibleImagePath,
    mimeForMediaPath,
    formatBytes,
    formatSizeChange,
    countReplace,
    replaceTypefacesInString,
    countTypefaces,
    findNextSpShapeStart,
    findBalancedSpBlockEnd,
    replaceFontsInStyleBlocks,
    replaceFontsContextual,
    getZipEntrySize,
    getZipEntryCompressedSize,
    zipEntryExists,
    parseRelationships,
    resolveZipPath,
    isMediaRelationship,
    ownerPathFromRelsPath,
    getPresentationSlideOrder,
    formatSlideLabel,
    addMediaUsage,
    isOrphanUsageLabel,
    sortUsageLabels,
    analyzeMediaUsage,
    escapeRegex,
    partFileName,
    relsPathForPart,
    buildDeckStructure,
    collectAllReferencedMedia,
    computePackageOrphanMedia,
    collectMediaOwners,
    computeStructureFreedMedia,
    computeNotesInfo,
    readXmlElementText,
    setXmlElementText,
    scanDocumentProperties,
    clearDocumentProperties,
    removeNotesFromZip,
    computeCleanupPlan,
    removeRelationshipById,
    removeXmlElementsByRId,
    removeContentTypeOverrides,
    deleteZipEntry,
    removeLayoutFromMaster,
    removeMasterFromPresentation,
    removeUnusedStructure,
    removePackageOrphanMedia,
    removeSlideOrphanMedia,
    applyFontReplaceToZip,
    zipEntryCompressedBytes,
    computeReductionEstimate,
    EMU_PER_INCH,
    SRC_RECT_BASE,
    DEFAULT_IMAGE_PPI,
    DEFAULT_JPEG_QUALITY,
    IMAGE_DIM_PREFIX_BYTES,
    WARN_INPUT_FILE_BYTES,
    MAX_INPUT_FILE_BYTES,
    assessInputFileSize,
    readZipEntryPrefix,
    readImageDimensionsFromEntry,
    readImageDimensions,
    parseSrcRect,
    parseXfrmExtents,
    parseGroupScale,
    neededPixelsForUse,
    collectImageCompressUsages,
    buildImageCompressJobs,
    collectImageCompressJobs,
    estimateImageCompressBytes,
    compressImagesInZip,
    replaceFontsInXml,
    buildEmbeddedFontDetail,
    mapEmbeddedFontFileSizes,
    extractFontsFromZipAsync,
  };
}
