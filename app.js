// Reference Board Builder_14
// Stages A-D:
// - read folder tree
// - analyze structure
// - calculate layout
// - preview widgets on board
// - build final board with images

const { board } = window.miro;

const SAT_CODE_MAX = 99;
const SAT_BOOST = 4.0;
const SAT_GROUP_THRESHOLD = 35;
const NO_COLOR_KEY = "__no_color__";
const IMAGE_EXTENSIONS = new Set(["jpg", "jpeg", "png", "webp", "bmp", "gif", "avif"]);
const APP_VERSION = "Reference Board Builder_14";
const APP_META_ID = "reference-board-builder";
const FRAME_VERTICAL_GAP = 1200;
const COLUMN_HEADER_FILL = "#f4d44d";
const SUBTYPE_HEADER_FILL = "#f4d44d";
const OUTLINE_COLOR = "#111111";
const TEXT_BOX_FILL = "#ffffff";
const TEXT_BOX_BORDER = "#cbd5e1";
const PREVIEW_FILL = "#d1d5db";
const PREVIEW_BORDER = "#9ca3af";
const FRAME_FILL = "#808080";
const OUTLINE_HEADER_OFFSET_X = 150000;
const OUTLINE_HEADER_OFFSET_Y = 45000;
const OUTLINE_CATEGORY_WIDTH = 6000;
const OUTLINE_CATEGORY_HEIGHT = 2500;
const OUTLINE_SUBTYPE_WIDTH = 3500;
const OUTLINE_SUBTYPE_HEIGHT = 1500;
const OUTLINE_SUBTYPE_GAP_X = 600;
const OUTLINE_SUBTYPE_GAP_Y = 350;
const OUTLINE_SECTION_GAP_Y = 1500;
const COLUMN_HEADER_BORDER_WIDTH = 24;
const COLUMN_HEADER_FONT_SIZE = 700;
const SUBTYPE_HEADER_FONT_SIZE = 200;
const IMAGE_CREATE_CONCURRENCY = 3;
const DEFAULT_CATEGORY_COLORS = [
  "#D6BF52",
  "#98B65A",
  "#D98A56",
  "#7FB7D6",
  "#7EA37A",
  "#9A85C8",
  "#5E88C4",
  "#5F7A3E",
];
const FALLBACK_CATEGORY_COLOR = "#ffffff";

const state = {
  files: [],
  tree: null,
  layout: null,
  navigation: [],
  isRendering: false,
};

const imageInfoCache = new WeakMap();
const scaledImageCache = new WeakMap();

function clampNotificationMessage(message, fallback = "Operation failed") {
  const raw = (message == null ? "" : String(message)).replace(/\s+/g, " ").trim();
  const safe = raw || fallback;
  if (safe.length <= 80) return safe;
  return `${safe.slice(0, 79).trimEnd()}…`;
}

async function notify(kind, message, details) {
  const safeMessage = clampNotificationMessage(message);
  try {
    if (details !== undefined) {
      const logger = kind === "showError" ? console.error : kind === "showWarning" ? console.warn : console.log;
      logger(`[${APP_VERSION}]`, safeMessage, details);
    }

    const notifications = board && board.notifications ? board.notifications : null;
    if (!notifications) return;

    if (typeof notifications[kind] === "function") {
      await notifications[kind](safeMessage);
      return;
    }

    if (typeof notifications.show === "function") {
      await notifications.show({
        message: safeMessage,
        type: kind === "showError" ? "error" : "info",
      });
    }
  } catch (error) {
    console.error(`[${APP_VERSION}] notification failed`, { safeMessage, error, details });
  }
}

const notifyInfo = (message, details) => notify("showInfo", message, details);
const notifyWarning = (message, details) => notify("showWarning", message, details);
const notifyError = (message, details) => notify("showError", message, details);

function byName(a, b) {
  return String(a.name || "").localeCompare(String(b.name || ""), "ru", {
    numeric: true,
    sensitivity: "base",
  });
}


function parseOrderedPart(rawValue) {
  const raw = String(rawValue == null ? "" : rawValue).trim();
  const match = raw.match(/^(\d+)[_\-\s]*(.+)$/u);
  if (!match) {
    return {
      raw,
      order: null,
      displayName: raw,
    };
  }
  const displayName = String(match[2] || "").trim() || raw;
  return {
    raw,
    order: Number.parseInt(match[1], 10),
    displayName,
  };
}

function sanitizeDisplayName(rawValue) {
  return parseOrderedPart(rawValue).displayName;
}

function sanitizeHierarchicalName(rawValue) {
  return String(rawValue == null ? "" : rawValue)
    .split("/")
    .map((part) => sanitizeDisplayName(part))
    .join(" / ");
}

function compareOrderedEntries(a, b) {
  const aMeta = a && a.orderMeta ? a.orderMeta : parseOrderedPart(a && a.rawName ? a.rawName : a && a.name ? a.name : "");
  const bMeta = b && b.orderMeta ? b.orderMeta : parseOrderedPart(b && b.rawName ? b.rawName : b && b.name ? b.name : "");

  if (aMeta.order != null && bMeta.order != null && aMeta.order !== bMeta.order) {
    return aMeta.order - bMeta.order;
  }
  if (aMeta.order != null && bMeta.order == null) return -1;
  if (aMeta.order == null && bMeta.order != null) return 1;

  return String(aMeta.displayName || "").localeCompare(String(bMeta.displayName || ""), "ru", {
    numeric: true,
    sensitivity: "base",
  });
}

function getCategoryDefaultColor(index) {
  return DEFAULT_CATEGORY_COLORS[index] || FALLBACK_CATEGORY_COLOR;
}

function formatNumber(value) {
  return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(Math.round(value || 0));
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isImageFile(file) {
  const type = String(file.type || "").toLowerCase();
  if (type.startsWith("image/")) return true;
  const name = String(file.name || "");
  const ext = name.includes(".") ? name.split(".").pop().toLowerCase() : "";
  return IMAGE_EXTENSIONS.has(ext);
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.crossOrigin = "anonymous";
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}

async function decodeImageFromFile(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    return await loadImage(objectUrl);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function getBrightnessAndSaturationFromImageElement(
  image,
  smallSize = 50,
  blurPx = 3,
  cropTopRatio = 0.3,
  cropSideRatio = 0.2
) {
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  const width = smallSize;
  const height = smallSize;
  canvas.width = width;
  canvas.height = height;

  const prevFilter = ctx.filter || "none";
  try {
    ctx.filter = `blur(${blurPx}px)`;
  } catch (_) {}
  ctx.drawImage(image, 0, 0, width, height);
  ctx.filter = prevFilter;

  const cropY = Math.floor(height * cropTopRatio);
  const cropH = height - cropY;
  const cropX = Math.floor(width * cropSideRatio);
  const cropW = width - 2 * cropX;
  if (cropH <= 0 || cropW <= 0) return null;

  let imageData;
  try {
    imageData = ctx.getImageData(cropX, cropY, cropW, cropH);
  } catch (error) {
    console.warn("getImageData failed:", error);
    return null;
  }

  const data = imageData.data;
  const totalPixels = cropW * cropH;
  let sumY = 0;
  let sumDiff = 0;

  for (let i = 0; i < data.length; i += 4) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
    sumY += y;
    const maxv = Math.max(r, g, b);
    const minv = Math.min(r, g, b);
    sumDiff += maxv - minv;
  }

  return {
    brightness: (sumY / totalPixels) / 255,
    saturation: (sumDiff / totalPixels) / 255,
  };
}

async function getImageInfo(file) {
  const cached = imageInfoCache.get(file);
  if (cached) return cached;

  const promise = (async () => {
    const image = await decodeImageFromFile(file);
    const width = image.naturalWidth || image.width || 1;
    const height = image.naturalHeight || image.height || 1;
    let brightness = 0.5;
    let saturation = 0;

    try {
      const metrics = getBrightnessAndSaturationFromImageElement(image);
      if (metrics) {
        brightness = metrics.brightness;
        saturation = metrics.saturation;
      }
    } catch (error) {
      console.warn("Color analysis failed", { fileName: file.name, error });
    }

    try {
      image.src = "";
    } catch (_) {}

    const briCode = Math.max(0, Math.min(999, Math.round((1 - brightness) * 999)));
    const satCode = Math.max(
      0,
      Math.min(SAT_CODE_MAX, Math.round(Math.min(1, saturation * SAT_BOOST) * SAT_CODE_MAX))
    );

    return {
      width,
      height,
      brightness,
      saturation,
      briCode,
      satCode,
      satGroup: satCode <= SAT_GROUP_THRESHOLD ? 0 : 1,
    };
  })();

  imageInfoCache.set(file, promise);
  return promise;
}

function readConfig() {
  const getNumber = (id, fallback) => {
    const el = document.getElementById(id);
    const value = el ? Number(el.value) : fallback;
    return Number.isFinite(value) ? value : fallback;
  };

  return {
    columnShapeWidth: getNumber("columnShapeWidth", 6500),
    columnShapeHeight: getNumber("columnShapeHeight", 2600),
    frameShapeWidth: getNumber("frameShapeWidth", 2000),
    frameShapeHeight: getNumber("frameShapeHeight", 800),
    textBoxWidth: getNumber("textBoxWidth", 856),
    textBoxHeight: getNumber("textBoxHeight", 550),
    textFontSize: getNumber("textFontSize", 130),
    frameWidth: getNumber("frameWidth", 12000),
    imageGap: getNumber("imageGap", 25),
    groupGap: getNumber("groupGap", 300),
    innerPadding: getNumber("innerPadding", 120),
    columnGap: getNumber("columnGap", 13000),
    headerToFramesGap: getNumber("headerToFramesGap", 4000),
    outlineOffsetX: getNumber("outlineOffsetX", OUTLINE_HEADER_OFFSET_X),
    outlineOffsetY: getNumber("outlineOffsetY", OUTLINE_HEADER_OFFSET_Y),
    outlineCategoryWidth: getNumber("outlineCategoryWidth", OUTLINE_CATEGORY_WIDTH),
    outlineCategoryHeight: getNumber("outlineCategoryHeight", OUTLINE_CATEGORY_HEIGHT),
    outlineSubtypeWidth: getNumber("outlineSubtypeWidth", OUTLINE_SUBTYPE_WIDTH),
    outlineSubtypeHeight: getNumber("outlineSubtypeHeight", OUTLINE_SUBTYPE_HEIGHT),
    outlineSubtypeGapX: getNumber("outlineSubtypeGapX", OUTLINE_SUBTYPE_GAP_X),
    outlineSubtypeGapY: getNumber("outlineSubtypeGapY", OUTLINE_SUBTYPE_GAP_Y),
    outlineSectionGapY: getNumber("outlineSectionGapY", OUTLINE_SECTION_GAP_Y),
    sortByColor: !!document.getElementById("sortByColor")?.checked,
  };
}

function setStatus(message, tone = "info") {
  const box = document.getElementById("statusBox");
  if (!box) return;
  box.className = `status${tone === "info" ? "" : ` ${tone}`}`;
  box.textContent = String(message || "");
}

function updateStats(summary) {
  document.getElementById("statCategories").textContent = String(summary.categories || 0);
  document.getElementById("statFrames").textContent = String(summary.frames || 0);
  document.getElementById("statGroups").textContent = String(summary.groups || 0);
  document.getElementById("statImages").textContent = String(summary.images || 0);
}

function resetResults() {
  document.getElementById("structureResult").innerHTML = "";
  document.getElementById("layoutResult").innerHTML = "";
  document.getElementById("layoutNote").textContent = "";
  updateStats({ categories: 0, frames: 0, groups: 0, images: 0 });
  state.navigation = [];
  renderNavigationPanel();
}


function parseReferenceTree(files) {
  const categoryMap = new Map();
  let rootName = "";
  const skipped = [];
  let imageCount = 0;

  for (const file of files) {
    if (!isImageFile(file)) continue;

    const relPath = String(file.webkitRelativePath || file.name || "");
    const parts = relPath.split("/").filter(Boolean);
    if (!parts.length) {
      skipped.push({ fileName: file.name || "image", reason: "empty-relative-path" });
      continue;
    }

    if (!rootName) rootName = parts[0];
    const dirs = parts.slice(1, -1);
    if (dirs.length < 2) {
      skipped.push({ fileName: relPath, reason: "path-must-have-category-and-subtype" });
      continue;
    }

    const rawCategoryName = dirs[0];
    const rawSubtypeName = dirs[1];
    const rawColorName = dirs.length >= 3 ? dirs.slice(2).join(" / ") : null;

    const categoryMeta = parseOrderedPart(rawCategoryName);
    const subtypeMeta = parseOrderedPart(rawSubtypeName);
    const colorMeta = rawColorName
      ? {
          ...parseOrderedPart(rawColorName.split("/")[0]),
          displayName: sanitizeHierarchicalName(rawColorName),
        }
      : null;

    let category = categoryMap.get(rawCategoryName);
    if (!category) {
      category = {
        rawName: rawCategoryName,
        name: categoryMeta.displayName,
        orderMeta: categoryMeta,
        subtypesMap: new Map(),
      };
      categoryMap.set(rawCategoryName, category);
    }

    let subtype = category.subtypesMap.get(rawSubtypeName);
    if (!subtype) {
      subtype = {
        rawName: rawSubtypeName,
        name: subtypeMeta.displayName,
        orderMeta: subtypeMeta,
        groupsMap: new Map(),
      };
      category.subtypesMap.set(rawSubtypeName, subtype);
    }

    const groupKey = rawColorName ? rawColorName : NO_COLOR_KEY;
    let group = subtype.groupsMap.get(groupKey);
    if (!group) {
      group = {
        key: groupKey,
        rawName: rawColorName,
        name: colorMeta ? colorMeta.displayName : null,
        orderMeta: colorMeta,
        hasColorFolder: !!rawColorName,
        images: [],
      };
      subtype.groupsMap.set(groupKey, group);
    }

    group.images.push({
      file,
      name: file.name || "image",
      relativePath: relPath,
      rootName,
      categoryName: category.name,
      subtypeName: subtype.name,
      colorName: group.name,
    });
    imageCount += 1;
  }

  const categories = Array.from(categoryMap.values())
    .map((category) => ({
      rawName: category.rawName,
      name: category.name,
      orderMeta: category.orderMeta,
      subtypes: Array.from(category.subtypesMap.values())
        .map((subtype) => ({
          rawName: subtype.rawName,
          name: subtype.name,
          orderMeta: subtype.orderMeta,
          groups: Array.from(subtype.groupsMap.values())
            .map((group) => ({
              key: group.key,
              rawName: group.rawName,
              name: group.name,
              orderMeta: group.orderMeta,
              hasColorFolder: group.hasColorFolder,
              images: group.images.sort((a, b) => byName(a, b)),
            }))
            .sort((a, b) => {
              if (a.hasColorFolder && !b.hasColorFolder) return -1;
              if (!a.hasColorFolder && b.hasColorFolder) return 1;
              return compareOrderedEntries(a, b);
            }),
        }))
        .sort(compareOrderedEntries),
    }))
    .sort(compareOrderedEntries);

  const summary = {
    categories: categories.length,
    frames: categories.reduce((sum, category) => sum + category.subtypes.length, 0),
    groups: categories.reduce(
      (sum, category) => sum + category.subtypes.reduce((inner, subtype) => inner + subtype.groups.length, 0),
      0
    ),
    images: imageCount,
    skipped,
  };

  return {
    rootName,
    categories,
    summary,
  };
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.crossOrigin = "anonymous";
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}

async function decodeImageFromFile(file) {
  const objectUrl = URL.createObjectURL(file);
  try {
    return await loadImage(objectUrl);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function getBrightnessAndSaturationFromImageElement(
  image,
  smallSize = 50,
  blurPx = 3,
  cropTopRatio = 0.3,
  cropSideRatio = 0.2
) {
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  const width = smallSize;
  const height = smallSize;
  canvas.width = width;
  canvas.height = height;

  const prevFilter = ctx.filter || "none";
  try {
    ctx.filter = `blur(${blurPx}px)`;
  } catch (_) {}
  ctx.drawImage(image, 0, 0, width, height);
  ctx.filter = prevFilter;

  const cropY = Math.floor(height * cropTopRatio);
  const cropH = height - cropY;
  const cropX = Math.floor(width * cropSideRatio);
  const cropW = width - 2 * cropX;
  if (cropH <= 0 || cropW <= 0) return null;

  let imageData;
  try {
    imageData = ctx.getImageData(cropX, cropY, cropW, cropH);
  } catch (error) {
    console.warn("getImageData failed:", error);
    return null;
  }

  const data = imageData.data;
  const totalPixels = cropW * cropH;
  let sumY = 0;
  let sumDiff = 0;

  for (let i = 0; i < data.length; i += 4) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
    sumY += y;
    const maxv = Math.max(r, g, b);
    const minv = Math.min(r, g, b);
    sumDiff += maxv - minv;
  }

  return {
    brightness: (sumY / totalPixels) / 255,
    saturation: (sumDiff / totalPixels) / 255,
  };
}

async function getImageInfo(file) {
  const cached = imageInfoCache.get(file);
  if (cached) return cached;

  const promise = (async () => {
    const image = await decodeImageFromFile(file);
    const width = image.naturalWidth || image.width || 1;
    const height = image.naturalHeight || image.height || 1;
    let brightness = 0.5;
    let saturation = 0;

    try {
      const metrics = getBrightnessAndSaturationFromImageElement(image);
      if (metrics) {
        brightness = metrics.brightness;
        saturation = metrics.saturation;
      }
    } catch (error) {
      console.warn("Color analysis failed", { fileName: file.name, error });
    }

    try {
      image.src = "";
    } catch (_) {}

    const briCode = Math.max(0, Math.min(999, Math.round((1 - brightness) * 999)));
    const satCode = Math.max(
      0,
      Math.min(SAT_CODE_MAX, Math.round(Math.min(1, saturation * SAT_BOOST) * SAT_CODE_MAX))
    );

    return {
      width,
      height,
      brightness,
      saturation,
      briCode,
      satCode,
      satGroup: satCode <= SAT_GROUP_THRESHOLD ? 0 : 1,
    };
  })();

  imageInfoCache.set(file, promise);
  return promise;
}

function readConfig() {
  const getNumber = (id, fallback) => {
    const el = document.getElementById(id);
    const value = el ? Number(el.value) : fallback;
    return Number.isFinite(value) ? value : fallback;
  };

  return {
    columnShapeWidth: getNumber("columnShapeWidth", 6500),
    columnShapeHeight: getNumber("columnShapeHeight", 2600),
    frameShapeWidth: getNumber("frameShapeWidth", 2000),
    frameShapeHeight: getNumber("frameShapeHeight", 800),
    textBoxWidth: getNumber("textBoxWidth", 856),
    textBoxHeight: getNumber("textBoxHeight", 550),
    textFontSize: getNumber("textFontSize", 130),
    frameWidth: getNumber("frameWidth", 12000),
    imageGap: getNumber("imageGap", 25),
    groupGap: getNumber("groupGap", 300),
    innerPadding: getNumber("innerPadding", 120),
    columnGap: getNumber("columnGap", 13000),
    headerToFramesGap: getNumber("headerToFramesGap", 4000),
    outlineOffsetX: getNumber("outlineOffsetX", OUTLINE_HEADER_OFFSET_X),
    outlineOffsetY: getNumber("outlineOffsetY", OUTLINE_HEADER_OFFSET_Y),
    outlineCategoryWidth: getNumber("outlineCategoryWidth", OUTLINE_CATEGORY_WIDTH),
    outlineCategoryHeight: getNumber("outlineCategoryHeight", OUTLINE_CATEGORY_HEIGHT),
    outlineSubtypeWidth: getNumber("outlineSubtypeWidth", OUTLINE_SUBTYPE_WIDTH),
    outlineSubtypeHeight: getNumber("outlineSubtypeHeight", OUTLINE_SUBTYPE_HEIGHT),
    outlineSubtypeGapX: getNumber("outlineSubtypeGapX", OUTLINE_SUBTYPE_GAP_X),
    outlineSubtypeGapY: getNumber("outlineSubtypeGapY", OUTLINE_SUBTYPE_GAP_Y),
    outlineSectionGapY: getNumber("outlineSectionGapY", OUTLINE_SECTION_GAP_Y),
    sortByColor: !!document.getElementById("sortByColor")?.checked,
  };
}

function setStatus(message, tone = "info") {
  const box = document.getElementById("statusBox");
  if (!box) return;
  box.className = `status${tone === "info" ? "" : ` ${tone}`}`;
  box.textContent = String(message || "");
}

function updateStats(summary) {
  document.getElementById("statCategories").textContent = String(summary.categories || 0);
  document.getElementById("statFrames").textContent = String(summary.frames || 0);
  document.getElementById("statGroups").textContent = String(summary.groups || 0);
  document.getElementById("statImages").textContent = String(summary.images || 0);
}

function resetResults() {
  document.getElementById("structureResult").innerHTML = "";
  document.getElementById("layoutResult").innerHTML = "";
  document.getElementById("layoutNote").textContent = "";
  updateStats({ categories: 0, frames: 0, groups: 0, images: 0 });
  state.navigation = [];
  renderNavigationPanel();
}

function parseReferenceTree(files) {
  const categoryMap = new Map();
  let rootName = "";
  const skipped = [];
  let imageCount = 0;

  for (const file of files) {
    if (!isImageFile(file)) continue;

    const relPath = String(file.webkitRelativePath || file.name || "");
    const parts = relPath.split("/").filter(Boolean);
    if (!parts.length) {
      skipped.push({ fileName: file.name || "image", reason: "empty-relative-path" });
      continue;
    }

    if (!rootName) rootName = parts[0];
    const dirs = parts.slice(1, -1);
    if (dirs.length < 2) {
      skipped.push({ fileName: relPath, reason: "path-must-have-category-and-subtype" });
      continue;
    }

    const categoryName = dirs[0];
    const subtypeName = dirs[1];
    const colorName = dirs.length >= 3 ? dirs.slice(2).join(" / ") : null;

    let category = categoryMap.get(categoryName);
    if (!category) {
      category = { name: categoryName, subtypesMap: new Map() };
      categoryMap.set(categoryName, category);
    }

    let subtype = category.subtypesMap.get(subtypeName);
    if (!subtype) {
      subtype = { name: subtypeName, groupsMap: new Map() };
      category.subtypesMap.set(subtypeName, subtype);
    }

    const groupKey = colorName ? colorName : NO_COLOR_KEY;
    let group = subtype.groupsMap.get(groupKey);
    if (!group) {
      group = {
        key: groupKey,
        name: colorName,
        hasColorFolder: !!colorName,
        images: [],
      };
      subtype.groupsMap.set(groupKey, group);
    }

    group.images.push({
      file,
      name: file.name || "image",
      relativePath: relPath,
      rootName,
      categoryName,
      subtypeName,
      colorName,
    });
    imageCount += 1;
  }

  const categories = Array.from(categoryMap.values())
    .map((category) => ({
      name: category.name,
      subtypes: Array.from(category.subtypesMap.values())
        .map((subtype) => ({
          name: subtype.name,
          groups: Array.from(subtype.groupsMap.values())
            .map((group) => ({
              key: group.key,
              name: group.name,
              hasColorFolder: group.hasColorFolder,
              images: group.images.sort((a, b) => byName(a, b)),
            }))
            .sort((a, b) => {
              if (a.hasColorFolder && !b.hasColorFolder) return -1;
              if (!a.hasColorFolder && b.hasColorFolder) return 1;
              return byName(a, b);
            }),
        }))
        .sort(byName),
    }))
    .sort(byName);

  const summary = {
    categories: categories.length,
    frames: categories.reduce((sum, category) => sum + category.subtypes.length, 0),
    groups: categories.reduce(
      (sum, category) => sum + category.subtypes.reduce((inner, subtype) => inner + subtype.groups.length, 0),
      0
    ),
    images: imageCount,
    skipped,
  };

  return {
    rootName,
    categories,
    summary,
  };
}

function measureWrappedTextLines(text, width, fontSize) {
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  ctx.font = `${fontSize}px Arial`;
  const words = String(text || "").split(/\s+/).filter(Boolean);
  if (!words.length) return 1;

  const innerWidth = Math.max(1, width - 24);
  let line = "";
  let lines = 1;

  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (ctx.measureText(next).width <= innerWidth) {
      line = next;
      continue;
    }
    if (!line) {
      line = word;
      lines += 1;
      line = "";
      continue;
    }
    lines += 1;
    line = word;
  }

  return Math.max(1, lines);
}

function sortGroupImagesByColor(imagesWithInfo) {
  return [...imagesWithInfo].sort((a, b) => {
    if (a.info.satGroup !== b.info.satGroup) return a.info.satGroup - b.info.satGroup;
    if (a.info.briCode !== b.info.briCode) return a.info.briCode - b.info.briCode;
    if (a.info.satCode !== b.info.satCode) return a.info.satCode - b.info.satCode;
    return String(a.image.name || "").localeCompare(String(b.image.name || ""), "ru", {
      numeric: true,
      sensitivity: "base",
    });
  });
}

function packRows(items, maxWidth, gap) {
  const rows = [];
  let currentRow = [];
  let currentWidth = 0;

  for (const item of items) {
    const itemWidth = item.scaledWidth;
    const nextWidth = currentRow.length ? currentWidth + gap + itemWidth : currentWidth + itemWidth;

    if (currentRow.length && nextWidth > maxWidth) {
      rows.push({ items: currentRow, width: currentWidth });
      currentRow = [item];
      currentWidth = itemWidth;
      continue;
    }

    currentRow.push(item);
    currentWidth = nextWidth;
  }

  if (currentRow.length) rows.push({ items: currentRow, width: currentWidth });
  return rows;
}

async function buildLayout(tree, config) {
  const layout = {
    rootName: tree.rootName,
    config,
    categories: [],
    summary: {
      categories: tree.summary.categories,
      frames: tree.summary.frames,
      groups: tree.summary.groups,
      images: tree.summary.images,
      totalRows: 0,
      maxFrameHeight: 0,
      maxColumnHeight: 0,
    },
  };

  for (let categoryIndex = 0; categoryIndex < tree.categories.length; categoryIndex += 1) {
    const category = tree.categories[categoryIndex];
    const layoutCategory = {
      name: category.name,
      index: categoryIndex,
      header: {
        width: config.columnShapeWidth,
        height: config.columnShapeHeight,
      },
      frames: [],
      totalHeight: 0,
    };

    for (const subtype of category.subtypes) {
      const layoutFrame = {
        name: subtype.name,
        width: config.frameWidth,
        height: 0,
        relativeTop: 0,
        groups: [],
      };

      let currentY = config.innerPadding;

      for (const group of subtype.groups) {
        const showText = group.hasColorFolder;
        const wrappedLines = showText ? measureWrappedTextLines(group.name, config.textBoxWidth, config.textFontSize) : 0;
        const extraLines = Math.max(0, wrappedLines - 1);
        const extraLineHeight = config.textFontSize * 1.35;
        const textBlockHeight = showText ? config.textBoxHeight + extraLines * extraLineHeight : 0;

        const imageTargetHeight = config.textBoxHeight;
        const availableWidth = Math.max(
          1,
          config.frameWidth - config.innerPadding - config.innerPadding - (showText ? config.textBoxWidth + config.innerPadding : 0)
        );

        const imagesWithInfo = await Promise.all(
          group.images.map(async (image) => ({ image, info: await getImageInfo(image.file) }))
        );
        const orderedImages = config.sortByColor ? sortGroupImagesByColor(imagesWithInfo) : imagesWithInfo;
        const scaledItems = orderedImages.map((item) => ({
          image: item.image,
          info: item.info,
          scaledWidth: item.info.height > 0 ? item.info.width * (imageTargetHeight / item.info.height) : imageTargetHeight,
        }));

        const rows = packRows(scaledItems, availableWidth, config.imageGap);
        const rowsWithPositions = rows.map((row, rowIndex) => {
          let cursorX = 0;
          const items = row.items.map((item, itemIndex) => {
            const placed = {
              name: item.image.name,
              file: item.image.file,
              width: item.scaledWidth,
              height: imageTargetHeight,
              xOffset: cursorX + item.scaledWidth / 2,
              yOffset: rowIndex * (imageTargetHeight + config.imageGap) + imageTargetHeight / 2,
              satCode: item.info.satCode,
              briCode: item.info.briCode,
              title: item.image.name,
              index: itemIndex,
            };
            cursorX += item.scaledWidth + config.imageGap;
            return placed;
          });
          return {
            index: rowIndex,
            width: row.width,
            count: row.items.length,
            items,
          };
        });

        const imageBlockHeight = rows.length
          ? rows.length * imageTargetHeight + Math.max(0, rows.length - 1) * config.imageGap
          : 0;
        const blockHeight = Math.max(textBlockHeight, imageBlockHeight, imageTargetHeight);

        layout.summary.totalRows += rows.length;

        layoutFrame.groups.push({
          key: group.key,
          name: group.name,
          hasColorFolder: group.hasColorFolder,
          imageCount: group.images.length,
          y: currentY,
          blockHeight,
          text: {
            visible: showText,
            width: showText ? config.textBoxWidth : 0,
            height: textBlockHeight,
            wrappedLines,
          },
          images: {
            targetHeight: imageTargetHeight,
            availableWidth,
            blockHeight: imageBlockHeight,
            rows: rowsWithPositions,
          },
        });

        currentY += blockHeight + config.groupGap;
      }

      const hasGroups = layoutFrame.groups.length > 0;
      layoutFrame.height = hasGroups ? currentY - config.groupGap + config.innerPadding : config.innerPadding * 2 + config.textBoxHeight;
      layout.summary.maxFrameHeight = Math.max(layout.summary.maxFrameHeight, layoutFrame.height);
      layoutCategory.frames.push(layoutFrame);
    }

    let frameTopCursor = config.columnShapeHeight + config.headerToFramesGap;
    layoutCategory.frames.forEach((frame, frameIndex) => {
      frame.relativeTop = frameTopCursor;
      frame.relativeCenterY = frameTopCursor + frame.height / 2;
      frame.index = frameIndex;
      frameTopCursor += frame.height + FRAME_VERTICAL_GAP;
    });

    layoutCategory.totalHeight = layoutCategory.frames.length
      ? frameTopCursor - FRAME_VERTICAL_GAP
      : config.columnShapeHeight;
    layout.summary.maxColumnHeight = Math.max(layout.summary.maxColumnHeight, layoutCategory.totalHeight);
    layout.categories.push(layoutCategory);
  }

  return layout;
}

function renderStructure(tree) {
  const skippedHtml = tree.summary.skipped.length
    ? `<div class="layout-note">Пропущено файлов: ${tree.summary.skipped.length}</div>`
    : "";

  return `
    <div><strong>Корневая папка:</strong> ${escapeHtml(tree.rootName || "—")}</div>
    ${skippedHtml}
    ${tree.categories.map((category) => `
      <details open>
        <summary>${escapeHtml(category.name)} <span class="pill">${category.subtypes.length} фрейм.</span></summary>
        ${category.subtypes.map((subtype) => `
          <details open>
            <summary>${escapeHtml(subtype.name)} <span class="pill">${subtype.groups.length} групп</span></summary>
            ${subtype.groups.map((group) => `
              <div>
                ${group.hasColorFolder ? escapeHtml(group.name) : "Без цветовых папок"}
                <span class="pill">${group.images.length} карт.</span>
              </div>
            `).join("")}
          </details>
        `).join("")}
      </details>
    `).join("")}
  `;
}

function renderLayout(layout) {
  return layout.categories.map((category) => `
    <details open>
      <summary>${escapeHtml(category.name)} <span class="pill">${category.frames.length} фрейм.</span></summary>
      ${category.frames.map((frame) => `
        <details open>
          <summary>
            ${escapeHtml(frame.name)}
            <span class="pill">${formatNumber(frame.width)} × ${formatNumber(frame.height)}</span>
          </summary>
          ${frame.groups.map((group) => `
            <div style="margin-bottom:8px;">
              <strong>${group.hasColorFolder ? escapeHtml(group.name) : "Без цветовых папок"}</strong>
              <span class="pill">${group.imageCount} карт.</span>
              <span class="pill">${group.images.rows.length} ряд.</span>
              <span class="pill">блок ${formatNumber(group.blockHeight)}</span>
              ${group.text.visible ? `<div class="muted">Текст: ${group.text.wrappedLines} строк, ${formatNumber(group.text.height)} по высоте</div>` : `<div class="muted">Текстовый блок пропущен</div>`}
              <div class="muted">Зона картинок: ${formatNumber(group.images.availableWidth)} по ширине</div>
            </div>
          `).join("")}
        </details>
      `).join("")}
    </details>
  `).join("");
}

function describeLayout(layout) {
  return [
    `Всего рядов картинок: ${layout.summary.totalRows}`,
    `Максимальная высота фрейма: ${formatNumber(layout.summary.maxFrameHeight)}`,
    `Максимальная высота колонки: ${formatNumber(layout.summary.maxColumnHeight)}`,
    `Ширина фрейма: ${formatNumber(layout.config.frameWidth)}`,
    `Сортировка по цвету: ${layout.config.sortByColor ? "включена" : "выключена"}`,
    `Используй вкладку статистики для проверки структуры и раскладки.`,
  ].join("\n");
}


async function handleFolderSelected(fileList) {
  const files = Array.from(fileList || []).filter(isImageFile);
  state.files = files;
  state.tree = null;
  state.layout = null;
  resetResults();

  const label = document.getElementById("folderLabel");
  if (!files.length) {
    if (label) label.textContent = "В выбранной папке не найдено изображений";
    setStatus("В выбранной папке не нашлось файлов изображений.", "warning");
    return;
  }

  const root = String(files[0].webkitRelativePath || "").split("/").filter(Boolean)[0] || "Выбранная папка";
  if (label) label.textContent = `${root} · ${files.length} файлов`;

  try {
    const layout = await ensureLayout();
    if (!layout) return;
    setStatus(
      `Папка загружена: ${root}
Категорий: ${layout.summary.categories}
Фреймов: ${layout.summary.frames}
Картинок: ${layout.summary.images}`,
      "info"
    );
  } catch (error) {
    console.error(error);
    setStatus("Не удалось обработать папку. Подробности в консоли.", "error");
  }
}


async function ensureTree() {
  if (!state.files.length) {
    setStatus("Сначала выбери папку с изображениями.", "warning");
    await notifyWarning("Сначала выбери папку");
    return null;
  }
  const tree = state.tree || parseReferenceTree(state.files);
  state.tree = tree;
  updateStats(tree.summary);
  document.getElementById("structureResult").innerHTML = renderStructure(tree);
  return tree;
}

async function ensureLayout() {
  const tree = await ensureTree();
  if (!tree) return null;
  const layout = await buildLayout(tree, readConfig());
  state.layout = layout;
  document.getElementById("layoutResult").innerHTML = renderLayout(layout);
  document.getElementById("layoutNote").textContent = describeLayout(layout);
  return layout;
}

let refreshTimer = null;
function scheduleStatsRefresh() {
  if (!state.files.length || state.isRendering) return;
  if (refreshTimer) clearTimeout(refreshTimer);
  refreshTimer = setTimeout(async () => {
    try {
      await ensureLayout();
      setStatus("Параметры обновлены. Статистика пересчитана.", "info");
    } catch (error) {
      console.error(error);
      setStatus("Не удалось пересчитать статистику. Подробности в консоли.", "error");
    }
  }, 180);
}

async function handleAnalyzeStructure() {
  try {
    const tree = await ensureTree();
    if (!tree) return;

    document.getElementById("layoutResult").innerHTML = "";
    document.getElementById("layoutNote").textContent = "";
    state.layout = null;

    const skippedCount = tree.summary.skipped.length;
    if (!tree.summary.images) {
      setStatus("Структура не собрана. Нужны изображения минимум на уровне Категория / Подтип / файл.", "error");
      await notifyError("Структура пустая");
      return;
    }

    setStatus(
      `Структура собрана.\nКатегорий: ${tree.summary.categories}\nФреймов: ${tree.summary.frames}\nГрупп: ${tree.summary.groups}\n${skippedCount ? `Пропущено файлов: ${skippedCount}` : "Ошибок структуры нет."}`,
      skippedCount ? "warning" : "info"
    );
    await notifyInfo("Структура проверена");
  } catch (error) {
    console.error(error);
    setStatus("Не удалось разобрать структуру папки. Подробности в консоли.", "error");
    await notifyError("Ошибка разбора структуры", error);
  }
}

async function handleAnalyzeLayout() {
  try {
    setStatus("Идет расчет раскладки…", "info");
    const layout = await ensureLayout();
    if (!layout) return;

    setStatus(
      `Раскладка рассчитана.\nВсего рядов: ${layout.summary.totalRows}\nМакс. высота фрейма: ${formatNumber(layout.summary.maxFrameHeight)}\nМожно создавать превью и итоговую доску.`,
      "info"
    );
    await notifyInfo("Раскладка рассчитана");
  } catch (error) {
    console.error(error);
    setStatus("Не удалось рассчитать раскладку. Подробности в консоли.", "error");
    await notifyError("Ошибка расчета раскладки", error);
  }
}


function buildScene(layout, viewport) {
  const config = layout.config;
  const centerX = viewport.x + viewport.width / 2;
  const centerY = viewport.y + viewport.height / 2;
  const count = layout.categories.length;
  const firstCenterX = centerX - ((Math.max(0, count - 1) * config.columnGap) / 2);
  const top = centerY - layout.summary.maxColumnHeight / 2;

  const categories = layout.categories.map((category, categoryIndex) => {
    const x = firstCenterX + categoryIndex * config.columnGap;
    const color = getCategoryDefaultColor(categoryIndex);
    const header = {
      x,
      y: top + config.columnShapeHeight / 2,
      width: config.columnShapeWidth,
      height: config.columnShapeHeight,
      title: category.name,
      color,
    };

    const frames = category.frames.map((frame) => {
      const frameLeft = x - frame.width / 2;
      const frameTop = top + frame.relativeTop;
      const frameScene = {
        x,
        y: frameTop + frame.height / 2,
        width: frame.width,
        height: frame.height,
        top: frameTop,
        left: frameLeft,
        name: frame.name,
        color,
        subtitleShape: {
          x: frameLeft - config.frameShapeWidth / 2,
          y: frameTop + config.frameShapeHeight / 2,
          width: config.frameShapeWidth,
          height: config.frameShapeHeight,
          title: frame.name,
          color,
        },
        groups: [],
      };

      frame.groups.forEach((group) => {
        const blockTop = frameTop + group.y;
        const groupScene = {
          name: group.name,
          hasColorFolder: group.hasColorFolder,
          text: null,
          images: [],
        };

        const imagesLeft = frameLeft + config.innerPadding + (group.text.visible ? config.textBoxWidth + config.innerPadding : 0);

        if (group.text.visible) {
          groupScene.text = {
            x: frameLeft + config.innerPadding + config.textBoxWidth / 2,
            y: blockTop + group.text.height / 2,
            width: config.textBoxWidth,
            height: group.text.height,
            title: group.name,
          };
        }

        group.images.rows.forEach((row) => {
          row.items.forEach((item) => {
            groupScene.images.push({
              file: item.file,
              title: item.title,
              x: imagesLeft + item.xOffset,
              y: blockTop + item.yOffset,
              width: item.width,
              height: item.height,
            });
          });
        });

        frameScene.groups.push(groupScene);
      });

      return frameScene;
    });

    return { name: category.name, color, header, frames };
  });

  let outline = [];
  if (categories.length > 0) {
    const firstHeader = categories[0].header;
    const firstHeaderLeft = firstHeader.x - firstHeader.width / 2;
    const firstHeaderTop = firstHeader.y - firstHeader.height / 2;
    const outlineLeft = firstHeaderLeft - config.outlineOffsetX;
    const outlineTop = firstHeaderTop - config.outlineOffsetY;
    let sectionTop = outlineTop;

    outline = categories.map((category) => {
      const categoryLeft = outlineLeft;
      const categoryTop = sectionTop;
      const section = {
        title: category.name,
        color: category.color,
        headerTargetTitle: category.name,
        headerShape: {
          x: categoryLeft + config.outlineCategoryWidth / 2,
          y: categoryTop + config.outlineCategoryHeight / 2,
          width: config.outlineCategoryWidth,
          height: config.outlineCategoryHeight,
          title: category.name,
          color: category.color,
        },
        subtypeShapes: [],
      };

      const subtypeLeft = categoryLeft + config.outlineCategoryWidth + config.outlineSubtypeGapX;
      category.frames.forEach((frame, index) => {
        section.subtypeShapes.push({
          x: subtypeLeft + config.outlineSubtypeWidth / 2,
          y: categoryTop + config.outlineSubtypeHeight / 2 + index * (config.outlineSubtypeHeight + config.outlineSubtypeGapY),
          width: config.outlineSubtypeWidth,
          height: config.outlineSubtypeHeight,
          title: frame.name,
          color: category.color,
          targetCategoryTitle: category.name,
          targetSubtypeTitle: frame.name,
        });
      });

      const subtypeStackHeight = section.subtypeShapes.length
        ? config.outlineSubtypeHeight * section.subtypeShapes.length + config.outlineSubtypeGapY * Math.max(0, section.subtypeShapes.length - 1)
        : 0;
      const sectionHeight = Math.max(config.outlineCategoryHeight, subtypeStackHeight);
      sectionTop += sectionHeight + config.outlineSectionGapY;

      return section;
    });
  }

  return { categories, outline, config };
}

function makeHeaderContent(title) {
  return `<p><strong>${escapeHtml(title)}</strong></p>`;
}

function makeLinkedHeaderContent(title, url) {
  return makeHeaderContent(title);
}

function getBoardIdFromInfo(info) {
  if (!info || typeof info !== "object") return null;
  return info.id || info.boardId || info.board_id || info.shortId || null;
}

async function getBoardInfoSafe() {
  try {
    if (board && typeof board.getInfo === "function") {
      return await board.getInfo();
    }
  } catch (error) {
    console.warn("[Reference Board Builder] board.getInfo failed", error);
  }
  return null;
}

function buildWidgetDeepLink(boardInfo, targetWidget) {
  const targetId = targetWidget && targetWidget.id ? String(targetWidget.id) : "";
  const boardId = getBoardIdFromInfo(boardInfo);
  if (!targetId || !boardId) return null;
  return `https://miro.com/app/board/${encodeURIComponent(boardId)}/?moveToWidget=${encodeURIComponent(targetId)}&cot=14`;
}

async function trySyncAfterTemporaryPropChange(widget, propName, value) {
  if (!widget || !(propName in widget) || typeof widget.sync !== "function") return false;

  const previousValue = widget[propName];
  try {
    widget[propName] = value;
    await widget.sync();
    return true;
  } catch (error) {
    console.warn("[Reference Board Builder] whole-item link prop sync failed", {
      propName,
      error,
    });

    try {
      widget[propName] = previousValue;
      await widget.sync();
    } catch (_) {}

    return false;
  }
}

async function tryApplyMiroLink(sourceWidget, targetWidget, url, metadata = {}) {
  let linkApplied = false;
  const targetId = targetWidget && targetWidget.id ? targetWidget.id : null;

  try {
    if (sourceWidget && typeof sourceWidget.setMetadata === "function") {
      await sourceWidget.setMetadata(APP_META_ID, {
        ...metadata,
        targetId,
        targetUrl: url || null,
        linkApplied: false,
        linkMode: "whole-item",
        app: APP_VERSION,
      });
    }
  } catch (_) {}

  if (url) {
    const attempts = [
      async () => {
        if (sourceWidget && typeof sourceWidget.setLink === "function") {
          await sourceWidget.setLink(url);
          return true;
        }
        return false;
      },
      async () => {
        if (sourceWidget && typeof sourceWidget.setLink === "function") {
          await sourceWidget.setLink({ url });
          return true;
        }
        return false;
      },
      async () => {
        if (sourceWidget && typeof sourceWidget.addLink === "function") {
          await sourceWidget.addLink({ url });
          return true;
        }
        return false;
      },
      async () => {
        if (sourceWidget && typeof sourceWidget.setLinks === "function") {
          await sourceWidget.setLinks([{ url }]);
          return true;
        }
        return false;
      },
      async () => trySyncAfterTemporaryPropChange(sourceWidget, "link", url),
      async () => trySyncAfterTemporaryPropChange(sourceWidget, "link", { url }),
      async () => trySyncAfterTemporaryPropChange(sourceWidget, "links", [{ url }]),
      async () => trySyncAfterTemporaryPropChange(sourceWidget, "url", url),
    ];

    for (const attempt of attempts) {
      try {
        const result = await attempt();
        if (result) {
          linkApplied = true;
          break;
        }
      } catch (error) {
        console.warn("[Reference Board Builder] whole-item link attempt failed", error);
      }
    }
  }

  try {
    if (sourceWidget && typeof sourceWidget.setMetadata === "function") {
      await sourceWidget.setMetadata(APP_META_ID, {
        ...metadata,
        targetId,
        targetUrl: url || null,
        linkApplied,
        linkMode: "whole-item",
        app: APP_VERSION,
      });
    }
  } catch (_) {}

  return linkApplied;
}

function makeTextBoxContent(title) {
  return `<p>&nbsp;</p><p>${escapeHtml(title)}</p><p>&nbsp;</p>`;
}

async function createFrameSafe(params) {
  const baseParams = {
    ...params,
    title: "",
    style: {
      ...(params.style || {}),
      fillColor: (params.style && params.style.fillColor) || FRAME_FILL,
    },
  };

  let frameWidget;
  try {
    frameWidget = await board.createFrame(baseParams);
  } catch (_) {
    frameWidget = await board.createFrame({ ...params, title: "" });
  }

  try {
    if ("title" in frameWidget) frameWidget.title = "";
    if (frameWidget.style && typeof frameWidget.style === "object") {
      frameWidget.style.fillColor = (baseParams.style && baseParams.style.fillColor) || FRAME_FILL;
    }
    if (typeof frameWidget.sync === "function") {
      await frameWidget.sync();
    }
  } catch (_) {}

  return frameWidget;
}

function stripHtmlLinks(content) {
  return String(content || "").replace(/<a\b[^>]*>(.*?)<\/a>/gis, "$1");
}

async function createShapeSafe(params) {
  try {
    return await board.createShape(params);
  } catch (error) {
    const content = String(params && params.content ? params.content : "");
    if (content.includes("<a")) {
      return await board.createShape({
        ...params,
        content: stripHtmlLinks(content),
      });
    }
    throw error;
  }
}

async function createShapeWithWholeItemLink(params, url) {
  if (!url) return createShapeSafe(params);

  const variants = [
    { link: url },
    { link: { url } },
    { links: [{ url }] },
    { url },
  ];

  for (const variant of variants) {
    try {
      return await createShapeSafe({
        ...params,
        ...variant,
      });
    } catch (error) {
      console.warn("[Reference Board Builder] create shape with whole-item link failed", {
        variant,
        error,
      });
    }
  }

  return createShapeSafe(params);
}

async function createImageWithRetry(params, maxRetries = 2) {
  let attempt = 0;
  let lastError = null;
  while (attempt <= maxRetries) {
    try {
      return await board.createImage(params);
    } catch (error) {
      lastError = error;
      attempt += 1;
      if (attempt > maxRetries) break;
      await new Promise((resolve) => setTimeout(resolve, 300 * attempt));
    }
  }
  throw lastError;
}

async function runWithConcurrency(items, limit, worker) {
  const safeLimit = Math.max(1, Number(limit) || 1);
  let cursor = 0;
  const runners = new Array(Math.min(safeLimit, items.length || 1)).fill(0).map(async () => {
    while (true) {
      const index = cursor;
      cursor += 1;
      if (index >= items.length) break;
      await worker(items[index], index);
    }
  });
  await Promise.all(runners);
}

async function getScaledImageDataUrl(file, width, height) {
  let perFile = scaledImageCache.get(file);
  if (!perFile) {
    perFile = new Map();
    scaledImageCache.set(file, perFile);
  }

  const key = `${Math.round(width)}x${Math.round(height)}`;
  if (perFile.has(key)) return perFile.get(key);

  const promise = (async () => {
    const image = await decodeImageFromFile(file);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(width));
    canvas.height = Math.max(1, Math.round(height));
    const ctx = canvas.getContext("2d");
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
    const mimeType = String(file.type || "").toLowerCase().includes("png") ? "image/png" : "image/jpeg";
    const dataUrl = mimeType === "image/png" ? canvas.toDataURL(mimeType) : canvas.toDataURL(mimeType, 0.9);
    try {
      image.src = "";
    } catch (_) {}
    return dataUrl;
  })();

  perFile.set(key, promise);
  return promise;
}


async function renderScene(scene, mode) {
  const config = scene.config;
  const created = [];
  let createdImages = 0;
  let createdFrames = 0;
  let createdShapes = 0;
  let createdTextBoxes = 0;
  const navigation = [];

  const navigationByCategory = new Map();
  const linkTargets = {
    categories: new Map(),
    subtypes: new Map(),
  };

  for (const category of scene.categories) {
    const categoryColor = category.color || FALLBACK_CATEGORY_COLOR;
    const navCategory = {
      name: category.name,
      color: categoryColor,
      headerWidget: null,
      subtypeWidgets: [],
      outlineWidgets: [],
    };
    navigation.push(navCategory);
    navigationByCategory.set(category.name, navCategory);

    const headerWidget = await createShapeSafe({
      shape: "round_rectangle",
      x: category.header.x,
      y: category.header.y,
      width: category.header.width,
      height: category.header.height,
      content: makeHeaderContent(category.header.title),
      style: {
        fillColor: categoryColor,
        borderColor: OUTLINE_COLOR,
        borderWidth: COLUMN_HEADER_BORDER_WIDTH,
        color: "#111111",
        fontSize: COLUMN_HEADER_FONT_SIZE,
        textAlign: "center",
        textAlignVertical: "middle",
      },
    });
    created.push(headerWidget);
    createdShapes += 1;
    navCategory.headerWidget = headerWidget;
    linkTargets.categories.set(category.name, headerWidget);

    for (const frame of category.frames) {
      const frameWidget = await createFrameSafe({
        x: frame.x,
        y: frame.y,
        width: frame.width,
        height: frame.height,
        title: "",
        style: {
          fillColor: FRAME_FILL,
        },
      });
      created.push(frameWidget);
      createdFrames += 1;

      const subtypeWidget = await createShapeSafe({
        shape: "round_rectangle",
        x: frame.subtitleShape.x,
        y: frame.subtitleShape.y,
        width: frame.subtitleShape.width,
        height: frame.subtitleShape.height,
        content: makeHeaderContent(frame.subtitleShape.title),
        style: {
          fillColor: categoryColor,
          borderColor: OUTLINE_COLOR,
          borderWidth: 3,
          color: "#111111",
          fontSize: SUBTYPE_HEADER_FONT_SIZE,
          textAlign: "center",
          textAlignVertical: "middle",
        },
      });
      created.push(subtypeWidget);
      createdShapes += 1;
      navCategory.subtypeWidgets.push({
        name: frame.name,
        widget: subtypeWidget,
      });
      linkTargets.subtypes.set(`${category.name}:::${frame.name}`, subtypeWidget);

      const frameContentJobs = [];

      for (const group of frame.groups) {
        if (group.text) {
          const textWidget = await createShapeSafe({
            shape: "rectangle",
            x: group.text.x,
            y: group.text.y,
            width: group.text.width,
            height: group.text.height,
            content: makeTextBoxContent(group.text.title),
            style: {
              fillColor: TEXT_BOX_FILL,
              borderColor: TEXT_BOX_BORDER,
              borderWidth: 1,
              color: "#111111",
              fontSize: config.textFontSize,
              textAlign: "center",
              textAlignVertical: "middle",
            },
          });
          created.push(textWidget);
          createdTextBoxes += 1;
        }

        for (const image of group.images) {
          frameContentJobs.push(async () => {
            if (mode === "preview") {
              const placeholder = await createShapeSafe({
                shape: "rectangle",
                x: image.x,
                y: image.y,
                width: image.width,
                height: image.height,
                content: "<p>&nbsp;</p>",
                style: {
                  fillColor: PREVIEW_FILL,
                  borderColor: PREVIEW_BORDER,
                  borderWidth: 1,
                  color: "#111111",
                },
              });
              created.push(placeholder);
              createdShapes += 1;
              return;
            }

            const dataUrl = await getScaledImageDataUrl(image.file, image.width, image.height);
            const imageWidget = await createImageWithRetry({
              url: dataUrl,
              x: image.x,
              y: image.y,
              width: image.width,
              height: image.height,
              title: image.title,
            });
            try {
              await imageWidget.setMetadata(APP_META_ID, {
                sourceFileName: image.file.name || image.title,
                app: APP_VERSION,
              });
            } catch (_) {}
            created.push(imageWidget);
            createdImages += 1;
          });
        }
      }

      await runWithConcurrency(frameContentJobs, IMAGE_CREATE_CONCURRENCY, async (job) => {
        await job();
      });
    }
  }

  for (const section of scene.outline || []) {
    const categoryColor = section.color || FALLBACK_CATEGORY_COLOR;
    const navCategory = navigationByCategory.get(section.headerTargetTitle) || null;

    const categoryOutlineWidget = await createShapeSafe({
      shape: "round_rectangle",
      x: section.headerShape.x,
      y: section.headerShape.y,
      width: section.headerShape.width,
      height: section.headerShape.height,
      content: makeHeaderContent(section.headerShape.title),
      style: {
        fillColor: categoryColor,
        borderColor: OUTLINE_COLOR,
        borderWidth: 3,
        color: "#111111",
        fontSize: 300,
        textAlign: "center",
        textAlignVertical: "middle",
      },
    });
    created.push(categoryOutlineWidget);
    createdShapes += 1;
    if (navCategory) navCategory.outlineWidgets.push(categoryOutlineWidget);

    try {
      await categoryOutlineWidget.setMetadata(APP_META_ID, {
        role: "outline-category",
        targetType: "category-header",
        targetId: linkTargets.categories.get(section.headerTargetTitle)?.id || null,
        targetTitle: section.headerTargetTitle,
        app: APP_VERSION,
      });
    } catch (_) {}

    for (const subtypeShape of section.subtypeShapes) {
      const subtypeOutlineWidget = await createShapeSafe({
        shape: "round_rectangle",
        x: subtypeShape.x,
        y: subtypeShape.y,
        width: subtypeShape.width,
        height: subtypeShape.height,
        content: makeHeaderContent(subtypeShape.title),
        style: {
          fillColor: categoryColor,
          borderColor: OUTLINE_COLOR,
          borderWidth: 2,
          color: "#111111",
          fontSize: 130,
          textAlign: "center",
          textAlignVertical: "middle",
        },
      });
      created.push(subtypeOutlineWidget);
      createdShapes += 1;
      if (navCategory) navCategory.outlineWidgets.push(subtypeOutlineWidget);

      try {
        await subtypeOutlineWidget.setMetadata(APP_META_ID, {
          role: "outline-subtype",
          targetType: "subtype-header",
          targetId: linkTargets.subtypes.get(`${subtypeShape.targetCategoryTitle}:::${subtypeShape.targetSubtypeTitle}`)?.id || null,
          targetCategoryTitle: subtypeShape.targetCategoryTitle,
          targetSubtypeTitle: subtypeShape.targetSubtypeTitle,
          app: APP_VERSION,
        });
      } catch (_) {}
    }
  }

  state.navigation = navigation;
  renderNavigationPanel();

  return {
    widgets: created,
    stats: {
      images: createdImages,
      frames: createdFrames,
      shapes: createdShapes,
      textBoxes: createdTextBoxes,
    },
  };
}


function getNavigationCategory(index) {
  const numericIndex = Number.parseInt(String(index), 10);
  if (!Number.isFinite(numericIndex)) return null;
  return state.navigation[numericIndex] || null;
}

function getNavigationSubtype(categoryIndex, subtypeIndex) {
  const category = getNavigationCategory(categoryIndex);
  if (!category) return null;
  const numericSubtypeIndex = Number.parseInt(String(subtypeIndex), 10);
  if (!Number.isFinite(numericSubtypeIndex)) return null;
  return category.subtypeWidgets[numericSubtypeIndex] || null;
}

function renderNavigationPanel() {
  const box = document.getElementById("navigationResult");
  if (!box) return;

  if (!state.navigation.length) {
    box.innerHTML = '<div class="muted">Навигация появится после создания доски.</div>';
    return;
  }

  box.innerHTML = state.navigation.map((category, categoryIndex) => `
    <details class="nav-section" ${categoryIndex === 0 ? "open" : ""}>
      <summary>${escapeHtml(category.name)}</summary>
      <div class="nav-section-body">
        <label class="nav-color-row">
          <span>Цвет</span>
          <input type="color" class="nav-color-input" data-nav-action="color" data-category-index="${categoryIndex}" value="${escapeHtml(category.color || FALLBACK_CATEGORY_COLOR)}" />
        </label>
        <button type="button" class="nav-button" data-nav-action="zoom-category" data-category-index="${categoryIndex}">
          Главный shape
        </button>
        <div class="nav-subtype-list">
          ${category.subtypeWidgets.map((subtype, subtypeIndex) => `
            <button type="button" class="nav-button nav-button-light" data-nav-action="zoom-subtype" data-category-index="${categoryIndex}" data-subtype-index="${subtypeIndex}">
              ${escapeHtml(subtype.name)}
            </button>
          `).join("")}
        </div>
      </div>
    </details>
  `).join("");
}

async function zoomToNavigationWidget(widget) {
  if (!widget) {
    await notifyWarning("Элемент не найден");
    return;
  }

  try {
    await board.viewport.zoomTo([widget]);
  } catch (error) {
    console.warn("[Reference Board Builder] zoomTo failed", error);
    await notifyError("Не удалось приблизить элемент", error);
  }
}

async function setWidgetFillColor(widget, color) {
  if (!widget || !color) return false;

  try {
    if (typeof widget.setStyle === "function") {
      await widget.setStyle({ fillColor: color });
      return true;
    }
  } catch (error) {
    console.warn("[Reference Board Builder] setStyle fillColor failed", error);
  }

  try {
    if (!widget.style || typeof widget.style !== "object") widget.style = {};
    widget.style.fillColor = color;
    if (typeof widget.sync === "function") {
      await widget.sync();
      return true;
    }
  } catch (error) {
    console.warn("[Reference Board Builder] style fillColor sync failed", error);
  }

  return false;
}

async function applyNavigationCategoryColor(categoryIndex, color) {
  const category = getNavigationCategory(categoryIndex);
  if (!category) return;

  category.color = color;
  const widgets = [
    category.headerWidget,
    ...category.subtypeWidgets.map((subtype) => subtype.widget),
    ...category.outlineWidgets,
  ].filter(Boolean);

  await Promise.all(widgets.map((widget) => setWidgetFillColor(widget, color)));
}

function initNavigationPanel() {
  renderNavigationPanel();

  const box = document.getElementById("navigationResult");
  if (!box) return;

  box.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-nav-action]");
    if (!button || button.tagName === "INPUT") return;

    const action = button.dataset.navAction;
    if (action === "zoom-category") {
      const category = getNavigationCategory(button.dataset.categoryIndex);
      await zoomToNavigationWidget(category?.headerWidget);
      return;
    }

    if (action === "zoom-subtype") {
      const subtype = getNavigationSubtype(button.dataset.categoryIndex, button.dataset.subtypeIndex);
      await zoomToNavigationWidget(subtype?.widget);
    }
  });

  box.addEventListener("input", async (event) => {
    const input = event.target.closest("input[data-nav-action='color']");
    if (!input) return;
    await applyNavigationCategoryColor(input.dataset.categoryIndex, input.value);
  });
}

function setBuilderButtonsDisabled(isDisabled) {
  const buttonIds = ["folderButton", "buildButton"];
  buttonIds.forEach((id) => {
    const element = document.getElementById(id);
    if (element) element.disabled = isDisabled;
  });
}

async function handleCreate(mode) {
  if (state.isRendering) return;
  state.isRendering = true;
  setBuilderButtonsDisabled(true);

  try {
    setStatus("Создание доски…", "info");
    const layout = state.layout || await ensureLayout();
    if (!layout) return;

    const viewport = await board.viewport.get();
    const scene = buildScene(layout, viewport);
    const result = await renderScene(scene, mode);

    if (result.widgets.length) {
      try {
        await board.viewport.zoomTo(result.widgets);
      } catch (error) {
        console.warn("zoomTo failed", error);
      }
    }

    if (mode === "preview") {
      setStatus(
        `Превью создано.\nФреймов: ${result.stats.frames}\nФигур: ${result.stats.shapes}\nТекстовых блоков: ${result.stats.textBoxes}`,
        "info"
      );
      await notifyInfo("Превью создано");
    } else {
      setStatus(
        `Доска создана.\nФреймов: ${result.stats.frames}\nКартинок: ${result.stats.images}\nФигур: ${result.stats.shapes}`,
        "info"
      );
      await notifyInfo("Доска создана");
    }
  } catch (error) {
    console.error(error);
    setStatus(mode === "preview" ? "Не удалось создать превью. Подробности в консоли." : "Не удалось создать доску. Подробности в консоли.", "error");
    await notifyError(mode === "preview" ? "Ошибка создания превью" : "Ошибка создания доски", error);
  } finally {
    state.isRendering = false;
    setBuilderButtonsDisabled(false);
  }
}

function initTabs() {
  const buttons = Array.from(document.querySelectorAll(".tab-btn"));
  const contents = {
    builder: document.getElementById("tab-builder"),
    stats: document.getElementById("tab-stats"),
  };

  function activate(name) {
    buttons.forEach((button) => button.classList.toggle("active", button.dataset.tab === name));
    Object.entries(contents).forEach(([key, element]) => {
      if (!element) return;
      element.classList.toggle("active", key === name);
    });
  }

  buttons.forEach((button) => {
    button.addEventListener("click", () => activate(button.dataset.tab));
  });
}


window.addEventListener("DOMContentLoaded", () => {
  resetResults();
  setStatus("");
  initTabs();
  initNavigationPanel();

  const folderButton = document.getElementById("folderButton");
  const folderInput = document.getElementById("folderInput");
  const buildButton = document.getElementById("buildButton");
  const configInputs = Array.from(document.querySelectorAll("input[type='number'], input[type='checkbox']"));

  folderButton?.addEventListener("click", () => folderInput?.click());
  folderInput?.addEventListener("change", async (event) => {
    await handleFolderSelected(event.target.files);
  });

  configInputs.forEach((element) => {
    const eventName = element.type === "checkbox" ? "change" : "input";
    element.addEventListener(eventName, scheduleStatsRefresh);
  });

  buildButton?.addEventListener("click", async () => {
    await handleCreate("build");
  });
});
