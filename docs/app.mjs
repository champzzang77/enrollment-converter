import {
  PROFILES,
  PROFILE_FAMILIES,
  PROFILE_FAMILY_BY_PROFILE_ID,
  PROFILE_FAMILY_DEFAULT_PROFILE_ID,
  STRUCTURE_PATTERNS,
} from "./data.mjs";
import { runProfile } from "./engine.mjs";

const STORAGE_KEY = "enrollment-upload-static-usage-log-v2";
const LEGACY_STORAGE_KEYS = ["enrollment-upload-static-usage-log-v1"];
const ALLOWED_EXTENSIONS = new Set([".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"]);
const MANUAL_COURSE_FILE_EXTENSIONS = new Set([
  ".txt",
  ".csv",
  ".tsv",
  ".xlsx",
  ".xlsm",
  ".xltx",
  ".xltm",
  ".xls",
]);

const familySelect = document.getElementById("familySelect");
const fileInput = document.getElementById("fileInput");
const convertButton = document.getElementById("convertButton");
const downloadLink = document.getElementById("downloadLink");
const statusBox = document.getElementById("statusBox");
const errorBox = document.getElementById("errorBox");
const totalRows = document.getElementById("totalRows");
const emailUndefined = document.getElementById("emailUndefined");
const mobileUndefined = document.getElementById("mobileUndefined");
const guideTitle = document.getElementById("guideTitle");
const guideDescription = document.getElementById("guideDescription");
const guideUseWhen = document.getElementById("guideUseWhen");
const guideExample = document.getElementById("guideExample");
const guideVariant = document.getElementById("guideVariant");
const guideHints = document.getElementById("guideHints");
const recommendBadge = document.getElementById("recommendBadge");
const manualCourseField = document.getElementById("manualCourseField");
const manualCourseLabel = document.getElementById("manualCourseLabel");
const manualCourseInput = document.getElementById("manualCourseInput");
const manualCourseNote = document.getElementById("manualCourseNote");
const manualCoursePathInput = document.getElementById("manualCoursePathInput");
const manualCourseFileInput = document.getElementById("manualCourseFileInput");
const manualCourseSourceNote = document.getElementById("manualCourseSourceNote");
const recentUsageBox = document.getElementById("recentUsageBox");
const totalUsageCount = document.getElementById("totalUsageCount");
const todayUsageCount = document.getElementById("todayUsageCount");

let latestDownloadUrl = "";
let recommendedProfileId = "";
let recommendedProfileReason = "";
let selectedProfileId = "";
let latestWorkbookData = [];
let latestInputFileName = "";

function clearLegacyUsageData() {
  LEGACY_STORAGE_KEYS.forEach((key) => {
    try {
      localStorage.removeItem(key);
    } catch (_error) {
      // Ignore storage cleanup issues.
    }
  });
}

function getProfileById(profileId) {
  return PROFILES.find((item) => item.id === profileId) || null;
}

function getProfileFamilyId(profileOrId) {
  const profileId = typeof profileOrId === "string" ? profileOrId : profileOrId?.id;
  return PROFILE_FAMILY_BY_PROFILE_ID[profileId] || "";
}

function getSelectedFamily() {
  return PROFILE_FAMILIES.find((item) => item.id === familySelect.value) || null;
}

function getFamilyProfiles(familyId) {
  return PROFILES.filter((profile) => getProfileFamilyId(profile) === familyId);
}

function sortProfiles(profiles) {
  return [...profiles].sort((left, right) => {
    const orderDiff = getProfileDisplayOrder(left) - getProfileDisplayOrder(right);
    if (orderDiff !== 0) {
      return orderDiff;
    }
    return String(left.label || "").localeCompare(String(right.label || ""), "ko");
  });
}

function getDefaultProfileForFamily(familyId) {
  const defaultProfileId = PROFILE_FAMILY_DEFAULT_PROFILE_ID[familyId];
  const defaultProfile = getProfileById(defaultProfileId);
  if (defaultProfile) {
    return defaultProfile;
  }
  return sortProfiles(getFamilyProfiles(familyId))[0] || null;
}

function getSelectedProfile() {
  const family = getSelectedFamily();
  const selectedProfile = getProfileById(selectedProfileId);
  if (selectedProfile && family && getProfileFamilyId(selectedProfile) === family.id) {
    return selectedProfile;
  }
  return family ? getDefaultProfileForFamily(family.id) : null;
}

function getProfileDisplayOrder(profile) {
  const match = String(profile?.label || "").match(/^(\d+)\./);
  return match ? Number(match[1]) : Number.POSITIVE_INFINITY;
}

function renderFamilies() {
  PROFILE_FAMILIES.forEach((profileFamily) => {
    const option = document.createElement("option");
    option.value = profileFamily.id;
    option.textContent = profileFamily.label;
    familySelect.appendChild(option);
  });
  if (!familySelect.value && PROFILE_FAMILIES[0]) {
    familySelect.value = PROFILE_FAMILIES[0].id;
  }
  selectedProfileId = getDefaultProfileForFamily(familySelect.value)?.id || "";
  updateProfileGuide();
}

function updateProfileGuide() {
  const family = getSelectedFamily();
  const profile = getSelectedProfile();
  if (!family) {
    guideTitle.textContent = "상위 유형을 선택해 주세요";
    guideDescription.textContent = "선택한 상위 유형 설명이 여기에 표시됩니다.";
    guideUseWhen.textContent = "-";
    guideExample.textContent = "-";
    guideVariant.textContent = "-";
    guideHints.innerHTML = "";
    recommendBadge.textContent = "직접 선택";
    recommendBadge.className = "badge badge-neutral";
    manualCourseField.hidden = true;
    return;
  }

  guideTitle.textContent = family.label;
  guideDescription.textContent = family.description;
  guideUseWhen.textContent = family.use_when;
  guideExample.textContent = family.example_file;
  guideVariant.textContent = profile
    ? `${profile.label} - ${profile.short_description}`
    : "파일을 올리면 내부 세부 유형을 자동으로 고릅니다.";
  guideHints.innerHTML = "";
  family.hints.forEach((hint) => {
    const item = document.createElement("li");
    item.textContent = hint;
    guideHints.appendChild(item);
  });

  const recommendedFamilyId = getProfileFamilyId(recommendedProfileId);
  if (recommendedProfileId && family.id === recommendedFamilyId) {
    recommendBadge.textContent =
      profile && recommendedProfileId === profile.id
        ? (recommendedProfileReason === "structure" ? "파일 구조 기준 추천" : "파일명 기준 추천")
        : "추천된 상위 유형";
    recommendBadge.className = "badge badge-recommend";
  } else {
    recommendBadge.textContent = "직접 선택";
    recommendBadge.className = "badge badge-neutral";
  }

  const manualCourseConfig = profile?.manual_course_input;
  if (manualCourseConfig) {
    manualCourseField.hidden = false;
    manualCourseLabel.textContent = manualCourseConfig.label || "선택 입력. 과정코드 입력 방식";
    manualCourseNote.textContent =
      manualCourseConfig.note ||
      "과정코드를 따로 받은 경우에만 `직접 입력`, `첨부파일 경로 지정`, `첨부파일 선택` 중 하나를 사용해 주세요.";
    manualCourseInput.placeholder = manualCourseConfig.placeholder || "";
    manualCoursePathInput.placeholder =
      manualCourseConfig.path_placeholder || "/Users/parkchamin/Downloads/과정코드 첨부파일.xlsx";
    manualCourseSourceNote.textContent =
      manualCourseConfig.source_note ||
      "이 영역은 `과정코드 직접 입력`, `첨부파일 경로 지정`, `첨부파일 선택` 중 하나로 사용할 수 있습니다. 경로 읽기는 로컬 오프라인 앱이나 파일 접근이 허용된 브라우저에서만 동작할 수 있고, 막히면 바로 아래 파일 선택을 사용해 주세요. 화면에 직접 적은 내용이 있으면 그 값을 우선 적용합니다.";
  } else {
    manualCourseField.hidden = true;
  }
}

function normalizeFileNameMatch(value) {
  return String(value || "")
    .normalize("NFC")
    .toLowerCase()
    .replace(/\s+/g, "");
}

function normalizeHeaderText(value) {
  return String(value || "")
    .normalize("NFC")
    .toLowerCase()
    .replace(/\n/g, "")
    .replace(/\s+/g, "")
    .replace(/\xa0/g, "")
    .replace(/\*/g, "");
}

function pickProfileByFileName(fileName, candidates = PROFILES) {
  const normalizedName = normalizeFileNameMatch(fileName);
  let matched = null;
  let matchedScore = -1;

  candidates.forEach((profile) => {
    (profile.filename_keywords || []).forEach((keyword) => {
      const normalizedKeyword = normalizeFileNameMatch(keyword);
      if (!normalizedKeyword || !normalizedName.includes(normalizedKeyword)) {
        return;
      }

      const score = normalizedKeyword.length;
      if (score > matchedScore) {
        matched = profile;
        matchedScore = score;
      }
    });
  });

  return matched;
}

function getCell(rows, rowIndex, columnIndex) {
  const row = rows[rowIndex] || [];
  return columnIndex < row.length ? row[columnIndex] : null;
}

function matchesStructureHeaderChecks(rows, checks = []) {
  return checks.every((check) =>
    normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
  );
}

function matchesStructureRowMetric(row, metric, metrics) {
  const anyFilledColumns = Array.isArray(metric.any_filled_columns) ? metric.any_filled_columns : [];
  const allFilledColumns = Array.isArray(metric.all_filled_columns) ? metric.all_filled_columns : [];
  const emptyColumns = Array.isArray(metric.empty_columns) ? metric.empty_columns : [];

  if (metric.requires_metric_positive && Number(metrics[metric.requires_metric_positive] || 0) <= 0) {
    return false;
  }

  if (anyFilledColumns.length && !anyFilledColumns.some((columnIndex) => hasFilledValue(row[columnIndex]))) {
    return false;
  }

  if (allFilledColumns.length && !allFilledColumns.every((columnIndex) => hasFilledValue(row[columnIndex]))) {
    return false;
  }

  if (emptyColumns.length && !emptyColumns.every((columnIndex) => !hasFilledValue(row[columnIndex]))) {
    return false;
  }

  return true;
}

function collectStructureRowMetrics(rows, rowScan = {}) {
  const metricConfigs = Array.isArray(rowScan.metrics) ? rowScan.metrics : [];
  const metrics = Object.fromEntries(metricConfigs.map((metric) => [metric.id, 0]));
  const startRow = Number(rowScan.start_row || 0);

  for (let rowIndex = startRow; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    metricConfigs.forEach((metric) => {
      if (matchesStructureRowMetric(row, metric, metrics)) {
        metrics[metric.id] = Number(metrics[metric.id] || 0) + 1;
      }
    });
  }

  return metrics;
}

function compareStructureValues(leftValue, operator, rightValue) {
  switch (operator) {
    case ">=":
      return leftValue >= rightValue;
    case ">":
      return leftValue > rightValue;
    case "<=":
      return leftValue <= rightValue;
    case "<":
      return leftValue < rightValue;
    case "===":
    case "==":
      return leftValue === rightValue;
    default:
      return false;
  }
}

function matchesStructureConditions(metrics, conditions = []) {
  return conditions.every((condition) => {
    const leftValue = Number(metrics[condition.left] || 0);
    const rightValue =
      typeof condition.right_metric === "string"
        ? Number(metrics[condition.right_metric] || 0)
        : Number(condition.right || 0);
    return compareStructureValues(leftValue, condition.operator, rightValue);
  });
}

function evaluateSheetStructurePattern(rows, pattern) {
  if (!pattern) {
    return false;
  }

  if (pattern.mode === "header_checks") {
    return matchesStructureHeaderChecks(rows, pattern.checks || []);
  }

  if (pattern.mode === "header_checks_with_row_scan") {
    if (!matchesStructureHeaderChecks(rows, pattern.checks || [])) {
      return false;
    }
    const rowMetrics = collectStructureRowMetrics(rows, pattern.row_scan || {});
    return matchesStructureConditions(rowMetrics, pattern.row_scan?.conditions || []);
  }

  return false;
}

function evaluateWorkbookStructurePattern(workbookData, pattern) {
  if (!pattern || !Array.isArray(workbookData) || workbookData.length === 0) {
    return false;
  }

  if (pattern.mode === "required_sheet_names") {
    const sheetNames = new Set(workbookData.map((sheet) => sheet.name));
    return (pattern.required_sheet_names || []).every((sheetName) => sheetNames.has(sheetName));
  }

  if (pattern.mode === "named_sheet_pattern") {
    const targetSheetNames = new Set(pattern.sheet_names || []);
    const matchedSheets = workbookData.filter((sheet) => targetSheetNames.has(sheet.name));
    if (!matchedSheets.length) {
      return false;
    }

    const matchedCount = matchedSheets.filter((sheet) =>
      evaluateSheetStructurePattern(sheet.rows || [], pattern.sheet_pattern || null)
    ).length;

    if (pattern.match_mode === "all") {
      return matchedCount === matchedSheets.length;
    }

    return matchedCount >= Number(pattern.min_matches || 1);
  }

  return false;
}

function matchesStructurePattern(patternId, input) {
  const pattern = STRUCTURE_PATTERNS[patternId];
  if (!pattern) {
    return false;
  }

  if (pattern.scope === "sheet") {
    return evaluateSheetStructurePattern(Array.isArray(input) ? input : [], pattern);
  }

  if (pattern.scope === "workbook") {
    return evaluateWorkbookStructurePattern(Array.isArray(input) ? input : [], pattern);
  }

  return false;
}

function hasManualFixedCourseData(rows) {
  for (let rowIndex = 18; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const courseName = String(row[2] || "").trim();
    const userId = String(row[3] || "").trim();
    const name = String(row[5] || "").trim();
    const mobile = String(row[8] || "").trim();
    if (courseName || userId || name || mobile) {
      return true;
    }
  }
  return false;
}

function getStructureRecommendationThreshold(config, workbookData) {
  const requestedMatches = Number(config.min_matches || 1);
  if (config.min_match_mode === "up_to_two") {
    return Math.min(requestedMatches, workbookData.length);
  }
  if (config.min_match_mode === "all") {
    return workbookData.length;
  }
  return requestedMatches;
}

function buildStructureRecommendation(profile, config, reason = "structure") {
  return {
    profile,
    reason,
    message: `<strong>${profile.label}</strong> 형식으로 추천했습니다.<br>${
      config.message_detail || "파일 구조가 확인되었습니다."
    }`,
  };
}

function evaluateStructureRecommendation(profile, workbookData) {
  const config = profile.structure_recommendation;
  if (!config || !Array.isArray(workbookData) || workbookData.length === 0) {
    return null;
  }

  if (config.single_sheet_only && workbookData.length !== 1) {
    return null;
  }

  if (config.multi_sheet_only && workbookData.length < 2) {
    return null;
  }

  if (config.mode === "header_alias_sheet") {
    const sourceConfig = profile.source || {};
    const matchedSheets = workbookData.filter((sheet) => {
      try {
        const summary = summarizeHeaderAliasData(sheet.rows || [], sourceConfig);
        if (summary.dataRowCount <= 0) {
          return false;
        }
        if (config.require_course_code && summary.courseCodeValueCount <= 0) {
          return false;
        }
        if (config.require_missing_course_code && summary.courseCodeValueCount > 0) {
          return false;
        }
        return true;
      } catch (_error) {
        return false;
      }
    });

    const threshold = getStructureRecommendationThreshold(config, workbookData);
    if (matchedSheets.length >= threshold) {
      return buildStructureRecommendation(profile, config);
    }
    return null;
  }

  const pattern = STRUCTURE_PATTERNS[config.pattern_id];
  if (!pattern) {
    return null;
  }

  if (pattern.scope === "workbook") {
    if (matchesStructurePattern(config.pattern_id, workbookData)) {
      return buildStructureRecommendation(profile, config);
    }
    return null;
  }

  if (pattern.scope === "sheet") {
    const matchedSheets = workbookData.filter((sheet) =>
      matchesStructurePattern(config.pattern_id, sheet.rows || [])
    );
    const threshold = getStructureRecommendationThreshold(config, workbookData);
    if (matchedSheets.length >= threshold) {
      return buildStructureRecommendation(profile, config);
    }
  }

  return null;
}

function recommendProfileByWorkbookStructure(workbookData) {
  if (!Array.isArray(workbookData) || workbookData.length === 0) {
    return null;
  }

  const recommendationProfiles = [...PROFILES]
    .filter((profile) => profile.structure_recommendation)
    .sort((left, right) => {
      const priorityDiff =
        Number(right.structure_recommendation?.priority || 0) -
        Number(left.structure_recommendation?.priority || 0);
      if (priorityDiff !== 0) {
        return priorityDiff;
      }
      return getProfileDisplayOrder(left) - getProfileDisplayOrder(right);
    });

  for (const profile of recommendationProfiles) {
    const recommendation = evaluateStructureRecommendation(profile, workbookData);
    if (recommendation) {
      return recommendation;
    }
  }

  return null;
}

function formatRecommendationMessage(profile, message) {
  const family = PROFILE_FAMILIES.find((item) => item.id === getProfileFamilyId(profile));
  const intro = `<strong>${family?.label || profile.label}</strong> 유형으로 추천했습니다.<br>내부 세부 유형은 <strong>${profile.label}</strong>로 맞췄습니다.`;
  const detail = String(message || "").replace(/^<strong>.*?<\/strong>\s*(?:형식|유형)으로 추천했습니다\.<br>/, "");
  return detail ? `${intro}<br>${detail}` : intro;
}

function applyRecommendation(profile, reason, message) {
  recommendedProfileId = profile.id;
  recommendedProfileReason = reason;
  selectedProfileId = profile.id;
  familySelect.value = getProfileFamilyId(profile);
  updateProfileGuide();
  statusBox.innerHTML = formatRecommendationMessage(profile, message);
}

function resolveProfileForFamily(familyId, workbookData = [], fileName = "") {
  if (!familyId) {
    return null;
  }

  const familyProfiles = sortProfiles(getFamilyProfiles(familyId));
  if (!familyProfiles.length) {
    return null;
  }

  const structureRecommendation = recommendProfileByWorkbookStructure(workbookData);
  if (
    structureRecommendation?.profile &&
    getProfileFamilyId(structureRecommendation.profile) === familyId
  ) {
    return structureRecommendation.profile;
  }

  const matchedByFileName = pickProfileByFileName(fileName, familyProfiles);
  if (matchedByFileName) {
    return matchedByFileName;
  }

  return getDefaultProfileForFamily(familyId);
}

function recommendProfileByFileName(fileName) {
  const matched = pickProfileByFileName(fileName, PROFILES);

  if (!matched) {
    const fallback = getProfileById("generic_auto_header");
    if (fallback) {
      applyRecommendation(
        fallback,
        "filename",
        `<strong>${PROFILE_FAMILIES.find((item) => item.id === getProfileFamilyId(fallback))?.label || fallback.label}</strong> 유형으로 추천했습니다.<br>파일명만으로 정확한 세부 형식을 찾지 못해 일반 명단 파일 묶음을 먼저 선택했습니다.`
      );
      return;
    }

    recommendedProfileId = "";
    recommendedProfileReason = "";
    updateProfileGuide();
    statusBox.innerHTML = "파일을 선택했습니다. <strong>상위 유형</strong>을 확인한 뒤 <strong>변환 시작</strong>을 눌러 주세요.";
    return;
  }

  if (matched.manual_course_input?.required) {
    applyRecommendation(
      matched,
      "filename",
      `<strong>${PROFILE_FAMILIES.find((item) => item.id === getProfileFamilyId(matched))?.label || matched.label}</strong> 유형으로 추천했습니다.<br>예시 파일명을 확인한 뒤, 필요한 경우 아래 입력칸에 과정코드도 함께 넣어 주세요.`
    );
    return;
  }

  applyRecommendation(
    matched,
    "filename",
    `<strong>${PROFILE_FAMILIES.find((item) => item.id === getProfileFamilyId(matched))?.label || matched.label}</strong> 유형으로 추천했습니다.<br>예시 파일명과 설명이 맞는지만 한 번 확인해 주세요.`
  );
}

function setSummary(summary) {
  totalRows.textContent = summary?.total_rows ?? "-";
  emailUndefined.textContent = summary?.email_undefined ?? "-";
  mobileUndefined.textContent = summary?.mobile_undefined ?? "-";
}

function clearDownload() {
  if (latestDownloadUrl) {
    URL.revokeObjectURL(latestDownloadUrl);
  }
  latestDownloadUrl = "";
  downloadLink.href = "#";
  downloadLink.download = "";
  downloadLink.classList.add("disabled");
}

function normalizeSheetNameKey(value) {
  return normalizeFileNameMatch(value).replace(/[()_\-]/g, "");
}

function normalizeManualCoursePathToUrl(rawPath) {
  const value = String(rawPath || "").trim();
  if (!value) {
    return "";
  }

  if (value.startsWith("file://")) {
    return value;
  }

  if (value.startsWith("/")) {
    return encodeURI(`file://${value}`);
  }

  if (value.startsWith("./") || value.startsWith("../")) {
    return new URL(value, window.location.href).href;
  }

  throw new Error("과정코드 첨부파일 경로는 `/Users/...` 또는 `file:///...` 형태로 넣어 주세요.");
}

function hasFilledValue(value) {
  return String(value ?? "").trim() !== "";
}

function ensureSupportedManualCourseFileName(fileName) {
  const extension = getFileExtension(fileName);
  if (!MANUAL_COURSE_FILE_EXTENSIONS.has(extension)) {
    throw new Error("지원하지 않는 과정코드 첨부파일 형식입니다.");
  }
  return extension;
}

function extractManualCourseLinesFromWorkbook(workbook) {
  const lines = [];

  workbook.SheetNames.forEach((sheetName) => {
    const rows = buildSheetRowsPreservingLayout(workbook.Sheets[sheetName]);
    rows.forEach((row) => {
      const values = (row || []).map((cell) => String(cell ?? "").trim()).filter(Boolean);
      if (!values.length) {
        return;
      }

      const codeCell = values.find((value) => /\bHL[A-Za-z0-9]+\b/i.test(value)) || "";
      const labelParts = values.filter((value) => value !== codeCell);

      if (codeCell && labelParts.length) {
        lines.push(labelParts.join(" "));
        lines.push(codeCell);
        return;
      }

      lines.push(...values);
    });
  });

  return lines.join("\n").trim();
}

async function readTextFromFile(file) {
  if (typeof file.text === "function") {
    return file.text();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("첨부파일을 읽는 중 오류가 발생했습니다."));
    reader.readAsText(file);
  });
}

async function readArrayBufferFromUrl(url) {
  const response = await fetch(url);
  return response.arrayBuffer();
}

async function readTextFromUrl(url) {
  const response = await fetch(url);
  return response.text();
}

async function readManualCourseAttachmentFromFile(file) {
  if (!file) {
    return "";
  }

  const extension = ensureSupportedManualCourseFileName(file.name);
  if (extension === ".txt" || extension === ".csv" || extension === ".tsv") {
    return String(await readTextFromFile(file)).trim();
  }

  const data = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(data, {
    type: "array",
    cellDates: false,
    cellText: false,
  });
  return extractManualCourseLinesFromWorkbook(workbook);
}

async function readManualCourseAttachmentFromPath(rawPath) {
  const value = String(rawPath || "").trim();
  if (!value) {
    return "";
  }

  const extension = ensureSupportedManualCourseFileName(value);
  const url = normalizeManualCoursePathToUrl(value);

  try {
    if (extension === ".txt" || extension === ".csv" || extension === ".tsv") {
      return String(await readTextFromUrl(url)).trim();
    }

    const data = await readArrayBufferFromUrl(url);
    const workbook = XLSX.read(data, {
      type: "array",
      cellDates: false,
      cellText: false,
    });
    return extractManualCourseLinesFromWorkbook(workbook);
  } catch (error) {
    throw new Error(`과정코드 첨부파일 경로를 읽지 못했습니다. ${error?.message || error}`);
  }
}

async function buildManualCourseSourceText(rawText, rawPath, attachmentFile) {
  const parts = [];

  const pathText = await readManualCourseAttachmentFromPath(rawPath);
  if (pathText) {
    parts.push(pathText);
  }

  const fileText = await readManualCourseAttachmentFromFile(attachmentFile);
  if (fileText) {
    parts.push(fileText);
  }

  const typedText = String(rawText || "").trim();
  if (typedText) {
    parts.push(typedText);
  }

  return parts.join("\n").trim();
}

function parseManualCourseAssignments(rawText, knownSheetNames = []) {
  const textValue = String(rawText || "").trim();
  const sheetCourseCodes = {};
  const orderedEntries = [];

  if (!textValue) {
    return { sheetCourseCodes, orderedEntries };
  }

  const sheetsByNormalizedName = new Map(
    knownSheetNames.map((sheetName) => [normalizeSheetNameKey(sheetName), sheetName])
  );

  const lines = textValue
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  let pendingName = "";

  lines.forEach((line) => {
    const normalizedLine = line.replace(/^[-•]/, "").trim();
    const codeOnlyMatch = normalizedLine.match(/^(?:ㄴ\s*)?과정코드\s*[:=]\s*([A-Za-z0-9]+)\s*$/i);
    if (codeOnlyMatch) {
      orderedEntries.push({
        name: pendingName || null,
        code: codeOnlyMatch[1].toUpperCase(),
      });
      pendingName = "";
      return;
    }

    const pairMatch = normalizedLine.match(/^(.*?)\s*(?:=|:|\|)\s*([A-Za-z0-9]+)\s*$/);
    if (pairMatch) {
      const left = pairMatch[1].trim();
      const code = pairMatch[2].toUpperCase();
      const matchedSheetName = sheetsByNormalizedName.get(normalizeSheetNameKey(left));
      if (matchedSheetName) {
        sheetCourseCodes[matchedSheetName] = code;
      } else {
        orderedEntries.push({ name: left, code });
      }
      pendingName = "";
      return;
    }

    const trailingParenCodeMatch = normalizedLine.match(/^(.*)\s*[\(\[（]\s*(HL[A-Za-z0-9]+)\s*[\)\]）]\s*$/i);
    if (trailingParenCodeMatch) {
      const left = trailingParenCodeMatch[1].trim();
      const code = trailingParenCodeMatch[2].toUpperCase();
      const matchedSheetName = sheetsByNormalizedName.get(normalizeSheetNameKey(left));
      if (matchedSheetName) {
        sheetCourseCodes[matchedSheetName] = code;
      } else {
        orderedEntries.push({ name: left || null, code });
      }
      pendingName = "";
      return;
    }

    const bareCodeMatch = normalizedLine.match(/^(HL[A-Za-z0-9]+)$/i);
    if (bareCodeMatch) {
      orderedEntries.push({
        name: pendingName || null,
        code: bareCodeMatch[1].toUpperCase(),
      });
      pendingName = "";
      return;
    }

    pendingName = normalizedLine;
  });

  return { sheetCourseCodes, orderedEntries };
}

function normalizeManualInputSheetSelection(sheetNames, sourceConfig = {}) {
  const skipSet = new Set(sourceConfig.skip_sheets || []);
  const selected = (sourceConfig.include_sheets || []).length
    ? sheetNames.filter((name) => sourceConfig.include_sheets.includes(name))
    : [...sheetNames];
  return selected.filter((name) => !skipSet.has(name));
}

function chooseProfileSheetNames(sheetNames, sourceConfig = {}) {
  const preferredSheets = sourceConfig.preferred_sheets || [];
  if (preferredSheets.length) {
    const matched = preferredSheets.filter((name) => sheetNames.includes(name));
    if (matched.length) {
      return matched;
    }
    if (sourceConfig.fallback_first_sheet && sheetNames.length) {
      return [sheetNames[0]];
    }
  }

  const selected = normalizeManualInputSheetSelection(sheetNames, sourceConfig);
  return sourceConfig.first_sheet_only ? selected.slice(0, 1) : selected;
}

function findManualInputHeaderRowByKeywords(rows, requiredKeywords) {
  const normalizedKeywords = requiredKeywords.map((keyword) => normalizeHeaderText(keyword));

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const joined = row.map((cell) => normalizeHeaderText(cell)).join("|");
    if (normalizedKeywords.every((keyword) => joined.includes(keyword))) {
      return { rowIndex, headerRow: row };
    }
  }

  throw new Error(`헤더 행을 찾지 못했습니다: ${requiredKeywords.join(", ")}`);
}

function findManualInputHeaderRowWithConfig(rows, sourceConfig = {}) {
  const keywordGroups = Array.isArray(sourceConfig.header_keyword_groups) && sourceConfig.header_keyword_groups.length
    ? sourceConfig.header_keyword_groups
    : [sourceConfig.header_keywords || []];

  let lastError = null;
  for (const keywords of keywordGroups) {
    try {
      return findManualInputHeaderRowByKeywords(rows, keywords);
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("헤더 행을 찾지 못했습니다.");
}

function getManualInputHeaderIndexMap(headerRow) {
  const entries = [];
  headerRow.forEach((cell, index) => {
    const key = normalizeHeaderText(cell);
    if (key) {
      entries.push([key, index]);
    }
  });
  return Object.fromEntries(entries);
}

function scoreManualInputAliasMatch(headerKey, aliasKey) {
  if (!headerKey || !aliasKey || !headerKey.includes(aliasKey)) {
    return Number.NEGATIVE_INFINITY;
  }

  if (
    aliasKey === "주소" &&
    (headerKey.includes("이메일") || headerKey.includes("메일"))
  ) {
    return Number.NEGATIVE_INFINITY;
  }

  let score = aliasKey.length * 10;
  if (headerKey === aliasKey) {
    score += 1000;
  } else {
    if (headerKey.startsWith(aliasKey)) {
      score += 300;
    }
    if (headerKey.endsWith(aliasKey)) {
      score += 200;
    }
  }

  score -= Math.max(0, headerKey.length - aliasKey.length);
  return score;
}

function pickManualInputValueByAliases(row, indexMap, aliases = []) {
  let bestValue = null;
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const alias of aliases) {
    const aliasKey = normalizeHeaderText(alias);
    for (const [headerKey, index] of Object.entries(indexMap)) {
      if (index >= row.length) {
        continue;
      }

      const value = row[index];
      const score = scoreManualInputAliasMatch(headerKey, aliasKey) + (hasFilledValue(value) ? 1 : 0);
      if (score > bestScore) {
        bestScore = score;
        bestValue = value;
      }
    }
  }

  return hasFilledValue(bestValue) ? bestValue : null;
}

function pickManualInputIndexByAliases(indexMap, aliases = []) {
  let bestIndex = null;
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const alias of aliases) {
    const aliasKey = normalizeHeaderText(alias);
    for (const [headerKey, index] of Object.entries(indexMap)) {
      const score = scoreManualInputAliasMatch(headerKey, aliasKey);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    }
  }

  return bestIndex;
}

function buildManualInputHeaderAliasRowData(rawRow, indexMap, fieldAliases = {}) {
  return Object.fromEntries(
    Object.entries(fieldAliases).map(([field, aliases]) => [
      field,
      pickManualInputValueByAliases(rawRow || [], indexMap, aliases),
    ])
  );
}

function hasRequiredHeaderAliasData(rows, sourceConfig = {}) {
  const headerInfo = findManualInputHeaderRowWithConfig(rows, sourceConfig);
  const indexMap = getManualInputHeaderIndexMap(headerInfo.headerRow);
  const requiredAny = sourceConfig.required_any || ["user_id", "name", "email", "mobile"];
  const fieldAliases = sourceConfig.field_aliases || {};

  return rows.slice(headerInfo.rowIndex + 1).some((rawRow) => {
    const extracted = buildManualInputHeaderAliasRowData(rawRow || [], indexMap, fieldAliases);
    return requiredAny.some((field) => hasFilledValue(extracted[field]));
  });
}

function summarizeHeaderAliasData(rows, sourceConfig = {}) {
  const headerInfo = findManualInputHeaderRowWithConfig(rows, sourceConfig);
  const indexMap = getManualInputHeaderIndexMap(headerInfo.headerRow);
  const requiredAny = sourceConfig.required_any || ["user_id", "name", "email", "mobile"];
  const fieldAliases = sourceConfig.field_aliases || {};

  let dataRowCount = 0;
  let courseCodeValueCount = 0;

  rows.slice(headerInfo.rowIndex + 1).forEach((rawRow) => {
    const extracted = buildManualInputHeaderAliasRowData(rawRow || [], indexMap, fieldAliases);
    const hasData = requiredAny.some((field) => hasFilledValue(extracted[field]));
    if (!hasData) {
      return;
    }
    dataRowCount += 1;
    if (hasFilledValue(extracted.course_code)) {
      courseCodeValueCount += 1;
    }
  });

  return {
    headerRowIndex: headerInfo.rowIndex,
    dataRowCount,
    courseCodeValueCount,
  };
}

function getManualCourseCandidateSheets(profile, workbookData) {
  if (!profile?.manual_course_input?.required || !Array.isArray(workbookData)) {
    return [];
  }

  const sourceConfig = profile.source || {};
  const selectedSheetNames = chooseProfileSheetNames(
    workbookData.map((sheet) => sheet.name),
    sourceConfig
  );

  if (profile.id === "manual_fixed_sheet_course_codes") {
    return workbookData
      .filter((sheet) => selectedSheetNames.includes(sheet.name))
      .filter((sheet) => matchesStructurePattern("manual_fixed_course_sheet", sheet.rows || []))
      .filter((sheet) => hasManualFixedCourseData(sheet.rows || []))
      .map((sheet) => sheet.name);
  }

  if (sourceConfig.mode === "header_alias") {
    return workbookData
      .filter((sheet) => selectedSheetNames.includes(sheet.name))
      .filter((sheet) => {
        try {
          return hasRequiredHeaderAliasData(sheet.rows || [], sourceConfig);
        } catch (_error) {
          return false;
        }
      })
      .map((sheet) => sheet.name);
  }

  return selectedSheetNames;
}

function getManualCourseCandidateGroups(profile, workbookData) {
  if (!profile?.manual_course_input?.required || !Array.isArray(workbookData)) {
    return {};
  }

  const sourceConfig = profile.source || {};
  const groupAliases = Array.isArray(sourceConfig.manual_course_group_aliases)
    ? sourceConfig.manual_course_group_aliases
    : [];
  if (!groupAliases.length || sourceConfig.mode !== "header_alias") {
    return {};
  }

  const requiredAny = sourceConfig.required_any || ["user_id", "name", "email", "mobile"];
  const fieldAliases = sourceConfig.field_aliases || {};
  const selectedSheetNames = chooseProfileSheetNames(
    workbookData.map((sheet) => sheet.name),
    sourceConfig
  );

  return Object.fromEntries(
    workbookData
      .filter((sheet) => selectedSheetNames.includes(sheet.name))
      .map((sheet) => {
        try {
          const headerInfo = findManualInputHeaderRowWithConfig(sheet.rows || [], sourceConfig);
          const indexMap = getManualInputHeaderIndexMap(headerInfo.headerRow);
          const groupColumnIndex = pickManualInputIndexByAliases(indexMap, groupAliases);
          if (groupColumnIndex === null || groupColumnIndex === undefined) {
            return [sheet.name, []];
          }

          const seen = new Set();
          const groups = [];
          (sheet.rows || []).slice(headerInfo.rowIndex + 1).forEach((rawRow) => {
            const extracted = buildManualInputHeaderAliasRowData(rawRow || [], indexMap, fieldAliases);
            const hasData = requiredAny.some((field) => hasFilledValue(extracted[field]));
            if (!hasData) {
              return;
            }

            const groupValue = rawRow[groupColumnIndex];
            if (!hasFilledValue(groupValue)) {
              return;
            }

            const groupName = String(groupValue).trim();
            const groupKey = normalizeSheetNameKey(groupName);
            if (!groupKey || seen.has(groupKey)) {
              return;
            }

            seen.add(groupKey);
            groups.push(groupName);
          });

          return [sheet.name, groups];
        } catch (_error) {
          return [sheet.name, []];
        }
      })
      .filter(([, groups]) => groups.length > 0)
  );
}

function resolveManualCourseAssignmentsForProfile(profile, workbookData, rawText) {
  const allSheetNames = Array.isArray(workbookData) ? workbookData.map((sheet) => sheet.name) : [];
  const parsedAssignments = parseManualCourseAssignments(rawText, allSheetNames);
  const candidateSheetNames = getManualCourseCandidateSheets(profile, workbookData);
  const candidateGroupsBySheet = getManualCourseCandidateGroups(profile, workbookData);
  const runtimeSheetCourseCodes = { ...parsedAssignments.sheetCourseCodes };
  const runtimeGroupCourseCodes = {};
  const remainingEntries = [];

  parsedAssignments.orderedEntries.forEach((entry) => {
    if (!entry.name) {
      remainingEntries.push(entry);
      return;
    }

    const matchedTargets = [];
    Object.entries(candidateGroupsBySheet).forEach(([sheetName, groupNames]) => {
      groupNames.forEach((groupName) => {
        if (normalizeSheetNameKey(groupName) === normalizeSheetNameKey(entry.name)) {
          matchedTargets.push({ sheetName, groupName });
        }
      });
    });

    if (matchedTargets.length === 1) {
      const target = matchedTargets[0];
      runtimeGroupCourseCodes[target.sheetName] = {
        ...(runtimeGroupCourseCodes[target.sheetName] || {}),
        [target.groupName]: entry.code,
      };
      return;
    }

    remainingEntries.push(entry);
  });

  const unmatchedGroupTargets = Object.entries(candidateGroupsBySheet).flatMap(([sheetName, groupNames]) =>
    groupNames
      .filter((groupName) => !runtimeGroupCourseCodes[sheetName]?.[groupName])
      .map((groupName) => ({ sheetName, groupName }))
  );
  const hasSingleGroupedSheet = Object.keys(candidateGroupsBySheet).length === 1;
  const canAssignRemainingEntriesToGroupsInOrder =
    remainingEntries.length > 0 &&
    unmatchedGroupTargets.length === remainingEntries.length &&
    (
      remainingEntries.every((entry) => !entry.name) ||
      hasSingleGroupedSheet
    );

  if (canAssignRemainingEntriesToGroupsInOrder) {
    remainingEntries.forEach((entry, index) => {
      const target = unmatchedGroupTargets[index];
      if (!target) {
        return;
      }
      runtimeGroupCourseCodes[target.sheetName] = {
        ...(runtimeGroupCourseCodes[target.sheetName] || {}),
        [target.groupName]: entry.code,
      };
    });
    remainingEntries.length = 0;
  } else {
    const orderedTargetSheetNames = (candidateSheetNames.length ? candidateSheetNames : allSheetNames)
      .filter((sheetName) => !runtimeSheetCourseCodes[sheetName]);

    if (remainingEntries.length === 1 && orderedTargetSheetNames.length === 1) {
      runtimeSheetCourseCodes[orderedTargetSheetNames[0]] = remainingEntries[0].code;
    } else if (remainingEntries.length > 1) {
      remainingEntries.forEach((entry, index) => {
        const sheetName = orderedTargetSheetNames[index];
        if (sheetName) {
          runtimeSheetCourseCodes[sheetName] = entry.code;
        }
      });
    }
  }

  return {
    sheetCourseCodes: runtimeSheetCourseCodes,
    groupCourseCodes: runtimeGroupCourseCodes,
    orderedEntries: remainingEntries,
    candidateSheetNames,
  };
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function createTimestamp() {
  const now = new Date();
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
}

function formatLogTimestamp(value) {
  return String(value || "").trim() || "-";
}

function loadUsageEntries() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return [];
    }
    const items = JSON.parse(raw);
    return Array.isArray(items) ? items : [];
  } catch (_error) {
    return [];
  }
}

function saveUsageEntry(entry) {
  const items = loadUsageEntries();
  const nextItems = [entry, ...items].slice(0, 200);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(nextItems));
}

function buildUsageCounts(items) {
  const today = createTimestamp().slice(0, 10);
  return {
    total_count: items.length,
    today_count: items.filter((item) => String(item.timestamp || "").startsWith(today)).length,
  };
}

function renderUsageCounts(counts) {
  totalUsageCount.textContent = counts?.total_count ?? 0;
  todayUsageCount.textContent = counts?.today_count ?? 0;
}

function renderRecentUsage(items) {
  if (!Array.isArray(items) || items.length === 0) {
    recentUsageBox.className = "recent-usage-panel empty-log";
    recentUsageBox.textContent = "아직 이 브라우저에 저장된 사용 내역이 없습니다.";
    renderUsageCounts({ total_count: 0, today_count: 0 });
    return;
  }

  recentUsageBox.className = "recent-usage-panel";
  const list = document.createElement("ul");
  list.className = "log-list";

  items.forEach((item) => {
    const entry = document.createElement("li");
    entry.className = "log-item";

    const top = document.createElement("div");
    top.className = "log-top";

    const time = document.createElement("div");
    time.className = "log-time";
    time.textContent = formatLogTimestamp(item.timestamp);

    const status = document.createElement("div");
    const isSuccess = item.status === "success";
    status.className = `log-status ${isSuccess ? "log-status-success" : "log-status-error"}`;
    status.textContent = isSuccess ? "성공" : "오류";

    top.appendChild(time);
    top.appendChild(status);

    const main = document.createElement("div");
    main.className = "log-main";
    main.textContent = item.profile_label || "유형 정보 없음";

    const sub = document.createElement("div");
    sub.className = "log-sub";
    if (isSuccess) {
      sub.textContent = `변환 행 수: ${item.total_rows ?? "-"} | 이메일 공란 처리: ${item.email_undefined ?? "-"} | 휴대폰 공란 처리: ${item.mobile_undefined ?? "-"}`;
    } else {
      sub.textContent = `오류: ${item.error || "알 수 없는 오류"}`;
    }

    entry.appendChild(top);
    entry.appendChild(main);
    entry.appendChild(sub);
    list.appendChild(entry);
  });

  recentUsageBox.replaceChildren(list);
  renderUsageCounts(buildUsageCounts(items));
}

function refreshUsageView() {
  renderRecentUsage(loadUsageEntries());
}

function getFileExtension(fileName) {
  const match = String(fileName || "").toLowerCase().match(/\.[^.]+$/);
  return match ? match[0] : "";
}

function ensureSupportedFile(file) {
  const extension = getFileExtension(file?.name);
  if (!ALLOWED_EXTENSIONS.has(extension)) {
    throw new Error("지원하지 않는 엑셀 형식입니다.");
  }
}

async function readFileAsArrayBuffer(file) {
  if (typeof file.arrayBuffer === "function") {
    return file.arrayBuffer();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
    reader.readAsArrayBuffer(file);
  });
}

function extractCellValue(sheet, rowIndex, columnIndex) {
  const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  const cell = sheet[cellRef];
  return cell && Object.prototype.hasOwnProperty.call(cell, "v") ? cell.v : null;
}

function buildSheetRowsPreservingLayout(sheet) {
  const ref = sheet["!ref"];
  if (!ref) {
    return [];
  }

  const range = XLSX.utils.decode_range(ref);
  const rows = [];

  // We intentionally start from A1 so fixed Excel column letters like C, K, T
  // continue to match even when the worksheet begins with fully blank columns.
  for (let rowIndex = 0; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    for (let columnIndex = 0; columnIndex <= range.e.c; columnIndex += 1) {
      row.push(extractCellValue(sheet, rowIndex, columnIndex));
    }
    rows.push(row);
  }

  return rows;
}

async function extractWorkbookData(file) {
  ensureSupportedFile(file);
  const data = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(data, {
    type: "array",
    cellDates: false,
    cellText: false,
  });

  return workbook.SheetNames.map((sheetName) => ({
    name: sheetName,
    rows: buildSheetRowsPreservingLayout(workbook.Sheets[sheetName]),
  }));
}

function createWorkbookDownload(headers, rowMatrix) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rowMatrix]);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  const bytes = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });
  return new Blob(
    [bytes],
    {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
  );
}

function explainError(error) {
  const message = String(error?.message || error || "");
  if (message.includes("선택한 파일 유형과 업로드한 파일 구조가 다릅니다")) {
    return "선택한 상위 유형이 현재 파일과 맞지 않습니다.\n다른 상위 유형으로 바꿔서 다시 시도해 주세요.";
  }
  if (message.includes("시트별 과정코드를 직접 입력해야 합니다")) {
    return "이 유형은 시트별 과정코드를 함께 입력해야 합니다.\n예: `사무직 = HLAP21561`처럼 시트명과 과정코드를 한 줄씩 넣어 주세요.";
  }
  if (message.includes("구분값별 과정코드를 직접 입력해야 합니다")) {
    return "이 파일은 한 장 안에서 직종/구분별로 과정코드를 나눠 넣어야 합니다.\n예: `사무직 = HLAP21561`와 `비사무직 = HLAP21547`처럼 한 줄씩 입력해 주세요.";
  }
  if (message.includes("과정코드 첨부파일 경로를 읽지 못했습니다")) {
    return "과정코드 첨부파일 경로를 읽지 못했습니다.\n일반 브라우저 보안으로 막힌 경우가 많으니, 같은 파일을 아래 첨부파일 선택으로 넣어 주세요.";
  }
  if (message.includes("지원하지 않는 과정코드 첨부파일 형식")) {
    return "과정코드 첨부파일은 .txt, .csv, .tsv, .xlsx, .xlsm, .xltx, .xltm, .xls 형식만 읽을 수 있습니다.";
  }
  if (message.includes("헤더 행을 찾지 못했습니다")) {
    return "선택한 상위 유형이 현재 파일과 맞지 않습니다.\n다른 상위 유형으로 바꿔서 다시 시도해 주세요.";
  }
  if (message.includes("명단으로 보이는 데이터 행을 찾지 못했습니다")) {
    return "표 안에서 실제 대상자 명단을 찾지 못했습니다.\n다른 상위 유형으로 바꾸거나 `2. 일반 명단 파일` 유형으로 다시 시도해 주세요.";
  }
  if (message.includes("지원하지 않는 엑셀 형식")) {
    return "지원하는 확장자는 .xlsx, .xlsm, .xltx, .xltm, .xls 입니다.";
  }
  if (message.includes("파일 유형을 먼저 선택") || message.includes("상위 유형을 먼저 선택")) {
    return "상위 유형을 먼저 선택해 주세요.";
  }
  return "변환 중 오류가 발생했습니다.\n원본 메시지: " + message;
}

function saveSuccessLog(profile, summary) {
  saveUsageEntry({
    timestamp: createTimestamp(),
    profile_id: profile.id,
    profile_label: profile.label,
    status: "success",
    total_rows: summary.total_rows,
    email_undefined: summary.email_undefined,
    mobile_undefined: summary.mobile_undefined,
  });
}

function saveErrorLog(profile, errorMessage) {
  saveUsageEntry({
    timestamp: createTimestamp(),
    profile_id: profile?.id || "",
    profile_label: profile?.label || "유형 정보 없음",
    status: "error",
    error: errorMessage,
  });
}

function buildSuccessStatus(summary) {
  const unresolvedCount = (summary.sheet_stats || []).reduce(
    (total, item) => total + ((item.unresolved_course_names || []).length),
    0
  );
  if (unresolvedCount > 0) {
    return `<strong>변환은 완료되었습니다.</strong><br>과정코드가 연결되지 않은 항목이 ${unresolvedCount}건 있습니다. 요약을 확인한 뒤 <strong>결과 파일 다운로드</strong>를 눌러 받아 주세요.`;
  }
  return "<strong>변환이 완료되었습니다.</strong><br>결과 파일이 준비되었습니다. <strong>결과 파일 다운로드</strong>를 눌러 받아 주세요.";
}

async function convertFile() {
  const file = fileInput.files?.[0];
  const family = getSelectedFamily();
  clearDownload();
  errorBox.value = "";
  setSummary(null);

  if (!file) {
    statusBox.textContent = "먼저 변환할 엑셀 파일을 선택해 주세요.";
    return;
  }

  convertButton.disabled = true;
  statusBox.innerHTML = "<strong>변환 중입니다.</strong><br>브라우저 안에서 파일을 읽고 결과 파일을 만들고 있습니다.";

  try {
    const workbookData = await extractWorkbookData(file);
    const profile = resolveProfileForFamily(family?.id, workbookData, file.name);
    if (!profile) {
      throw new Error("상위 유형을 먼저 선택해 주세요.");
    }
    selectedProfileId = profile.id;
    updateProfileGuide();
    const manualCourseSourceText = profile.manual_course_input?.required
      ? await buildManualCourseSourceText(
          manualCourseInput.value,
          manualCoursePathInput.value,
          manualCourseFileInput.files?.[0] || null
        )
      : "";
    const manualCourseAssignments = resolveManualCourseAssignmentsForProfile(
      profile,
      workbookData,
      manualCourseSourceText
    );

    const result = runProfile(profile, workbookData, file.name, {
      manual_sheet_course_codes: manualCourseAssignments.sheetCourseCodes,
      manual_group_course_codes: manualCourseAssignments.groupCourseCodes,
    });
    if (!result.summary?.total_rows) {
      throw new Error("명단으로 보이는 데이터 행을 찾지 못했습니다.");
    }
    const blob = createWorkbookDownload(result.headers, result.rowMatrix);
    latestDownloadUrl = URL.createObjectURL(blob);
    downloadLink.href = latestDownloadUrl;
    downloadLink.download = result.outputFileName;
    downloadLink.classList.remove("disabled");
    setSummary(result.summary);
    statusBox.innerHTML = buildSuccessStatus(result.summary);
    saveSuccessLog(profile, result.summary);
    refreshUsageView();
  } catch (error) {
    const friendlyMessage = explainError(error);
    statusBox.textContent = "변환에 실패했습니다. 오류 안내 내용을 확인해 주세요.";
    errorBox.value = friendlyMessage;
    saveErrorLog(profile, friendlyMessage);
    refreshUsageView();
  } finally {
    convertButton.disabled = false;
  }
}

familySelect.addEventListener("change", () => {
  const family = getSelectedFamily();
  if (!family) {
    selectedProfileId = "";
    updateProfileGuide();
    return;
  }

  const resolved = resolveProfileForFamily(family.id, latestWorkbookData, latestInputFileName);
  selectedProfileId = resolved?.id || getDefaultProfileForFamily(family.id)?.id || "";
  updateProfileGuide();

  if (fileInput.files?.[0] && getProfileFamilyId(recommendedProfileId) !== family.id) {
    statusBox.innerHTML = `<strong>${family.label}</strong> 유형으로 변경했습니다.<br>변환할 때 현재 파일에 맞는 내부 세부 유형을 이 묶음 안에서 다시 자동 선택합니다.`;
  }
});
fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) {
    latestWorkbookData = [];
    latestInputFileName = "";
    recommendedProfileId = "";
    recommendedProfileReason = "";
    selectedProfileId = getDefaultProfileForFamily(familySelect.value)?.id || "";
    updateProfileGuide();
    return;
  }

  latestInputFileName = file.name;
  statusBox.innerHTML = "<strong>파일을 확인 중입니다.</strong><br>파일명과 시트 구조를 함께 보고 가장 비슷한 상위 유형과 내부 세부 유형을 찾고 있습니다.";

  try {
    const workbookData = await extractWorkbookData(file);
    latestWorkbookData = workbookData;
    const structureRecommendation = recommendProfileByWorkbookStructure(workbookData);
    if (structureRecommendation) {
      applyRecommendation(
        structureRecommendation.profile,
        structureRecommendation.reason,
        structureRecommendation.message
      );
      return;
    }
  } catch (_error) {
    latestWorkbookData = [];
    // Ignore recommendation-time parsing issues and fall back to filename only.
  }

  recommendProfileByFileName(file.name);
});
convertButton.addEventListener("click", convertFile);

clearLegacyUsageData();
renderFamilies();
refreshUsageView();
