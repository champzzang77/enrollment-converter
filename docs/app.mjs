import { PROFILES } from "./data.mjs";
import { runProfile } from "./engine.mjs";

const STORAGE_KEY = "enrollment-upload-static-usage-log-v2";
const LEGACY_STORAGE_KEYS = ["enrollment-upload-static-usage-log-v1"];
const ALLOWED_EXTENSIONS = new Set([".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"]);

const profileSelect = document.getElementById("profileSelect");
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
const guideHints = document.getElementById("guideHints");
const recommendBadge = document.getElementById("recommendBadge");
const manualCourseField = document.getElementById("manualCourseField");
const manualCourseLabel = document.getElementById("manualCourseLabel");
const manualCourseInput = document.getElementById("manualCourseInput");
const manualCourseNote = document.getElementById("manualCourseNote");
const recentUsageBox = document.getElementById("recentUsageBox");
const totalUsageCount = document.getElementById("totalUsageCount");
const todayUsageCount = document.getElementById("todayUsageCount");

let latestDownloadUrl = "";
let recommendedProfileId = "";
let recommendedProfileReason = "";

function clearLegacyUsageData() {
  LEGACY_STORAGE_KEYS.forEach((key) => {
    try {
      localStorage.removeItem(key);
    } catch (_error) {
      // Ignore storage cleanup issues.
    }
  });
}

function getSelectedProfile() {
  return PROFILES.find((item) => item.id === profileSelect.value) || null;
}

function getProfileDisplayOrder(profile) {
  const match = String(profile?.label || "").match(/^(\d+)\./);
  return match ? Number(match[1]) : Number.POSITIVE_INFINITY;
}

function renderProfiles() {
  const sortedProfiles = [...PROFILES].sort((left, right) => {
    const orderDiff = getProfileDisplayOrder(left) - getProfileDisplayOrder(right);
    if (orderDiff !== 0) {
      return orderDiff;
    }
    return String(left.label || "").localeCompare(String(right.label || ""), "ko");
  });

  sortedProfiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile.id;
    option.textContent = `${profile.label} - ${profile.short_description}`;
    profileSelect.appendChild(option);
  });
  if (!profileSelect.value && sortedProfiles[0]) {
    profileSelect.value = sortedProfiles[0].id;
  }
  updateProfileGuide();
}

function updateProfileGuide() {
  const profile = getSelectedProfile();
  if (!profile) {
    guideTitle.textContent = "파일 유형을 선택해 주세요";
    guideDescription.textContent = "선택한 유형에 대한 설명이 여기에 표시됩니다.";
    guideUseWhen.textContent = "-";
    guideExample.textContent = "-";
    guideHints.innerHTML = "";
    recommendBadge.textContent = "직접 선택";
    recommendBadge.className = "badge badge-neutral";
    manualCourseField.hidden = true;
    return;
  }

  guideTitle.textContent = profile.label;
  guideDescription.textContent = profile.description;
  guideUseWhen.textContent = profile.use_when;
  guideExample.textContent = profile.example_file;
  guideHints.innerHTML = "";
  profile.hints.forEach((hint) => {
    const item = document.createElement("li");
    item.textContent = hint;
    guideHints.appendChild(item);
  });

  if (recommendedProfileId && recommendedProfileId === profile.id) {
    recommendBadge.textContent = recommendedProfileReason === "structure"
      ? "파일 구조 기준 추천"
      : "파일명 기준 추천";
    recommendBadge.className = "badge badge-recommend";
  } else {
    recommendBadge.textContent = "직접 선택";
    recommendBadge.className = "badge badge-neutral";
  }

  const manualCourseConfig = profile.manual_course_input;
  if (manualCourseConfig) {
    manualCourseField.hidden = false;
    manualCourseLabel.textContent = manualCourseConfig.label || "선택 입력. 시트별 과정코드 직접 입력";
    manualCourseNote.textContent = manualCourseConfig.note || "과정코드를 따로 받은 경우에만 입력해 주세요.";
    manualCourseInput.placeholder = manualCourseConfig.placeholder || "";
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

function getCell(rows, rowIndex, columnIndex) {
  const row = rows[rowIndex] || [];
  return columnIndex < row.length ? row[columnIndex] : null;
}

function checkCompletedApplicationSheet(rows) {
  const checks = [
    { row: 17, col: 2, includes: "과정코드" },
    { row: 17, col: 3, includes: "과정명" },
    { row: 17, col: 4, includes: "이름" },
    { row: 17, col: 5, includes: "희망id" },
    { row: 17, col: 10, includes: "이메일" },
    { row: 17, col: 11, includes: "휴대폰" },
    { row: 17, col: 13, includes: "부서" },
  ];

  return checks.every((check) =>
    normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
  );
}

function checkGroupedSingleSheetApplication(rows) {
  const headerChecks = [
    { row: 17, col: 2, includes: "과정코드" },
    { row: 17, col: 3, includes: "과정명" },
    { row: 17, col: 4, includes: "이름" },
    { row: 17, col: 5, includes: "희망id" },
  ];

  const headerMatched = headerChecks.every((check) =>
    normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
  );

  if (!headerMatched) {
    return false;
  }

  let courseRows = 0;
  let anchorRows = 0;
  for (let rowIndex = 18; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const courseCode = String(row[2] || "").trim();
    const name = String(row[4] || "").trim();
    const userId = String(row[5] || "").trim();
    if (courseCode) {
      courseRows += 1;
    }
    if (name || userId) {
      anchorRows += 1;
    }
  }

  return courseRows >= 5 && anchorRows >= 2 && anchorRows < courseRows;
}

function checkCmcTrainingTeamBundle(workbookData) {
  const targetSheets = ["재직직원", "신규직원"];
  const matchedSheets = workbookData.filter((sheet) => targetSheets.includes(sheet.name));
  if (!matchedSheets.length) {
    return false;
  }

  return matchedSheets.some((sheet) => {
    const rows = sheet.rows || [];
    const checks = [
      { row: 17, col: 2, includes: "과정코드" },
      { row: 17, col: 3, includes: "과정명" },
      { row: 17, col: 4, includes: "이름" },
      { row: 17, col: 5, includes: "희망id" },
      { row: 17, col: 10, includes: "이메일" },
      { row: 17, col: 11, includes: "휴대폰" },
      { row: 17, col: 13, includes: "부서" },
    ];

    const headerMatched = checks.every((check) =>
      normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
    );

    if (!headerMatched) {
      return false;
    }

    let seedCourseCount = 0;
    let blankAfterSeed = 0;
    for (let rowIndex = 18; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const courseCode = String(row[2] || "").trim();
      const name = String(row[4] || "").trim();
      const userId = String(row[5] || row[6] || "").trim();

      if (courseCode) {
        seedCourseCount += 1;
        continue;
      }

      if ((name || userId) && seedCourseCount > 0) {
        blankAfterSeed += 1;
      }
    }

    return seedCourseCount >= 5 && blankAfterSeed >= 5;
  });
}

function checkManualCourseSheet(rows) {
  const checks = [
    { row: 1, col: 2, includes: "이름" },
    { row: 1, col: 3, includes: "희망id" },
    { row: 1, col: 5, includes: "이메일" },
    { row: 1, col: 6, includes: "휴대폰" },
    { row: 1, col: 8, includes: "부서" },
  ];

  return checks.every((check) =>
    normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
  );
}

function checkSingleSheetManualCourseSheet(rows) {
  const checks = [
    { row: 17, col: 3, includes: "과정명" },
    { row: 17, col: 6, includes: "이름" },
    { row: 17, col: 7, includes: "전화번호" },
    { row: 17, col: 8, includes: "사번" },
    { row: 17, col: 9, includes: "부서" },
  ];

  return checks.every((check) =>
    normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
  );
}

function checkManualFixedCourseSheet(rows) {
  const checks = [
    { row: 17, col: 2, includes: "과정명" },
    { row: 17, col: 3, includes: "사원번호" },
    { row: 17, col: 5, includes: "이름" },
    { row: 17, col: 7, includes: "이메일" },
    { row: 17, col: 8, includes: "휴대폰" },
    { row: 17, col: 9, includes: "회사명" },
    { row: 17, col: 10, includes: "근무부서" },
  ];

  return checks.every((check) =>
    normalizeHeaderText(getCell(rows, check.row, check.col)).includes(check.includes)
  );
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

function checkInternationalStMaryBundle(workbookData) {
  const sheetNames = new Set((workbookData || []).map((sheet) => sheet.name));
  return (
    sheetNames.has("요청사항") &&
    sheetNames.has("1. 4주기_의료인증제필수교육") &&
    sheetNames.has("2. 법정의무교육")
  );
}

function recommendProfileByWorkbookStructure(workbookData) {
  const completedProfile = PROFILES.find((profile) => profile.id === "multi_sheet_completed_application");
  const manualProfile = PROFILES.find((profile) => profile.id === "manual_sheet_course_header");
  const manualFixedProfile = PROFILES.find((profile) => profile.id === "manual_fixed_sheet_course_codes");
  const internationalStMaryProfile = PROFILES.find((profile) => profile.id === "international_stmary_group_enrollment");
  const singleSheetManualProfile = PROFILES.find((profile) => profile.id === "single_sheet_manual_course_header");
  const groupedSingleSheetProfile = PROFILES.find((profile) => profile.id === "grouped_single_sheet_application");
  const cmcTrainingTeamProfile = PROFILES.find((profile) => profile.id === "cmc_training_team_seed_bundle");

  if (!Array.isArray(workbookData) || workbookData.length === 0) {
    return null;
  }

  if (internationalStMaryProfile && checkInternationalStMaryBundle(workbookData)) {
    return {
      profile: internationalStMaryProfile,
      reason: "structure",
      message: `<strong>${internationalStMaryProfile.label}</strong> 형식으로 추천했습니다.<br>국제성모병원 단체입과명단 구조가 확인되어 과정코드까지 자동 연결하도록 맞췄습니다.`,
    };
  }

  if (cmcTrainingTeamProfile && checkCmcTrainingTeamBundle(workbookData)) {
    return {
      profile: cmcTrainingTeamProfile,
      reason: "structure",
      message: `<strong>${cmcTrainingTeamProfile.label}</strong> 형식으로 추천했습니다.<br>재직직원/신규직원 시트 상단의 과정코드 묶음을 모든 대상자에게 자동 확장하도록 맞췄습니다.`,
    };
  }

  const groupedSingleSheetMatches = workbookData.filter((sheet) =>
    checkGroupedSingleSheetApplication(sheet.rows || [])
  );
  if (groupedSingleSheetProfile && groupedSingleSheetMatches.length >= 1) {
    return {
      profile: groupedSingleSheetProfile,
      reason: "structure",
      message: `<strong>${groupedSingleSheetProfile.label}</strong> 형식으로 추천했습니다.<br>한 사람 아래에 여러 과정이 묶인 입과 신청서 구조가 확인되어, 이름과 ID를 아래 과정들에 자동으로 이어 붙이도록 맞췄습니다.`,
    };
  }

  const singleSheetManualMatches = workbookData.filter((sheet) =>
    checkSingleSheetManualCourseSheet(sheet.rows || [])
  );
  if (singleSheetManualProfile && singleSheetManualMatches.length >= 1) {
    return {
      profile: singleSheetManualProfile,
      reason: "structure",
      message: `<strong>${singleSheetManualProfile.label}</strong> 형식으로 추천했습니다.<br>한 장짜리 추가 명단표 구조가 확인되었습니다. 아래 입력칸에 과정코드만 한 번 넣어 주세요.`,
    };
  }

  const completedMatches = workbookData.filter((sheet) => checkCompletedApplicationSheet(sheet.rows || []));
  if (completedProfile && completedMatches.length >= Math.min(2, workbookData.length)) {
    return {
      profile: completedProfile,
      reason: "structure",
      message: `<strong>${completedProfile.label}</strong> 형식으로 추천했습니다.<br>시트 여러 장에서 완성된 입과 신청서 표 구조가 확인되어 이 유형이 더 잘 맞아 보입니다.`,
    };
  }

  const manualMatches = workbookData.filter((sheet) => checkManualCourseSheet(sheet.rows || []));
  if (manualProfile && manualMatches.length >= Math.min(2, workbookData.length)) {
    return {
      profile: manualProfile,
      reason: "structure",
      message: `<strong>${manualProfile.label}</strong> 형식으로 추천했습니다.<br>시트별 명단 구조가 확인되었습니다. 과정코드를 따로 받은 경우 아래 입력칸에 함께 넣어 주세요.`,
    };
  }

  const manualFixedMatches = workbookData.filter((sheet) => checkManualFixedCourseSheet(sheet.rows || []));
  if (manualFixedProfile && manualFixedMatches.length >= 2) {
    return {
      profile: manualFixedProfile,
      reason: "structure",
      message: `<strong>${manualFixedProfile.label}</strong> 형식으로 추천했습니다.<br>여러 시트에서 단체 입과 신청양식 구조가 확인되었습니다. 과정코드를 따로 받은 경우 아래 입력칸에 시트별로 넣어 주세요.`,
    };
  }

  return null;
}

function applyRecommendation(profile, reason, message) {
  recommendedProfileId = profile.id;
  recommendedProfileReason = reason;
  profileSelect.value = profile.id;
  updateProfileGuide();
  statusBox.innerHTML = message;
}

function recommendProfileByFileName(fileName) {
  const normalizedName = normalizeFileNameMatch(fileName);
  let matched = null;
  let matchedScore = -1;

  PROFILES.forEach((profile) => {
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

  if (!matched) {
    const fallback = PROFILES.find((profile) => profile.id === "generic_auto_header");
    if (fallback) {
      applyRecommendation(
        fallback,
        "filename",
        `<strong>${fallback.label}</strong> 형식으로 추천했습니다.<br>파일명만으로 정확한 유형을 찾지 못해 범용 자동 인식 유형을 먼저 선택했습니다.`
      );
      return;
    }

    recommendedProfileId = "";
    recommendedProfileReason = "";
    updateProfileGuide();
    statusBox.innerHTML = "파일을 선택했습니다. <strong>파일 유형</strong>을 확인한 뒤 <strong>변환 시작</strong>을 눌러 주세요.";
    return;
  }

  if (matched.manual_course_input?.required) {
    applyRecommendation(
      matched,
      "filename",
      `<strong>${matched.label}</strong> 형식으로 추천했습니다.<br>예시 파일명을 확인한 뒤, 아래 입력칸에 시트별 과정코드도 함께 넣어 주세요.`
    );
    return;
  }

  applyRecommendation(
    matched,
    "filename",
    `<strong>${matched.label}</strong> 형식으로 추천했습니다.<br>예시 파일명과 설명이 맞는지만 한 번 확인해 주세요.`
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

function hasFilledValue(value) {
  return String(value ?? "").trim() !== "";
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

function pickManualInputValueByAliases(row, indexMap, aliases = []) {
  for (const alias of aliases) {
    const aliasKey = normalizeHeaderText(alias);
    for (const [headerKey, index] of Object.entries(indexMap)) {
      if (headerKey.includes(aliasKey) && index < row.length) {
        const value = row[index];
        if (hasFilledValue(value)) {
          return value;
        }
      }
    }
  }
  return null;
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
      .filter((sheet) => checkManualFixedCourseSheet(sheet.rows || []))
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

function resolveManualCourseAssignmentsForProfile(profile, workbookData, rawText) {
  const allSheetNames = Array.isArray(workbookData) ? workbookData.map((sheet) => sheet.name) : [];
  const parsedAssignments = parseManualCourseAssignments(rawText, allSheetNames);
  const candidateSheetNames = getManualCourseCandidateSheets(profile, workbookData);
  const runtimeSheetCourseCodes = { ...parsedAssignments.sheetCourseCodes };
  const orderedTargetSheetNames = (candidateSheetNames.length ? candidateSheetNames : allSheetNames)
    .filter((sheetName) => !runtimeSheetCourseCodes[sheetName]);

  if (parsedAssignments.orderedEntries.length === 1 && orderedTargetSheetNames.length === 1) {
    runtimeSheetCourseCodes[orderedTargetSheetNames[0]] = parsedAssignments.orderedEntries[0].code;
  } else if (parsedAssignments.orderedEntries.length > 1) {
    parsedAssignments.orderedEntries.forEach((entry, index) => {
      const sheetName = orderedTargetSheetNames[index];
      if (sheetName) {
        runtimeSheetCourseCodes[sheetName] = entry.code;
      }
    });
  }

  return {
    sheetCourseCodes: runtimeSheetCourseCodes,
    orderedEntries: parsedAssignments.orderedEntries,
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
    return "선택한 파일 유형이 현재 파일과 맞지 않습니다.\n파일 유형을 바꿔서 다시 시도해 주세요.";
  }
  if (message.includes("시트별 과정코드를 직접 입력해야 합니다")) {
    return "이 유형은 시트별 과정코드를 함께 입력해야 합니다.\n예: `사무직 = HLAP21561`처럼 시트명과 과정코드를 한 줄씩 넣어 주세요.";
  }
  if (message.includes("헤더 행을 찾지 못했습니다")) {
    return "선택한 파일 유형이 현재 파일과 맞지 않습니다.\n파일 유형을 바꿔서 다시 시도해 주세요.";
  }
  if (message.includes("명단으로 보이는 데이터 행을 찾지 못했습니다")) {
    return "표 안에서 실제 대상자 명단을 찾지 못했습니다.\n다른 파일 유형으로 바꾸거나 `8. 일반 명단 파일 자동 인식` 유형으로 다시 시도해 주세요.";
  }
  if (message.includes("지원하지 않는 엑셀 형식")) {
    return "지원하는 확장자는 .xlsx, .xlsm, .xltx, .xltm, .xls 입니다.";
  }
  if (message.includes("파일 유형을 먼저 선택")) {
    return "파일 유형을 먼저 선택해 주세요.";
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
  const profile = getSelectedProfile();
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
    const manualCourseAssignments = resolveManualCourseAssignmentsForProfile(
      profile,
      workbookData,
      manualCourseInput.value
    );

    const result = runProfile(profile, workbookData, file.name, {
      manual_sheet_course_codes: manualCourseAssignments.sheetCourseCodes,
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

profileSelect.addEventListener("change", updateProfileGuide);
fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) {
    recommendedProfileId = "";
    recommendedProfileReason = "";
    updateProfileGuide();
    return;
  }

  statusBox.innerHTML = "<strong>파일을 확인 중입니다.</strong><br>파일명과 시트 구조를 함께 보고 가장 비슷한 유형을 찾고 있습니다.";

  try {
    const workbookData = await extractWorkbookData(file);
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
    // Ignore recommendation-time parsing issues and fall back to filename only.
  }

  recommendProfileByFileName(file.name);
});
convertButton.addEventListener("click", convertFile);

clearLegacyUsageData();
renderProfiles();
refreshUsageView();
