import { GUIDE_HEADERS, OUTPUT_FIELDS } from "./data.mjs";

export function text(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "boolean") {
    return value ? "Y" : "N";
  }
  if (typeof value === "number") {
    if (Number.isInteger(value)) {
      return String(value);
    }
    return String(value).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  const normalized = String(value).trim();
  return normalized || null;
}

export function normalizeHeader(value) {
  const raw = text(value) || "";
  return raw
    .replace(/\n/g, "")
    .replace(/ /g, "")
    .replace(/\xa0/g, "")
    .replace(/\*/g, "")
    .replace(/[①②③④⑤⑥()]/g, "")
    .toLowerCase();
}

export function normalizeMatchKey(value) {
  const raw = text(value) || "";
  return raw.replace(/\xa0/g, " ").replace(/\s+/g, "").toLowerCase();
}

export function columnRefToIndex(ref) {
  if (typeof ref === "number") {
    return ref >= 1 ? ref - 1 : ref;
  }

  const raw = String(ref || "").trim().toUpperCase();
  if (/^\d+$/.test(raw)) {
    return Number(raw) - 1;
  }

  let total = 0;
  for (const char of raw) {
    if (char < "A" || char > "Z") {
      throw new Error(`잘못된 컬럼 표기입니다: ${ref}`);
    }
    total = total * 26 + (char.charCodeAt(0) - 64);
  }
  return total - 1;
}

function normalizeDefaults(defaults = {}) {
  return Object.fromEntries(
    Object.entries(defaults).map(([key, value]) => [key, text(value)])
  );
}

function normalizeKeywordList(values = []) {
  return values.map((value) => normalizeHeader(value)).filter(Boolean);
}

function normalizeLookup(lookup = {}) {
  return Object.fromEntries(
    Object.entries(lookup).map(([key, value]) => [normalizeMatchKey(key), text(value)])
  );
}

function hasRequiredValue(rowData, requiredFields) {
  return requiredFields.some((field) => text(rowData[field]));
}

function applyDefaults(rowData, defaults = {}) {
  const merged = { ...rowData };
  for (const [key, value] of Object.entries(defaults)) {
    if (!text(merged[key])) {
      merged[key] = text(value);
    }
  }
  return merged;
}

function finalizeRow(rowData, copyIfMissing = {}, undefinedIfMissing = []) {
  const finalized = { method: "0", ...rowData };

  for (const [destination, source] of Object.entries(copyIfMissing)) {
    if (!text(finalized[destination]) && text(finalized[source])) {
      finalized[destination] = finalized[source];
    }
  }

  for (const field of undefinedIfMissing) {
    if (!text(finalized[field])) {
      finalized[field] = "undefined";
    }
  }

  for (const field of ["sms", "mail"]) {
    if (text(finalized[field])) {
      finalized[field] = String(finalized[field]).toUpperCase();
    }
  }

  return Object.fromEntries(
    OUTPUT_FIELDS.map((field) => [field, text(finalized[field]) || null])
  );
}

function countUndefined(rows, fieldName) {
  return rows.filter((row) => row[fieldName] === "undefined").length;
}

function normalizeSheetSelection(sheetNames, includeSheets = [], skipSheets = []) {
  const skipSet = new Set(skipSheets);
  const selected = includeSheets.length
    ? sheetNames.filter((name) => includeSheets.includes(name))
    : [...sheetNames];
  return selected.filter((name) => !skipSet.has(name));
}

function chooseSheetNames(sheetNames, sourceConfig) {
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

  const selected = normalizeSheetSelection(
    sheetNames,
    sourceConfig.include_sheets || [],
    sourceConfig.skip_sheets || []
  );
  return sourceConfig.first_sheet_only ? selected.slice(0, 1) : selected;
}

function resolveCourseExpansions(sheetName, sourceConfig, courseLookup) {
  const directConfig = sourceConfig.sheet_course_codes?.[sheetName];
  if (directConfig) {
    return {
      expansions: directConfig.map((item) => {
        if (typeof item === "object" && item) {
          return {
            name: text(item.name),
            code: text(item.code),
          };
        }
        return {
          name: null,
          code: text(item),
        };
      }),
      unresolvedNames: [],
    };
  }

  const names = sourceConfig.sheet_course_names?.[sheetName] || [];
  if (!names.length) {
    return {
      expansions: [{ name: null, code: null }],
      unresolvedNames: [],
    };
  }

  const expansions = [];
  const unresolvedNames = [];
  for (const name of names) {
    const code = courseLookup[normalizeMatchKey(name)] || null;
    expansions.push({ name: text(name), code });
    if (!code) {
      unresolvedNames.push(name);
    }
  }
  return { expansions, unresolvedNames };
}

function findHeaderRowByKeywords(rows, requiredKeywords) {
  const normalizedKeywords = requiredKeywords.map((keyword) => normalizeHeader(keyword));
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const joined = row.map((cell) => normalizeHeader(cell)).join("|");
    if (normalizedKeywords.every((keyword) => joined.includes(keyword))) {
      return { rowIndex, headerRow: row };
    }
  }
  throw new Error(`헤더 행을 찾지 못했습니다: ${requiredKeywords.join(", ")}`);
}

function findHeaderRowWithConfig(rows, sourceConfig) {
  const keywordGroups = Array.isArray(sourceConfig.header_keyword_groups) && sourceConfig.header_keyword_groups.length
    ? sourceConfig.header_keyword_groups
    : [sourceConfig.header_keywords || []];

  let lastError = null;
  for (const keywords of keywordGroups) {
    try {
      return findHeaderRowByKeywords(rows, keywords);
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("헤더 행을 찾지 못했습니다.");
}

function getHeaderIndexMap(headerRow) {
  const entries = [];
  headerRow.forEach((cell, index) => {
    const key = normalizeHeader(cell);
    if (key) {
      entries.push([key, index]);
    }
  });
  return Object.fromEntries(entries);
}

function pickByAliases(row, indexMap, aliases) {
  for (const alias of aliases) {
    const aliasKey = normalizeHeader(alias);
    for (const [headerKey, index] of Object.entries(indexMap)) {
      if (headerKey.includes(aliasKey) && index < row.length) {
        return text(row[index]);
      }
    }
  }
  return null;
}

function scanRowLabelValue(rows, labels) {
  const targets = new Set(labels.map((label) => normalizeHeader(label)));
  for (const row of rows) {
    const normalizedRow = row.map((cell) => normalizeHeader(cell));
    for (let index = 0; index < normalizedRow.length; index += 1) {
      if (!targets.has(normalizedRow[index])) {
        continue;
      }
      for (const candidate of row.slice(index + 1)) {
        const value = text(candidate);
        if (value) {
          return value;
        }
      }
    }
  }
  return null;
}

function validateFixedColumnLayout(rows, sourceConfig, sheetName) {
  const validation = sourceConfig.layout_validation;
  if (!validation?.column_checks?.length) {
    return;
  }

  const derivedHeaderRow = Number(sourceConfig.start_row || 2) - 1;
  const headerRowNumber = Number(validation.header_row || derivedHeaderRow);
  const headerRowIndex = headerRowNumber - 1;
  const headerRow = rows[headerRowIndex] || [];
  const mismatchMessages = [];

  validation.column_checks.forEach((check) => {
    const columnIndex = columnRefToIndex(check.column);
    const rawHeader = text(columnIndex < headerRow.length ? headerRow[columnIndex] : null) || "";
    const normalizedHeader = normalizeHeader(rawHeader);
    const includes = normalizeKeywordList(check.includes || []);
    const anyOf = normalizeKeywordList(check.any_of || []);

    const includesMatched = includes.every((keyword) => normalizedHeader.includes(keyword));
    const anyOfMatched = !anyOf.length || anyOf.some((keyword) => normalizedHeader.includes(keyword));

    if (!includesMatched || !anyOfMatched) {
      const expectedParts = [];
      if (includes.length) {
        expectedParts.push(`포함: ${check.includes.join(", ")}`);
      }
      if (anyOf.length) {
        expectedParts.push(`다음 중 하나: ${check.any_of.join(", ")}`);
      }
      mismatchMessages.push(
        `${check.column}열 기대값(${expectedParts.join(" / ")}) != 실제값(${rawHeader || "빈칸"})`
      );
    }
  });

  if (mismatchMessages.length) {
    throw new Error(
      `선택한 파일 유형과 업로드한 파일 구조가 다릅니다. ${sheetName} 시트의 헤더를 확인해 주세요.\n` +
      mismatchMessages.join("\n")
    );
  }
}

function buildRowsFixedColumns(profile, workbookData, courseLookup) {
  const sourceConfig = profile.source;
  const headers = profile.guide_headers || GUIDE_HEADERS;
  if (headers.length !== GUIDE_HEADERS.length) {
    throw new Error("guide_headers 길이는 기본 업로드 양식과 같아야 합니다.");
  }

  const sheetMap = new Map(workbookData.map((sheet) => [sheet.name, sheet.rows]));
  const sheetNames = workbookData.map((sheet) => sheet.name);
  const columnMap = Object.fromEntries(
    Object.entries(sourceConfig.column_map).map(([field, columnRef]) => [field, columnRefToIndex(columnRef)])
  );
  const requiredAny = sourceConfig.required_any || ["user_id", "name"];
  const defaults = normalizeDefaults(profile.defaults || {});
  const sheetDefaults = sourceConfig.sheet_defaults || {};
  const selectedSheets = normalizeSheetSelection(
    sheetNames,
    sourceConfig.include_sheets || [],
    sourceConfig.skip_sheets || []
  );
  if (!selectedSheets.length) {
    throw new Error("선택한 파일 유형과 업로드한 파일 구조가 다릅니다. 필요한 시트를 찾지 못했습니다.");
  }
  const startRow = Number(sourceConfig.start_row || 2);
  const copyIfMissing = profile.copy_if_missing || {};
  const undefinedIfMissing = profile.undefined_if_missing || [];

  const rows = [];
  const sheetStats = [];

  for (const sheetName of selectedSheets) {
    const rawRows = sheetMap.get(sheetName) || [];
    validateFixedColumnLayout(rawRows, sourceConfig, sheetName);
    const { expansions, unresolvedNames } = resolveCourseExpansions(sheetName, sourceConfig, courseLookup);
    const perSheetDefaults = {
      ...defaults,
      ...normalizeDefaults(sheetDefaults[sheetName] || {}),
    };

    let sourcePersonCount = 0;
    let rowCount = 0;
    let emailUndefined = 0;
    let mobileUndefined = 0;

    for (let rowIndex = startRow - 1; rowIndex < rawRows.length; rowIndex += 1) {
      const rawRow = rawRows[rowIndex] || [];
      const extracted = Object.fromEntries(
        Object.entries(columnMap).map(([field, columnIndex]) => [
          field,
          text(columnIndex < rawRow.length ? rawRow[columnIndex] : null),
        ])
      );

      if (!hasRequiredValue(extracted, requiredAny)) {
        continue;
      }

      sourcePersonCount += 1;
      const expandedCourses = expansions.length ? expansions : [{ name: null, code: null }];
      for (const expansion of expandedCourses) {
        let rowData = { ...extracted };
        rowData.course_code = text(expansion.code) || rowData.course_code || null;
        rowData = applyDefaults(rowData, perSheetDefaults);
        const uploadRow = finalizeRow(rowData, copyIfMissing, undefinedIfMissing);
        rows.push(uploadRow);
        rowCount += 1;
        if (uploadRow.email === "undefined") {
          emailUndefined += 1;
        }
        if (uploadRow.mobile === "undefined") {
          mobileUndefined += 1;
        }
      }
    }

    sheetStats.push({
      sheet_name: sheetName,
      course_names: expansions.map((item) => item.name).filter(Boolean),
      course_codes: expansions.map((item) => item.code).filter(Boolean),
      unresolved_course_names: unresolvedNames,
      source_person_count: sourcePersonCount,
      row_count: rowCount,
      email_undefined: emailUndefined,
      mobile_undefined: mobileUndefined,
    });
  }

  return { rows, sheetStats };
}

function buildRowsHeaderAlias(profile, workbookData) {
  const sourceConfig = profile.source;
  const headers = profile.guide_headers || GUIDE_HEADERS;
  if (headers.length !== GUIDE_HEADERS.length) {
    throw new Error("guide_headers 길이는 기본 업로드 양식과 같아야 합니다.");
  }

  const sheetMap = new Map(workbookData.map((sheet) => [sheet.name, sheet.rows]));
  const sheetNames = workbookData.map((sheet) => sheet.name);
  const defaults = normalizeDefaults(profile.defaults || {});
  const sheetDefaults = sourceConfig.sheet_defaults || {};
  const copyIfMissing = profile.copy_if_missing || {};
  const undefinedIfMissing = profile.undefined_if_missing || [];
  const requiredAny = sourceConfig.required_any || ["user_id", "name", "email", "mobile"];
  const fieldAliases = sourceConfig.field_aliases || {};
  const headerKeywords = sourceConfig.header_keywords || [];
  const sheetContextLabels = sourceConfig.sheet_context_labels || {};
  const selectedSheets = chooseSheetNames(sheetNames, sourceConfig);
  if (!selectedSheets.length) {
    throw new Error("선택한 파일 유형과 업로드한 파일 구조가 다릅니다. 필요한 시트를 찾지 못했습니다.");
  }

  const rows = [];
  const sheetStats = [];

  for (const sheetName of selectedSheets) {
    const rawRows = sheetMap.get(sheetName) || [];
    let headerInfo;
    try {
      headerInfo = findHeaderRowWithConfig(rawRows, sourceConfig);
    } catch (error) {
      if (sourceConfig.ignore_missing_header) {
        sheetStats.push({
          sheet_name: sheetName,
          skipped_reason: "header_not_found",
          row_count: 0,
        });
        continue;
      }
      throw error;
    }

    const { rowIndex: headerRowIndex, headerRow } = headerInfo;
    const indexMap = getHeaderIndexMap(headerRow);
    const preHeaderRows = rawRows.slice(0, headerRowIndex);
    const sheetContext = Object.fromEntries(
      Object.entries(sheetContextLabels).map(([field, labels]) => [
        field,
        scanRowLabelValue(preHeaderRows, labels),
      ])
    );
    const perSheetDefaults = {
      ...defaults,
      ...normalizeDefaults(sheetDefaults[sheetName] || {}),
    };

    let sourcePersonCount = 0;
    let rowCount = 0;
    let emailUndefined = 0;
    let mobileUndefined = 0;

    for (const rawRow of rawRows.slice(headerRowIndex + 1)) {
      const extracted = Object.fromEntries(
        Object.entries(fieldAliases).map(([field, aliases]) => [
          field,
          pickByAliases(rawRow || [], indexMap, aliases),
        ])
      );

      if (!hasRequiredValue(extracted, requiredAny)) {
        continue;
      }

      let rowData = applyDefaults(extracted, sheetContext);
      rowData = applyDefaults(rowData, perSheetDefaults);
      const uploadRow = finalizeRow(rowData, copyIfMissing, undefinedIfMissing);
      rows.push(uploadRow);
      sourcePersonCount += 1;
      rowCount += 1;

      if (uploadRow.email === "undefined") {
        emailUndefined += 1;
      }
      if (uploadRow.mobile === "undefined") {
        mobileUndefined += 1;
      }
    }

    sheetStats.push({
      sheet_name: sheetName,
      header_row: headerRowIndex + 1,
      source_person_count: sourcePersonCount,
      row_count: rowCount,
      email_undefined: emailUndefined,
      mobile_undefined: mobileUndefined,
    });
  }

  return { rows, sheetStats };
}

function sanitizeDownloadName(fileName) {
  const raw = String(fileName || "result").replace(/\.[^.]+$/, "");
  const safe = raw.replace(/[\\/:*?"<>|]+/g, "_").trim() || "result";
  return `${safe}_lms_일괄등록양식.xlsx`;
}

export function buildSummary(profile, inputFileName, outputFileName, rows, sheetStats) {
  return {
    job_name: profile.id,
    input_file: inputFileName,
    output_file: outputFileName,
    total_rows: rows.length,
    email_undefined: countUndefined(rows, "email"),
    mobile_undefined: countUndefined(rows, "mobile"),
    assumptions: profile.assumptions || [],
    sheet_stats: sheetStats,
  };
}

export function runProfile(profile, workbookData, inputFileName) {
  const sourceMode = profile?.source?.mode;
  if (!profile || !sourceMode) {
    throw new Error("파일 유형을 먼저 선택해 주세요.");
  }

  const courseLookup = normalizeLookup(profile.course_lookup || {});
  let result;
  if (sourceMode === "fixed_columns") {
    result = buildRowsFixedColumns(profile, workbookData, courseLookup);
  } else if (sourceMode === "header_alias") {
    result = buildRowsHeaderAlias(profile, workbookData);
  } else {
    throw new Error(`아직 지원하지 않는 source.mode 입니다: ${sourceMode}`);
  }

  const headers = profile.guide_headers || GUIDE_HEADERS;
  const outputFileName = sanitizeDownloadName(inputFileName);
  const rowMatrix = result.rows.map((row) => OUTPUT_FIELDS.map((field) => row[field] ?? null));
  const summary = buildSummary(profile, inputFileName, outputFileName, result.rows, result.sheetStats);
  return {
    headers,
    rows: result.rows,
    rowMatrix,
    summary,
    outputFileName,
  };
}
