import { AFFILIATE_CODE_LOOKUP, GUIDE_HEADERS, OUTPUT_FIELDS } from "./data.mjs";

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

const AFFILIATE_LOOKUP = normalizeLookup(AFFILIATE_CODE_LOOKUP);

function normalizeRuntimeSheetMap(sheetMap = {}) {
  return Object.fromEntries(
    Object.entries(sheetMap).map(([key, value]) => [String(key), text(value)])
  );
}

function normalizeRuntimeGroupCodeMap(groupMap = {}) {
  return Object.fromEntries(
    Object.entries(groupMap).map(([sheetName, groupCodes]) => [
      String(sheetName),
      Object.fromEntries(
        Object.entries(groupCodes || {})
          .map(([groupName, code]) => [normalizeMatchKey(groupName), text(code)])
          .filter(([, code]) => code)
      ),
    ])
  );
}

function normalizeColumnRefs(columnRef) {
  if (Array.isArray(columnRef)) {
    return columnRef.map((item) => columnRefToIndex(item));
  }
  return [columnRefToIndex(columnRef)];
}

function getCellValueByRef(rows, cellRef) {
  if (!cellRef) {
    return null;
  }
  const match = String(cellRef).trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }

  const columnIndex = columnRefToIndex(match[1]);
  const rowIndex = Number(match[2]) - 1;
  if (rowIndex < 0 || rowIndex >= rows.length) {
    return null;
  }
  const row = rows[rowIndex] || [];
  return text(columnIndex < row.length ? row[columnIndex] : null);
}

function getPrimaryColumnIndex(columnIndexes) {
  if (Array.isArray(columnIndexes)) {
    return columnIndexes[0];
  }
  return columnIndexes;
}

function hasRequiredValue(rowData, requiredFields) {
  return requiredFields.some((field) => text(rowData[field]));
}

function valueMatchesRule(fieldValue, matcher) {
  const normalizedValue = text(fieldValue);
  if (matcher && typeof matcher === "object" && !Array.isArray(matcher)) {
    if (matcher.equals !== undefined && normalizedValue !== text(matcher.equals)) {
      return false;
    }
    if (matcher.includes !== undefined) {
      const expected = String(text(matcher.includes) || "").toLowerCase();
      const actual = String(normalizedValue || "").toLowerCase();
      if (!actual.includes(expected)) {
        return false;
      }
    }
    if (Array.isArray(matcher.any_of) && matcher.any_of.length) {
      const matched = matcher.any_of.some((item) => normalizedValue === text(item));
      if (!matched) {
        return false;
      }
    }
    return true;
  }
  return normalizedValue === text(matcher);
}

function shouldSkipExtractedRow(rowData, sourceConfig = {}) {
  const skipRules = Array.isArray(sourceConfig.skip_rows_if) ? sourceConfig.skip_rows_if : [];
  return skipRules.some((rule) =>
    Object.entries(rule).every(([field, matcher]) => valueMatchesRule(rowData[field], matcher))
  );
}

function sheetHasFixedColumnData(rawRows, startRowIndex, columnMap, requiredFields, sourceConfig = {}) {
  for (let rowIndex = startRowIndex; rowIndex < rawRows.length; rowIndex += 1) {
    const rawRow = rawRows[rowIndex] || [];
    const extracted = Object.fromEntries(
      Object.entries(columnMap).map(([field, columnIndexes]) => [
        field,
        columnIndexes
          .map((columnIndex) => text(columnIndex < rawRow.length ? rawRow[columnIndex] : null))
          .find((value) => value) || null,
      ])
    );

    if (!hasRequiredValue(extracted, requiredFields)) {
      continue;
    }

    if (shouldSkipExtractedRow(extracted, sourceConfig)) {
      continue;
    }

    return true;
  }

  return false;
}

function applyCarryForwardContext(rowData, carryForwardContext, sourceConfig = {}) {
  const fields = Array.isArray(sourceConfig.carry_forward_fields)
    ? sourceConfig.carry_forward_fields
    : [];

  if (!fields.length) {
    return {
      rowData,
      carryForwardContext,
    };
  }

  const anchorFields = Array.isArray(sourceConfig.carry_forward_anchor_fields) &&
    sourceConfig.carry_forward_anchor_fields.length
    ? sourceConfig.carry_forward_anchor_fields
    : fields;

  const isAnchorRow = anchorFields.some((field) => text(rowData[field]));
  const nextContext = { ...carryForwardContext };
  const nextRowData = { ...rowData };

  if (isAnchorRow) {
    fields.forEach((field) => {
      nextContext[field] = text(rowData[field]) || null;
    });
    return {
      rowData: nextRowData,
      carryForwardContext: nextContext,
    };
  }

  fields.forEach((field) => {
    if (!text(nextRowData[field]) && text(nextContext[field])) {
      nextRowData[field] = nextContext[field];
    }
  });

  return {
    rowData: nextRowData,
    carryForwardContext: nextContext,
  };
}

function resolveFixedFieldPrefixContext(rawRows, sourceConfig = {}) {
  const prefixRules = sourceConfig.prefix_fields || {};
  const resolved = {};

  Object.entries(prefixRules).forEach(([field, config]) => {
    const prefix = getCellValueByRef(rawRows, config?.from_cell);
    if (prefix) {
      resolved[field] = {
        prefix,
        when_regex: text(config?.when_regex),
      };
    }
  });

  return resolved;
}

function applyFieldPrefixes(rowData, prefixContext = {}) {
  const merged = { ...rowData };

  Object.entries(prefixContext).forEach(([field, config]) => {
    const value = text(merged[field]);
    const prefix = text(config?.prefix);
    if (!value || !prefix) {
      return;
    }

    if (value.startsWith(prefix)) {
      return;
    }

    const pattern = text(config?.when_regex);
    if (pattern) {
      const regex = new RegExp(pattern);
      if (!regex.test(value)) {
        return;
      }
    }

    merged[field] = `${prefix}${value}`;
  });

  return merged;
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

function applyAffiliateCode(rowData) {
  const merged = { ...rowData };
  if (text(merged.affiliate_code)) {
    return merged;
  }

  const companyKey = normalizeMatchKey(merged.company);
  if (!companyKey) {
    return merged;
  }

  const affiliateCode = AFFILIATE_LOOKUP[companyKey];
  if (affiliateCode) {
    merged.affiliate_code = affiliateCode;
  }
  return merged;
}

function formatKoreanPhoneNumber(value) {
  const raw = text(value);
  if (!raw || raw === "undefined") {
    return raw;
  }

  if (String(raw).includes("-")) {
    return raw;
  }

  const digits = String(raw).replace(/\D/g, "");
  if (!digits) {
    return raw;
  }

  if (digits.startsWith("02")) {
    if (digits.length === 9) {
      return `02-${digits.slice(2, 5)}-${digits.slice(5)}`;
    }
    if (digits.length === 10) {
      return `02-${digits.slice(2, 6)}-${digits.slice(6)}`;
    }
    return raw;
  }

  if (digits.length === 10) {
    return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`;
  }

  if (digits.length === 11) {
    return `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`;
  }

  return raw;
}

function finalizeRow(rowData, copyIfMissing = {}, undefinedIfMissing = []) {
  const finalized = applyAffiliateCode({ method: "0", ...rowData });

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

  for (const field of ["phone", "mobile"]) {
    if (text(finalized[field])) {
      finalized[field] = formatKoreanPhoneNumber(finalized[field]);
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

function resolveSeedCourseExpansions(rawRows, sourceConfig, columnMap) {
  const seedConfig = sourceConfig.seed_courses_from_top_rows;
  if (!seedConfig) {
    return [];
  }

  const codeField = seedConfig.code_field || "course_code";
  const nameField = seedConfig.name_field || "course_name";
  const codeIndex = getPrimaryColumnIndex(columnMap[codeField]);
  const nameIndex = getPrimaryColumnIndex(columnMap[nameField]);

  if (codeIndex === undefined && nameIndex === undefined) {
    return [];
  }

  const startRow = Number(seedConfig.start_row || sourceConfig.start_row || 2);
  const maxRows = Number(seedConfig.max_rows || 20);
  const expansions = [];

  for (
    let rowIndex = startRow - 1;
    rowIndex < rawRows.length && rowIndex < startRow - 1 + maxRows;
    rowIndex += 1
  ) {
    const rawRow = rawRows[rowIndex] || [];
    const code = codeIndex === undefined ? null : text(codeIndex < rawRow.length ? rawRow[codeIndex] : null);
    const name = nameIndex === undefined ? null : text(nameIndex < rawRow.length ? rawRow[nameIndex] : null);

    if (code || name) {
      expansions.push({ code, name });
      continue;
    }

    if (expansions.length) {
      break;
    }
  }

  return expansions;
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

function scoreAliasMatch(headerKey, aliasKey) {
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

function pickByAliases(row, indexMap, aliases) {
  let bestValue = null;
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const alias of aliases) {
    const aliasKey = normalizeHeader(alias);
    for (const [headerKey, index] of Object.entries(indexMap)) {
      if (index >= row.length) {
        continue;
      }

      const value = text(row[index]);
      const score = scoreAliasMatch(headerKey, aliasKey) + (value ? 1 : 0);
      if (score > bestScore) {
        bestScore = score;
        bestValue = value;
      }
    }
  }

  return bestValue;
}

function pickHeaderIndexByAliases(indexMap, aliases = []) {
  let bestIndex = null;
  let bestScore = Number.NEGATIVE_INFINITY;

  for (const alias of aliases) {
    const aliasKey = normalizeHeader(alias);
    for (const [headerKey, index] of Object.entries(indexMap)) {
      const score = scoreAliasMatch(headerKey, aliasKey);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    }
  }

  return bestIndex;
}

function buildHeaderAliasRowData(rawRow, indexMap, fieldAliases = {}) {
  return Object.fromEntries(
    Object.entries(fieldAliases).map(([field, aliases]) => [
      field,
      pickByAliases(rawRow || [], indexMap, aliases),
    ])
  );
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
    const candidateColumns = Array.isArray(check.columns) && check.columns.length
      ? check.columns
      : [check.column];
    const includes = normalizeKeywordList(check.includes || []);
    const anyOf = normalizeKeywordList(check.any_of || []);
    const matchedColumn = candidateColumns.find((columnRef) => {
      const columnIndex = columnRefToIndex(columnRef);
      const rawHeader = text(columnIndex < headerRow.length ? headerRow[columnIndex] : null) || "";
      const normalizedHeader = normalizeHeader(rawHeader);
      const includesMatched = includes.every((keyword) => normalizedHeader.includes(keyword));
      const anyOfMatched = !anyOf.length || anyOf.some((keyword) => normalizedHeader.includes(keyword));
      return includesMatched && anyOfMatched;
    });

    if (!matchedColumn) {
      const rawHeaders = candidateColumns.map((columnRef) => {
        const columnIndex = columnRefToIndex(columnRef);
        const rawHeader = text(columnIndex < headerRow.length ? headerRow[columnIndex] : null) || "빈칸";
        return `${columnRef}:${rawHeader}`;
      });
      const expectedParts = [];
      if (includes.length) {
        expectedParts.push(`포함: ${check.includes.join(", ")}`);
      }
      if (anyOf.length) {
        expectedParts.push(`다음 중 하나: ${check.any_of.join(", ")}`);
      }
      mismatchMessages.push(
        `${candidateColumns.join("/")}열 기대값(${expectedParts.join(" / ")}) != 실제값(${rawHeaders.join(", ")})`
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

function buildRowsFixedColumns(profile, workbookData, courseLookup, runtimeOptions = {}) {
  const sourceConfig = profile.source;
  const headers = profile.guide_headers || GUIDE_HEADERS;
  if (headers.length !== GUIDE_HEADERS.length) {
    throw new Error("guide_headers 길이는 기본 업로드 양식과 같아야 합니다.");
  }

  const sheetMap = new Map(workbookData.map((sheet) => [sheet.name, sheet.rows]));
  const sheetNames = workbookData.map((sheet) => sheet.name);
  const columnMap = Object.fromEntries(
    Object.entries(sourceConfig.column_map).map(([field, columnRef]) => [field, normalizeColumnRefs(columnRef)])
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
  const manualSheetCourseCodes = normalizeRuntimeSheetMap(runtimeOptions.manual_sheet_course_codes || {});

  const rows = [];
  const sheetStats = [];
  const missingManualCourseSheets = [];

  for (const sheetName of selectedSheets) {
    const rawRows = sheetMap.get(sheetName) || [];
    try {
      validateFixedColumnLayout(rawRows, sourceConfig, sheetName);
    } catch (error) {
      if (sourceConfig.ignore_invalid_layout) {
        sheetStats.push({
          sheet_name: sheetName,
          skipped_reason: "layout_mismatch",
          row_count: 0,
        });
        continue;
      }
      throw error;
    }
    const { expansions, unresolvedNames } = resolveCourseExpansions(sheetName, sourceConfig, courseLookup);
    const seedExpansions = resolveSeedCourseExpansions(rawRows, sourceConfig, columnMap);
    const effectiveExpansions = seedExpansions.length ? seedExpansions : expansions;
    const hasDataRows = sheetHasFixedColumnData(
      rawRows,
      startRow - 1,
      columnMap,
      requiredAny,
      sourceConfig
    );
    if (!hasDataRows) {
      sheetStats.push({
        sheet_name: sheetName,
        course_names: effectiveExpansions.map((item) => item.name).filter(Boolean),
        course_codes: effectiveExpansions.map((item) => item.code).filter(Boolean),
        unresolved_course_names: unresolvedNames,
        source_person_count: 0,
        row_count: 0,
        email_undefined: 0,
        mobile_undefined: 0,
      });
      continue;
    }
    if (sourceConfig.require_manual_course_codes && !manualSheetCourseCodes[sheetName]) {
      missingManualCourseSheets.push(sheetName);
      continue;
    }
    const perSheetDefaults = {
      ...defaults,
      ...normalizeDefaults(sheetDefaults[sheetName] || {}),
      ...normalizeDefaults(
        manualSheetCourseCodes[sheetName]
          ? { course_code: manualSheetCourseCodes[sheetName] }
          : {}
      ),
    };
    const prefixContext = resolveFixedFieldPrefixContext(rawRows, sourceConfig);

    let sourcePersonCount = 0;
    let rowCount = 0;
    let emailUndefined = 0;
    let mobileUndefined = 0;
    let carryForwardContext = {};

    for (let rowIndex = startRow - 1; rowIndex < rawRows.length; rowIndex += 1) {
      const rawRow = rawRows[rowIndex] || [];
      let extracted = Object.fromEntries(
        Object.entries(columnMap).map(([field, columnIndexes]) => [
          field,
          columnIndexes
            .map((columnIndex) => text(columnIndex < rawRow.length ? rawRow[columnIndex] : null))
            .find((value) => value) || null,
        ])
      );

      const carryForwardResult = applyCarryForwardContext(extracted, carryForwardContext, sourceConfig);
      extracted = carryForwardResult.rowData;
      carryForwardContext = carryForwardResult.carryForwardContext;
      extracted = applyFieldPrefixes(extracted, prefixContext);

      if (!hasRequiredValue(extracted, requiredAny)) {
        continue;
      }

      if (shouldSkipExtractedRow(extracted, sourceConfig)) {
        continue;
      }

      sourcePersonCount += 1;
      const expandedCourses = effectiveExpansions.length ? effectiveExpansions : [{ name: null, code: null }];
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
      course_names: effectiveExpansions.map((item) => item.name).filter(Boolean),
      course_codes: effectiveExpansions.map((item) => item.code).filter(Boolean),
      unresolved_course_names: unresolvedNames,
      source_person_count: sourcePersonCount,
      row_count: rowCount,
      email_undefined: emailUndefined,
      mobile_undefined: mobileUndefined,
    });
  }

  if (missingManualCourseSheets.length) {
    throw new Error(
      `선택한 파일 유형은 시트별 과정코드를 직접 입력해야 합니다.\n누락된 시트: ${missingManualCourseSheets.join(", ")}`
    );
  }

  return { rows, sheetStats };
}

function buildRowsHeaderAlias(profile, workbookData, runtimeOptions = {}) {
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
  const sheetContextLabels = sourceConfig.sheet_context_labels || {};
  const selectedSheets = chooseSheetNames(sheetNames, sourceConfig);
  const manualSheetCourseCodes = normalizeRuntimeSheetMap(runtimeOptions.manual_sheet_course_codes || {});
  const manualGroupCourseCodes = normalizeRuntimeGroupCodeMap(runtimeOptions.manual_group_course_codes || {});
  if (!selectedSheets.length) {
    throw new Error("선택한 파일 유형과 업로드한 파일 구조가 다릅니다. 필요한 시트를 찾지 못했습니다.");
  }

  const rows = [];
  const sheetStats = [];
  const missingManualCourseSheets = [];

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
    const groupAliases = Array.isArray(sourceConfig.manual_course_group_aliases)
      ? sourceConfig.manual_course_group_aliases
      : [];
    const groupCourseLookup = manualGroupCourseCodes[sheetName] || {};
    const hasGroupCourseCodes = Object.keys(groupCourseLookup).length > 0;
    const groupColumnIndex = hasGroupCourseCodes && groupAliases.length
      ? pickHeaderIndexByAliases(indexMap, groupAliases)
      : null;
    const hasDataRows = rawRows.slice(headerRowIndex + 1).some((rawRow) =>
      hasRequiredValue(buildHeaderAliasRowData(rawRow, indexMap, fieldAliases), requiredAny)
    );

    if (!hasDataRows) {
      sheetStats.push({
        sheet_name: sheetName,
        header_row: headerRowIndex + 1,
        source_person_count: 0,
        row_count: 0,
        email_undefined: 0,
        mobile_undefined: 0,
      });
      continue;
    }

    if (hasGroupCourseCodes && groupColumnIndex === null && !manualSheetCourseCodes[sheetName]) {
      throw new Error(
        `선택한 파일 유형과 업로드한 파일 구조가 다릅니다. ${sheetName} 시트에서 구분 열을 찾지 못했습니다.`
      );
    }

    if (
      sourceConfig.require_manual_course_codes &&
      !manualSheetCourseCodes[sheetName] &&
      !hasGroupCourseCodes
    ) {
      missingManualCourseSheets.push(sheetName);
      continue;
    }

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
      ...normalizeDefaults(
        manualSheetCourseCodes[sheetName]
          ? { course_code: manualSheetCourseCodes[sheetName] }
          : {}
      ),
    };

    let sourcePersonCount = 0;
    let rowCount = 0;
    let emailUndefined = 0;
    let mobileUndefined = 0;
    const appliedCourseCodes = new Set();
    const missingManualCourseGroups = new Set();

    for (const rawRow of rawRows.slice(headerRowIndex + 1)) {
      const extracted = buildHeaderAliasRowData(rawRow, indexMap, fieldAliases);

      if (!hasRequiredValue(extracted, requiredAny)) {
        continue;
      }

      let rowData = applyDefaults(extracted, sheetContext);
      const groupValue = groupColumnIndex === null || groupColumnIndex === undefined
        ? null
        : text(groupColumnIndex < rawRow.length ? rawRow[groupColumnIndex] : null);
      const groupCode = hasGroupCourseCodes
        ? groupCourseLookup[normalizeMatchKey(groupValue)] || null
        : null;
      if (!text(rowData.course_code) && groupCode) {
        rowData.course_code = groupCode;
      }
      rowData = applyDefaults(rowData, perSheetDefaults);
      if (sourceConfig.require_manual_course_codes && !text(rowData.course_code)) {
        missingManualCourseGroups.add(groupValue || "(빈 값)");
        continue;
      }
      const uploadRow = finalizeRow(rowData, copyIfMissing, undefinedIfMissing);
      rows.push(uploadRow);
      sourcePersonCount += 1;
      rowCount += 1;
      if (text(uploadRow.course_code)) {
        appliedCourseCodes.add(uploadRow.course_code);
      }

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
      course_codes: appliedCourseCodes.size
        ? [...appliedCourseCodes]
        : (manualSheetCourseCodes[sheetName] ? [manualSheetCourseCodes[sheetName]] : []),
      source_person_count: sourcePersonCount,
      row_count: rowCount,
      email_undefined: emailUndefined,
      mobile_undefined: mobileUndefined,
    });

    if (missingManualCourseGroups.size) {
      throw new Error(
        `선택한 파일 유형은 구분값별 과정코드를 직접 입력해야 합니다.\n${sheetName} 시트 누락 구분: ${[...missingManualCourseGroups].join(", ")}`
      );
    }
  }

  if (missingManualCourseSheets.length) {
    throw new Error(
      `선택한 파일 유형은 시트별 과정코드를 직접 입력해야 합니다.\n누락된 시트: ${missingManualCourseSheets.join(", ")}`
    );
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

export function runProfile(profile, workbookData, inputFileName, runtimeOptions = {}) {
  const sourceMode = profile?.source?.mode;
  if (!profile || !sourceMode) {
    throw new Error("파일 유형을 먼저 선택해 주세요.");
  }

  const courseLookup = normalizeLookup(profile.course_lookup || {});
  let result;
  if (sourceMode === "fixed_columns") {
    result = buildRowsFixedColumns(profile, workbookData, courseLookup, runtimeOptions);
  } else if (sourceMode === "header_alias") {
    result = buildRowsHeaderAlias(profile, workbookData, runtimeOptions);
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
