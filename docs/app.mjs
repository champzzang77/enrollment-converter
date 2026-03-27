import { PROFILES } from "./data.mjs";
import { runProfile } from "./engine.mjs";

const STORAGE_KEY = "enrollment-upload-static-usage-log-v1";
const ALLOWED_EXTENSIONS = new Set([".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"]);

const profileSelect = document.getElementById("profileSelect");
const fileInput = document.getElementById("fileInput");
const userNameInput = document.getElementById("userNameInput");
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
const recentUsageBox = document.getElementById("recentUsageBox");
const totalUsageCount = document.getElementById("totalUsageCount");
const todayUsageCount = document.getElementById("todayUsageCount");

let latestDownloadUrl = "";
let recommendedProfileId = "";

function getSelectedProfile() {
  return PROFILES.find((item) => item.id === profileSelect.value) || null;
}

function renderProfiles() {
  PROFILES.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile.id;
    option.textContent = `${profile.label} - ${profile.short_description}`;
    profileSelect.appendChild(option);
  });
  if (!profileSelect.value && PROFILES[0]) {
    profileSelect.value = PROFILES[0].id;
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
    recommendBadge.textContent = "파일명 기준 추천";
    recommendBadge.className = "badge badge-recommend";
  } else {
    recommendBadge.textContent = "직접 선택";
    recommendBadge.className = "badge badge-neutral";
  }
}

function recommendProfileByFileName(fileName) {
  const normalizedName = String(fileName || "").toLowerCase();
  const matched = PROFILES.find((profile) =>
    (profile.filename_keywords || []).some((keyword) =>
      normalizedName.includes(String(keyword).toLowerCase())
    )
  );

  if (!matched) {
    recommendedProfileId = "";
    updateProfileGuide();
    statusBox.innerHTML = "파일을 선택했습니다. <strong>파일 유형</strong>을 확인한 뒤 <strong>변환 시작</strong>을 눌러 주세요.";
    return;
  }

  recommendedProfileId = matched.id;
  profileSelect.value = matched.id;
  updateProfileGuide();
  statusBox.innerHTML = `<strong>${matched.label}</strong> 형식으로 추천했습니다.<br>예시 파일명과 설명이 맞는지만 한 번 확인해 주세요.`;
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
    main.textContent = `${item.user_name || "사용자명 없음"} | ${item.profile_label || "유형 정보 없음"}`;

    const sub = document.createElement("div");
    sub.className = "log-sub";
    if (isSuccess) {
      sub.textContent = `파일: ${item.file_name || "파일명 없음"} | 변환 행 수: ${item.total_rows ?? "-"}`;
    } else {
      sub.textContent = `파일: ${item.file_name || "파일명 없음"} | 오류: ${item.error || "알 수 없는 오류"}`;
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
    rows: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
    }),
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
  if (message.includes("사용자명을 입력")) {
    return "사용 내역을 구분할 수 있도록 사용자명을 먼저 입력해 주세요.";
  }
  if (message.includes("헤더 행을 찾지 못했습니다")) {
    return "선택한 파일 유형과 업로드한 파일 형식이 맞지 않습니다. 다른 유형으로 바꿔 다시 시도해 주세요.\n원본 메시지: " + message;
  }
  if (message.includes("지원하지 않는 엑셀 형식")) {
    return "지원하는 확장자는 .xlsx, .xlsm, .xltx, .xltm, .xls 입니다.";
  }
  if (message.includes("파일 유형을 먼저 선택")) {
    return "파일 유형을 먼저 선택해 주세요.";
  }
  return "변환 중 오류가 발생했습니다.\n원본 메시지: " + message;
}

function saveSuccessLog(profile, fileName, summary) {
  saveUsageEntry({
    timestamp: createTimestamp(),
    user_name: String(userNameInput.value || "").trim(),
    profile_id: profile.id,
    profile_label: profile.label,
    file_name: fileName,
    status: "success",
    total_rows: summary.total_rows,
    email_undefined: summary.email_undefined,
    mobile_undefined: summary.mobile_undefined,
  });
}

function saveErrorLog(profile, fileName, errorMessage) {
  saveUsageEntry({
    timestamp: createTimestamp(),
    user_name: String(userNameInput.value || "").trim(),
    profile_id: profile?.id || "",
    profile_label: profile?.label || "유형 정보 없음",
    file_name: fileName || "",
    status: "error",
    error: errorMessage,
  });
}

function buildSuccessStatus(fileName, summary) {
  const unresolvedCount = (summary.sheet_stats || []).reduce(
    (total, item) => total + ((item.unresolved_course_names || []).length),
    0
  );
  if (unresolvedCount > 0) {
    return `<strong>변환은 완료되었습니다.</strong><br>${fileName} 파일을 처리했지만 과정코드가 연결되지 않은 항목이 ${unresolvedCount}건 있습니다. 요약을 확인한 뒤 결과 파일을 내려받아 주세요.`;
  }
  return `<strong>변환이 완료되었습니다.</strong><br>${fileName} 파일이 준비되었습니다. 아래 버튼으로 결과 파일을 내려받아 주세요.`;
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

  if (!String(userNameInput.value || "").trim()) {
    statusBox.textContent = "사용자명을 먼저 입력해 주세요.";
    errorBox.value = "사용 내역을 구분할 수 있도록 사용자명을 입력해 주세요.";
    userNameInput.focus();
    return;
  }

  convertButton.disabled = true;
  statusBox.innerHTML = "<strong>변환 중입니다.</strong><br>브라우저 안에서 파일을 읽고 결과 파일을 만들고 있습니다.";

  try {
    const workbookData = await extractWorkbookData(file);
    const result = runProfile(profile, workbookData, file.name);
    const blob = createWorkbookDownload(result.headers, result.rowMatrix);
    latestDownloadUrl = URL.createObjectURL(blob);
    downloadLink.href = latestDownloadUrl;
    downloadLink.download = result.outputFileName;
    downloadLink.classList.remove("disabled");
    setSummary(result.summary);
    statusBox.innerHTML = buildSuccessStatus(file.name, result.summary);
    saveSuccessLog(profile, file.name, result.summary);
    refreshUsageView();
  } catch (error) {
    const friendlyMessage = explainError(error);
    statusBox.textContent = "변환에 실패했습니다. 아래 오류 안내를 확인해 주세요.";
    errorBox.value = friendlyMessage;
    saveErrorLog(profile, file?.name, friendlyMessage);
    refreshUsageView();
  } finally {
    convertButton.disabled = false;
  }
}

profileSelect.addEventListener("change", updateProfileGuide);
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (!file) {
    recommendedProfileId = "";
    updateProfileGuide();
    return;
  }
  recommendProfileByFileName(file.name);
});
convertButton.addEventListener("click", convertFile);

renderProfiles();
refreshUsageView();
