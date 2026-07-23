const STORAGE_KEY = "specificDayWorkManager_v1";
const FLOOR_UNIT = 30;
const MAX_DAILY_USAGE_MINUTES = 240;
const USAGE_HISTORY_VISIBLE_DAYS = 3;
const LUNCH_START_MINUTES = 12 * 60;
const LUNCH_END_MINUTES = 13 * 60;
const WORK_TYPES = {
  A: { label: "A형", start: "07:00", end: "16:00" },
  "A-1": { label: "A-1형", start: "07:30", end: "16:30" },
  B: { label: "B형", start: "08:00", end: "17:00" },
  "B-1": { label: "B-1형", start: "08:30", end: "17:30" },
  C: { label: "C형", start: "09:00", end: "18:00" },
  "C-1": { label: "C-1형", start: "09:30", end: "18:30" },
  D: { label: "D형", start: "10:00", end: "19:00" }
};
const EMPTY_TEMPLATE = document.getElementById("emptyStateTemplate");
let state = { attendanceRecords: [], usageRecords: [] };
let currentUser = null; // { email, name, picture, idToken }
let appInitialized = false;
let syncInFlight = null;
let syncQueued = false;
const elements = {
  authGate: document.getElementById("authGate"),
  googleSignInButton: document.getElementById("googleSignInButton"),
  authStatusText: document.getElementById("authStatusText"),
  localModeBanner: document.getElementById("localModeBanner"),
  appShell: document.getElementById("appShell"),
  userBadge: document.getElementById("userBadge"),
  userAvatar: document.getElementById("userAvatar"),
  userName: document.getElementById("userName"),
  syncStatus: document.getElementById("syncStatus"),
  refreshDataBtn: document.getElementById("refreshDataBtn"),
  signOutBtn: document.getElementById("signOutBtn"),
  totalRemainingLabel: document.getElementById("totalRemainingLabel"),
  validEntryCountLabel: document.getElementById("validEntryCountLabel"),
  leaveTimeLabel: document.getElementById("leaveTimeLabel"),
  leaveTimeMeta: document.getElementById("leaveTimeMeta"),
  todayWorkType: document.getElementById("todayWorkType"),
  attendanceForm: document.getElementById("attendanceForm"),
  attendanceId: document.getElementById("attendanceId"),
  attendanceDate: document.getElementById("attendanceDate"),
  attendanceWorkType: document.getElementById("attendanceWorkType"),
  actualStart: document.getElementById("actualStart"),
  actualEnd: document.getElementById("actualEnd"),
  overtimeChecked: document.getElementById("overtimeChecked"),
  computedEarnedLabel: document.getElementById("computedEarnedLabel"),
  attendanceFormMode: document.getElementById("attendanceFormMode"),
  cancelAttendanceEditBtn: document.getElementById("cancelAttendanceEditBtn"),
  usageForm: document.getElementById("usageForm"),
  usageId: document.getElementById("usageId"),
  usageDate: document.getElementById("usageDate"),
  usageAttendanceDate1: document.getElementById("usageAttendanceDate1"),
  usageAttendanceDate2: document.getElementById("usageAttendanceDate2"),
  usageAttendanceDate3: document.getElementById("usageAttendanceDate3"),
  usageAttendanceDate4: document.getElementById("usageAttendanceDate4"),
  selectedAttendanceSummary: document.getElementById("selectedAttendanceSummary"),
  selectedAttendanceTotalLabel: document.getElementById("selectedAttendanceTotalLabel"),
  usageWorkType: document.getElementById("usageWorkType"),
  usageStartHour: document.getElementById("usageStartHour"),
  usageStartMinute: document.getElementById("usageStartMinute"),
  usageStart: document.getElementById("usageStart"),
  usageEndHour: document.getElementById("usageEndHour"),
  usageEndMinute: document.getElementById("usageEndMinute"),
  usageEnd: document.getElementById("usageEnd"),
  computedUsageLabel: document.getElementById("computedUsageLabel"),
  usageFormMode: document.getElementById("usageFormMode"),
  cancelUsageEditBtn: document.getElementById("cancelUsageEditBtn"),
  attendanceList: document.getElementById("attendanceList"),
  timelineList: document.getElementById("timelineList"),
  attendanceCount: document.getElementById("attendanceCount"),
  excelFile: document.getElementById("excelFile"),
  uploadBtn: document.getElementById("uploadBtn"),
  uploadResult: document.getElementById("uploadResult"),
  seedDemoBtn: document.getElementById("seedDemoBtn"),
  resetBtn: document.getElementById("resetBtn"),
  matchingSummaryBtn: document.getElementById("matchingSummaryBtn"),
  matchingSummaryDialog: document.getElementById("matchingSummaryDialog"),
  matchingSummaryContent: document.getElementById("matchingSummaryContent"),
  closeMatchingSummaryBtn: document.getElementById("closeMatchingSummaryBtn"),
  helpBtn: document.getElementById("helpBtn"),
  helpDialog: document.getElementById("helpDialog"),
  closeHelpBtn: document.getElementById("closeHelpBtn")
};
boot();
function isServerModeEnabled() {
  return Boolean(window.APP_CONFIG && window.APP_CONFIG.GOOGLE_CLIENT_ID && window.APP_CONFIG.API_BASE_URL);
}
async function boot() {
  bindStaticEvents();
  if (!isServerModeEnabled()) return enterLocalOnlyMode();
  elements.authGate.hidden = false;
  elements.authStatusText.textContent = "로그인 준비 중...";
  try {
    const idLib = await waitForGoogleIdentity();
    idLib.initialize({ client_id: window.APP_CONFIG.GOOGLE_CLIENT_ID, callback: handleGoogleCredential });
    idLib.renderButton(elements.googleSignInButton, { theme: "outline", size: "large", text: "signin_with", shape: "pill", locale: "ko" });
    idLib.prompt();
    elements.authStatusText.textContent = "회사 구글 계정으로 로그인하세요.";
  } catch (error) {
    elements.authStatusText.textContent = error.message;
  }
}
function enterLocalOnlyMode() {
  state = loadState();
  elements.localModeBanner.hidden = false;
  elements.appShell.hidden = false;
  initialize();
}
function bindStaticEvents() {
  elements.signOutBtn.addEventListener("click", signOut);
  elements.refreshDataBtn.addEventListener("click", refreshFromServer);
}
function waitForGoogleIdentity(retries = 50) {
  return new Promise((resolve, reject) => {
    const check = (remaining) => {
      if (window.google?.accounts?.id) return resolve(window.google.accounts.id);
      if (remaining <= 0) return reject(new Error("Google 로그인 라이브러리를 불러오지 못했습니다. 새로고침 후 다시 시도하세요."));
      setTimeout(() => check(remaining - 1), 100);
    };
    check(retries);
  });
}
function decodeJwtPayload(token) {
  const base64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
  const json = decodeURIComponent(atob(base64).split("").map((c) => "%" + c.charCodeAt(0).toString(16).padStart(2, "0")).join(""));
  return JSON.parse(json);
}
async function handleGoogleCredential(response) {
  const idToken = response.credential;
  const payload = decodeJwtPayload(idToken);
  currentUser = { email: payload.email, name: payload.name || payload.email, picture: payload.picture || "", idToken };
  elements.authStatusText.textContent = "로그인 확인 중...";
  try {
    state = await apiLoad(idToken);
  } catch (error) {
    currentUser = null;
    elements.authStatusText.textContent = `로그인 실패: ${error.message}`;
    return;
  }
  await maybeImportLegacyLocalData();
  elements.authGate.hidden = true;
  elements.appShell.hidden = false;
  elements.userBadge.hidden = false;
  elements.userName.textContent = currentUser.name;
  elements.userAvatar.src = currentUser.picture;
  setSyncStatus("", "동기화 완료");
  initialize();
}
async function maybeImportLegacyLocalData() {
  const hasServerData = state.attendanceRecords.length > 0 || state.usageRecords.length > 0;
  if (hasServerData) return;
  const legacy = loadState();
  const hasLegacyData = legacy.attendanceRecords.length > 0 || legacy.usageRecords.length > 0;
  if (!hasLegacyData) return;
  const shouldImport = confirm(
    `이 브라우저에 저장되어 있던 기존 데이터(발생기록 ${legacy.attendanceRecords.length}건, 사용기록 ${legacy.usageRecords.length}건)를 서버로 가져올까요?`
  );
  if (!shouldImport) return;
  setSyncStatus("syncing", "기존 데이터 업로드 중...");
  try {
    await apiSave(currentUser.idToken, legacy.attendanceRecords, legacy.usageRecords);
    state = legacy;
    localStorage.removeItem(STORAGE_KEY);
    setSyncStatus("", "동기화 완료");
  } catch (error) {
    setSyncStatus("error", `가져오기 실패: ${error.message}`);
    alert(`기존 데이터를 서버로 옮기지 못했습니다: ${error.message}\n이 브라우저의 기존 데이터는 삭제하지 않았으니 다음 로그인 때 다시 시도할 수 있습니다.`);
  }
}
function signOut() {
  if (window.google?.accounts?.id) window.google.accounts.id.disableAutoSelect();
  currentUser = null;
  state = { attendanceRecords: [], usageRecords: [] };
  elements.appShell.hidden = true;
  elements.userBadge.hidden = true;
  elements.authGate.hidden = false;
  elements.authStatusText.textContent = "다시 로그인해주세요.";
}
async function refreshFromServer() {
  if (!currentUser) return;
  setSyncStatus("syncing", "불러오는 중...");
  try {
    state = await apiLoad(currentUser.idToken);
    renderAll();
    setSyncStatus("", "동기화 완료");
  } catch (error) {
    setSyncStatus("error", `불러오기 실패: ${error.message}`);
  }
}
function setSyncStatus(className, text) {
  if (!elements.syncStatus) return;
  elements.syncStatus.textContent = text;
  elements.syncStatus.className = `sync-status ${className}`.trim();
}
async function apiLoad(idToken) {
  const url = `${window.APP_CONFIG.API_BASE_URL}?action=load&idToken=${encodeURIComponent(idToken)}`;
  const response = await fetch(url, { method: "GET" });
  const data = await response.json();
  if (data.error) throw new Error(data.error);
  return { attendanceRecords: Array.isArray(data.attendanceRecords) ? data.attendanceRecords : [], usageRecords: Array.isArray(data.usageRecords) ? data.usageRecords : [] };
}
async function apiSave(idToken, attendanceRecords, usageRecords) {
  const response = await fetch(window.APP_CONFIG.API_BASE_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "save", idToken, attendanceRecords, usageRecords })
  });
  const data = await response.json();
  if (data.error) throw new Error(data.error);
  return data;
}
function syncStateToServer() {
  if (!currentUser) return;
  if (syncInFlight) { syncQueued = true; return; }
  setSyncStatus("syncing", "동기화 중...");
  syncInFlight = apiSave(currentUser.idToken, state.attendanceRecords, state.usageRecords)
    .then(() => setSyncStatus("", "동기화 완료"))
    .catch((error) => setSyncStatus("error", `동기화 실패: ${error.message}`))
    .finally(() => {
      syncInFlight = null;
      if (syncQueued) { syncQueued = false; syncStateToServer(); }
    });
}
function initialize() {
  populateWorkTypeSelects();
  populateUsageTimeSelects();
  populateUsageAttendanceOptions();
  elements.usageStart.step = "1800";
  elements.usageEnd.step = "1800";
  if (!appInitialized) {
    bindEvents();
    appInitialized = true;
  }
  resetForms();
  renderAll();
}
function populateWorkTypeSelects() {
  const optionsHtml = Object.entries(WORK_TYPES).map(([key, config]) => `<option value="${key}">${config.label} (${config.start} ~ ${config.end})</option>`).join("");
  elements.attendanceWorkType.innerHTML = optionsHtml;
  elements.usageWorkType.innerHTML = optionsHtml;
  elements.todayWorkType.innerHTML = optionsHtml;
  elements.attendanceWorkType.value = "C";
  elements.usageWorkType.value = "C";
  elements.todayWorkType.value = "C";
}
function populateUsageTimeSelects() {
  const hourOptions = Array.from({ length: 24 }, (_, hour) => `<option value="${String(hour).padStart(2, "0")}">${String(hour).padStart(2, "0")}시</option>`).join("");
  elements.usageStartHour.innerHTML = hourOptions;
  elements.usageEndHour.innerHTML = hourOptions;
  elements.usageStartMinute.value = "00";
  elements.usageEndMinute.value = "00";
}
function populateUsageAttendanceOptions(selectedIds = getSelectedAttendanceIdsFromForm()) {
  const usageDate = elements.usageDate.value;
  const options = ['<option value="">자동 선택(FIFO)</option>'];
  getSelectableAttendanceEntries(usageDate).forEach((entry) => {
    options.push(`<option value="${entry.id}">${formatSelectableAttendanceLabel(entry)}</option>`);
  });
  getUsageAttendanceSelectElements().forEach((select, index) => {
    select.innerHTML = options.join("");
    const selectedId = selectedIds[index] || "";
    const selectedOptionExists = Array.from(select.options).some((option) => option.value === selectedId);
    select.value = selectedOptionExists ? selectedId : "";
  });
  renderSelectedAttendanceSummary();
}
function bindEvents() {
  elements.attendanceForm.addEventListener("submit", handleAttendanceSubmit);
  elements.usageForm.addEventListener("submit", handleUsageSubmit);
  elements.cancelAttendanceEditBtn.addEventListener("click", resetAttendanceForm);
  elements.cancelUsageEditBtn.addEventListener("click", resetUsageForm);
  elements.attendanceDate.addEventListener("input", renderAttendancePreview);
  elements.attendanceWorkType.addEventListener("change", renderAttendancePreview);
  elements.actualStart.addEventListener("input", renderAttendancePreview);
  elements.actualEnd.addEventListener("input", renderAttendancePreview);
  elements.overtimeChecked.addEventListener("change", renderAttendancePreview);
  elements.usageStartHour.addEventListener("change", () => handleUsageTimeInput("start"));
  elements.usageStartMinute.addEventListener("change", () => handleUsageTimeInput("start"));
  elements.usageEndHour.addEventListener("change", () => handleUsageTimeInput("end"));
  elements.usageEndMinute.addEventListener("change", () => handleUsageTimeInput("end"));
  elements.usageDate.addEventListener("input", () => populateUsageAttendanceOptions());
  getUsageAttendanceSelectElements().forEach((select) => select.addEventListener("change", () => {
    normalizeUsageAttendanceSelects(select);
    renderSelectedAttendanceSummary();
  }));
  elements.todayWorkType.addEventListener("change", () => {
    elements.usageWorkType.value = elements.todayWorkType.value;
    renderSummary();
    renderUsagePreview();
  });
  elements.uploadBtn.addEventListener("click", () => handleExcelUpload());
  elements.seedDemoBtn.addEventListener("click", seedDemoData);
  elements.resetBtn.addEventListener("click", resetAllData);
  elements.matchingSummaryBtn.addEventListener("click", openMatchingSummaryDialog);
  elements.closeMatchingSummaryBtn.addEventListener("click", closeMatchingSummaryDialog);
  elements.matchingSummaryDialog.addEventListener("click", (event) => {
    if (event.target === elements.matchingSummaryDialog) closeMatchingSummaryDialog();
  });
  elements.helpBtn.addEventListener("click", openHelpDialog);
  elements.closeHelpBtn.addEventListener("click", closeHelpDialog);
  elements.helpDialog.addEventListener("click", (event) => {
    if (event.target === elements.helpDialog) closeHelpDialog();
  });
}
function openHelpDialog() {
  if (typeof elements.helpDialog.showModal === "function") {
    elements.helpDialog.showModal();
  } else {
    elements.helpDialog.setAttribute("open", "");
  }
}
function closeHelpDialog() {
  if (typeof elements.helpDialog.close === "function") {
    elements.helpDialog.close();
  } else {
    elements.helpDialog.removeAttribute("open");
  }
}
function handleUsageTimeInput(target) {
  if (target === "start") syncUsageTimeField("start");
  if (target === "end") syncUsageTimeField("end");
  renderUsagePreview();
}
function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return { attendanceRecords: Array.isArray(parsed.attendanceRecords) ? parsed.attendanceRecords : [], usageRecords: Array.isArray(parsed.usageRecords) ? parsed.usageRecords : [] };
  } catch (error) {
    console.error("저장 데이터 로드 실패", error);
    return { attendanceRecords: [], usageRecords: [] };
  }
}
function saveState() {
  if (!isServerModeEnabled()) return localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  syncStateToServer();
}
function saveAttendanceRecords() { saveState(); }
function rerenderAll() { renderAll(); }
function resetForms() { resetAttendanceForm(); resetUsageForm(); }
function resetAttendanceForm() {
  elements.attendanceForm.reset();
  elements.attendanceId.value = "";
  elements.attendanceWorkType.value = "C";
  elements.overtimeChecked.checked = false;
  elements.attendanceFormMode.textContent = "신규 등록";
  elements.attendanceFormMode.className = "pill neutral";
  elements.computedEarnedLabel.textContent = "0분";
}
function resetUsageForm() {
  elements.usageForm.reset();
  elements.usageId.value = "";
  elements.usageWorkType.value = "C";
  setUsageTimeControl("start", "");
  setUsageTimeControl("end", "");
  populateUsageAttendanceOptions([]);
  elements.usageFormMode.textContent = "신규 등록";
  elements.usageFormMode.className = "pill neutral";
  elements.computedUsageLabel.textContent = "0분";
  elements.selectedAttendanceTotalLabel.textContent = "0분";
  elements.selectedAttendanceSummary.textContent = "자동 선택(FIFO)";
}
function syncUsageTimeField(target) {
  const hourElement = target === "start" ? elements.usageStartHour : elements.usageEndHour;
  const minuteElement = target === "start" ? elements.usageStartMinute : elements.usageEndMinute;
  const inputElement = target === "start" ? elements.usageStart : elements.usageEnd;
  const rawValue = `${hourElement.value}:${minuteElement.value}`;
  inputElement.value = normalizeUsageTimeInput(rawValue);
  const [normalizedHour = hourElement.value, normalizedMinute = minuteElement.value] = (inputElement.value || rawValue).split(":");
  hourElement.value = normalizedHour;
  minuteElement.value = normalizedMinute;
}
function setUsageTimeControl(target, timeString) {
  const hourElement = target === "start" ? elements.usageStartHour : elements.usageEndHour;
  const minuteElement = target === "start" ? elements.usageStartMinute : elements.usageEndMinute;
  const inputElement = target === "start" ? elements.usageStart : elements.usageEnd;
  if (!timeString) {
    hourElement.value = "09";
    minuteElement.value = "00";
    inputElement.value = "";
    return;
  }
  const normalized = normalizeUsageTimeInput(timeString);
  const [hour = "09", minute = "00"] = normalized.split(":");
  hourElement.value = hour;
  minuteElement.value = minute;
  inputElement.value = normalized;
}
function renderAll() {
  populateUsageAttendanceOptions();
  renderAttendancePreview();
  renderUsagePreview();
  renderSummary();
  renderAttendanceList();
  renderTimelineList();
}
function getSelectableAttendanceEntries(usageDate = "") {
  const excludeUsageId = elements.usageId?.value || "";
  const referenceDate = usageDate || getTodayString();
  return buildLedger({ excludeUsageId }).generatedEntries
    .filter((entry) => entry.earnedMinutes > 0 && entry.remainingMinutes > 0)
    .filter((entry) => compareDate(entry.date, referenceDate) <= 0)
    .filter((entry) => !isEntryExpired(entry, referenceDate))
    .sort(sortByDateThenId);
}
function formatSelectableAttendanceLabel(entry) {
  const schedule = WORK_TYPES[resolveWorkTypeKey(entry.workType)];
  return `${entry.date} · ${schedule.label} · ${entry.segmentLabel} · 사용 ${formatDuration(entry.usedMinutes)} / 잔여 ${formatDuration(entry.remainingMinutes)}`;
}
function getSelectedAttendanceLabel(selectedAttendanceIds) {
  const selectedIds = normalizeAttendanceIds(selectedAttendanceIds, Infinity);
  if (!selectedIds.length) return "자동 선택(FIFO)";
  const ledger = buildLedger();
  return resolveSelectedAttendanceEntries(selectedIds, ledger.generatedEntries).map((entryOrToken) => {
    if (typeof entryOrToken === "string") return `${entryOrToken} · 삭제된 기록`;
    return formatSelectableAttendanceLabel(entryOrToken);
  }).join(" / ");
}
function handleAttendanceSubmit(event) {
  event.preventDefault();
  const record = { id: elements.attendanceId.value || createId("attendance"), date: elements.attendanceDate.value, workType: elements.attendanceWorkType.value, actualStart: elements.actualStart.value, actualEnd: elements.actualEnd.value, overtime: elements.overtimeChecked.checked, overtimeChecked: elements.overtimeChecked.checked, source: "manual" };
  const validationMessage = validateAttendanceRecord(record);
  if (validationMessage) return alert(validationMessage);
  if (isWeekendDate(record.date)) return alert("주말 근무는 계산에서 제외되므로 발생기록으로 저장할 수 없습니다.");
  const generatedMinutes = calculateGeneratedMinutes(record);
  if (generatedMinutes <= 0) return alert("발생시간이 0분인 기록은 저장할 수 없습니다.");
  const normalizedRecord = { ...record, generatedMinutes };
  const existingIndex = state.attendanceRecords.findIndex((item) => item.id === record.id);
  const sameDateIndex = state.attendanceRecords.findIndex((item) => item.date === record.date && item.id !== record.id);
  if (sameDateIndex >= 0) {
    const shouldReplace = confirm("같은 날짜의 발생기록이 이미 있습니다. 기존 기록을 이 내용으로 업데이트할까요?");
    if (!shouldReplace) return;
    normalizedRecord.id = state.attendanceRecords[sameDateIndex].id;
    state.attendanceRecords.splice(sameDateIndex, 1, normalizedRecord);
    if (existingIndex >= 0 && existingIndex !== sameDateIndex) state.attendanceRecords.splice(existingIndex > sameDateIndex ? existingIndex : existingIndex + 1, 1);
  } else if (existingIndex >= 0) {
    state.attendanceRecords.splice(existingIndex, 1, normalizedRecord);
  } else {
    state.attendanceRecords.push(normalizedRecord);
  }
  refreshAutoUsageSelections();
  saveAttendanceRecords();
  resetAttendanceForm();
  populateUsageAttendanceOptions();
  rerenderAll();
}
function handleUsageSubmit(event) {
  event.preventDefault();
  syncUsageTimeField("start");
  syncUsageTimeField("end");
  const normalizedStartTime = normalizeUsageTimeInput(elements.usageStart.value);
  const normalizedEndTime = normalizeUsageTimeInput(elements.usageEnd.value);
  setUsageTimeControl("start", normalizedStartTime);
  setUsageTimeControl("end", normalizedEndTime);
  const selectedAttendanceIds = getSelectedAttendanceIdsFromForm();
  const record = { id: elements.usageId.value || createId("usage"), date: elements.usageDate.value, workType: elements.usageWorkType.value, startTime: normalizedStartTime, endTime: normalizedEndTime, selectedAttendanceIds };
  const validationMessage = validateUsageRecord(record);
  if (validationMessage) return alert(validationMessage);
  if (isWeekendDate(record.date)) return alert("주말 사용기록은 계산에서 제외되므로 저장할 수 없습니다.");
  const durationMinutes = calculateUsageMinutes(record.startTime, record.endTime);
  if (durationMinutes <= 0) return alert("사용시간이 0분입니다. 30분 단위 기준으로 다시 입력하세요.");
  const simulation = buildLedger({ usageRecordOverride: { ...record, durationMinutes } });
  if (simulation.invalidUsageIds.includes(record.id)) return alert("사용 가능한 특정일 시간이 부족합니다. 유효한 발생시간 범위를 확인하세요.");
  const lockedAttendanceIds = selectedAttendanceIds.length
    ? []
    : getAllocationAttendanceIds(simulation.usageAllocations[record.id] || []);
  const existingIndex = state.usageRecords.findIndex((item) => item.id === record.id);
  if (existingIndex >= 0) state.usageRecords.splice(existingIndex, 1, { ...record, durationMinutes, lockedAttendanceIds });
  else state.usageRecords.push({ ...record, durationMinutes, lockedAttendanceIds });
  refreshAutoUsageSelections();
  saveState();
  resetUsageForm();
  populateUsageAttendanceOptions();
  renderAll();
}
function validateAttendanceRecord(record) {
  if (!record.date || !record.workType || !record.actualStart || !record.actualEnd) return "발생기록의 모든 필드를 입력하세요.";
  if (toMinutes(record.actualEnd) <= toMinutes(record.actualStart)) return "퇴근시간은 출근시간보다 늦어야 합니다.";
  return "";
}
function validateUsageRecord(record) {
  if (!record.date || !record.workType || !record.startTime || !record.endTime) return "사용기록의 모든 필드를 입력하세요.";
  if (toMinutes(record.endTime) <= toMinutes(record.startTime)) return "사용 종료시간은 시작시간보다 늦어야 합니다.";
  const selectedAttendanceIds = normalizeSelectedAttendanceIds(record.selectedAttendanceIds ?? record.selectedAttendanceDates);
  if (selectedAttendanceIds.length > 4) return "차감할 발생기록은 최대 4건까지 선택할 수 있습니다.";
  if (new Set(selectedAttendanceIds).size !== selectedAttendanceIds.length) return "같은 발생기록을 중복 선택할 수 없습니다.";
  const durationMinutes = calculateUsageMinutes(record.startTime, record.endTime);
  if (durationMinutes <= 0) return "사용시간이 0분입니다. 30분 단위 기준으로 다시 입력하세요.";
  const dailyUsageMinutes = getDailyUsageMinutes(record.date, record.id);
  if (dailyUsageMinutes + durationMinutes > MAX_DAILY_USAGE_MINUTES) {
    return `하루 최대 사용 가능 시간은 ${formatDuration(MAX_DAILY_USAGE_MINUTES)}입니다. 같은 날짜의 총 사용시간을 확인하세요.`;
  }
  const resolvedEntries = resolveSelectedAttendanceEntries(selectedAttendanceIds, buildLedger().generatedEntries);
  for (const resolvedEntry of resolvedEntries) {
    if (typeof resolvedEntry === "string") return "선택한 발생기록을 찾을 수 없습니다.";
    if (compareDate(resolvedEntry.date, record.date) > 0) return "사용일보다 미래의 발생기록은 선택할 수 없습니다.";
  }
  return "";
}
function getDailyUsageMinutes(targetDate, excludeUsageId = "") {
  return state.usageRecords
    .filter((usageRecord) => usageRecord.date === targetDate && usageRecord.id !== excludeUsageId && !isWeekendDate(usageRecord.date))
    .reduce((sum, usageRecord) => sum + (usageRecord.durationMinutes ?? calculateUsageMinutes(usageRecord.startTime, usageRecord.endTime)), 0);
}
function renderAttendancePreview() {
  if (!elements.attendanceWorkType.value || !elements.actualStart.value || !elements.actualEnd.value) return elements.computedEarnedLabel.textContent = "0분";
  const minutes = calculateEarnedMinutes({ date: elements.attendanceDate.value, workType: elements.attendanceWorkType.value, actualStart: elements.actualStart.value, actualEnd: elements.actualEnd.value, overtime: elements.overtimeChecked.checked, overtimeChecked: elements.overtimeChecked.checked });
  elements.computedEarnedLabel.textContent = formatDuration(minutes);
}
function renderUsagePreview() {
  syncUsageTimeField("start");
  syncUsageTimeField("end");
  if (!elements.usageStart.value || !elements.usageEnd.value) return elements.computedUsageLabel.textContent = "0분";
  const duration = calculateUsageMinutes(elements.usageStart.value, elements.usageEnd.value);
  elements.computedUsageLabel.textContent = duration > 0 ? formatDuration(duration) : "0분";
  renderSelectedAttendanceSummary();
}
function calculateEarnedMinutes(record) {
  if (isWeekendDate(record.date)) return 0;
  const schedule = WORK_TYPES[resolveWorkTypeKey(record.workType)];
  if (!schedule) return 0;
  if (!record.actualStart || !record.actualEnd) return 0;
  const scheduledStart = toMinutes(schedule.start);
  const scheduledEnd = toMinutes(schedule.end);
  const actualStart = toMinutes(record.actualStart);
  const actualEnd = toMinutes(record.actualEnd);
  if (actualEnd <= actualStart) return 0;
  const earlyEarned = floorToUnit(Math.max(0, scheduledStart - actualStart), FLOOR_UNIT);
  const lateBase = Boolean(record.overtime ?? record.overtimeChecked) ? scheduledEnd + 150 : scheduledEnd;
  const lateEarned = floorToUnit(Math.max(0, actualEnd - lateBase), FLOOR_UNIT);
  return earlyEarned + lateEarned;
}
function calculateGeneratedMinutes(record) { return calculateEarnedMinutes(record); }
function buildGeneratedTimeRanges(record) {
  if (isWeekendDate(record.date)) return [];
  const schedule = WORK_TYPES[resolveWorkTypeKey(record.workType)];
  if (!schedule) return [];
  if (!record.actualStart || !record.actualEnd) return [];
  const scheduledStart = toMinutes(schedule.start);
  const scheduledEnd = toMinutes(schedule.end);
  const actualStart = toMinutes(record.actualStart);
  const actualEnd = toMinutes(record.actualEnd);
  if (actualEnd <= actualStart) return [];
  const ranges = [];
  const earlyEarned = floorToUnit(Math.max(0, scheduledStart - actualStart), FLOOR_UNIT);
  if (earlyEarned > 0) ranges.push({ start: scheduledStart - earlyEarned, end: scheduledStart, minutes: earlyEarned, segmentType: "early", segmentLabel: "출근 전" });
  const lateBase = Boolean(record.overtime ?? record.overtimeChecked) ? scheduledEnd + 150 : scheduledEnd;
  const lateEarned = floorToUnit(Math.max(0, actualEnd - lateBase), FLOOR_UNIT);
  if (lateEarned > 0) ranges.push({ start: lateBase, end: lateBase + lateEarned, minutes: lateEarned, segmentType: "late", segmentLabel: "퇴근 후" });
  return ranges;
}
function calculateUsageMinutes(startTime, endTime) {
  const normalized = normalizeRangeToHalfHour(startTime, endTime);
  return normalized ? calculateUsageMinutesExcludingLunch(normalized.start, normalized.end) : 0;
}
function buildLedger(options = {}) {
  const attendanceRecords = cloneRecords(options.attendanceRecordOverride ? upsertClone(state.attendanceRecords, options.attendanceRecordOverride) : state.attendanceRecords)
    .filter((record) => !isWeekendDate(record.date))
    .sort(sortByDateThenId);
  const sourceUsageRecords = options.usageRecordOverride ? upsertClone(state.usageRecords, options.usageRecordOverride) : state.usageRecords;
  const filteredUsageRecords = options.excludeUsageId ? sourceUsageRecords.filter((item) => item.id !== options.excludeUsageId) : sourceUsageRecords;
  const usageRecords = cloneRecords(filteredUsageRecords)
    .filter((item) => !isWeekendDate(item.date))
    .map((item) => {
      const normalizedSelection = normalizeSelectedAttendanceIds(item.selectedAttendanceIds ?? item.selectedAttendanceDates ?? item.selectedAttendanceDate);
      return {
        ...item,
        selectedAttendanceIds: normalizedSelection,
        lockedAttendanceIds: normalizeAttendanceIds(item.lockedAttendanceIds, Infinity),
        startTime: normalizeUsageTimeInput(item.startTime),
        endTime: normalizeUsageTimeInput(item.endTime),
        durationMinutes: item.durationMinutes ?? calculateUsageMinutes(item.startTime, item.endTime)
      };
    })
    .sort(sortByDateThenId);
  const generatedEntries = attendanceRecords.flatMap((record) => {
    const generatedRanges = buildGeneratedTimeRanges(record);
    if (!generatedRanges.length) {
      return [{
        ...record,
        id: `${record.id}__empty`,
        sourceAttendanceId: record.id,
        earnedMinutes: 0,
        segmentType: "empty",
        segmentLabel: "발생 없음",
        generatedRanges: [],
        remainingRangeBuckets: [],
        expiryDate: addDays(record.date, 30),
        usedMinutes: 0,
        remainingMinutes: 0,
        allocations: []
      }];
    }
    return generatedRanges.map((range, index) => ({
      ...record,
      id: `${record.id}__${range.segmentType || index}`,
      sourceAttendanceId: record.id,
      earnedMinutes: range.minutes,
      segmentType: range.segmentType || `segment-${index + 1}`,
      segmentLabel: range.segmentLabel || `구간 ${index + 1}`,
      generatedRanges: [range],
      remainingRangeBuckets: [{ ...range }],
      expiryDate: addDays(record.date, 30),
      usedMinutes: 0,
      remainingMinutes: range.minutes,
      allocations: []
    }));
  });
  const invalidUsageIds = [];
  const usageAllocations = {};
  for (const usage of usageRecords) {
    let remainingUsage = usage.durationMinutes;
    const allocations = [];
    const selectedAttendanceIds = normalizeSelectedAttendanceIds(usage.selectedAttendanceIds);
    const lockedAttendanceIds = normalizeAttendanceIds(usage.lockedAttendanceIds, Infinity);
    const effectiveAttendanceIds = selectedAttendanceIds.length ? selectedAttendanceIds : lockedAttendanceIds;
    const selectableEntries = effectiveAttendanceIds.length
      ? resolveSelectedAttendanceEntries(effectiveAttendanceIds, generatedEntries).filter((entry) => typeof entry !== "string")
      : generatedEntries;
    for (const entry of selectableEntries) {
      if (remainingUsage <= 0) break;
      if (entry.earnedMinutes <= 0 || entry.remainingMinutes <= 0) continue;
      if (isEntryExpired(entry, usage.date)) continue;
      if (compareDate(entry.date, usage.date) > 0) continue;
      const allocated = Math.min(entry.remainingMinutes, remainingUsage);
      if (allocated <= 0) continue;
      entry.usedMinutes += allocated;
      entry.remainingMinutes -= allocated;
      remainingUsage -= allocated;
      const attendanceRangeLabel = allocateFromGeneratedRanges(entry, allocated);
      allocations.push({ attendanceId: entry.id, attendanceDate: entry.date, attendanceRangeLabel, usageId: usage.id, usageDate: usage.date, usageStart: usage.startTime, usageEnd: usage.endTime, minutes: allocated, attendanceSegmentLabel: entry.segmentLabel });
      entry.allocations.push({ usageId: usage.id, usageDate: usage.date, usageStart: usage.startTime, usageEnd: usage.endTime, minutes: allocated });
    }
    if (remainingUsage > 0) {
      allocations.forEach((allocation) => {
        const entry = generatedEntries.find((item) => item.id === allocation.attendanceId);
        if (!entry) return;
        entry.usedMinutes -= allocation.minutes;
        entry.remainingMinutes += allocation.minutes;
        entry.allocations = entry.allocations.filter((item) => item.usageId !== usage.id);
      });
      invalidUsageIds.push(usage.id);
      continue;
    }
    usageAllocations[usage.id] = allocations;
  }
  return { attendanceRecords, usageRecords, generatedEntries, invalidUsageIds, usageAllocations };
}
function renderSummary() {
  const ledger = buildLedger();
  const today = getTodayString();
  const validEntries = ledger.generatedEntries.filter((entry) => entry.earnedMinutes > 0 && entry.remainingMinutes > 0 && !isEntryExpired(entry, today));
  const totalRemaining = validEntries.reduce((sum, item) => sum + item.remainingMinutes, 0);
  const usableToday = Math.min(MAX_DAILY_USAGE_MINUTES, totalRemaining);
  const schedule = WORK_TYPES[elements.todayWorkType.value];
  const leaveMinutes = calculateLeaveTimeExcludingLunch(elements.todayWorkType.value, usableToday);
  elements.totalRemainingLabel.textContent = formatDuration(totalRemaining);
  elements.validEntryCountLabel.textContent = `유효한 발생 기록 ${validEntries.length}건`;
  elements.leaveTimeLabel.textContent = usableToday > 0 ? formatTime(leaveMinutes) : schedule.end;
  elements.leaveTimeMeta.textContent = usableToday > 0 ? `오늘 최대 ${formatDuration(usableToday)} 사용 기준` : "사용 가능한 잔여 특정일 시간이 없습니다.";
}
function renderAttendanceList() {
  const ledger = buildLedger();
  const visibleEntries = ledger.generatedEntries.filter((entry) => shouldShowAttendanceEntry(entry));
  elements.attendanceCount.textContent = `${visibleEntries.length}건`;
  renderCollection(elements.attendanceList, visibleEntries.map(renderAttendanceItem));
}
function renderTimelineList() {
  const ledger = buildLedger();
  const rows = buildLedgerRows(ledger);
  renderCollection(elements.timelineList, rows.map(renderTimelineItem));
}
function renderCollection(container, nodes) {
  container.innerHTML = "";
  if (!nodes.length) return container.appendChild(EMPTY_TEMPLATE.content.firstElementChild.cloneNode(true));
  nodes.forEach((node) => container.appendChild(node));
}
function renderAttendanceItem(entry) {
  const item = document.createElement("article");
  const status = getEntryStatus(entry);
  const schedule = WORK_TYPES[resolveWorkTypeKey(entry.workType)];
  const overtimeChecked = Boolean(entry.overtime ?? entry.overtimeChecked);
  const visibleAllocations = entry.allocations.filter((allocation) => isUsageHistoryVisible(allocation.usageDate));
  const usedDetails = visibleAllocations.length
    ? visibleAllocations.map((allocation) => formatUsageRange({ date: allocation.usageDate, startTime: allocation.usageStart, endTime: allocation.usageEnd }, allocation.minutes)).join(" / ")
    : (entry.allocations.length ? "표시 기간 종료" : "아직 사용되지 않음");
  item.className = `list-item ${status.className}`.trim();
  item.innerHTML = `
    <div class="item-row"><div><div class="item-title">${entry.date} · ${schedule.label} · ${entry.segmentLabel}</div><div class="item-subtitle">실제 ${entry.actualStart || "--:--"} ~ ${entry.actualEnd || "--:--"} · <label class="checkbox-field"><input type="checkbox" data-role="overtime-checkbox" data-id="${entry.sourceAttendanceId}" ${overtimeChecked ? "checked" : ""}><span>시간외근무</span></label></div></div><div class="status-row"><span class="pill ${status.pillClass}">${status.label}</span><span class="pill neutral">만료 ${entry.expiryDate}</span></div></div>
    <div class="detail-grid"><div class="detail-box"><span>발생내역</span><strong>${formatGeneratedRanges(entry)}</strong></div><div class="detail-box"><span>발생시간</span><strong>${formatDuration(entry.earnedMinutes)}</strong></div><div class="detail-box"><span>남은시간</span><strong>${formatDuration(entry.remainingMinutes)}</strong></div><div class="detail-box"><span>차감 내역</span><strong>${usedDetails}</strong></div></div>
    <div class="item-actions"><button class="mini-btn" type="button" data-action="save-overtime" data-id="${entry.sourceAttendanceId}">시간외근무 저장</button><button class="mini-btn" type="button" data-action="edit-attendance" data-id="${entry.sourceAttendanceId}">수정</button><button class="mini-btn danger" type="button" data-action="delete-attendance" data-id="${entry.sourceAttendanceId}">삭제</button></div>`;
  bindItemActions(item);
  return item;
}
function renderUsageItem(usage, ledger) {
  const item = document.createElement("article");
  const isInvalid = ledger.invalidUsageIds.includes(usage.id);
  const allocations = ledger.usageAllocations[usage.id] || [];
  const deductionText = allocations.length ? allocations.map((allocation) => `${allocation.attendanceRangeLabel}에서 ${formatDuration(allocation.minutes)}`).join(" / ") : "차감 내역 없음";
  const detailColumns = "repeat(auto-fit, minmax(150px, 1fr))";
  const schedule = WORK_TYPES[resolveWorkTypeKey(usage.workType)];
  item.className = `list-item ${isInvalid ? "warning" : ""}`.trim();
  item.innerHTML = `
    <div class="item-row"><div><div class="item-title">${usage.date} · ${schedule.label}</div><div class="item-subtitle">${formatUsageRange(usage)}</div></div><div class="status-row"><span class="pill ${isInvalid ? "warning" : "info"}">${isInvalid ? "차감 불가" : "차감 완료"}</span></div></div>
    <div class="detail-grid" style="grid-template-columns:${detailColumns}"><div class="detail-box"><span>사용내역</span><strong>${formatUsageRange(usage)}</strong></div><div class="detail-box"><span>사용시간</span><strong>${formatDuration(usage.durationMinutes)}</strong></div><div class="detail-box"><span>선택 발생기록</span><strong>${getSelectedAttendanceLabel(getEffectiveUsageAttendanceIds(usage))}</strong></div><div class="detail-box"><span>차감 출처</span><strong>${deductionText}</strong></div><div class="detail-box"><span>상태</span><strong>${isInvalid ? "유효한 발생시간 부족" : "정상 저장"}</strong></div></div>
    <div class="item-actions"><button class="mini-btn" type="button" data-action="edit-usage" data-id="${usage.id}">수정</button><button class="mini-btn danger" type="button" data-action="delete-usage" data-id="${usage.id}">삭제</button></div>`;
  bindItemActions(item);
  return item;
}
function openMatchingSummaryDialog() {
  elements.matchingSummaryContent.innerHTML = renderMatchingSummaryTable(buildLedger());
  if (typeof elements.matchingSummaryDialog.showModal === "function") {
    elements.matchingSummaryDialog.showModal();
  } else {
    elements.matchingSummaryDialog.setAttribute("open", "");
  }
}
function closeMatchingSummaryDialog() {
  if (typeof elements.matchingSummaryDialog.close === "function") {
    elements.matchingSummaryDialog.close();
  } else {
    elements.matchingSummaryDialog.removeAttribute("open");
  }
}
function renderMatchingSummaryTable(ledger) {
  const rows = buildMatchingSummaryRows(ledger);
  if (!rows.length) {
    return `<div class="empty-state"><strong>표시할 매칭 내역이 없습니다.</strong><span>유효한 평일 발생/사용 기록이 없습니다.</span></div>`;
  }
  const rowHtml = rows.map((row) => `
    <tr>
      <td>${row.attendanceDate}</td>
      <td>${row.attendanceRange}</td>
      <td>${formatDuration(row.earnedMinutes)}</td>
      <td>${row.usageDate || "-"}</td>
      <td>${row.usageRange || "-"}</td>
      <td>${formatDuration(row.usedMinutes)}</td>
      <td>${formatDuration(row.remainingMinutes)}</td>
      <td><span class="pill ${row.statusClass}">${row.status}</span></td>
    </tr>`).join("");
  return `
    <div class="table-scroll">
      <table class="matching-table">
        <thead>
          <tr>
            <th>발생일</th>
            <th>발생구간</th>
            <th>발생시간</th>
            <th>사용일</th>
            <th>사용시간</th>
            <th>차감시간</th>
            <th>잔여시간</th>
            <th>상태</th>
          </tr>
        </thead>
        <tbody>${rowHtml}</tbody>
      </table>
    </div>`;
}
function buildMatchingSummaryRows(ledger) {
  const rows = [];
  ledger.generatedEntries
    .filter((entry) => entry.earnedMinutes > 0)
    .sort(sortByDateThenId)
    .forEach((entry) => {
      const allocations = entry.allocations || [];
      if (!allocations.length) {
        rows.push(createMatchingSummaryRow(entry, null, 0, entry.remainingMinutes, "미사용", "info"));
        return;
      }
      allocations.forEach((allocation, index) => {
        const remainingMinutes = index === allocations.length - 1 ? entry.remainingMinutes : 0;
        rows.push(createMatchingSummaryRow(entry, allocation, allocation.minutes, remainingMinutes, remainingMinutes > 0 ? "일부사용" : "사용완료", remainingMinutes > 0 ? "warning" : "neutral"));
      });
    });
  return rows;
}
function createMatchingSummaryRow(entry, allocation, usedMinutes, remainingMinutes, status, statusClass) {
  return {
    attendanceDate: entry.date,
    attendanceRange: `${entry.segmentLabel} · ${formatGeneratedRanges(entry)}`,
    earnedMinutes: entry.earnedMinutes,
    usageDate: allocation?.usageDate || "",
    usageRange: allocation ? formatUsageRange({ date: allocation.usageDate, startTime: allocation.usageStart, endTime: allocation.usageEnd }, allocation.minutes) : "",
    usedMinutes,
    remainingMinutes,
    status,
    statusClass
  };
}
function buildLedgerRows(ledger) {
  const today = getTodayString();
  const rows = [];
  ledger.generatedEntries
    .filter((entry) => entry.earnedMinutes > 0 && entry.remainingMinutes > 0 && !isEntryExpired(entry, today))
    .forEach((entry) => rows.push({
      type: entry.allocations.length ? "잔여" : "미사용",
      usageId: `remaining_${entry.id}`,
      usageLabel: entry.allocations.length ? `${entry.date} 발생분 잔여` : `${entry.date} 발생분 미사용`,
      attendanceItems: [entry.allocations.length ? formatRemainingRanges(entry) : `${entry.segmentLabel} · ${formatGeneratedRanges(entry)}`],
      flowLabel: entry.allocations.length ? "잔여" : "적립",
      minutes: entry.remainingMinutes,
      subtitle: entry.allocations.length ? "사용 후 남아 있는 특정일 시간" : "아직 사용되지 않은 특정일 시간"
    }));
  ledger.usageRecords.filter((usage) => isUsageHistoryVisible(usage.date)).forEach((usage) => {
    const allocations = (ledger.usageAllocations[usage.id] || []).filter((allocation) => !isExpiredAttendanceDate(allocation.attendanceDate, today));
    if (!allocations.length) return;
    const visibleMinutes = allocations.reduce((sum, allocation) => sum + allocation.minutes, 0);
    rows.push({ type: "차감", usageId: usage.id || `${usage.date}_${usage.startTime}_${usage.endTime}`, usageLabel: formatUsageRange(usage), attendanceItems: allocations.map((allocation) => allocation.attendanceRangeLabel), flowLabel: "차감", minutes: visibleMinutes, canDelete: true, subtitle: "사용기록 1건 기준 연결 내역" });
  });
  return groupLedgerRowsByUsage(rows);
}
function renderTimelineItem(row) {
  const item = document.createElement("article");
  const actionHtml = row.canDelete
  ? `<div class="item-actions"><button class="mini-btn danger" type="button" data-action="delete-usage" data-id="${row.usageId}">삭제</button></div>`
  : "";
  const flowPillClass = row.flowLabel === "차감" ? "warning" : "info";
  const amountLabel = row.flowLabel === "차감" ? "차감시간" : "남은시간";
  item.className = "list-item";
  item.innerHTML = `
    <div class="item-row"><div><div class="item-title">${row.type}</div><div class="item-subtitle">${row.subtitle || "사용기록 1건 기준 연결 내역"}</div></div><span class="pill ${flowPillClass}">${row.flowLabel}</span></div>
    <div class="detail-grid"><div class="detail-box"><span>사용내역</span><strong>${row.usageLabel}</strong></div><div class="detail-box"><span>발생내역</span><strong>${row.attendanceItems.join("<br>")}</strong></div><div class="detail-box"><span>${amountLabel}</span><strong>${formatDuration(row.minutes)}</strong></div><div class="detail-box"><span>흐름</span><strong>${row.flowLabel}</strong></div></div>
    ${actionHtml}`;
  bindItemActions(item);
  return item;
}
function bindItemActions(container) {
  container.querySelectorAll("[data-action]").forEach((button) => button.addEventListener("click", () => {
    const action = button.dataset.action;
    const id = button.dataset.id;
    if (action === "save-overtime") {
      const checkbox = container.querySelector(`[data-role="overtime-checkbox"][data-id="${id}"]`);
      return saveOvertimeFromList(id, Boolean(checkbox?.checked));
    }
    if (action === "edit-attendance") startAttendanceEdit(id);
    if (action === "delete-attendance") deleteAttendance(id);
    if (action === "edit-usage") startUsageEdit(id);
    if (action === "delete-usage") deleteUsage(id);
  }));
}
function saveOvertimeFromList(recordId, checked) {
  const recordIndex = state.attendanceRecords.findIndex((item) => item.id === recordId);
  if (recordIndex < 0) return;
  const updatedRecord = { ...state.attendanceRecords[recordIndex], overtime: checked, overtimeChecked: checked };
  const generatedMinutes = calculateGeneratedMinutes(updatedRecord);
  state.attendanceRecords.splice(recordIndex, 1, { ...updatedRecord, generatedMinutes });
  if (elements.attendanceId.value === recordId) elements.overtimeChecked.checked = checked;
  refreshAutoUsageSelections();
  saveAttendanceRecords();
  rerenderAll();
}
function startAttendanceEdit(id) {
  const record = state.attendanceRecords.find((item) => item.id === id);
  if (!record) return;
  elements.attendanceId.value = record.id;
  elements.attendanceDate.value = record.date;
  elements.attendanceWorkType.value = resolveWorkTypeKey(record.workType);
  elements.actualStart.value = record.actualStart;
  elements.actualEnd.value = record.actualEnd;
  elements.overtimeChecked.checked = Boolean(record.overtime ?? record.overtimeChecked);
  elements.attendanceFormMode.textContent = "편집 중";
  elements.attendanceFormMode.className = "pill info";
  renderAttendancePreview();
  window.scrollTo({ top: 0, behavior: "smooth" });
}
function deleteAttendance(id) {
  if (!confirm("이 발생기록을 삭제할까요? 연결된 사용 차감 결과도 다시 계산됩니다.")) return;
  state.attendanceRecords = state.attendanceRecords.filter((item) => item.id !== id);
  refreshAutoUsageSelections();
  saveState();
  populateUsageAttendanceOptions();
  renderAll();
}
function startUsageEdit(id) {
  const record = state.usageRecords.find((item) => item.id === id);
  if (!record) return;
  elements.usageId.value = record.id;
  elements.usageDate.value = record.date;
  populateUsageAttendanceOptions(normalizeSelectedAttendanceIds(record.selectedAttendanceIds ?? record.selectedAttendanceDates ?? record.selectedAttendanceDate));
  elements.usageWorkType.value = record.workType;
  setUsageTimeControl("start", record.startTime);
  setUsageTimeControl("end", record.endTime);
  elements.usageFormMode.textContent = "편집 중";
  elements.usageFormMode.className = "pill info";
  renderUsagePreview();
  window.scrollTo({ top: 0, behavior: "smooth" });
}
function deleteUsage(id) {
  if (!confirm("이 사용기록을 삭제할까요? FIFO 차감 결과가 다시 계산됩니다.")) return;
  state.usageRecords = state.usageRecords.filter((item) => item.id !== id);
  refreshAutoUsageSelections();
  saveState();
  renderAll();
}
async function handleExcelUpload(file = elements.excelFile.files[0]) {
  if (!file) return alert("업로드할 파일을 선택하세요.");
  try {
    const parsed = await parseSpreadsheetFile(file);
    if (!parsed.parsedRecords.length) throw new Error("업로드 가능한 근태 데이터가 없습니다.");
    const existingRecordsLength = state.attendanceRecords.length;
    let zeroMinuteExcludedCount = 0;
    let duplicateSkippedCount = 0;
    let weekendExcludedCount = 0;
    const existingDates = new Set(state.attendanceRecords.map((record) => record.date));
    const appendedRecords = [];
    parsed.parsedRecords.forEach((record) => {
      if (isWeekendDate(record.date)) {
        weekendExcludedCount += 1;
        return;
      }
      const generatedMinutes = calculateGeneratedMinutes(record);
      if (generatedMinutes <= 0) {
        zeroMinuteExcludedCount += 1;
        return;
      }
      if (existingDates.has(record.date)) {
        duplicateSkippedCount += 1;
        return;
      }
      existingDates.add(record.date);
      appendedRecords.push({ ...record, generatedMinutes });
    });
    state.attendanceRecords = [...state.attendanceRecords, ...appendedRecords];
    refreshAutoUsageSelections();
    saveAttendanceRecords();
    populateUsageAttendanceOptions();
    rerenderAll();
    const appendedCount = appendedRecords.length;
    const finalSavedCount = state.attendanceRecords.length;
    const finalRenderedCount = buildLedger().generatedEntries.filter((entry) => entry.earnedMinutes > 0).length;
    const sheetLabel = parsed.sheetName ? `, 선택 시트 ${parsed.sheetName}` : "";
    elements.uploadResult.textContent = `업로드 완료: 실제 데이터 ${parsed.dataRowCount}건, 유효 파싱 ${parsed.parsedRecords.length}건, 주말 제외 ${weekendExcludedCount}건, 0분 제외 ${zeroMinuteExcludedCount}건, 기존 날짜 중복 ${duplicateSkippedCount}건, 신규 추가 ${appendedCount}건, 최종 저장 ${finalSavedCount}건, 최종 표시 ${finalRenderedCount}건${sheetLabel}`;
  } catch (error) {
    console.error(error);
    elements.uploadResult.textContent = `업로드 실패: ${error.message}`;
    alert(`업로드 실패: ${error.message}`);
  }
}
async function parseSpreadsheetFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  const supportedExtensions = new Set(["xls", "xlsx", "xml", "csv", "html", "htm"]);
  if (!supportedExtensions.has(extension)) throw new Error("지원하지 않는 파일 형식입니다. 근태 파일(.xls/.xlsx/.xml/.csv/.html)을 업로드하세요.");
  return parseExcelWithSheetJs(file);
}
async function parseExcelWithSheetJs(file) {
  if (typeof XLSX === "undefined") throw new Error("엑셀 라이브러리를 불러오지 못했습니다. xlsx.full.min.js 로드 여부를 확인하세요.");
  const data = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(data, { type: "array", codepage: 949, cellDates: false });
  if (!workbook.SheetNames.length) throw new Error("시트를 찾지 못했습니다.");
  let bestParsed = null;
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });
    const parsed = parseAttendanceXlsRows(rows);
    if (!bestParsed || isBetterParsedSheet(parsed, bestParsed)) bestParsed = { ...parsed, sheetName };
  });
  if (!bestParsed) throw new Error("업로드 가능한 시트를 찾지 못했습니다.");
  return bestParsed;
}
function isBetterParsedSheet(candidate, currentBest) {
  if (candidate.parsedRecords.length !== currentBest.parsedRecords.length) {
    return candidate.parsedRecords.length > currentBest.parsedRecords.length;
  }
  const candidateUsableCount = candidate.parsedRecords.filter((record) => calculateGeneratedMinutes(record) > 0).length;
  const currentUsableCount = currentBest.parsedRecords.filter((record) => calculateGeneratedMinutes(record) > 0).length;
  if (candidateUsableCount !== currentUsableCount) return candidateUsableCount > currentUsableCount;
  if (candidate.dataRowCount !== currentBest.dataRowCount) return candidate.dataRowCount > currentBest.dataRowCount;
  return candidate.excludedRowCount < currentBest.excludedRowCount;
}
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
    reader.readAsArrayBuffer(file);
  });
}
function parseAttendanceXlsRows(rows) {
  if (!Array.isArray(rows) || !rows.length) return { parsedRecords: [], dataRowCount: 0, excludedRowCount: 0 };
  const rowEntries = rows
    .map((row, index) => ({ row: Array.isArray(row) ? row : [], index }))
    .filter(({ row }) => row.some((cell) => String(cell ?? "").trim() !== ""));
  if (!rowEntries.length) return { parsedRecords: [], dataRowCount: 0, excludedRowCount: 0 };
  const detectedHeader = findAttendanceHeaderRow(rowEntries);
  if (!detectedHeader) return { parsedRecords: [], dataRowCount: 0, excludedRowCount: rowEntries.length };
  const columnIndexes = detectedHeader.columnIndexes;
  const dataRows = rowEntries.slice(detectedHeader.rowEntryIndex + 1).map(({ row }) => row);
  const parsedRecords = [];
  let excludedRowCount = 0;
  dataRows.forEach((row) => {
    const cells = Array.isArray(row) ? row : [];
    const date = normalizeExcelDate(cells[columnIndexes.date]);
    const workType = extractWorkType(cells[columnIndexes.workType]);
    const actualStart = normalizeExcelTime(cells[columnIndexes.actualStart]);
    const actualEnd = normalizeExcelTime(cells[columnIndexes.actualEnd]);
    if (!date) {
      excludedRowCount += 1;
      return;
    }
    if (!actualStart || !actualEnd) {
      excludedRowCount += 1;
      return;
    }
    if (toMinutes(actualEnd) <= toMinutes(actualStart)) {
      excludedRowCount += 1;
      return;
    }
    parsedRecords.push({ id: createId("attendance"), date, workType: workType || "C형", actualStart, actualEnd, overtime: false, overtimeChecked: false, source: "import" });
  });
  return { parsedRecords, dataRowCount: dataRows.length, excludedRowCount };
}
function findAttendanceHeaderRow(rowEntries) {
  for (let rowEntryIndex = 0; rowEntryIndex < rowEntries.length; rowEntryIndex += 1) {
    const { row } = rowEntries[rowEntryIndex];
    const nextRow = rowEntries[rowEntryIndex + 1]?.row || [];
    const normalizedHeaders = buildNormalizedAttendanceHeaders(row, nextRow);
    const columnIndexes = {
      date: findHeaderColumnIndex(normalizedHeaders, ["날짜", "근무일자", "근무일", "일자", "date", "workdate"]),
      workType: findHeaderColumnIndex(normalizedHeaders, ["근무유형", "근무형", "근무", "유형", "worktype", "shift"]),
      actualStart: findHeaderColumnIndex(normalizedHeaders, ["출퇴근카드출근일시", "실제출근시간", "출근일시", "출근시간", "실제출근", "출근", "출근시각", "starttime", "clockin", "intime"]),
      actualEnd: findHeaderColumnIndex(normalizedHeaders, ["출퇴근카드퇴근일시", "실제퇴근시간", "퇴근일시", "퇴근시간", "실제퇴근", "퇴근", "퇴근시각", "endtime", "clockout", "outtime"])
    };
    if (Number.isInteger(columnIndexes.date) && Number.isInteger(columnIndexes.actualStart) && Number.isInteger(columnIndexes.actualEnd)) {
      return { rowEntryIndex, columnIndexes };
    }
  }
  return null;
}
function buildNormalizedAttendanceHeaders(row, nextRow = []) {
  const maxLength = Math.max(row.length, nextRow.length);
  return Array.from({ length: maxLength }, (_, index) => {
    const primary = normalizeAttendanceHeaderCell(row[index]);
    const secondary = normalizeAttendanceHeaderCell(nextRow[index]);
    if (primary && secondary && primary !== secondary) return `${primary}${secondary}`;
    return primary || secondary;
  });
}
function getLegacyAttendanceColumnIndexes() {
  return { date: 5, workType: 8, actualStart: 12, actualEnd: 13 };
}
function normalizeAttendanceHeaderCell(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s/g, "")
    .replace(/[()[\]{}_.:/\\-]/g, "");
}
function findHeaderColumnIndex(headers, candidates) {
  const validHeaders = headers.map((header, index) => ({ header, index })).filter(({ header }) => header);
  for (const candidate of candidates) {
    const normalizedCandidate = normalizeAttendanceHeaderCell(candidate);
    const exactMatch = validHeaders.find(({ header }) => header === normalizedCandidate);
    if (exactMatch) return exactMatch.index;
  }
  for (const candidate of candidates) {
    const normalizedCandidate = normalizeAttendanceHeaderCell(candidate);
    const partialMatch = validHeaders.find(({ header }) => header.includes(normalizedCandidate) || normalizedCandidate.includes(header));
    if (partialMatch) return partialMatch.index;
  }
  return -1;
}
function normalizeExcelDate(value) {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toDateString(value);
  if (typeof value === "number" && Number.isFinite(value) && value > 20000) return excelSerialToDate(value);
  const text = String(value).trim();
  const koreanMatch = text.match(/(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
  if (koreanMatch) return `${koreanMatch[1]}-${koreanMatch[2].padStart(2, "0")}-${koreanMatch[3].padStart(2, "0")}`;
  const match = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
  const shortMatch = text.match(/(^|[^\d])(\d{1,2})[./-](\d{1,2})(?:[^\d]|$)/);
  if (shortMatch) {
    const inferredYear = new Date().getFullYear();
    return `${inferredYear}-${shortMatch[2].padStart(2, "0")}-${shortMatch[3].padStart(2, "0")}`;
  }
  if (/^\d{5}$/.test(text)) return excelSerialToDate(Number(text));
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : toDateString(parsed);
}
function normalizeExcelTime(value) {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
  if (typeof value === "number" && Number.isFinite(value)) {
    const fraction = value % 1;
    if (fraction > 0 || (value > 0 && value < 1)) return formatTime(Math.round((fraction || value) * 24 * 60));
  }
  const text = String(value).trim();
  const koreanMeridiemMatch = text.match(/(오전|오후)\s*(\d{1,2})[:시]\s*(\d{1,2})?/);
  if (koreanMeridiemMatch) {
    let hours = Number(koreanMeridiemMatch[2]) % 12;
    if (koreanMeridiemMatch[1] === "오후") hours += 12;
    const minutes = Number(koreanMeridiemMatch[3] ?? "0");
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }
  const koreanTimeMatch = text.match(/(\d{1,2})\s*시(?:\s*(\d{1,2})\s*분?)?/);
  if (koreanTimeMatch) return `${koreanTimeMatch[1].padStart(2, "0")}:${String(Number(koreanTimeMatch[2] ?? "0")).padStart(2, "0")}`;
  const match = text.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (match) return `${match[1].padStart(2, "0")}:${match[2]}`;
  return "";
}
function extractWorkType(value) {
  const text = String(value ?? "").trim();
  const match = text.match(/([A-D](?:-1)?형)/i);
  if (match) {
    const normalized = match[1].toUpperCase().replace("형", "");
    return `${normalized}형`;
  }
  return normalizeWorkType(text);
}
function normalizeWorkType(value) {
  const raw = String(value ?? "").trim().toUpperCase().replace(/\s/g, "").replace("형", "");
  if (!raw) return "C형";
  const matchedKey = Object.keys(WORK_TYPES).find((key) => key.toUpperCase() === raw);
  if (matchedKey) return WORK_TYPES[matchedKey].label;
  const matchedLabel = Object.values(WORK_TYPES).find((item) => item.label.toUpperCase().replace(/\s/g, "") === raw);
  return matchedLabel ? matchedLabel.label : "C형";
}
function resolveWorkTypeKey(workTypeValue) {
  const raw = String(workTypeValue ?? "").trim().toUpperCase().replace(/\s/g, "");
  if (!raw) return "C";
  const directKey = Object.keys(WORK_TYPES).find((key) => key.toUpperCase() === raw);
  if (directKey) return directKey;
  const matchedEntry = Object.entries(WORK_TYPES).find(([, config]) => config.label.toUpperCase().replace(/\s/g, "") === raw);
  return matchedEntry ? matchedEntry[0] : "C";
}
function formatGeneratedRanges(entry) {
  const ranges = entry.generatedRanges || buildGeneratedTimeRanges(entry);
  if (!ranges.length) return "0분";
  return ranges.map((range) => `${formatShortDate(entry.date)} ${formatTime(range.start)}~${formatTime(range.end)}(${formatDuration(range.minutes)})`).join(", ");
}
function formatUsageRange(usageRecord, minutes = null) {
  const normalized = normalizeRangeToHalfHour(usageRecord.startTime, usageRecord.endTime);
  const startLabel = normalized ? formatTime(normalized.start) : usageRecord.startTime;
  const endLabel = normalized ? formatTime(normalized.end) : usageRecord.endTime;
  const duration = minutes ?? usageRecord.durationMinutes ?? (normalized ? calculateUsageMinutesExcludingLunch(normalized.start, normalized.end) : 0);
  return `${formatShortDate(usageRecord.date)} ${startLabel}~${endLabel}(${formatDuration(duration)})`;
}
function formatRemainingRanges(entry) {
  const buckets = (entry.remainingRangeBuckets || []).filter((bucket) => bucket.end > bucket.start);
  if (!buckets.length) return `${entry.segmentLabel} · 잔여 없음`;
  return buckets.map((bucket) => `${entry.segmentLabel} · ${formatShortDate(entry.date)} ${formatTime(bucket.start)}~${formatTime(bucket.end)}(${formatDuration(bucket.end - bucket.start)})`).join(", ");
}
function allocateFromGeneratedRanges(entry, minutesToAllocate) {
  const labels = [];
  let remaining = minutesToAllocate;
  for (const bucket of entry.remainingRangeBuckets || []) {
    if (remaining <= 0) break;
    const available = bucket.end - bucket.start;
    if (available <= 0) continue;
    const slice = Math.min(available, remaining);
    labels.push(`${formatShortDate(entry.date)} ${formatTime(bucket.start)}~${formatTime(bucket.start + slice)}(${formatDuration(slice)})`);
    bucket.start += slice;
    remaining -= slice;
  }
  return labels.length ? labels.join(", ") : formatGeneratedRanges(entry);
}
function floorToHalfHour(timeString) { return formatTime(floorToUnit(toMinutes(timeString), FLOOR_UNIT)); }
function floorTimeToHalfHour(timeString) { return timeString ? floorToHalfHour(timeString) : ""; }
function normalizeUsageTimeInput(timeString) { return floorTimeToHalfHour(timeString); }
// Shared helper so usage deduction and leave-time calculation exclude the same lunch window.
function getLunchOverlapMinutes(startMinutes, endMinutes) {
  const overlapStart = Math.max(startMinutes, LUNCH_START_MINUTES);
  const overlapEnd = Math.min(endMinutes, LUNCH_END_MINUTES);
  return overlapEnd > overlapStart ? overlapEnd - overlapStart : 0;
}
function calculateUsageMinutesExcludingLunch(startMinutes, endMinutes) {
  if (startMinutes >= endMinutes) return 0;
  const totalUsage = endMinutes - startMinutes;
  const lunchOverlap = getLunchOverlapMinutes(startMinutes, endMinutes);
  return Math.max(0, floorToUnit(totalUsage - lunchOverlap, FLOOR_UNIT));
}
// Walk backward from scheduled end and skip lunch so early-leave time matches actual usable minutes.
function calculateLeaveTimeExcludingLunch(workType, usableMinutes) {
  const schedule = WORK_TYPES[resolveWorkTypeKey(workType)];
  if (!schedule) return toMinutes("18:00");
  const scheduleStart = toMinutes(schedule.start);
  let cursor = toMinutes(schedule.end);
  let remainingUsage = Math.max(0, floorToUnit(usableMinutes, FLOOR_UNIT));
  if (remainingUsage <= 0) return cursor;
  if (cursor > LUNCH_END_MINUTES) {
    const afternoonStart = Math.max(scheduleStart, LUNCH_END_MINUTES);
    const availableAfternoon = Math.max(0, cursor - afternoonStart);
    const usedAfternoon = Math.min(availableAfternoon, remainingUsage);
    cursor -= usedAfternoon;
    remainingUsage -= usedAfternoon;
  }
  if (remainingUsage > 0) {
    cursor = Math.min(cursor, LUNCH_START_MINUTES);
    const availableMorning = Math.max(0, cursor - scheduleStart);
    const usedMorning = Math.min(availableMorning, remainingUsage);
    cursor -= usedMorning;
  }
  return Math.max(scheduleStart, cursor);
}
function normalizeRangeToHalfHour(startTime, endTime) {
  if (!startTime || !endTime) return null;
  const start = floorToUnit(toMinutes(startTime), FLOOR_UNIT);
  const end = floorToUnit(toMinutes(endTime), FLOOR_UNIT);
  return end <= start ? null : { start, end };
}
function groupLedgerRowsByUsage(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const key = row.usageId || row.usageLabel;
    if (!grouped.has(key)) return grouped.set(key, { ...row, attendanceItems: [...row.attendanceItems] });
    grouped.get(key).attendanceItems.push(...row.attendanceItems);
  });
  return Array.from(grouped.values());
}
function getEntryStatus(entry) {
  const today = getTodayString();
  const remainingDays = diffDays(today, entry.expiryDate);
  if (compareDate(entry.expiryDate, today) < 0) return { label: "만료", pillClass: "expired", className: "expired", remainingDaysLabel: "만료됨" };
  if (remainingDays <= 3) return { label: "임박", pillClass: "warning", className: "warning", remainingDaysLabel: `${remainingDays}일` };
  return { label: "정상", pillClass: "info", className: "", remainingDaysLabel: `${remainingDays}일` };
}
function seedDemoData() {
  if (!confirm("샘플 데이터를 현재 저장소에 추가할까요? 기존 데이터는 유지됩니다.")) return;
  const today = getTodayString();
  const sampleAttendance = [
    { id: createId("attendance"), date: addDays(today, -12), workType: "C", actualStart: "08:00", actualEnd: "19:10", overtime: false, overtimeChecked: false, source: "manual" },
    { id: createId("attendance"), date: addDays(today, -8), workType: "A", actualStart: "06:20", actualEnd: "16:40", overtime: false, overtimeChecked: false, source: "manual" },
    { id: createId("attendance"), date: addDays(today, -2), workType: "B-1", actualStart: "08:00", actualEnd: "20:40", overtime: true, overtimeChecked: true, source: "manual" }
  ];
  const sampleUsage = [{ id: createId("usage"), date: addDays(today, -1), workType: "C", startTime: "14:00", endTime: "16:00", durationMinutes: 120 }];
  sampleAttendance.forEach((record) => {
    if (!state.attendanceRecords.some((item) => item.date === record.date) && calculateEarnedMinutes(record) > 0) state.attendanceRecords.push(record);
  });
  sampleUsage.forEach((record) => state.usageRecords.push(record));
  saveState();
  populateUsageAttendanceOptions();
  renderAll();
}
function getUsageAttendanceSelectElements() {
  return [elements.usageAttendanceDate1, elements.usageAttendanceDate2, elements.usageAttendanceDate3, elements.usageAttendanceDate4];
}
function normalizeAttendanceIds(value, maxCount = 4) {
  if (!value) return [];
  const rawValues = Array.isArray(value) ? value : [value];
  const normalizedValues = rawValues.map((item) => String(item || "").trim()).filter(Boolean);
  return Number.isFinite(maxCount) ? normalizedValues.slice(0, maxCount) : normalizedValues;
}
function normalizeSelectedAttendanceIds(value) {
  return normalizeAttendanceIds(value, 4);
}
function getSelectedAttendanceIdsFromForm() {
  return normalizeSelectedAttendanceIds(getUsageAttendanceSelectElements().map((select) => select.value));
}
function normalizeUsageAttendanceSelects(changedSelect) {
  if (!changedSelect.value) return;
  getUsageAttendanceSelectElements().forEach((select) => {
    if (select === changedSelect) return;
    if (select.value === changedSelect.value) select.value = "";
  });
}
function renderSelectedAttendanceSummary() {
  const usageDate = elements.usageDate.value;
  const selectedIds = getSelectedAttendanceIdsFromForm();
  if (!selectedIds.length) {
    elements.selectedAttendanceSummary.textContent = "자동 선택(FIFO)";
    elements.selectedAttendanceTotalLabel.textContent = "0분";
    return;
  }
  const availableEntries = getSelectableAttendanceEntries(usageDate);
  const selectedEntries = resolveSelectedAttendanceEntries(selectedIds, availableEntries).filter((entry) => typeof entry !== "string");
  const totalMinutes = selectedEntries.reduce((sum, entry) => sum + entry.remainingMinutes, 0);
  elements.selectedAttendanceSummary.textContent = `선택 ${selectedEntries.length}건: ${selectedEntries.map((entry) => `${entry.date} ${entry.segmentLabel}`).join(", ")}`;
  elements.selectedAttendanceTotalLabel.textContent = formatDuration(totalMinutes);
}
function resolveSelectedAttendanceEntries(selectedIds, generatedEntries) {
  const resolved = [];
  const usedEntryIds = new Set();
  normalizeSelectedAttendanceIds(selectedIds).forEach((selectedId) => {
    const exactEntry = generatedEntries.find((entry) => entry.id === selectedId);
    if (exactEntry && !usedEntryIds.has(exactEntry.id)) {
      resolved.push(exactEntry);
      usedEntryIds.add(exactEntry.id);
      return;
    }
    const legacyEntry = generatedEntries.find((entry) => entry.date === selectedId && !usedEntryIds.has(entry.id));
    if (legacyEntry) {
      resolved.push(legacyEntry);
      usedEntryIds.add(legacyEntry.id);
      return;
    }
    resolved.push(selectedId);
  });
  return resolved;
}
function getEffectiveUsageAttendanceIds(usageRecord) {
  const selectedIds = normalizeSelectedAttendanceIds(usageRecord.selectedAttendanceIds ?? usageRecord.selectedAttendanceDates ?? usageRecord.selectedAttendanceDate);
  return selectedIds.length ? selectedIds : normalizeAttendanceIds(usageRecord.lockedAttendanceIds, Infinity);
}
function getAllocationAttendanceIds(allocations) {
  return Array.from(new Set((allocations || []).map((allocation) => allocation.attendanceId).filter(Boolean)));
}
function refreshAutoUsageSelections() {
  state.usageRecords = state.usageRecords.map((usageRecord) => {
    const manualSelection = normalizeSelectedAttendanceIds(usageRecord.selectedAttendanceIds ?? usageRecord.selectedAttendanceDates ?? usageRecord.selectedAttendanceDate);
    if (manualSelection.length || !normalizeAttendanceIds(usageRecord.lockedAttendanceIds, Infinity).length) return usageRecord;
    const { lockedAttendanceIds, ...rest } = usageRecord;
    return rest;
  });
  freezeAutoUsageSelections();
}
function freezeAutoUsageSelections() {
  const ledger = buildLedger();
  let hasChanges = false;
  state.usageRecords = state.usageRecords.map((usageRecord) => {
    const manualSelection = normalizeSelectedAttendanceIds(usageRecord.selectedAttendanceIds ?? usageRecord.selectedAttendanceDates ?? usageRecord.selectedAttendanceDate);
    if (manualSelection.length) return usageRecord;
    const lockedAttendanceIds = getAllocationAttendanceIds(ledger.usageAllocations[usageRecord.id] || []);
    if (!lockedAttendanceIds.length) return usageRecord;
    const currentLockedIds = normalizeAttendanceIds(usageRecord.lockedAttendanceIds, Infinity);
    const isSame = currentLockedIds.length === lockedAttendanceIds.length && currentLockedIds.every((id, index) => id === lockedAttendanceIds[index]);
    if (isSame) return usageRecord;
    hasChanges = true;
    return { ...usageRecord, lockedAttendanceIds };
  });
  if (hasChanges) saveState();
}
function shouldShowAttendanceEntry(entry, referenceDate = getTodayString()) {
  if (isEntryExpired(entry, referenceDate)) return false;
  if (entry.remainingMinutes > 0) return true;
  if (isAttendanceHistoryVisible(entry.date, referenceDate)) return true;
  return entry.allocations.some((allocation) => isUsageHistoryVisible(allocation.usageDate, referenceDate));
}
function isEntryExpired(entry, referenceDate = getTodayString()) {
  return compareDate(entry.expiryDate, referenceDate) <= 0;
}
function isExpiredAttendanceDate(attendanceDate, referenceDate = getTodayString()) {
  if (!attendanceDate) return true;
  return compareDate(addDays(attendanceDate, 30), referenceDate) <= 0;
}
function isAttendanceHistoryVisible(attendanceDate, referenceDate = getTodayString()) {
  return compareDate(referenceDate, addDays(attendanceDate, USAGE_HISTORY_VISIBLE_DAYS)) < 0;
}
function isUsageHistoryVisible(usageDate, referenceDate = getTodayString()) {
  return compareDate(referenceDate, addDays(usageDate, USAGE_HISTORY_VISIBLE_DAYS)) < 0;
}
function isWeekendDate(dateString) {
  if (!dateString) return false;
  const day = new Date(`${dateString}T00:00:00`).getDay();
  return day === 0 || day === 6;
}
function resetAllData() {
  if (!confirm("모든 데이터를 삭제할까요? 이 작업은 되돌릴 수 없습니다.")) return;
  state.attendanceRecords = [];
  state.usageRecords = [];
  saveState();
  resetForms();
  populateUsageAttendanceOptions();
  renderAll();
  elements.uploadResult.textContent = "";
}
function upsertClone(list, record) {
  const cloned = cloneRecords(list);
  const index = cloned.findIndex((item) => item.id === record.id);
  if (index >= 0) cloned.splice(index, 1, record); else cloned.push(record);
  return cloned;
}
function cloneRecords(list) { return JSON.parse(JSON.stringify(list)); }
function sortByDateThenId(a, b) { const d = compareDate(a.date, b.date); return d !== 0 ? d : a.id.localeCompare(b.id); }
function createId(prefix) { return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`; }
function toMinutes(timeString) { const [hours, minutes] = timeString.split(":").map(Number); return hours * 60 + minutes; }
function formatTime(totalMinutes) { const normalized = ((totalMinutes % (24 * 60)) + (24 * 60)) % (24 * 60); return `${String(Math.floor(normalized / 60)).padStart(2, "0")}:${String(normalized % 60).padStart(2, "0")}`; }
function floorToUnit(minutes, unit) { return Math.floor(minutes / unit) * unit; }
function formatDuration(minutes) {
  if (!minutes) return "0분";
  const hours = Math.floor(minutes / 60);
  const remainMinutes = minutes % 60;
  if (hours && remainMinutes) return `${hours}시간 ${remainMinutes}분`;
  if (hours) return `${hours}시간`;
  return `${remainMinutes}분`;
}
function formatShortDate(dateString) { const [, month, day] = dateString.split("-"); return `${Number(month)}/${Number(day)}`; }
function addDays(dateString, days) { const date = new Date(`${dateString}T00:00:00`); date.setDate(date.getDate() + days); return toDateString(date); }
function diffDays(fromDateString, toDateString) { return Math.round((new Date(`${toDateString}T00:00:00`) - new Date(`${fromDateString}T00:00:00`)) / 86400000); }
function compareDate(a, b) { return a === b ? 0 : (a < b ? -1 : 1); }
function toDateString(date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`; }
function getTodayString() { return toDateString(new Date()); }
function excelSerialToDate(serial) { return toDateString(new Date(Math.floor(serial - 25569) * 86400 * 1000)); }
