const STORAGE_KEY = "specificDayWorkManager_v1";
const FLOOR_UNIT = 30;
const MAX_DAILY_USAGE_MINUTES = 240;
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
const state = loadState();
const elements = {
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
  usageAttendanceDate: document.getElementById("usageAttendanceDate"),
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
  usageList: document.getElementById("usageList"),
  timelineList: document.getElementById("timelineList"),
  attendanceCount: document.getElementById("attendanceCount"),
  usageCount: document.getElementById("usageCount"),
  excelFile: document.getElementById("excelFile"),
  uploadBtn: document.getElementById("uploadBtn"),
  uploadResult: document.getElementById("uploadResult"),
  seedDemoBtn: document.getElementById("seedDemoBtn"),
  resetBtn: document.getElementById("resetBtn")
};
initialize();
function initialize() {
  populateWorkTypeSelects();
  populateUsageTimeSelects();
  populateUsageAttendanceOptions();
  elements.usageStart.step = "1800";
  elements.usageEnd.step = "1800";
  bindEvents();
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
function populateUsageAttendanceOptions(selectedDate = elements.usageAttendanceDate?.value || "") {
  const usageDate = elements.usageDate.value;
  const options = ['<option value="">자동 선택(FIFO)</option>'];
  getSelectableAttendanceEntries(usageDate).forEach((entry) => {
    const isFutureRecord = Boolean(usageDate) && compareDate(entry.date, usageDate) > 0;
    const isExpiredRecord = Boolean(usageDate) && compareDate(entry.expiryDate, usageDate) < 0;
    const isDisabled = isFutureRecord || isExpiredRecord;
    options.push(`<option value="${entry.date}" ${isDisabled ? "disabled" : ""}>${formatSelectableAttendanceLabel(entry)}</option>`);
  });
  elements.usageAttendanceDate.innerHTML = options.join("");
  const selectedOptionExists = Array.from(elements.usageAttendanceDate.options).some((option) => option.value === selectedDate && !option.disabled);
  elements.usageAttendanceDate.value = selectedOptionExists ? selectedDate : "";
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
  elements.todayWorkType.addEventListener("change", renderSummary);
  elements.uploadBtn.addEventListener("click", () => handleExcelUpload());
  elements.seedDemoBtn.addEventListener("click", seedDemoData);
  elements.resetBtn.addEventListener("click", resetAllData);
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
function saveState() { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
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
  populateUsageAttendanceOptions("");
  elements.usageFormMode.textContent = "신규 등록";
  elements.usageFormMode.className = "pill neutral";
  elements.computedUsageLabel.textContent = "0분";
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
  renderUsageList();
  renderTimelineList();
}
function getSelectableAttendanceEntries(usageDate = "") {
  const excludeUsageId = elements.usageId?.value || "";
  return buildLedger({ excludeUsageId }).generatedEntries
    .filter((entry) => entry.earnedMinutes > 0 && entry.remainingMinutes > 0)
    .filter((entry) => !usageDate || compareDate(entry.date, usageDate) <= 0)
    .filter((entry) => !usageDate || compareDate(entry.expiryDate, usageDate) >= 0)
    .sort(sortByDateThenId);
}
function formatSelectableAttendanceLabel(entry) {
  const schedule = WORK_TYPES[resolveWorkTypeKey(entry.workType)];
  return `${entry.date} · ${schedule.label} · 사용 ${formatDuration(entry.usedMinutes)} / 잔여 ${formatDuration(entry.remainingMinutes)}`;
}
function getSelectedAttendanceLabel(selectedAttendanceDate) {
  if (!selectedAttendanceDate) return "자동 선택(FIFO)";
  const record = state.attendanceRecords.find((item) => item.date === selectedAttendanceDate);
  if (!record) return `${selectedAttendanceDate} · 삭제된 기록`;
  const entry = buildLedger().generatedEntries.find((item) => item.date === selectedAttendanceDate);
  if (!entry) {
    const generatedMinutes = record.generatedMinutes ?? calculateGeneratedMinutes(record);
    return `${record.date} · ${WORK_TYPES[resolveWorkTypeKey(record.workType)].label} · 사용 0분 / 잔여 ${formatDuration(generatedMinutes)}`;
  }
  return formatSelectableAttendanceLabel(entry);
}
function handleAttendanceSubmit(event) {
  event.preventDefault();
  const record = { id: elements.attendanceId.value || createId("attendance"), date: elements.attendanceDate.value, workType: elements.attendanceWorkType.value, actualStart: elements.actualStart.value, actualEnd: elements.actualEnd.value, overtime: elements.overtimeChecked.checked, overtimeChecked: elements.overtimeChecked.checked, source: "manual" };
  const validationMessage = validateAttendanceRecord(record);
  if (validationMessage) return alert(validationMessage);
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
  const record = { id: elements.usageId.value || createId("usage"), date: elements.usageDate.value, workType: elements.usageWorkType.value, startTime: normalizedStartTime, endTime: normalizedEndTime, selectedAttendanceDate: elements.usageAttendanceDate.value };
  const validationMessage = validateUsageRecord(record);
  if (validationMessage) return alert(validationMessage);
  const durationMinutes = calculateUsageMinutes(record.startTime, record.endTime);
  if (durationMinutes <= 0) return alert("사용시간이 0분입니다. 30분 단위 기준으로 다시 입력하세요.");
  const simulation = buildLedger({ usageRecordOverride: { ...record, durationMinutes } });
  if (simulation.invalidUsageIds.includes(record.id)) return alert("사용 가능한 특정일 시간이 부족합니다. 유효한 발생시간 범위를 확인하세요.");
  const existingIndex = state.usageRecords.findIndex((item) => item.id === record.id);
  if (existingIndex >= 0) state.usageRecords.splice(existingIndex, 1, { ...record, durationMinutes });
  else state.usageRecords.push({ ...record, durationMinutes });
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
  if (record.selectedAttendanceDate && compareDate(record.selectedAttendanceDate, record.date) > 0) return "사용일보다 미래의 발생기록은 선택할 수 없습니다.";
  if (record.selectedAttendanceDate && !state.attendanceRecords.some((item) => item.date === record.selectedAttendanceDate)) return "선택한 발생기록 날짜를 찾을 수 없습니다.";
  return "";
}
function renderAttendancePreview() {
  if (!elements.attendanceWorkType.value || !elements.actualStart.value || !elements.actualEnd.value) return elements.computedEarnedLabel.textContent = "0분";
  const minutes = calculateEarnedMinutes({ workType: elements.attendanceWorkType.value, actualStart: elements.actualStart.value, actualEnd: elements.actualEnd.value, overtime: elements.overtimeChecked.checked, overtimeChecked: elements.overtimeChecked.checked });
  elements.computedEarnedLabel.textContent = formatDuration(minutes);
}
function renderUsagePreview() {
  syncUsageTimeField("start");
  syncUsageTimeField("end");
  if (!elements.usageStart.value || !elements.usageEnd.value) return elements.computedUsageLabel.textContent = "0분";
  const duration = calculateUsageMinutes(elements.usageStart.value, elements.usageEnd.value);
  elements.computedUsageLabel.textContent = duration > 0 ? formatDuration(duration) : "0분";
}
function calculateEarnedMinutes(record) {
  const schedule = WORK_TYPES[resolveWorkTypeKey(record.workType)];
  if (!schedule) return 0;
  const scheduledStart = toMinutes(schedule.start);
  const scheduledEnd = toMinutes(schedule.end);
  const actualStart = record.actualStart ? toMinutes(record.actualStart) : scheduledStart;
  const actualEnd = record.actualEnd ? toMinutes(record.actualEnd) : scheduledEnd;
  const earlyEarned = floorToUnit(Math.max(0, scheduledStart - actualStart), FLOOR_UNIT);
  const lateBase = Boolean(record.overtime ?? record.overtimeChecked) ? scheduledEnd + 150 : scheduledEnd;
  const lateEarned = floorToUnit(Math.max(0, actualEnd - lateBase), FLOOR_UNIT);
  return earlyEarned + lateEarned;
}
function calculateGeneratedMinutes(record) { return calculateEarnedMinutes(record); }
function buildGeneratedTimeRanges(record) {
  const schedule = WORK_TYPES[resolveWorkTypeKey(record.workType)];
  if (!schedule) return [];
  const scheduledStart = toMinutes(schedule.start);
  const scheduledEnd = toMinutes(schedule.end);
  const actualStart = record.actualStart ? toMinutes(record.actualStart) : scheduledStart;
  const actualEnd = record.actualEnd ? toMinutes(record.actualEnd) : scheduledEnd;
  const ranges = [];
  const earlyEarned = floorToUnit(Math.max(0, scheduledStart - actualStart), FLOOR_UNIT);
  if (earlyEarned > 0) ranges.push({ start: scheduledStart - earlyEarned, end: scheduledStart, minutes: earlyEarned });
  const lateBase = Boolean(record.overtime ?? record.overtimeChecked) ? scheduledEnd + 150 : scheduledEnd;
  const lateEarned = floorToUnit(Math.max(0, actualEnd - lateBase), FLOOR_UNIT);
  if (lateEarned > 0) ranges.push({ start: lateBase, end: lateBase + lateEarned, minutes: lateEarned });
  return ranges;
}
function calculateUsageMinutes(startTime, endTime) {
  const normalized = normalizeRangeToHalfHour(startTime, endTime);
  return normalized ? calculateUsageMinutesExcludingLunch(normalized.start, normalized.end) : 0;
}
function buildLedger(options = {}) {
  const attendanceRecords = cloneRecords(options.attendanceRecordOverride ? upsertClone(state.attendanceRecords, options.attendanceRecordOverride) : state.attendanceRecords).sort(sortByDateThenId);
  const sourceUsageRecords = options.usageRecordOverride ? upsertClone(state.usageRecords, options.usageRecordOverride) : state.usageRecords;
  const filteredUsageRecords = options.excludeUsageId ? sourceUsageRecords.filter((item) => item.id !== options.excludeUsageId) : sourceUsageRecords;
  const usageRecords = cloneRecords(filteredUsageRecords)
    .map((item) => ({ ...item, startTime: normalizeUsageTimeInput(item.startTime), endTime: normalizeUsageTimeInput(item.endTime), durationMinutes: item.durationMinutes ?? calculateUsageMinutes(item.startTime, item.endTime) }))
    .sort(sortByDateThenId);
  const generatedEntries = attendanceRecords.map((record) => {
    const generatedRanges = buildGeneratedTimeRanges(record);
    const earnedMinutes = generatedRanges.reduce((sum, range) => sum + range.minutes, 0);
    return { ...record, earnedMinutes, generatedRanges, remainingRangeBuckets: generatedRanges.map((range) => ({ ...range })), expiryDate: addDays(record.date, 30), usedMinutes: 0, remainingMinutes: earnedMinutes, allocations: [] };
  });
  const invalidUsageIds = [];
  const usageAllocations = {};
  for (const usage of usageRecords) {
    let remainingUsage = usage.durationMinutes;
    const allocations = [];
    for (const entry of generatedEntries) {
      if (remainingUsage <= 0) break;
      if (entry.earnedMinutes <= 0 || entry.remainingMinutes <= 0) continue;
      if (usage.selectedAttendanceDate && entry.date !== usage.selectedAttendanceDate) continue;
      if (compareDate(entry.expiryDate, usage.date) < 0) continue;
      if (compareDate(entry.date, usage.date) > 0) continue;
      const allocated = Math.min(entry.remainingMinutes, remainingUsage);
      if (allocated <= 0) continue;
      entry.usedMinutes += allocated;
      entry.remainingMinutes -= allocated;
      remainingUsage -= allocated;
      const attendanceRangeLabel = allocateFromGeneratedRanges(entry, allocated);
      allocations.push({ attendanceId: entry.id, attendanceDate: entry.date, attendanceRangeLabel, usageId: usage.id, usageDate: usage.date, usageStart: usage.startTime, usageEnd: usage.endTime, minutes: allocated });
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
  const validEntries = ledger.generatedEntries.filter((entry) => entry.earnedMinutes > 0 && entry.remainingMinutes > 0 && compareDate(entry.expiryDate, today) >= 0);
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
  const visibleEntries = ledger.generatedEntries.filter((entry) => entry.earnedMinutes > 0);
  elements.attendanceCount.textContent = `${visibleEntries.length}건`;
  renderCollection(elements.attendanceList, visibleEntries.map(renderAttendanceItem));
}
function renderUsageList() {
  const ledger = buildLedger();
  elements.usageCount.textContent = `${ledger.usageRecords.length}건`;
  renderCollection(elements.usageList, ledger.usageRecords.map((usage) => renderUsageItem(usage, ledger)));
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
  const usedDetails = entry.allocations.length ? entry.allocations.map((allocation) => formatUsageRange({ date: allocation.usageDate, startTime: allocation.usageStart, endTime: allocation.usageEnd }, allocation.minutes)).join(" / ") : "아직 사용되지 않음";
  item.className = `list-item ${status.className}`.trim();
  item.innerHTML = `
    <div class="item-row"><div><div class="item-title">${entry.date} · ${schedule.label}</div><div class="item-subtitle">실제 ${entry.actualStart || "--:--"} ~ ${entry.actualEnd || "--:--"} · <label class="checkbox-field"><input type="checkbox" data-action="toggle-overtime" data-id="${entry.id}" ${overtimeChecked ? "checked" : ""}><span>시간외근무</span></label></div></div><div class="status-row"><span class="pill ${status.pillClass}">${status.label}</span><span class="pill neutral">만료 ${entry.expiryDate}</span></div></div>
    <div class="detail-grid"><div class="detail-box"><span>발생내역</span><strong>${formatGeneratedRanges(entry)}</strong></div><div class="detail-box"><span>발생시간</span><strong>${formatDuration(entry.earnedMinutes)}</strong></div><div class="detail-box"><span>남은시간</span><strong>${formatDuration(entry.remainingMinutes)}</strong></div><div class="detail-box"><span>차감 내역</span><strong>${usedDetails}</strong></div></div>
    <div class="item-actions"><button class="mini-btn" type="button" data-action="edit-attendance" data-id="${entry.id}">수정</button><button class="mini-btn danger" type="button" data-action="delete-attendance" data-id="${entry.id}">삭제</button></div>`;
  bindItemActions(item);
  return item;
}
function renderUsageItem(usage, ledger) {
  const item = document.createElement("article");
  const isInvalid = ledger.invalidUsageIds.includes(usage.id);
  const allocations = ledger.usageAllocations[usage.id] || [];
  const deductionText = allocations.length ? allocations.map((allocation) => `${allocation.attendanceRangeLabel}에서 ${formatDuration(allocation.minutes)}`).join(" / ") : "차감 내역 없음";
  const detailColumns = "repeat(auto-fit, minmax(150px, 1fr))";
  item.className = `list-item ${isInvalid ? "warning" : ""}`.trim();
  item.innerHTML = `
    <div class="item-row"><div><div class="item-title">${usage.date} · ${WORK_TYPES[usage.workType].label}</div><div class="item-subtitle">${formatUsageRange(usage)}</div></div><div class="status-row"><span class="pill ${isInvalid ? "warning" : "info"}">${isInvalid ? "차감 불가" : "차감 완료"}</span></div></div>
    <div class="detail-grid" style="grid-template-columns:${detailColumns}"><div class="detail-box"><span>사용내역</span><strong>${formatUsageRange(usage)}</strong></div><div class="detail-box"><span>사용시간</span><strong>${formatDuration(usage.durationMinutes)}</strong></div><div class="detail-box"><span>선택 발생기록</span><strong>${getSelectedAttendanceLabel(usage.selectedAttendanceDate)}</strong></div><div class="detail-box"><span>차감 출처</span><strong>${deductionText}</strong></div><div class="detail-box"><span>상태</span><strong>${isInvalid ? "유효한 발생시간 부족" : "정상 저장"}</strong></div></div>
    <div class="item-actions"><button class="mini-btn" type="button" data-action="edit-usage" data-id="${usage.id}">수정</button><button class="mini-btn danger" type="button" data-action="delete-usage" data-id="${usage.id}">삭제</button></div>`;
  bindItemActions(item);
  return item;
}
function buildLedgerRows(ledger) {
  const rows = [];
  ledger.generatedEntries.filter((entry) => entry.earnedMinutes > 0 && !entry.allocations.length).forEach((entry) => rows.push({ type: "미사용", usageId: `unused_${entry.id}`, usageLabel: "미사용", attendanceItems: [formatGeneratedRanges(entry)], flowLabel: "적립", minutes: entry.earnedMinutes }));
  ledger.usageRecords.forEach((usage) => {
    const allocations = ledger.usageAllocations[usage.id] || [];
    if (!allocations.length) return;
    rows.push({ type: "차감", usageId: usage.id || `${usage.date}_${usage.startTime}_${usage.endTime}`, usageLabel: formatUsageRange(usage), attendanceItems: allocations.map((allocation) => allocation.attendanceRangeLabel), flowLabel: "차감", minutes: usage.durationMinutes });
  });
  return groupLedgerRowsByUsage(rows);
}
function renderTimelineItem(row) {
  const item = document.createElement("article");
  item.className = "list-item";
  item.innerHTML = `
    <div class="item-row"><div><div class="item-title">${row.type}</div><div class="item-subtitle">사용기록 1건 기준 연결 내역</div></div><span class="pill ${row.flowLabel === "적립" ? "info" : "warning"}">${row.flowLabel}</span></div>
    <div class="detail-grid"><div class="detail-box"><span>사용내역</span><strong>${row.usageLabel}</strong></div><div class="detail-box"><span>발생내역</span><strong>${row.attendanceItems.join("<br>")}</strong></div><div class="detail-box"><span>차감시간</span><strong>${formatDuration(row.minutes)}</strong></div><div class="detail-box"><span>흐름</span><strong>${row.flowLabel}</strong></div></div>`;
  return item;
}
function bindItemActions(container) {
  container.querySelectorAll("[data-action]").forEach((button) => button.addEventListener("click", () => {
    const action = button.dataset.action;
    const id = button.dataset.id;
    if (action === "toggle-overtime") return;
    if (action === "edit-attendance") startAttendanceEdit(id);
    if (action === "delete-attendance") deleteAttendance(id);
    if (action === "edit-usage") startUsageEdit(id);
    if (action === "delete-usage") deleteUsage(id);
  }));
  container.querySelectorAll('[data-action="toggle-overtime"]').forEach((checkbox) => checkbox.addEventListener("change", (event) => {
    toggleOvertimeFromList(event.target.dataset.id, event.target.checked);
  }));
}
function toggleOvertimeFromList(recordId, checked) {
  const recordIndex = state.attendanceRecords.findIndex((item) => item.id === recordId);
  if (recordIndex < 0) return;
  const updatedRecord = { ...state.attendanceRecords[recordIndex], overtime: checked, overtimeChecked: checked };
  const generatedMinutes = calculateGeneratedMinutes(updatedRecord);
  if (generatedMinutes <= 0) {
    state.attendanceRecords.splice(recordIndex, 1);
    if (elements.attendanceId.value === recordId) resetAttendanceForm();
  } else {
    state.attendanceRecords.splice(recordIndex, 1, { ...updatedRecord, generatedMinutes });
    if (elements.attendanceId.value === recordId) elements.overtimeChecked.checked = checked;
  }
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
  saveState();
  populateUsageAttendanceOptions();
  renderAll();
}
function startUsageEdit(id) {
  const record = state.usageRecords.find((item) => item.id === id);
  if (!record) return;
  elements.usageId.value = record.id;
  elements.usageDate.value = record.date;
  populateUsageAttendanceOptions(record.selectedAttendanceDate || "");
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
  saveState();
  renderAll();
}
async function handleExcelUpload(file = elements.excelFile.files[0]) {
  if (!file) return alert("업로드할 파일을 선택하세요.");
  try {
    const rows = await parseSpreadsheetFile(file);
    const parsed = parseAttendanceXlsRows(rows);
    if (!parsed.parsedRecords.length) throw new Error("업로드 가능한 근태 데이터가 없습니다.");
    const existingRecordsLength = state.attendanceRecords.length;
    let zeroMinuteExcludedCount = 0;
    let duplicateSkippedCount = 0;
    const existingDates = new Set(state.attendanceRecords.map((record) => record.date));
    const appendedRecords = [];
    parsed.parsedRecords.forEach((record) => {
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
    saveAttendanceRecords();
    populateUsageAttendanceOptions();
    rerenderAll();
    const appendedCount = appendedRecords.length;
    const finalSavedCount = state.attendanceRecords.length;
    const finalRenderedCount = buildLedger().generatedEntries.filter((entry) => entry.earnedMinutes > 0).length;
    console.log("existingRecords.length", existingRecordsLength);
    console.log("parsedRecords.length", parsed.parsedRecords.length);
    console.log("zeroMinuteExcludedCount", zeroMinuteExcludedCount);
    console.log("duplicateSkippedCount", duplicateSkippedCount);
    console.log("appendedCount", appendedCount);
    console.log("finalSavedCount", finalSavedCount);
    console.log("finalRenderedCount", finalRenderedCount);
    elements.uploadResult.textContent = `업로드 완료: 실제 데이터 ${parsed.dataRowCount}건, 유효 파싱 ${parsed.parsedRecords.length}건, 0분 제외 ${zeroMinuteExcludedCount}건, 기존 날짜 중복 ${duplicateSkippedCount}건, 신규 추가 ${appendedCount}건, 최종 저장 ${finalSavedCount}건, 최종 표시 ${finalRenderedCount}건`;
  } catch (error) {
    console.error(error);
    elements.uploadResult.textContent = `업로드 실패: ${error.message}`;
    alert(`업로드 실패: ${error.message}`);
  }
}
async function parseSpreadsheetFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  if (extension !== "xls" && extension !== "xlsx") throw new Error("지원하지 않는 파일 형식입니다. 회사 근태 엑셀(.xls/.xlsx) 파일을 업로드하세요.");
  return parseExcelWithSheetJs(file);
}
async function parseExcelWithSheetJs(file) {
  if (typeof XLSX === "undefined") throw new Error("엑셀 라이브러리를 불러오지 못했습니다. xlsx.full.min.js 로드 여부를 확인하세요.");
  const data = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(data, { type: "array", codepage: 949, cellDates: false });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) throw new Error("첫 번째 시트를 찾지 못했습니다.");
  const worksheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });
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
  if (!Array.isArray(rows) || rows.length <= 2) return { parsedRecords: [], dataRowCount: 0, excludedRowCount: 0 };
  const dataRows = rows.slice(2).filter((row) => Array.isArray(row) && row.some((cell) => String(cell ?? "").trim() !== ""));
  const parsedRecords = [];
  let excludedRowCount = 0;
  dataRows.forEach((row) => {
    const cells = Array.isArray(row) ? row : [];
    const date = normalizeExcelDate(cells[5]);
    const workType = extractWorkType(cells[8]);
    const actualStart = normalizeExcelTime(cells[12]);
    const actualEnd = normalizeExcelTime(cells[13]);
    if (!date) {
      excludedRowCount += 1;
      return;
    }
    if (!actualStart && !actualEnd) {
      excludedRowCount += 1;
      return;
    }
    parsedRecords.push({ id: createId("attendance"), date, workType: workType || "C형", actualStart, actualEnd, overtime: false, overtimeChecked: false, source: "import" });
  });
  return { parsedRecords, dataRowCount: dataRows.length, excludedRowCount };
}
function normalizeExcelDate(value) {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return toDateString(value);
  if (typeof value === "number" && Number.isFinite(value) && value > 20000) return excelSerialToDate(value);
  const text = String(value).trim();
  const match = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
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
function resetAllData() {
  if (!confirm("모든 localStorage 데이터를 삭제할까요? 이 작업은 되돌릴 수 없습니다.")) return;
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
