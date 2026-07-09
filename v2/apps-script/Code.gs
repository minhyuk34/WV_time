/**
 * 특정일 근무시간 관리 - Google Apps Script 백엔드
 *
 * 이 스크립트는 이 스크립트가 바인딩된 Google Sheets 문서를 사용자별 데이터 저장소로 사용합니다.
 * "UserData" 시트에 사용자(email)당 한 행씩 발생기록/사용기록 JSON을 저장합니다.
 *
 * 설정 방법은 저장소의 setup-guide.md 를 참고하세요.
 */

// ===== 설정: 아래 두 값을 채우세요 =====
// index.html에서 로그인할 때 쓰는 OAuth 클라이언트 ID와 동일해야 합니다.
const OAUTH_CLIENT_ID = "252739518185-6gahljodsr32a38qa79nodrjs4v0kov7.apps.googleusercontent.com";
// 회사 구글 워크스페이스 도메인. 이 도메인 계정만 로그인 허용합니다. 제한하지 않으려면 "" 로 두세요.
const ALLOWED_DOMAIN = "";

const SHEET_NAME = "UserData";
const LOCK_TIMEOUT_MS = 10000;

function doGet(e) {
  return handleRequest(e.parameter || {});
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents || "{}");
  } catch (err) {
    return jsonResponse({ error: "요청 본문을 해석할 수 없습니다." });
  }
  return handleRequest(body);
}

function handleRequest(params) {
  try {
    const action = params.action;
    const email = verifyIdTokenAndGetEmail(params.idToken);

    if (action === "load") {
      return jsonResponse(loadUserData(email));
    }
    if (action === "save") {
      const attendanceRecords = Array.isArray(params.attendanceRecords) ? params.attendanceRecords : [];
      const usageRecords = Array.isArray(params.usageRecords) ? params.usageRecords : [];
      saveUserData(email, attendanceRecords, usageRecords);
      return jsonResponse({ ok: true });
    }
    return jsonResponse({ error: `지원하지 않는 action입니다: ${action}` });
  } catch (err) {
    return jsonResponse({ error: String((err && err.message) || err) });
  }
}

function verifyIdTokenAndGetEmail(idToken) {
  if (!idToken) throw new Error("로그인이 필요합니다.");
  const response = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (response.getResponseCode() !== 200) throw new Error("유효하지 않거나 만료된 로그인 토큰입니다. 다시 로그인해주세요.");
  const payload = JSON.parse(response.getContentText());
  if (OAUTH_CLIENT_ID && payload.aud !== OAUTH_CLIENT_ID) throw new Error("허용되지 않은 클라이언트에서 온 요청입니다.");
  if (payload.email_verified !== "true" && payload.email_verified !== true) throw new Error("이메일 인증이 확인되지 않은 계정입니다.");
  if (ALLOWED_DOMAIN && payload.hd !== ALLOWED_DOMAIN) throw new Error(`${ALLOWED_DOMAIN} 소속 계정으로만 로그인할 수 있습니다.`);
  return payload.email;
}

function getUserDataSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(["email", "updatedAt", "attendanceJson", "usageJson"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findUserRowIndex(sheet, email) {
  const emails = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (emails[i][0] === email) return i + 2; // 1-indexed, +1 for header row
  }
  return -1;
}

function loadUserData(email) {
  const sheet = getUserDataSheet();
  const rowIndex = findUserRowIndex(sheet, email);
  if (rowIndex < 0) return { attendanceRecords: [], usageRecords: [] };
  const row = sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
  return {
    attendanceRecords: safeParseJson(row[2], []),
    usageRecords: safeParseJson(row[3], [])
  };
}

function saveUserData(email, attendanceRecords, usageRecords) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const sheet = getUserDataSheet();
    const rowIndex = findUserRowIndex(sheet, email);
    const rowValues = [email, new Date().toISOString(), JSON.stringify(attendanceRecords), JSON.stringify(usageRecords)];
    if (rowIndex < 0) {
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIndex, 1, 1, 4).setValues([rowValues]);
    }
  } finally {
    lock.releaseLock();
  }
}

function safeParseJson(text, fallback) {
  if (!text) return fallback;
  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch (err) {
    return fallback;
  }
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
