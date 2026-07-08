// 배포 시 아래 두 값만 채우면 됩니다. 둘 다 비어 있으면 이 브라우저(localStorage)에만
// 저장되는 "로컬 전용 모드"로 자동 동작하며, Google 로그인 화면은 표시되지 않습니다.
window.APP_CONFIG = {
  // Google Cloud Console > API 및 서비스 > 사용자 인증 정보에서 발급한 OAuth 클라이언트 ID
  GOOGLE_CLIENT_ID: "",

  // Google Apps Script를 "웹 앱"으로 배포한 뒤 나오는 URL (…/exec 로 끝남)
  API_BASE_URL: ""
};
