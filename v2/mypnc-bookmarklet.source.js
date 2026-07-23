/**
 * 마이P&C "출퇴근관리" 화면에서 조회된 근태 표를 긁어서
 * v2 사이트(#import=...)로 자동 전달하는 북마크릿의 읽기 쉬운 원본.
 *
 * 실제로 브라우저 북마크에 등록할 때는 이 파일이 아니라 한 줄로 압축된
 * javascript: 버전을 사용해야 합니다 (setup-guide.md 참고).
 *
 * 대상 화면 DOM 구조 (IBSheet/Pumpkin 그리드):
 * - .pumpkinBodyLeft table.pumpkinSection : 고정열(번호~휴일)
 * - .pumpkinBodyMid  table.pumpkinSection : 스크롤열(근무유형~예외사유)
 * - 각 표의 첫 <tr>은 숨겨진 접근성용 <th> 헤더 목록 (열 이름 인식용)
 * - 데이터 행은 tr.pumpkinDataRow, Left/Mid가 같은 행 순서로 1:1 대응
 */
(function () {
  // 화면이 iframe(하위 프레임) 안에 있을 수 있어서, 최상위 문서부터 시작해 모든 프레임을 재귀적으로 뒤집니다.
  function findInDocument(doc, selector) {
    try {
      var found = doc.querySelector(selector);
      if (found) return found;
    } catch (e) {
      return null;
    }
    var frames = doc.querySelectorAll('iframe, frame');
    for (var i = 0; i < frames.length; i++) {
      try {
        var innerDoc = frames[i].contentDocument;
        if (!innerDoc) continue;
        var result = findInDocument(innerDoc, selector);
        if (result) return result;
      } catch (e) {
        // 다른 도메인 iframe이면 접근이 막히는데, 이 사이트 내부 프레임이라면 문제 없음
      }
    }
    return null;
  }
  function findTable(cls) {
    return findInDocument(document, '.' + cls + ' table.pumpkinSection');
  }
  var leftTable = findTable('pumpkinBodyLeft');
  var midTable = findTable('pumpkinBodyMid');
  if (!leftTable || !midTable) {
    alert('출퇴근관리 표를 찾을 수 없습니다. 이 화면(출퇴근관리)에서 조회를 먼저 실행한 뒤 다시 시도하세요.');
    return;
  }

  function headerIndex(table, keyword, useLast) {
    var headerRow = table.querySelector('tr');
    var ths = headerRow.querySelectorAll('th');
    var found = -1;
    var target = keyword.replace(/\s+/g, '');
    for (var i = 0; i < ths.length; i++) {
      var span = ths[i].querySelector('span');
      var text = (span ? span.textContent : ths[i].textContent || '').replace(/\s+/g, '');
      if (text.indexOf(target) !== -1) {
        found = i;
        if (!useLast) break;
      }
    }
    return found;
  }

  var dateIdx = headerIndex(leftTable, '근무일자', false);
  // 근무유형 헤더가 3번(코드/라벨/라벨) 나오는데, 사람이 읽는 라벨(A형(07:00~16:00))이 마지막에 있어 useLast=true
  var workTypeIdx = headerIndex(midTable, '근무유형', true);
  var startIdx = headerIndex(midTable, '출퇴근카드출근일시', false);
  var endIdx = headerIndex(midTable, '출퇴근카드퇴근일시', false);

  if (dateIdx < 0 || startIdx < 0 || endIdx < 0) {
    alert('표 열 구성을 인식하지 못했습니다 (화면 구조가 바뀐 것 같습니다). 개발자에게 문의하세요.');
    return;
  }

  function dataRows(table) {
    return Array.prototype.slice.call(table.querySelectorAll('tr.pumpkinDataRow'));
  }
  function cellText(row, idx) {
    var cells = row.querySelectorAll('td');
    var cell = cells[idx];
    if (!cell) return '';
    return cell.textContent.replace(/ /g, '').trim();
  }

  var leftRows = dataRows(leftTable);
  var midRows = dataRows(midTable);
  var records = [];
  for (var r = 0; r < leftRows.length; r++) {
    var midRow = midRows[r];
    if (!midRow) continue;
    var dateText = cellText(leftRows[r], dateIdx);
    var startText = cellText(midRow, startIdx);
    var endText = cellText(midRow, endIdx);
    var workTypeText = workTypeIdx >= 0 ? cellText(midRow, workTypeIdx) : '';
    // 출근 기록이 없는 행(휴일 등)은 건너뜀
    if (!dateText || !startText) continue;
    records.push([dateText, workTypeText, startText, endText]);
  }

  if (!records.length) {
    alert('가져올 근태 데이터가 없습니다. 출퇴근 기록이 있는 기간으로 조회했는지 확인하세요.');
    return;
  }

  // v2 script.js의 parseAttendanceXlsRows()가 그대로 인식하는 헤더+데이터 2차원 배열 형태
  var rows = [['근무일자', '근무유형', '실제출근시간', '실제퇴근시간']].concat(records);
  var json = JSON.stringify(rows);
  var b64 = btoa(unescape(encodeURIComponent(json)));
  var encoded = encodeURIComponent(b64);
  var url = 'https://minhyuk34.github.io/WV_time/v2/#import=' + encoded;

  if (confirm(records.length + '건의 근태 데이터를 찾았습니다. v2 사이트를 새 탭으로 열어서 반영할까요?')) {
    window.open(url, '_blank');
  }
})();
