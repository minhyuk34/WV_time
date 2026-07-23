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
 * - 데이터 행은 tr.pumpkinDataRow, Left/Mid가 같은 행 순서로 1:1 대응
 *
 * 주의: 이 그리드는 헤더 <th>가 "라벨+빈 스페이서" 2개씩 짝지어 있는 반면
 * 데이터 <td>는 colspan으로 합쳐져 1개씩만 있어서, 헤더 인덱스를 그대로
 * 데이터 인덱스로 쓸 수 없습니다. 그래서 열 위치는 실제 화면에서 확인한
 * 고정 인덱스를 기본으로 쓰고, 내용 패턴(날짜/시간 모양)으로 검증한 뒤
 * 안 맞으면 행 전체를 스캔해서 찾는 방식으로 이중 안전장치를 둡니다.
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

  function dataRows(table) {
    return Array.prototype.slice.call(table.querySelectorAll('tr.pumpkinDataRow'));
  }
  function cellsOf(row) {
    return Array.prototype.slice.call(row.querySelectorAll('td'));
  }
  function cellText(cell) {
    return cell ? cell.textContent.replace(/ /g, ' ').trim() : '';
  }
  var DATE_RE = /^\d{4}\.\d{2}\.\d{2}$/;
  var DATETIME_RE = /^\d{4}\.\d{2}\.\d{2}\s+\d{2}:\d{2}$/;
  var WORKTYPE_RE = /^[A-D](-1)?형\(\d{2}:\d{2}~\d{2}:\d{2}\)$/;

  // 근무일자: 실제 화면에서 확인한 고정 위치(12번째 td)를 먼저 시도하고,
  // 날짜 모양이 아니면 행 전체에서 날짜 모양 셀을 찾습니다.
  function findDate(cells) {
    var candidate = cellText(cells[12]);
    if (DATE_RE.test(candidate)) return candidate;
    for (var i = 0; i < cells.length; i++) {
      var text = cellText(cells[i]);
      if (DATE_RE.test(text)) return text;
    }
    return '';
  }

  // 출퇴근카드 출근/퇴근일시: 고정 위치(8, 9번째 td)를 먼저 시도.
  // 안 맞으면 행에서 날짜+시간 모양 셀들을 순서대로 모아서, 세 번째/네 번째를
  // 출퇴근카드 값으로 사용합니다 (첫 두 개는 "정상근무기준시간" 출근/퇴근).
  function findStartEnd(cells) {
    var start = cellText(cells[8]);
    var end = cellText(cells[9]);
    if (DATETIME_RE.test(start)) return { start: start, end: DATETIME_RE.test(end) ? end : '' };
    var matches = [];
    for (var i = 0; i < cells.length; i++) {
      var text = cellText(cells[i]);
      if (DATETIME_RE.test(text)) matches.push(text);
    }
    return { start: matches[2] || '', end: matches[3] || '' };
  }

  // 근무유형(예: "A형(07:00~16:00)"): 고정 위치(2 또는 3번째 td)를 먼저 시도.
  function findWorkType(cells) {
    var candidate = cellText(cells[2]) || cellText(cells[3]);
    if (WORKTYPE_RE.test(candidate)) return candidate;
    for (var i = 0; i < cells.length; i++) {
      var text = cellText(cells[i]);
      if (WORKTYPE_RE.test(text)) return text;
    }
    return '';
  }

  var leftRows = dataRows(leftTable);
  var midRows = dataRows(midTable);
  var records = [];
  for (var r = 0; r < leftRows.length; r++) {
    var midRow = midRows[r];
    if (!midRow) continue;
    var leftCells = cellsOf(leftRows[r]);
    var midCells = cellsOf(midRow);
    var dateText = findDate(leftCells);
    var times = findStartEnd(midCells);
    var workTypeText = findWorkType(midCells);
    // 출근 기록이 없는 행(휴일, 미래 날짜 등)은 건너뜀
    if (!dateText || !times.start) continue;
    records.push([dateText, workTypeText, times.start, times.end]);
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
