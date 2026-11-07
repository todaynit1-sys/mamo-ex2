/**
 * 푸른하늘 업무체험관 방문 통계 → 구글 시트 (매일 자동)
 *
 * 구글 애널리틱스(GA4) 숫자를 구글 시트 '일별 통계'에 날짜별로 적습니다.
 * 구글 서버에서 돌아가서 컴퓨터를 꺼 둬도 매일 아침 채워집니다.
 *
 * 처음 한 번 설정
 * 1. 구글 시트에서 확장 프로그램 → Apps Script를 열고 이 코드를 붙여 넣습니다.
 * 2. 왼쪽 '서비스' 옆 + 를 눌러 'Google Analytics Data API'를 추가합니다.
 * 3. 위쪽 함수 목록에서 setup을 골라 실행하고, 권한을 허용합니다.
 *    → '일별 통계' 시트가 생기고, 매일 아침 6~7시에 최근 7일 숫자를 다시 받아 채웁니다.
 * 시트 메뉴 '체험관 통계 → 지금 업데이트'로 오늘 숫자까지 바로 받아 올 수도 있습니다.
 */
const PROPERTY_ID = '563107418'; // 구글 애널리틱스 '푸른하늘 업무체험관' 속성
const SHEET_NAME = '일별 통계';
const BASE = [ // [GA 측정항목, 머리글]
  ['activeUsers', '방문자(명)'],
  ['newUsers', '처음 온 사람(명)'],
  ['sessions', '방문 횟수'],
  ['screenPageViews', '페이지뷰'],
];
const EVENTS = [ // [앱 이벤트, 머리글, count = 횟수 / users = 사람 수]
  ['game_start', '체험 시작(회)', 'count'],
  ['mission_done', '업무 완료(회)', 'count'],
  ['all_main_done', '대표 업무 4개 완료(명)', 'users'],
  ['all_16_done', '16가지 모두 완료(명)', 'users'],
  ['chief_meet', '청장님 소개(명)', 'users'],
  ['cert_photo', '인증샷 찍은 사람(명)', 'users'],
  ['cert_photo', '인증샷(장)', 'count'],
  ['chief_photo', '기념사진(장)', 'count'],
  ['photo_keep', '사진 저장·공유(회)', 'count'],
  ['event_entry', '이벤트 응모 클릭(명)', 'users'],
  ['share', '체험관 공유(회)', 'count'],
  ['play_together', '다른 사람과 만남(명)', 'users'],
  ['group_walk', '같이 다니기(회)', 'count'],
  ['dust_game', '미세먼지 잡기(회)', 'count'],
  ['meong', '멍때리기(회)', 'count'],
];
const HEAD = ['날짜'].concat(BASE.map((b) => b[1]), EVENTS.map((e) => e[1]));

function setup() { // 처음 한 번: 시트 만들기, 매일 실행 예약, 지난 30일 채우기
  getSheet_();
  ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === 'dailyUpdate').forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyUpdate').timeBased().everyDays(1).atHour(6).inTimezone('Asia/Seoul').create();
  update_(30, true);
  SpreadsheetApp.getActive().toast('매일 아침 6~7시에 자동으로 채워져요.', '체험관 통계 설정 완료', 8);
}
function dailyUpdate() { update_(7, false); } // 매일: 최근 7일을 다시 받음 (구글 집계가 하루 이틀 늦게 바뀌는 것까지 반영)
function updateNow() { // 시트 메뉴: 오늘 숫자까지 바로
  update_(7, true);
  SpreadsheetApp.getActive().toast('오늘 숫자까지 받아 왔어요. 오늘 숫자는 하루가 지나야 확정돼요.', '체험관 통계', 6);
}
function onOpen() { SpreadsheetApp.getUi().createMenu('체험관 통계').addItem('지금 업데이트', 'updateNow').addToUi(); }

function update_(days, withToday) {
  const range = { startDate: days + 'daysAgo', endDate: withToday ? 'today' : 'yesterday' };
  const prop = 'properties/' + PROPERTY_ID;
  const rows = {}; // 'yyyy-MM-dd' → 숫자들
  const blank = () => new Array(HEAD.length - 1).fill(0);
  const day = (v) => v.slice(0, 4) + '-' + v.slice(4, 6) + '-' + v.slice(6, 8);
  const base = AnalyticsData.Properties.runReport({
    dateRanges: [range], dimensions: [{ name: 'date' }], metrics: BASE.map((b) => ({ name: b[0] })), limit: 1000,
  }, prop);
  (base.rows || []).forEach((r) => {
    const d = day(r.dimensionValues[0].value); rows[d] = rows[d] || blank();
    BASE.forEach((b, i) => { rows[d][i] = Number(r.metricValues[i].value) || 0; });
  });
  const names = EVENTS.map((e) => e[0]).filter((n, i, a) => a.indexOf(n) === i);
  const ev = AnalyticsData.Properties.runReport({
    dateRanges: [range], dimensions: [{ name: 'date' }, { name: 'eventName' }], metrics: [{ name: 'eventCount' }, { name: 'totalUsers' }],
    dimensionFilter: { filter: { fieldName: 'eventName', inListFilter: { values: names } } }, limit: 10000,
  }, prop);
  (ev.rows || []).forEach((r) => {
    const d = day(r.dimensionValues[0].value), name = r.dimensionValues[1].value; rows[d] = rows[d] || blank();
    EVENTS.forEach((e, i) => { if (e[0] === name) rows[d][BASE.length + i] = Number(r.metricValues[e[2] === 'users' ? 1 : 0].value) || 0; });
  });
  write_(rows);
}

function write_(rows) { // 같은 날짜는 고쳐 쓰고, 새 날짜는 아래에 붙인 뒤 최근 날짜가 위로 오게 정렬
  const sh = getSheet_();
  const have = {};
  const last = sh.getLastRow();
  if (last >= 3) sh.getRange(3, 1, last - 2, 1).getDisplayValues().forEach((v, i) => { if (v[0]) have[v[0]] = i + 3; });
  Object.keys(rows).forEach((d) => {
    const line = [d].concat(rows[d]);
    if (have[d]) sh.getRange(have[d], 1, 1, HEAD.length).setValues([line]);
    else { sh.appendRow(line); have[d] = sh.getLastRow(); }
  });
  const n = sh.getLastRow() - 2;
  if (n > 1) sh.getRange(3, 1, n, HEAD.length).sort({ column: 1, ascending: false });
  sh.getRange(1, HEAD.length + 2).setValue('마지막 업데이트: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'));
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_NAME, 0);
  sh.getRange('A:A').setNumberFormat('@'); // 날짜를 글자로 둠 (비교·정렬이 쉬움)
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8f3f1').setWrap(true).setVerticalAlignment('middle');
  const sums = ['합계'].concat(HEAD.slice(1).map((_, i) => { const c = col_(i + 2); return '=SUM(' + c + '3:' + c + ')'; }));
  sh.getRange(2, 1, 1, HEAD.length).setValues([sums]).setFontWeight('bold').setBackground('#fff6e0');
  sh.setFrozenRows(2); sh.setFrozenColumns(1);
  sh.setRowHeight(1, 48); sh.setColumnWidth(1, 96); sh.setColumnWidths(2, HEAD.length - 1, 92);
  const empty = ss.getSheetByName('시트1'); if (empty && empty.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(empty); // 빈 기본 시트 정리
  return sh;
}
function col_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
