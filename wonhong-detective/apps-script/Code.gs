/**
 * 지구촌 탐정단 - 학급 기록 저장용 Google Apps Script
 * 학생 기록은 이 스프레드시트의 'Users' 시트에서 바로 볼 수 있어요. 사용법은 README.md 를 따라 주세요.
 *
 * @OnlyCurrentDoc  ← 이 스크립트가 '이 스프레드시트 하나'에만 접근하도록 권한 범위를 제한해요.
 */
const SHEET_NAME = 'Users';
const HEAD = ['학급','이름','pw해시','마지막저장','황금지구본','자연보석','우리나라','최종작전','별점합','도감','퀴즈정답','퀴즈전체','정답률(%)','코인','진행중','취약단원','단원별 정답/시도','data'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#ffe9a6');
    sh.getRange('A:C').setNumberFormat('@');   // '6-3' 같은 학급 코드가 날짜로 바뀌지 않게 텍스트로 고정
    sh.getRange('D:D').setNumberFormat('m/d hh:mm');
    sh.setColumnWidth(17, 360);
    const rule = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(60)
      .setBackground('#ffd6d6').setRanges([sh.getRange('M2:M')]).build();
    sh.setConditionalFormatRules([rule]);
    sh.hideColumns(3);    // 비밀번호 해시 (비번 초기화: 열 숨김 해제 후 해당 학생 칸을 비우면 돼요)
    sh.hideColumns(18);   // data (진행 상황 원본)
  }
  sh.getRange('A:C').setNumberFormat('@');   // 이미 만들어진 시트에도 글자 형식을 보장
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// 시트가 '6-2'를 날짜로 바꿔 놓았더라도 같은 값으로 비교해요
function norm_(x) {
  if (x instanceof Date) return (x.getMonth() + 1) + '-' + x.getDate();
  return String(x).trim();
}
function findRow_(sh, cls, name) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const v = sh.getRange(2, 1, n, 2).getValues();
  const c = norm_(cls), m = norm_(name);
  for (let i = 0; i < n; i++) if (norm_(v[i][0]) === c && norm_(v[i][1]) === m) return i + 2;
  return 0;
}
function rowFromData_(cls, name, pw, d) {
  const wins = d.wins || {}, best = d.best || {};
  const stars = Object.keys(best).reduce((a, k) => a + (best[k] || 0), 0);
  const topics = d.topics || {};
  let tr = 0, tt = 0, weak = null;
  const parts = [];
  Object.keys(topics).forEach(k => {
    const r = topics[k][0], t = topics[k][1];
    tr += r; tt += t; parts.push(k + ' ' + r + '/' + t);
    if (t >= 2 && (weak === null || r / t < weak[1])) weak = [k, r / t];
  });
  const rate = tt ? Math.round(tr / tt * 100) : '';
  const g = d.game;
  const names = { globe: '황금지구본', nature: '자연보석', korea: '우리나라', final: '최종작전' };
  const progress = g ? (names[g.case] || g.case) + ' ' + g.stop + '/' + (g.route ? g.route.length : '?') : '';
  let json = JSON.stringify(d);
  if (json.length > 45000) { const c = JSON.parse(json); c.game = null; json = JSON.stringify(c); }
  return [cls, name, pw, new Date(), wins.globe || 0, wins.nature || 0, wins.korea || 0, wins.final || 0,
    stars, (d.dex || []).length, tr, tt, rate, d.coins || 0, progress,
    weak && weak[1] < 0.7 ? weak[0] : '', parts.join(' · '), json];
}

// A~C열(학급·이름·비번)을 먼저 '글자' 형식으로 만든 뒤 써서, '6-2'가 날짜로 바뀌지 않게 해요
function writeRow_(sh, r, values) {
  sh.getRange(r, 1, 1, 3).setNumberFormat('@');
  sh.getRange(r, 1, 1, HEAD.length).setValues([values]);
}

function doGet() { return out_({ ok: true, msg: '지구촌 탐정단 서버가 동작 중이에요.' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const q = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const cls = String(q.cls || ''), name = String(q.name || '');
    if (q.a === 'register') {
      if (!name) return out_({ ok: false, err: '이름이 비어 있어요.' });
      if (findRow_(sh, cls, name)) {
        if (q.again) return out_({ ok: true, data: null });   // 자동 재등록 요청: 이미 있으면 그대로 성공
        return out_({ ok: false, err: '이미 있는 이름이에요. 로그인하거나 다른 이름을 써 주세요.' });
      }
      writeRow_(sh, sh.getLastRow() + 1, rowFromData_(cls, name, q.pw, {}));
      return out_({ ok: true, data: null });
    }
    if (q.a === 'login' || q.a === 'save') {
      const r = findRow_(sh, cls, name);
      if (!r) return out_({ ok: false, err: '등록되지 않은 이름이에요. ‘새 탐정 등록’을 눌러 주세요.' });
      const stored = String(sh.getRange(r, 3).getValue());
      if (stored !== '' && stored !== q.pw) return out_({ ok: false, err: '비밀번호가 맞지 않아요.' });
      if (stored === '') sh.getRange(r, 3).setValue(q.pw);   // 선생님이 초기화한 경우: 새 비밀번호로 설정
      if (q.a === 'login') {
        const raw = sh.getRange(r, 18).getValue();
        return out_({ ok: true, data: raw ? JSON.parse(raw) : null });
      }
      writeRow_(sh, r, rowFromData_(cls, name, q.pw, q.data || {}));
      return out_({ ok: true });
    }
    return out_({ ok: false, err: '알 수 없는 요청이에요.' });
  } catch (err) {
    return out_({ ok: false, err: '서버 오류: ' + err });
  } finally {
    lock.releaseLock();
  }
}

/** 시트 메뉴: 탐정단 → 중복 기록 정리 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('탐정단').addItem('중복 기록 정리', 'dedupeUsers').addToUi();
}

/** 같은 학급·이름이 여러 줄이면 '진행 기록이 있고 가장 최근에 저장된' 한 줄만 남겨요. */
function dedupeUsers() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rows = sh.getRange(2, 1, n, HEAD.length).getValues();
  const best = {};
  rows.forEach((r, i) => {
    const key = norm_(r[0]) + '|' + norm_(r[1]);
    const t = r[3] instanceof Date ? r[3].getTime() : 0;
    const score = (r[17] ? 1e15 : 0) + t;          // data(진행 기록)가 있는 줄을 우선, 그다음 최근 저장
    if (!(key in best) || score > best[key].score) best[key] = { i: i, score: score };
  });
  const keep = new Set(Object.keys(best).map(k => best[k].i));
  let removed = 0;
  for (let i = n - 1; i >= 0; i--) if (!keep.has(i)) { sh.deleteRow(i + 2); removed++; }
  // 남은 줄의 학급·이름을 글자로 정리 (날짜로 바뀐 '6월 2일' → '6-2')
  const left = sh.getLastRow() - 1;
  if (left >= 1) {
    const ab = sh.getRange(2, 1, left, 2).getValues().map(r => [norm_(r[0]), norm_(r[1])]);
    sh.getRange(2, 1, left, 2).setNumberFormat('@').setValues(ab);
  }
  SpreadsheetApp.getActive().toast(removed + '줄을 정리했어요.', '탐정단', 5);
}
