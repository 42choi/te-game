/**
 * て형 조립 챌린지 — 점수 저장 서버 (Google Apps Script)
 * 구글 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣으세요.
 *
 * 시트 2개를 자동으로 만듭니다.
 *  - 「기록」        : 모든 판을 한 줄씩 (학번·이름·점수·풀이 상세)
 *  - 「학생별 최고점」 : 학번마다 한 줄. 같은 학번은 가장 높은 점수만 남음
 */

// ★ 선생님 분석 페이지 비밀번호
const TEACHER_KEY = '여기에-비밀번호';

const LOG_SHEET = '기록';
const BEST_SHEET = '학생별 최고점';
const LOG_HEAD = ['시각', '날짜', '학번', '이름', '학년', '반', '번호', '점수', '정답', '오답', '최대콤보', '상세'];
const BEST_HEAD = ['학번', '이름', '학년', '반', '번호', '최고점', '최고점 날짜', '플레이 횟수', '최근 플레이'];
const VALID_GROUPS = ['うつる', 'むぶぬ', 'く', 'ぐ', 'す', 'いく'];

function sheet_(name, head, textCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    textCols.forEach(function (c) { sh.getRange(c + ':' + c).setNumberFormat('@'); }); // 글자 그대로 보관
    sh.appendRow(head);
    sh.setFrozenRows(1);
  }
  return sh;
}
function logSheet_() { return sheet_(LOG_SHEET, LOG_HEAD, ['B', 'C']); }
function bestSheet_() { return sheet_(BEST_SHEET, BEST_HEAD, ['A', 'G']); }

function today_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function dateStr_(v) { return (v && typeof v.getTime === 'function') ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd') : String(v).slice(0, 10); }

// 학번 4자리: 학년1 + 반1 + 번호2  (예: 2612 = 2학년 6반 12번)
function parseSid_(sid) {
  const m = /^([1-3])([1-9])(\d\d)$/.exec(String(sid || '').trim());
  if (!m || +m[3] < 1 || +m[3] > 40) return null;
  return { sid: m[0], grade: +m[1], cls: +m[2], num: +m[3] };
}

// 점수는 학생 기기가 보낸 값을 믿지 않고, 풀이 기록으로 다시 계산합니다 (게임과 같은 규칙)
function score_(details) {
  let combo = 0, maxCombo = 0, score = 0;
  details.forEach(function (d) {
    if (d.ok) { combo++; maxCombo = Math.max(maxCombo, combo); score += 10 + Math.min(combo - 1, 10) * 2; }
    else combo = 0;
  });
  return { score: score, maxCombo: maxCombo };
}

/* ---------- 결과 저장 ---------- */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const st = parseSid_(d.sid);
    const name = String(d.name || '').trim().slice(0, 10);
    if (!st || !name) return out_({ ok: false, error: '학번 또는 이름 오류' });

    const details = (Array.isArray(d.details) ? d.details : []).slice(0, 150)
      .filter(function (x) { return x && VALID_GROUPS.indexOf(x.g) >= 0; })
      .map(function (x) { return { v: String(x.v).slice(0, 8), g: x.g, a: String(x.a || '').slice(0, 4), ok: x.ok ? 1 : 0 }; });
    const ok = details.filter(function (x) { return x.ok; }).length;
    const s = score_(details);
    const now = new Date(), today = today_();

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    let best, isNew;
    try {
      logSheet_().appendRow([now, today, st.sid, name, st.grade, st.cls, st.num, s.score, ok, details.length - ok, s.maxCombo, JSON.stringify(details)]);

      // 학생별 최고점 갱신 (같은 학번은 한 줄, 높은 점수만)
      const bs = bestSheet_();
      const n = bs.getLastRow() - 1;
      let row = -1;
      if (n > 0) {
        const ids = bs.getRange(2, 1, n, 1).getValues();
        for (let i = 0; i < n; i++) if (String(ids[i][0]) === st.sid) { row = i + 2; break; }
      }
      if (row < 0) {
        best = s.score; isNew = true;
        bs.appendRow([st.sid, name, st.grade, st.cls, st.num, best, today, 1, now]);
      } else {
        const r = bs.getRange(row, 1, 1, 9).getValues()[0];
        best = Number(r[5]) || 0; let bestDate = r[6];
        isNew = s.score > best;
        if (isNew) { best = s.score; bestDate = today; }
        bs.getRange(row, 1, 1, 9).setValues([[st.sid, name, st.grade, st.cls, st.num, best, bestDate, (Number(r[7]) || 0) + 1, now]]);
      }
    } finally { lock.releaseLock(); }
    return out_({ ok: true, score: s.score, best: best, isNew: isNew });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ---------- 조회 ---------- */
function doGet(e) {
  const p = e.parameter || {};
  try {
    if (p.action === 'rank') return out_(rank_(p));
    if (p.action === 'stats') {
      if (String(p.key || '') !== TEACHER_KEY) return out_({ ok: false, error: '비밀번호가 맞지 않아요' });
      return out_(stats_(p));
    }
    return out_({ ok: true, msg: 'て형 조립 챌린지 서버가 동작 중입니다' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// 순위: 학생별 최고점(전체 기간) 기준. cls=all 이면 전체 반 통합
function rank_(p) {
  const cls = String(p.cls || ''), all = cls === 'all', me = String(p.me || '');
  const bs = bestSheet_();
  const n = bs.getLastRow() - 1;
  const rows = n > 0 ? bs.getRange(2, 1, n, 9).getValues() : [];
  const list = rows
    .map(function (r) { return { sid: String(r[0]), name: String(r[1]), cls: String(r[3]), score: Number(r[5]) || 0 }; })
    .filter(function (r) { return all || r.cls === cls; })
    .sort(function (a, b) { return b.score - a.score; })
    .slice(0, 50)
    .map(function (x) { // 학번은 공개하지 않음
      const o = { name: x.name, score: x.score, me: x.sid === me };
      if (all) o.cls = x.cls;
      return o;
    });
  return { ok: true, rank: list };
}

function logRows_() {
  const sh = logSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, 12).getValues().map(function (r) {
    let det = [];
    try { det = JSON.parse(r[11] || '[]'); } catch (e) {}
    return { date: dateStr_(r[0] || r[1]), sid: String(r[2]), name: String(r[3]), cls: String(r[5]), num: Number(r[6]),
             score: Number(r[7]), det: det };
  });
}

// 선생님용 통계
function stats_(p) {
  const cls = String(p.cls || 'all'), date = String(p.date || 'all');
  const rows = logRows_().filter(function (r) {
    return (cls === 'all' || r.cls === cls) && (date === 'all' || r.date === date);
  });
  const groups = {}; VALID_GROUPS.forEach(function (g) { groups[g] = { g: g, total: 0, wrong: 0 }; });
  const verbs = {}, students = {};
  rows.forEach(function (r) {
    const st = students[r.sid] || (students[r.sid] = { sid: r.sid, name: r.name, cls: r.cls, num: r.num, plays: 0, best: 0, total: 0, okc: 0, gw: {} });
    st.plays++; st.best = Math.max(st.best, r.score); st.name = r.name;
    r.det.forEach(function (d) {
      if (!groups[d.g]) return;
      groups[d.g].total++; st.total++;
      const vb = verbs[d.v] || (verbs[d.v] = { v: d.v, g: d.g, total: 0, wrong: 0 });
      vb.total++;
      if (d.ok) st.okc++;
      else { groups[d.g].wrong++; vb.wrong++; st.gw[d.g] = (st.gw[d.g] || 0) + 1; }
    });
  });
  const stList = Object.keys(students).map(function (k) {
    const s = students[k];
    let weak = '', wmax = 0;
    Object.keys(s.gw).forEach(function (g) { if (s.gw[g] > wmax) { wmax = s.gw[g]; weak = g; } });
    return { sid: s.sid, name: s.name, cls: s.cls, plays: s.plays, best: s.best,
             acc: s.total ? Math.round(s.okc / s.total * 100) : 0, weak: weak, weakCount: wmax };
  }).sort(function (a, b) { return a.sid.localeCompare(b.sid); });
  const scores = rows.map(function (r) { return r.score; });
  return {
    ok: true,
    summary: { plays: rows.length, students: stList.length,
               avg: scores.length ? Math.round(scores.reduce(function (a, b) { return a + b; }, 0) / scores.length) : 0 },
    groups: VALID_GROUPS.map(function (g) { return groups[g]; }),
    verbs: Object.keys(verbs).map(function (k) { return verbs[k]; })
      .filter(function (v) { return v.wrong > 0; })
      .sort(function (a, b) { return b.wrong / b.total - a.wrong / a.total || b.wrong - a.wrong; }).slice(0, 10),
    students: stList
  };
}
