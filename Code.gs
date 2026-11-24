/**
 * て형 조립 챌린지 — 점수 저장 서버 (Google Apps Script)
 * 구글 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣으세요.
 */

// ★ 선생님 분석 페이지 비밀번호 — 반드시 바꾸세요
const TEACHER_KEY = '여기에-비밀번호';

const SHEET_NAME = 'results';
const HEAD = ['시각', '날짜', '반', '번호', '닉네임', '점수', '정답', '오답', '최대콤보', '상세'];
const VALID_GROUPS = ['うつる', 'むぶぬ', 'く', 'ぐ', 'す', 'いく'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 날짜를 글자로 보관
  }
  return sh;
}
function today_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function dateStr_(v) { return (v && typeof v.getTime === 'function') ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd') : String(v).slice(0, 10); }

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
    const cls = String(d.cls || '');
    const num = parseInt(d.num, 10);
    const nick = String(d.nick || '').trim().slice(0, 10);
    if (!/^[1-6]$/.test(cls) || !(num >= 1 && num <= 40) || !nick) return out_({ ok: false, error: '입력값 오류' });

    // 3분 동안 현실적으로 풀 수 있는 양만 인정
    const details = (Array.isArray(d.details) ? d.details : []).slice(0, 150)
      .filter(function (x) { return x && VALID_GROUPS.indexOf(x.g) >= 0; })
      .map(function (x) { return { v: String(x.v).slice(0, 8), g: x.g, a: String(x.a || '').slice(0, 4), ok: x.ok ? 1 : 0 }; });
    const ok = details.filter(function (x) { return x.ok; }).length;
    const s = score_(details);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_().appendRow([new Date(), today_(), Number(cls), num, nick, s.score, ok, details.length - ok, s.maxCombo, JSON.stringify(details)]);
    } finally { lock.releaseLock(); }
    return out_({ ok: true, score: s.score });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function rows_() {
  const v = sheet_().getDataRange().getValues();
  v.shift();
  return v.map(function (r) {
    let det = [];
    try { det = JSON.parse(r[9] || '[]'); } catch (e) {}
    return { date: dateStr_(r[0] || r[1]), cls: String(r[2]), num: Number(r[3]), nick: String(r[4]), score: Number(r[5]),
             ok: Number(r[6]), ng: Number(r[7]), det: det };
  });
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

// 오늘 순위: 학생마다 최고점 1개만. cls=all 이면 전체 반 통합
function rank_(p) {
  const cls = String(p.cls || ''), all = cls === 'all', today = today_();
  const me = parseInt(p.me, 10), meCls = String(p.meCls || cls);
  const best = {};
  rows_().forEach(function (r) {
    if (r.date !== today || (!all && r.cls !== cls)) return;
    const k = r.cls + '-' + r.num;
    if (!best[k] || r.score > best[k].score) best[k] = { nick: r.nick, score: r.score, num: r.num, cls: r.cls };
  });
  const list = Object.keys(best).map(function (k) { return best[k]; })
    .sort(function (a, b) { return b.score - a.score; })
    .slice(0, 50)
    .map(function (x) { // 번호는 공개하지 않음
      const o = { nick: x.nick, score: x.score, me: x.num === me && x.cls === meCls };
      if (all) o.cls = x.cls;
      return o;
    });
  return { ok: true, rank: list };
}

// 선생님용 통계
function stats_(p) {
  const cls = String(p.cls || 'all'), date = String(p.date || 'all');
  const rows = rows_().filter(function (r) {
    return (cls === 'all' || r.cls === cls) && (date === 'all' || r.date === date);
  });
  const groups = {}; VALID_GROUPS.forEach(function (g) { groups[g] = { g: g, total: 0, wrong: 0 }; });
  const verbs = {}, students = {};
  rows.forEach(function (r) {
    const key = r.cls + '-' + r.num;
    const st = students[key] || (students[key] = { cls: r.cls, num: r.num, nick: r.nick, plays: 0, best: 0, total: 0, okc: 0, gw: {} });
    st.plays++; st.best = Math.max(st.best, r.score); st.nick = r.nick;
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
    return { cls: s.cls, num: s.num, nick: s.nick, plays: s.plays, best: s.best,
             acc: s.total ? Math.round(s.okc / s.total * 100) : 0, weak: weak, weakCount: wmax };
  }).sort(function (a, b) { return a.cls - b.cls || a.num - b.num; });
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
