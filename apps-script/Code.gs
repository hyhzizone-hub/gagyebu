/**
 * 우리집 가계부 API (Google Apps Script)
 * 통합가계부 스프레드시트에 붙여서 씁니다: 확장 프로그램 → Apps Script
 * 배포: 배포 → 새 배포 → 웹 앱 / 다음 사용자로 실행: 나 / 액세스 권한: 모든 사용자
 *
 * 로그인한 구글 계정이 '사용자관리' 시트에 있어야 하고,
 * 같은 가족코드를 가진 사람의 데이터만 주고받습니다 (앱시트 보안 필터와 같은 규칙).
 */

// Google Cloud 콘솔에서 만든 OAuth 클라이언트 ID를 넣으세요.
const CLIENT_ID = '373963034223-77jbp57flgvbv4vhcsl38tqaqmj4p4jf.apps.googleusercontent.com';

const SHEETS = { tx: '통합상세내역', cat: '카테고리', rec: '자동입력', base: '기초자산', users: '사용자관리', plan: '예정항목', ev: '이벤트' };

/* ---------- 입구 ---------- */
function doGet() {
  return json_({ ok: true, app: '우리집 가계부 API' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const me = auth_(req.token);
    if (req.action === 'session') return json_({ ok: true, data: { session: newSession_(me.email), days: SESSION_DAYS } });
    if (req.action === 'logout') { if (String(req.token).indexOf('.') < 0) PropertiesService.getScriptProperties().deleteProperty('sess:' + req.token); return json_({ ok: true, data: {} }); }
    const actions = { load: load_, add: add_, update: update_, remove: remove_,
      saveCategories: saveCategories_, saveRecurring: saveRecurring_, saveBase: saveBase_, savePlan: savePlan_, saveLoan: saveLoan_,
      saveEvents: saveEvents_ };
    const fn = actions[req.action];
    if (!fn) throw fail_(400, '알 수 없는 요청이에요: ' + req.action);
    return json_({ ok: true, data: fn(me, req) });
  } catch (x) {
    return json_({ ok: false, code: x.code || 500, error: String(x.message || x) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fail_(code, msg) { const e = new Error(msg); e.code = code; return e; }

/* ---------- 로그인 확인 ---------- */
const SESSION_DAYS = 30;
function auth_(token) {
  if (!token) throw fail_(401, '로그인이 필요해요');
  const email = String(token).indexOf('.') < 0 ? sessionEmail_(String(token)) : idTokenEmail_(token);
  const users = table_(SHEETS.users);
  const ce = users.col('이메일'), cf = users.col('가족코드');
  const famOf = {};
  for (let i = 1; i < users.values.length; i++) {
    const em = String(users.values[i][ce]).trim().toLowerCase();
    if (em) famOf[em] = String(users.values[i][cf]).trim();
  }
  const fam = famOf[email];
  if (!fam) throw fail_(403, email + ' 계정은 사용자관리 시트에 없어요');
  const members = new Set(Object.keys(famOf).filter(function (k) { return famOf[k] === fam; }));
  return { email: email, fam: fam, members: members };
}

function idTokenEmail_(token) {
  const cache = CacheService.getScriptCache();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token);
  const key = 'tok_' + Utilities.base64EncodeWebSafe(digest).slice(0, 43);
  let email = cache.get(key);
  if (!email) {
    const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw fail_(401, '로그인이 만료됐어요');
    const p = JSON.parse(res.getContentText());
    const now = Math.floor(Date.now() / 1000);
    if (p.aud !== CLIENT_ID) throw fail_(401, '이 앱용 로그인이 아니에요 (CLIENT_ID를 확인하세요)');
    if (String(p.email_verified) !== 'true' || Number(p.exp) <= now) throw fail_(401, '로그인 정보를 확인할 수 없어요');
    email = String(p.email).toLowerCase();
    cache.put(key, email, Math.max(60, Math.min(21600, Number(p.exp) - now - 30)));
  }
  return email;
}
/* 세션: 구글 로그인 한 번 후 30일 동안 다시 로그인하지 않도록 서버가 발급하는 토큰 */
function newSession_(email) {
  const tok = 'S' + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('sess:' + tok, JSON.stringify({ email: email, exp: Date.now() + SESSION_DAYS * 86400000 }));
  return tok;
}
function sessionEmail_(tok) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('sess:' + tok);
  if (!raw) throw fail_(401, '로그인이 만료됐어요');
  const sess = JSON.parse(raw);
  if (Number(sess.exp) < Date.now()) { props.deleteProperty('sess:' + tok); throw fail_(401, '로그인이 만료됐어요'); }
  return String(sess.email);
}

/* ---------- 시트 도우미 ---------- */
function table_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw fail_(500, "'" + name + "' 시트를 찾을 수 없어요");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(function (h) { return String(h).trim(); });
  return {
    sh: sh, values: values, head: head,
    col: function (n) { const i = head.indexOf(n); if (i < 0) throw fail_(500, name + " 시트에 '" + n + "' 컬럼이 없어요"); return i; },
    has: function (n) { return head.indexOf(n) >= 0; }
  };
}
const pad_ = function (n) { return ('0' + n).slice(-2); };
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function toYMD_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const m = String(v).trim().match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
  return m ? m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]) : '';
}
function toDate_(ymd) {
  const p = String(ymd).split('-').map(Number);
  if (p.length !== 3 || !p[0]) throw fail_(400, '날짜 형식이 잘못됐어요: ' + ymd);
  return new Date(p[0], p[1] - 1, p[2]);
}
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function isMine_(me, email) { return me.members.has(String(email).trim().toLowerCase()); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function cleanTx_(it) {
  const a = Number(it.a);
  if (!isFinite(a)) throw fail_(400, '금액이 잘못됐어요');
  return { d: toDate_(it.d), t: String(it.t || ''), c: String(it.c || ''), m: String(it.m || ''), a: a, ev: String(it.ev || '') };
}
// 통합상세내역에 '이벤트' 컬럼이 없으면 맨 오른쪽에 만듭니다 (거래를 이벤트에 묶기 위해)
function txEnsureEv_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.tx);
  if (!sh) return;
  const last = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, last).getValues()[0].map(function (h) { return String(h).trim(); });
  if (head.indexOf('이벤트') < 0) sh.getRange(1, last + 1).setValue('이벤트');
}

/* ---------- 읽기 ---------- */
function load_(me) {
  const T = table_(SHEETS.tx);
  const c = { id: T.col('ID'), d: T.col('날짜'), t: T.col('거래형태'), c: T.col('카테고리'), m: T.col('상세내역'), a: T.col('금액'), w: T.col('작성자_이메일') };
  const ev = T.has('이벤트') ? T.col('이벤트') : -1;
  const tx = [];
  for (let i = 1; i < T.values.length; i++) {
    const r = T.values[i];
    if (r[c.id] === '' || r[c.id] === null || !isMine_(me, r[c.w])) continue;
    const d = toYMD_(r[c.d]);
    if (!d) continue;
    tx.push({ id: String(r[c.id]), d: d, t: String(r[c.t]), c: String(r[c.c]), m: String(r[c.m]), a: Number(r[c.a]) || 0, w: String(r[c.w]).split('@')[0], ev: ev >= 0 ? String(r[ev] || '') : '' });
  }
  return { me: { email: me.email, fam: me.fam }, tx: tx, categories: categories_(me), recurring: recurring_(me), base: base_(me), plan: plan_(me), loan: loan_(me), events: events_(me) };
}
function categories_(me) {
  const T = table_(SHEETS.cat), out = [];
  const id = T.col('ID'), t = T.col('거래형태'), c = T.col('카테고리'), b = T.col('예상금액'), w = T.col('작성자_이메일');
  for (let i = 1; i < T.values.length; i++) {
    const r = T.values[i];
    if (!isMine_(me, r[w]) || !r[c]) continue;
    out.push({ id: String(r[id]), t: String(r[t]), c: String(r[c]), b: r[b] === '' ? null : Number(r[b]) });
  }
  return out;
}
function recurring_(me) {
  const T = table_(SHEETS.rec), out = [];
  const id = T.col('ID'), t = T.col('거래형태'), c = T.col('카테고리'), m = T.col('상세내역'), a = T.col('금액'), w = T.col('작성자_이메일');
  for (let i = 1; i < T.values.length; i++) {
    const r = T.values[i];
    if (!isMine_(me, r[w]) || !r[c]) continue;
    out.push({ id: String(r[id]), t: String(r[t]), c: String(r[c]), m: String(r[m]), a: Number(r[a]) || 0 });
  }
  return out;
}
function base_(me) {
  const T = table_(SHEETS.base), out = [];
  const n = T.col('항목'), a = T.col('금액'), w = T.col('작성자_이메일');
  for (let i = 1; i < T.values.length; i++) {
    const r = T.values[i];
    if (!isMine_(me, r[w]) || !r[n]) continue;
    out.push({ n: String(r[n]), a: Number(r[a]) || 0 });
  }
  return out;
}

/* ---------- 예정 일정 (예상 잔고 계산용, 내역에는 들어가지 않음) ---------- */
const PLAN_HEAD = ['ID', '날짜', '거래형태', '카테고리', '상세내역', '금액', '작성자_이메일'];
function planSheet_(create) {
  let sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.plan);
  if (!sh && create) {
    sh = SpreadsheetApp.getActive().insertSheet(SHEETS.plan);
    sh.getRange(1, 1, 1, PLAN_HEAD.length).setValues([PLAN_HEAD]);
  }
  return sh;
}
function plan_(me) {
  if (!planSheet_(false)) return [];
  const T = table_(SHEETS.plan), out = [];
  const id = T.col('ID'), d = T.col('날짜'), t = T.col('거래형태'), c = T.col('카테고리'), m = T.col('상세내역'), a = T.col('금액'), w = T.col('작성자_이메일');
  for (let i = 1; i < T.values.length; i++) {
    const r = T.values[i];
    if (r[id] === '' || r[id] === null || !isMine_(me, r[w])) continue;
    const dd = toYMD_(r[d]);
    if (!dd) continue;
    out.push({ id: String(r[id]), d: dd, t: String(r[t]), c: String(r[c]), m: String(r[m]), a: Number(r[a]) || 0 });
  }
  return out;
}
function savePlan_(me, req) {
  const items = (req.items || []).map(function (it) {
    const x = cleanTx_(it);
    return { id: /^[0-9a-f]{8}$/.test(it.id) ? it.id : newId_(), d: x.d, t: x.t, c: x.c, m: x.m, a: x.a };
  });
  return withLock_(function () {
    planSheet_(true);
    const T = table_(SHEETS.plan), w = T.col('작성자_이메일');
    // 내 행만 지우고 다시 씁니다 (가족 다른 사람의 예정 일정은 그대로 둠)
    for (let i = T.values.length - 1; i >= 1; i--) {
      if (isMine_(me, T.values[i][w])) T.sh.deleteRow(i + 1);
    }
    const rows = items.map(function (x) {
      const row = new Array(T.head.length).fill('');
      row[T.col('ID')] = x.id; row[T.col('날짜')] = x.d; row[T.col('거래형태')] = x.t; row[T.col('카테고리')] = x.c;
      row[T.col('상세내역')] = x.m; row[T.col('금액')] = x.a; row[w] = me.email;
      return row;
    });
    if (rows.length) T.sh.getRange(T.sh.getLastRow() + 1, 1, rows.length, T.head.length).setValues(rows);
    return plan_(me);
  });
}

/* ---------- 이벤트 (여행·돌잔치·결혼식 등, 거래를 묶어서 보는 용도) ---------- */
const EVENT_HEAD = ['ID', '이름', '종류', '시작일', '종료일', '메모', '작성자_이메일'];
function eventSheet_(create) {
  let sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ev);
  if (!sh && create) {
    sh = SpreadsheetApp.getActive().insertSheet(SHEETS.ev);
    sh.getRange(1, 1, 1, EVENT_HEAD.length).setValues([EVENT_HEAD]);
  }
  return sh;
}
function events_(me) {
  if (!eventSheet_(false)) return [];
  const T = table_(SHEETS.ev), out = [];
  const id = T.col('ID'), n = T.col('이름'), k = T.col('종류'), s = T.col('시작일'), e = T.col('종료일'), m = T.col('메모'), w = T.col('작성자_이메일');
  for (let i = 1; i < T.values.length; i++) {
    const r = T.values[i];
    if (r[id] === '' || r[id] === null || !isMine_(me, r[w])) continue;
    out.push({ id: String(r[id]), name: String(r[n]), kind: String(r[k]), start: toYMD_(r[s]), end: toYMD_(r[e]) || toYMD_(r[s]), memo: String(r[m]) });
  }
  return out;
}
function saveEvents_(me, req) {
  const items = (req.items || []).map(function (it) {
    const name = String(it.name || '').trim();
    if (!name) throw fail_(400, '이벤트 이름을 입력해 주세요');
    const start = toDate_(it.start);
    const end = it.end ? toDate_(it.end) : start;
    if (end < start) throw fail_(400, '종료일이 시작일보다 빨라요');
    return { id: /^[0-9a-f]{8}$/.test(it.id) ? it.id : newId_(), name: name, kind: String(it.kind || ''), start: start, end: end, memo: String(it.memo || '') };
  });
  return withLock_(function () {
    eventSheet_(true);
    const T = table_(SHEETS.ev), w = T.col('작성자_이메일');
    const keep = {};
    items.forEach(function (x) { keep[x.id] = true; });
    // 내 행만 다시 씁니다 (가족 다른 사람의 이벤트는 그대로 둠)
    const removed = {};
    for (let i = T.values.length - 1; i >= 1; i--) {
      const r = T.values[i];
      if (!isMine_(me, r[w])) continue;
      const rid = String(r[T.col('ID')]);
      if (!keep[rid]) removed[rid] = true;
      T.sh.deleteRow(i + 1);
    }
    const rows = items.map(function (x) {
      const row = new Array(T.head.length).fill('');
      row[T.col('ID')] = x.id; row[T.col('이름')] = x.name; row[T.col('종류')] = x.kind;
      row[T.col('시작일')] = x.start; row[T.col('종료일')] = x.end; row[T.col('메모')] = x.memo; row[w] = me.email;
      return row;
    });
    if (rows.length) T.sh.getRange(T.sh.getLastRow() + 1, 1, rows.length, T.head.length).setValues(rows);
    // 삭제된 이벤트에 묶여 있던 내 거래는 이벤트 표시를 지웁니다
    if (Object.keys(removed).length) {
      const X = table_(SHEETS.tx);
      if (X.has('이벤트')) {
        const ev = X.col('이벤트'), xw = X.col('작성자_이메일');
        for (let i = 1; i < X.values.length; i++) {
          const r = X.values[i];
          if (isMine_(me, r[xw]) && removed[String(r[ev])]) X.sh.getRange(i + 1, ev + 1).setValue('');
        }
      }
    }
    return events_(me);
  });
}

/* ---------- 내역 쓰기 ---------- */
function add_(me, req) {
  const items = (req.items || []).map(function (it) { const x = cleanTx_(it); x.id = /^[0-9a-f]{8}$/.test(it.id) ? it.id : newId_(); return x; });
  if (!items.length) return { ids: [] };
  return withLock_(function () {
    txEnsureEv_();
    const T = table_(SHEETS.tx);
    const col = { id: T.col('ID'), d: T.col('날짜'), t: T.col('거래형태'), c: T.col('카테고리'), m: T.col('상세내역'), a: T.col('금액'), w: T.col('작성자_이메일') };
    const sel = T.has('선택') ? T.col('선택') : -1, yr = T.has('연도') ? T.col('연도') : -1, evc = T.col('이벤트');
    const rows = items.map(function (x) {
      const row = new Array(T.head.length).fill('');
      row[col.id] = x.id; row[col.d] = x.d; row[col.t] = x.t; row[col.c] = x.c; row[col.m] = x.m; row[col.a] = x.a; row[col.w] = me.email;
      row[evc] = x.ev;
      if (sel >= 0) row[sel] = false;
      if (yr >= 0) row[yr] = x.d.getFullYear();
      return row;
    });
    const start = T.sh.getLastRow() + 1;
    T.sh.getRange(start, 1, rows.length, T.head.length).setValues(rows);
    if (yr >= 0) T.sh.getRange(start, yr + 1, rows.length, 1).setFormulaR1C1('=YEAR(SUBSTITUTE(RC[' + (col.d - yr) + '], ".", "-"))');
    return { ids: items.map(function (x) { return x.id; }) };
  });
}

function findTxRow_(T, me, id) {
  const ci = T.col('ID'), cw = T.col('작성자_이메일');
  for (let i = 1; i < T.values.length; i++) {
    if (String(T.values[i][ci]) === String(id)) {
      if (!isMine_(me, T.values[i][cw])) throw fail_(403, '다른 가족의 내역이에요');
      return i + 1;
    }
  }
  throw fail_(404, '내역을 찾을 수 없어요. 새로고침해 주세요');
}

function update_(me, req) {
  const x = cleanTx_(req.item);
  return withLock_(function () {
    txEnsureEv_();
    const T = table_(SHEETS.tx), row = findTxRow_(T, me, req.item.id);
    const set = function (name, v) { T.sh.getRange(row, T.col(name) + 1).setValue(v); };
    set('날짜', x.d); set('거래형태', x.t); set('카테고리', x.c); set('상세내역', x.m); set('금액', x.a); set('이벤트', x.ev);
    return { id: req.item.id };
  });
}

function remove_(me, req) {
  return withLock_(function () {
    const T = table_(SHEETS.tx);
    T.sh.deleteRow(findTxRow_(T, me, req.id));
    return { id: req.id };
  });
}

/* ---------- 설정 저장 (가족 행만 바꿈) ---------- */
function saveList_(me, name, items, keyCol, fields, keyOf) {
  // 빈 목록으로 저장하면 가족 설정이 모두 지워지므로 막아 둡니다
  if (!items.length) throw fail_(400, '빈 목록은 저장할 수 없어요. 새로고침 후 다시 시도해 주세요');
  return withLock_(function () {
    const T = table_(name), w = T.col('작성자_이메일'), k = T.col(keyCol);
    const fresh = items.filter(function (it) { return !keyOf(it); });
    const want = new Map(items.filter(keyOf).map(function (it) { return [String(keyOf(it)), it]; }));
    const del = [];
    for (let i = T.values.length - 1; i >= 1; i--) {
      const r = T.values[i];
      if (!isMine_(me, r[w])) continue;
      const key = String(r[k]);
      if (key === '') continue;
      const it = want.get(key);
      if (!it) { del.push(i + 1); continue; }
      Object.keys(fields).forEach(function (f) {
        const v = it[f] === null || it[f] === undefined ? '' : it[f];
        if (r[T.col(fields[f])] !== v) T.sh.getRange(i + 1, T.col(fields[f]) + 1).setValue(v);
      });
      want.delete(key);
    }
    del.forEach(function (row) { T.sh.deleteRow(row); }); // 아래 행부터 지움
    want.forEach(function (it) { fresh.push(it); });
    fresh.forEach(function (it) {
      const row = new Array(T.head.length).fill('');
      if (keyCol === 'ID') row[k] = newId_();
      Object.keys(fields).forEach(function (f) { row[T.col(fields[f])] = it[f] === null || it[f] === undefined ? '' : it[f]; });
      row[w] = me.email;
      T.sh.appendRow(row);
    });
  });
}
function saveCategories_(me, req) {
  saveList_(me, SHEETS.cat, req.items || [], 'ID', { t: '거래형태', c: '카테고리', b: '예상금액' }, function (it) { return it.id; });
  return categories_(me);
}
function saveRecurring_(me, req) {
  saveList_(me, SHEETS.rec, req.items || [], 'ID', { t: '거래형태', c: '카테고리', m: '상세내역', a: '금액' }, function (it) { return it.id; });
  return recurring_(me);
}
function saveBase_(me, req) {
  // 기초자산은 ID가 없어서 항목 이름으로 맞춥니다
  saveList_(me, SHEETS.base, req.items || [], '항목', { n: '항목', a: '금액' }, function (it) { return it.n; });
  return base_(me);
}

/* ---------- 매달 1일 0시 자동입력 (트리거) ---------- */
function autoRecurring_() {
  withLock_(function () {
    const ym = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
    const R = table_(SHEETS.rec), T = table_(SHEETS.tx); autoLoan_(ym);
    const rc = { t: R.col('거래형태'), c: R.col('카테고리'), m: R.col('상세내역'), a: R.col('금액'), w: R.col('작성자_이메일') };
    const tc = { id: T.col('ID'), d: T.col('날짜'), t: T.col('거래형태'), c: T.col('카테고리'), m: T.col('상세내역'), a: T.col('금액'), w: T.col('작성자_이메일') };
    const have = {};
    for (let i = 1; i < T.values.length; i++) {
      const r = T.values[i], d = toYMD_(r[tc.d]);
      if (!d || d.slice(0, 7) !== ym) continue;
      have[[String(r[tc.w]).trim().toLowerCase(), String(r[tc.t]), String(r[tc.c]), String(r[tc.m]), Number(r[tc.a])].join('|')] = true;
    }
    const items = [];
    for (let i = 1; i < R.values.length; i++) {
      const r = R.values[i];
      if (!r[rc.c] || r[rc.a] === '') continue;
      const owner = String(r[rc.w]).trim().toLowerCase(), t = String(r[rc.t]), c = String(r[rc.c]), m = String(r[rc.m]), a = Number(r[rc.a]) || 0;
      const k = [owner, t, c, m, a].join('|');
      if (have[k]) continue;
      have[k] = true;
      items.push({ owner: owner, t: t, c: c, m: m, a: a });
    }
    if (!items.length) return;
    const sel = T.has('선택') ? T.col('선택') : -1, yr = T.has('연도') ? T.col('연도') : -1;
    const rows = items.map(function (x) {
      const row = new Array(T.head.length).fill('');
      row[tc.id] = newId_(); row[tc.d] = toDate_(ym + '-01'); row[tc.t] = x.t; row[tc.c] = x.c; row[tc.m] = x.m; row[tc.a] = x.a; row[tc.w] = x.owner;
      if (sel >= 0) row[sel] = false;
      if (yr >= 0) row[yr] = Number(ym.slice(0, 4));
      return row;
    });
    const start = T.sh.getLastRow() + 1;
    T.sh.getRange(start, 1, rows.length, T.head.length).setValues(rows);
    if (yr >= 0) T.sh.getRange(start, yr + 1, rows.length, 1).setFormulaR1C1('=YEAR(SUBSTITUTE(RC[' + (tc.d - yr) + '], ".", "-"))');
  });
}
/* 한 번만 실행: 편집기에서 이 함수를 선택해 ▶ 실행하면 매달 1일 0시 트리거가 생깁니다 */
function 자동입력_트리거설치() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoRecurring_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoRecurring_').timeBased().onMonthDay(1).atHour(0).create();
  Logger.log('매달 1일 0시 자동입력 트리거를 설치했어요');
}

/* ---------- 설치 확인용: 편집기에서 이 함수를 실행해 보세요 ---------- */
function 점검() {
  const names = Object.keys(SHEETS).map(function (k) { return SHEETS[k]; });
  names.forEach(function (n) { const T = table_(n); Logger.log(n + ': ' + (T.values.length - 1) + '행, 컬럼 ' + T.head.join(', ')); });
  Logger.log('CLIENT_ID 설정: ' + (CLIENT_ID.indexOf('여기에') === 0 ? '아직 안 됨' : '됨'));
}

/* ---------- 주담대 설정 · 매달 자동 분리 ---------- */
function loanSheet_() { const ss = SpreadsheetApp.getActive(); let sh = ss.getSheetByName('주담대설정'); if (!sh) { sh = ss.insertSheet('주담대설정'); sh.appendRow(['금리', '월상환액', '자동입력', '작성자_이메일']); } return sh; }
function loan_(me) { const v = loanSheet_().getDataRange().getValues(); for (let i = 1; i < v.length; i++) { if (isMine_(me, v[i][3])) return { rate: Number(v[i][0]) || 0, monthly: Number(v[i][1]) || 0, auto: v[i][2] === true || String(v[i][2]).toUpperCase() === 'TRUE' }; } return { rate: 0, monthly: 0, auto: false }; }
function saveLoan_(me, req) { const L = req.loan || {}; return withLock_(function () { const sh = loanSheet_(), v = sh.getDataRange().getValues(); let row = -1; for (let i = 1; i < v.length; i++) { if (isMine_(me, v[i][3])) { row = i + 1; break; } } const vals = [Number(L.rate) || 0, Number(L.monthly) || 0, !!L.auto, me.email]; if (row < 0) sh.appendRow(vals); else sh.getRange(row, 1, 1, 4).setValues([vals]); return loan_(me); }); }
function autoLoan_(ym) {
  const v = loanSheet_().getDataRange().getValues(); const T = table_(SHEETS.tx), tv = T.values, h = T.head;
  const cD = h.indexOf('날짜'), cT = h.indexOf('거래형태'), cC = h.indexOf('카테고리'), cM = h.indexOf('상세내역'), cA = h.indexOf('금액'), cW = h.indexOf('작성자_이메일'), cId = h.indexOf('ID');
  for (let i = 1; i < v.length; i++) {
    const owner = String(v[i][3]).trim().toLowerCase(), rate = Number(v[i][0]) || 0, monthly = Number(v[i][1]) || 0;
    if (!(v[i][2] === true || String(v[i][2]).toUpperCase() === 'TRUE') || !owner || rate <= 0 || monthly <= 0) continue;
    const me = { email: owner, members: new Set([owner]) };
    let already = false, bal = 0; const b = base_(me).find(function (x) { return x.n === '주담대'; }); if (b) bal = b.a;
    for (let r = 1; r < tv.length; r++) {
      const d = toYMD_(tv[r][cD]); if (!d) continue;
      if (d.slice(0, 7) === ym && String(tv[r][cM]) === '주담대원금상환') already = true;
      if (d.slice(0, 7) >= ym) continue;
      if (String(tv[r][cT]) !== '자산이동') continue;
      const parts = String(tv[r][cC]).split('>'); const a = Number(tv[r][cA]) || 0;
      if (parts[1] === '주담대') bal += a; if (parts[0] === '주담대') bal -= a;
    }
    if (already) continue;
    const owed = Math.max(0, -bal), interest = Math.round(owed * rate / 100 / 12), principal = monthly - interest;
    const sel = T.has('선택') ? T.col('선택') : -1, yr = T.has('연도') ? T.col('연도') : -1;
    const mk = function (t, c, m, a) { const row = new Array(h.length).fill(''); row[cId] = newId_(); row[cD] = toDate_(ym + '-01'); row[cT] = t; row[cC] = c; row[cM] = m; row[cA] = a; row[cW] = owner; if (sel >= 0) row[sel] = false; if (yr >= 0) row[yr] = Number(ym.slice(0, 4)); return row; };
    const start = T.sh.getLastRow() + 1;
    T.sh.getRange(start, 1, 2, h.length).setValues([mk('고정지출', '주담대상환', '주담대이자', interest), mk('자산이동', '현금>주담대', '주담대원금상환', principal)]);
  }
}

