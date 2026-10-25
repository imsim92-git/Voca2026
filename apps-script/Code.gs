/**
 * 우리반 단어장 — 기록 서버 (Google Apps Script)
 *
 * 이 코드는 선생님의 구글 스프레드시트에 붙어서 동작합니다.
 *  - '학생' 시트: 학생당 한 줄. 학번, 이름, 가입·최근 접속 시각, 비밀번호(해시로만 저장),
 *                연습 횟수·정답률·마지막 연습·자주 틀린 단어·최근 연습 30회
 *  - '지난 기록' 시트: 예전 방식(연습 한 번에 한 줄)으로 쌓였던 기록. 더는 늘어나지 않음
 *  - '지문' 시트: 선생님이 등록한 지문 목록(교과서, 모의고사 1번 …)과 그 단어
 *  - '모둠', '레이스', '레이스 기록' 시트: 반·모둠 목록, 모둠 레이스 설정과 기록
 *
 * 단어표 사진 인식을 쓰려면: 편집기 왼쪽 '서비스' 옆 + → Drive API → 추가 (안내서 참고)
 *
 * 처음 한 번: 아래 ADMIN_PASSWORD_INIT에 관리자 비밀번호를 적고 setup 함수를 실행하세요.
 * 실행이 끝나면 비밀번호는 안전하게 저장되므로, 여기 적은 글자는 지워도 됩니다.
 */
const ADMIN_PASSWORD_INIT = '';

const SHEET_STUDENTS = '학생';
const SHEET_LOG = '기록';
const STUDENT_HEADERS = ['학번', '이름', '가입 시각', '최근 접속', 'salt', '비밀번호 해시',
  '연습 횟수', '푼 문제', '맞힌 문제', '정답률', '마지막 연습', '자주 틀린 단어', '최근 연습(데이터)', '틀린 단어(데이터)', '내 학습 데이터'];
const S_COLS = STUDENT_HEADERS.length;   // 15
const MY_COL = 15;                       // 헷갈리는 단어·페이지별 최근 점수·매칭 최고 기록 (어느 브라우저에서나 같게 보이도록)
const RECENT_KEEP = 30;                  // 학생 줄에 남기는 최근 연습 수
const WRONG_KEEP = 200;                  // 학생 줄에 남기는 틀린 단어 수
const LOG_HEADERS = ['시각', '학번', '이름', '학습지', '연습', '점수', '문항 수', '정답률', '틀린 단어', '페이지 ID'];
const SHEET_SETS = '지문';
const SET_HEADERS = ['ID', '이름', '순서', '단어 수', '단어(JSON)', '만든 시각', '고친 시각'];
const SHEET_TEAMS = '모둠';
const TEAM_HEADERS = ['반', '모둠', '순서'];
const SHEET_RACES = '레이스';
const RACE_HEADERS = ['ID', '이름', '반(JSON)', '지문(JSON)', '설정(JSON)', '상태', '만든 시각', '고친 시각'];
const SHEET_RACELOG = '레이스 기록';
const RACELOG_HEADERS = ['시작 시각', '레이스 ID', '레이스', '반', '모둠', '상태', '맞힌 수', '문항 수', '기록(초)', '벌점(초)', '틀린 수', '틀린 단어', '마지막 갱신'];
const TOKEN_DAYS = 180;          // 학생 로그인 유지 기간
const ADMIN_TOKEN_HOURS = 12;    // 관리자 로그인 유지 시간
const MAX_FAILS = 8;             // 비밀번호를 이만큼 틀리면
const LOCK_MINUTES = 10;         // 이 시간 동안 로그인 잠금
const HASH_ROUNDS_OLD = 300;     // 예전 방식 (로그인할 때 새 방식으로 바뀜)
const HASH_ROUNDS = 40;          // 새 방식: salt가 'v2:'로 시작
const STU_CACHE_SEC = 21600;     // 로그인 확인용 학생 정보(이름·비밀번호 해시)를 서버가 기억해 두는 시간(초, 6시간)
const CACHE_SEC = 600;           // 지문·모둠·레이스 목록을 서버가 기억해 두는 시간(초)

/* ---------------- 설치 ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = ensureSheet_(ss, SHEET_STUDENTS, STUDENT_HEADERS);
  st.getRange(1, 1, 1, S_COLS).setValues([STUDENT_HEADERS]).setFontWeight('bold');   // 새 칸 머리글
  ensureSheet_(ss, SHEET_SETS, SET_HEADERS);
  ensureSheet_(ss, SHEET_TEAMS, TEAM_HEADERS);
  ensureSheet_(ss, SHEET_RACES, RACE_HEADERS);
  ensureSheet_(ss, SHEET_RACELOG, RACELOG_HEADERS);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  props.setProperty('CACHE_GEN', String(Date.now())); GEN_ = null;   // 시트를 직접 고쳤을 때: 서버가 기억해 둔 것을 모두 버림
  if (ADMIN_PASSWORD_INIT) {
    const salt = newSalt_();
    props.setProperty('ADMIN_SALT', salt);
    props.setProperty('ADMIN_HASH', hashPw_(ADMIN_PASSWORD_INIT, salt));
  }
  if (!props.getProperty('ADMIN_HASH')) throw new Error('맨 위 ADMIN_PASSWORD_INIT에 관리자 비밀번호를 적고 다시 실행해 주세요.');
  ScriptApp.getProjectTriggers();   // 지난 기록 자동 옮기기에 필요한 권한을 미리 받아 둠
  if (props.getProperty('ARCHIVE_AUTO') === '1') setArchiveTrigger_(true);
  const moved = migrateLog_();
  return '설치 완료' + (moved ? ` (예전 기록 ${moved}줄을 학생별로 합쳤어요)` : '');
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (name === SHEET_STUDENTS) sh.getRange('A:A').setNumberFormat('@'); // 학번 앞자리 0 보존
  return sh;
}

/* ---------------- 웹 앱 입구 ---------------- */
function doGet() {
  return json_({ok: true, app: 'word-practice'});
}

function doPost(e) {
  GEN_ = null;   // 요청마다 캐시 세대 번호를 새로 읽음
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ok: false, error: '요청 형식이 올바르지 않아요.'}); }
  try {
    const fn = ACTIONS[req.action];
    if (!fn) return json_({ok: false, error: '알 수 없는 요청이에요.'});
    return json_(fn(req));
  } catch (err) {
    return json_({ok: false, code: (err && err.code) || '', unexpected: !(err && err.userMessage), error: (err && err.userMessage) || '서버에서 문제가 생겼어요. 잠시 후 다시 해 주세요.', detail: String(err && err.message || err)});
  }
}

const ACTIONS = {
  register: register_,
  login: login_,
  changePw: changePw_,
  mySync: mySync_,
  log: log_,
  me: me_,
  adminLogin: adminLogin_,
  admin: admin_,
  adminStudent: adminStudent_,
  adminReset: adminReset_,
  sets: sets_,
  adminSetSave: adminSetSave_,
  adminSetDelete: adminSetDelete_,
  adminSetOrder: adminSetOrder_,
  adminOcr: adminOcr_,
  teams: teams_,
  teamLogin: teamLogin_,
  races: races_,
  raceStart: raceStart_,
  raceProgress: raceProgress_,
  raceFinish: raceFinish_,
  adminTeamsSave: adminTeamsSave_,
  adminRaces: adminRaces_,
  adminRaceSave: adminRaceSave_,
  adminRaceStatus: adminRaceStatus_,
  adminRaceDelete: adminRaceDelete_,
  adminBoard: adminBoard_,
  adminRaceReset: adminRaceReset_,
  adminArchive: adminArchive_
};

/* ---------------- 학생 ---------------- */
function register_(req) {
  const sid = cleanSid_(req.sid), name = cleanName_(req.name), pw = cleanPin_(req.pw);
  return withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS);
    const row = findStudentRow_(sh, sid);
    if (row) {
      const hash = sh.getRange(row, 6).getValue();
      if (hash) fail_('이미 등록된 학번이에요. 로그인해 주세요. 비밀번호가 기억나지 않으면 선생님께 초기화를 부탁하세요.');
      // 선생님이 비밀번호를 초기화한 학생: 새 비밀번호로 다시 등록
      const salt = newSalt_(), now = new Date();
      sh.getRange(row, 2, 1, 5).setValues([[asText_(name), sh.getRange(row, 3).getValue() || now, now, asText_(salt), asText_(hashPw_(pw, salt))]]);
    } else {
      const salt = newSalt_(), now = new Date();
      sh.appendRow([asText_(sid), asText_(name), now, now, asText_(salt), asText_(hashPw_(pw, salt))]);
    }
    stuCacheDrop_(sid);
    // 등록은 '기록' 탭에 남기지 않음 ('학생' 탭의 가입 시각으로 확인)
    return {ok: true, token: makeToken_('s', sid, TOKEN_DAYS * 24), student: {sid: sid, name: name}, setsPack: setsFor_(req.v)};
  });
}

/* 로그인 확인용 학생 정보를 캐시에 두어, 로그인할 때 시트를 읽거나 쓰지 않게 함 (로그인이 빨라짐).
   비밀번호가 바뀌는 곳(등록·바꾸기·초기화)에서 지우거나 새로 넣음 */
/* 캐시 이름에 '세대' 번호를 붙여 둠. 선생님이 시트를 직접 고친 뒤 setup을 실행하면 번호가 바뀌어
   서버가 기억해 둔 것(로그인 정보·내 기록·지문·모둠·레이스 목록)을 모두 버리고 시트에서 새로 읽음 */
let GEN_ = null;
function gk_(key) {
  if (GEN_ === null) GEN_ = PropertiesService.getScriptProperties().getProperty('CACHE_GEN') || '0';
  return key + '@' + GEN_;
}
function stuCacheGet_(sid) { const v = CacheService.getScriptCache().get(gk_('stu_' + sid)); return v ? JSON.parse(v) : null; }
function stuCachePut_(sid, c) { CacheService.getScriptCache().put(gk_('stu_' + sid), JSON.stringify({name: c.name, salt: c.salt, hash: c.hash}), STU_CACHE_SEC); }
function stuCacheDrop_(sid) { CacheService.getScriptCache().remove(gk_('stu_' + sid)); }
function stuFromSheet_(sid) {
  const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
  if (!row) return null;
  const r = sh.getRange(row, 1, 1, 6).getValues()[0];
  const c = {sh: sh, row: row, name: String(r[1]), salt: String(r[4] || ''), hash: String(r[5] || '')};
  if (c.hash) stuCachePut_(sid, c); else stuCacheDrop_(sid);
  return c;
}

/* 학번과 비밀번호 확인: 맞으면 학생 정보를, 틀리면 잠금 횟수를 올리고 오류. needRow면 시트에서 줄 위치까지 읽음 */
function checkStudentPw_(sid, pw, needRow) {
  const cache = CacheService.getScriptCache(), key = 'fail_' + sid;
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) fail_(`비밀번호를 여러 번 틀려서 ${LOCK_MINUTES}분 동안 로그인할 수 없어요.`);
  let c = needRow ? null : stuCacheGet_(sid), fromCache = !!c;
  if (c && hashPw_(pw, c.salt) !== c.hash) { c = null; fromCache = false; }   // 캐시가 오래됐을 수 있어 시트로 한 번 더 확인
  if (!c) c = stuFromSheet_(sid);
  if (!c) fail_('등록되지 않은 학번이에요. 처음이라면 "처음 등록"을 눌러 주세요.');
  if (!c.hash) fail_('비밀번호가 초기화되었어요. "처음 등록"에서 새 비밀번호를 정해 주세요.');
  if (!fromCache && hashPw_(pw, c.salt) !== c.hash) {
    cache.put(key, String(fails + 1), LOCK_MINUTES * 60);
    fail_('학번이나 비밀번호가 맞지 않아요.');
  }
  if (fails) cache.remove(key);
  return c;
}

function login_(req) {
  const sid = cleanSid_(req.sid), pw = String(req.pw || '');
  const st = checkStudentPw_(sid, pw, false);
  if (!/^v2:/.test(st.salt)) {   // 예전 방식 비밀번호를 빠른 방식으로 바꿔 둠 (학생마다 한 번)
    withLock_(() => {
      const c = stuFromSheet_(sid), s2 = newSalt_(), h2 = hashPw_(pw, s2);
      c.sh.getRange(c.row, 5, 1, 2).setValues([[asText_(s2), asText_(h2)]]);
      stuCachePut_(sid, {name: c.name, salt: s2, hash: h2});
    });
  }
  // 최근 접속 시각은 연습 기록을 남길 때 함께 적음 (로그인할 때는 시트에 쓰지 않아 빠름)
  return {ok: true, token: makeToken_('s', sid, TOKEN_DAYS * 24), student: {sid: sid, name: st.name}, setsPack: setsFor_(req.v)};
}

/* 학생이 스스로 비밀번호 바꾸기: 지금 비밀번호를 확인하고 새 비밀번호(숫자 5자리)로 바꾼 뒤 로그인 */
function changePw_(req) {
  const sid = cleanSid_(req.sid), pw = String(req.pw || ''), newPw = cleanPin_(req.newPw);
  return withLock_(() => {
    const st = checkStudentPw_(sid, pw, true);
    const s2 = newSalt_(), h2 = hashPw_(newPw, s2);
    st.sh.getRange(st.row, 4, 1, 3).setValues([[new Date(), asText_(s2), asText_(h2)]]);
    stuCachePut_(sid, {name: st.name, salt: s2, hash: h2});
    return {ok: true, token: makeToken_('s', sid, TOKEN_DAYS * 24), student: {sid: sid, name: st.name}, setsPack: setsFor_(req.v)};
  });
}

function log_(req) {
  const sid = checkToken_(req.token, 's');
  const score = Math.max(0, Math.floor(Number(req.score) || 0));
  const total = Math.max(0, Math.floor(Number(req.total) || 0));
  const wrong = (Array.isArray(req.wrong) ? req.wrong : []).slice(0, 200).map(w => String(w).slice(0, 80)).filter(Boolean);
  const when = req.at ? new Date(req.at) : new Date();
  // 카드 넘기기처럼 점수가 없는 연습은 기록하지 않음. 함께 온 내 학습 데이터는 저장함
  const entry = total ? {at: (isNaN(when) ? new Date() : when).toISOString(), pageTitle: String(req.pageTitle || '').slice(0, 60), mode: String(req.mode || '').slice(0, 20),
    score: Math.min(score, total), total: total, page: String(req.page || '').slice(0, 30), wrong: wrong} : null;
  if (!entry && !req.my) return {ok: true};
  const out = addPractice_(sid, entry, req.my);
  // 학생 화면이 '내 기록'과 '내 학습 데이터'를 바로 새로 고칠 수 있게 함께 보냄
  return Object.assign({ok: true}, out.me ? {me: out.me} : {}, out.myPack ? {myPack: out.myPack} : {});
}

/* 내 기록: 연습할 때마다 서버 캐시에 새로 넣어 두므로 보통은 시트를 열지 않음 */
function me_(req) {
  const sid = checkToken_(req.token, 's');
  const c = CacheService.getScriptCache().get(gk_('me_' + sid));
  if (c) return Object.assign({ok: true}, JSON.parse(c));
  const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
  if (!row) fail_('학생 정보를 찾을 수 없어요. 다시 로그인해 주세요.', 'auth');
  const r = sh.getRange(row, 1, 1, S_COLS).getValues()[0];
  return Object.assign({ok: true}, meCachePut_(sid, String(r[1]), rowStats_(r)));
}
function meCachePut_(sid, name, st) {
  const me = {at: Date.now(), student: {sid: sid, name: name}, stats: {sessions: st.sessions, questions: st.questions, correct: st.correct}, wrong: st.wrong.slice(0, 60)};
  try { CacheService.getScriptCache().put(gk_('me_' + sid), JSON.stringify(me), STU_CACHE_SEC); } catch (e) {}
  return me;
}

/* ---------------- 내 학습 데이터: 헷갈리는 단어, 페이지별 최근 점수, 매칭 최고 기록 ----------------
   브라우저마다 따로 저장하던 것을 학생 줄의 한 칸에 모아 두고, 브라우저가 가진 것과 시각을 비교해 합침.
   선생님이 그 칸을 지우고 setup을 실행하면 모든 브라우저에서 비워짐 ('epoch'가 달라져 예전 것을 버림) */
const MY_LIMIT = {items: 600, removed: 300, last: 300, best: 300};
function sortedObj_(o, limit, by) {
  let keys = Object.keys(o || {});
  if (limit && keys.length > limit) keys = keys.sort((a, b) => by(o[b]) - by(o[a])).slice(0, limit);
  const out = {}; keys.sort().forEach(k => { out[k] = o[k]; }); return out;
}
function cleanMy_(m) {
  m = m && typeof m === 'object' ? m : {};
  const items = {}, removed = {}, last = {}, best = {};
  Object.keys(m.items || {}).forEach(k => { const x = m.items[k]; if (k.length <= 120 && x) { items[k] = {added: Number(x.added) || 0}; if (x.star) items[k].star = true; } });
  Object.keys(m.removed || {}).forEach(k => { if (k.length <= 120) removed[k] = Number(m.removed[k]) || 0; });
  Object.keys(m.last || {}).forEach(k => { const x = m.last[k]; if (k.length <= 60 && x) last[k] = {m: String(x.m || '').slice(0, 20), s: Number(x.s) || 0, n: Number(x.n) || 0, at: Number(x.at) || 0}; });
  Object.keys(m.best || {}).forEach(k => { const t = Number(m.best[k]); if (k.length <= 60 && t > 0) best[k] = t; });
  const out = {items: sortedObj_(items, MY_LIMIT.items, x => x.added), removed: sortedObj_(removed, MY_LIMIT.removed, x => x),
    last: sortedObj_(last, MY_LIMIT.last, x => x.at), best: sortedObj_(best, MY_LIMIT.best, x => -x)};
  if (m.epoch) out.epoch = String(m.epoch).slice(0, 20);
  return out;
}
/* 두 쪽을 합침: 넣은 시각과 뺀 시각 중 나중 것이 이김, 최근 점수는 나중 것, 최고 기록은 짧은 것 */
function mergeMy_(a, b) {
  const items = {}, removed = {}, last = {}, best = {};
  const keys = {}; [a.items, b.items, a.removed, b.removed].forEach(o => Object.keys(o || {}).forEach(k => { keys[k] = 1; }));
  Object.keys(keys).forEach(k => {
    const ia = (a.items || {})[k], ib = (b.items || {})[k];
    const add = Math.max(ia ? ia.added : -1, ib ? ib.added : -1), rem = Math.max((a.removed || {})[k] || 0, (b.removed || {})[k] || 0);
    if ((ia || ib) && add > rem) { items[k] = {added: add}; if ((ia && ia.star) || (ib && ib.star)) items[k].star = true; }
    else if (rem) removed[k] = rem;
  });
  [a.last, b.last].forEach(o => Object.keys(o || {}).forEach(k => { if (!last[k] || o[k].at > last[k].at) last[k] = o[k]; }));
  [a.best, b.best].forEach(o => Object.keys(o || {}).forEach(k => { if (!best[k] || o[k] < best[k]) best[k] = o[k]; }));
  return cleanMy_({items: items, removed: removed, last: last, best: best});
}
/* 브라우저가 보낸 것과 시트 칸의 것을 합침. epoch가 다르면 선생님이 비운 뒤라서 브라우저의 예전 것을 버림 */
function myMerge_(cell, cliMy) {
  let my = cleanMy_(parseJ_(cell, {})), changed = false;
  if (cliMy) {
    const cli = cleanMy_(cliMy);
    if (!cli.epoch || cli.epoch === my.epoch) {
      const merged = mergeMy_(my, cli);
      merged.epoch = my.epoch || Utilities.getUuid().slice(0, 8);
      if (JSON.stringify(merged) !== JSON.stringify(my)) { my = merged; changed = true; }
    }
  }
  return {my: my, changed: changed};
}
function myPack_(sid, my) {
  const v = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(my), Utilities.Charset.UTF_8)).slice(0, 12);
  try { CacheService.getScriptCache().put(gk_('my_' + sid), JSON.stringify({my: my, v: v}), STU_CACHE_SEC); } catch (e) {}
  return {my: my, v: v};
}
function mySync_(req) {
  const sid = checkToken_(req.token, 's');
  if (!req.my) {   // 가져가기만: 서버가 기억해 둔 것이 있으면 시트를 열지 않음
    const c = CacheService.getScriptCache().get(gk_('my_' + sid));
    if (c) { const o = JSON.parse(c); return o.v === req.v ? {ok: true, same: true, v: o.v} : {ok: true, my: o.my, v: o.v}; }
  }
  return withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
    if (!row) fail_('학생 정보를 찾을 수 없어요. 다시 로그인해 주세요.', 'auth');
    const mm = myMerge_(sh.getRange(row, MY_COL).getValue(), req.my);
    if (mm.changed) sh.getRange(row, MY_COL).setValue(JSON.stringify(mm.my));
    const pack = myPack_(sid, mm.my);
    if (!req.my && req.v === pack.v) return {ok: true, same: true, v: pack.v};
    return Object.assign({ok: true}, pack);
  });
}

/* ---------------- 관리자 ---------------- */
function adminLogin_(req) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache(), key = 'fail_admin';
  if (Number(cache.get(key) || 0) >= MAX_FAILS) fail_(`비밀번호를 여러 번 틀려서 ${LOCK_MINUTES}분 동안 잠겼어요.`);
  const pw = String(req.pw || ''), salt = props.getProperty('ADMIN_SALT');
  if (hashPw_(pw, salt) !== props.getProperty('ADMIN_HASH')) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCK_MINUTES * 60);
    fail_('관리자 비밀번호가 맞지 않아요.');
  }
  cache.remove(key);
  if (!/^v2:/.test(String(salt))) { const s2 = newSalt_(); props.setProperty('ADMIN_SALT', s2); props.setProperty('ADMIN_HASH', hashPw_(pw, s2)); }
  return {ok: true, token: makeToken_('a', 'admin', ADMIN_TOKEN_HOURS)};
}

function admin_(req) {
  checkToken_(req.token, 'a');
  const sh = sheet_(SHEET_STUDENTS), n = sh.getLastRow();
  const rows = n < 2 ? [] : sh.getRange(2, 1, n - 1, S_COLS).getValues().filter(r => String(r[0]).trim());
  const students = [], recent = [], wrongCount = {};
  rows.forEach(r => {
    const sid = String(r[0]).trim(), name = String(r[1]), st = rowStats_(r);
    students.push({sid: sid, name: name, joined: iso_(r[2]), lastSeen: iso_(r[3]), hasPw: !!r[5], sessions: st.sessions, questions: st.questions, correct: st.correct, lastPractice: st.last});
    st.recent.forEach(e => recent.push(Object.assign({sid: sid, name: name}, e)));
    st.wrong.forEach(w => { const k = w.page + '|' + w.en; const o = wrongCount[k] || (wrongCount[k] = {en: w.en, page: w.page, count: 0, students: 0}); o.count += w.count; o.students++; });
  });
  recent.sort((a, b) => a.at < b.at ? 1 : -1);
  return {
    ok: true,
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    students: students.sort((a, b) => String(a.sid).localeCompare(String(b.sid))),
    recent: recent.slice(0, 300),
    wrongTop: Object.keys(wrongCount).map(k => wrongCount[k]).sort((a, b) => b.count - a.count).slice(0, 50),
    archive: archiveInfo_()
  };
}

function adminStudent_(req) {
  checkToken_(req.token, 'a');
  const sid = cleanSid_(req.sid);
  return Object.assign({ok: true}, statsFor_(sid));
}

function adminReset_(req) {
  checkToken_(req.token, 'a');
  const sid = cleanSid_(req.sid);
  return withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
    if (!row) fail_('그 학번의 학생이 없어요.');
    sh.getRange(row, 5, 1, 2).setValues([['', '']]);
    CacheService.getScriptCache().remove('fail_' + sid);
    stuCacheDrop_(sid);
    return {ok: true};
  });
}

/* ---------------- 통계 ---------------- */
function statsFor_(sid) {
  const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
  if (!row) fail_('학생 정보를 찾을 수 없어요.', 'auth');
  const r = sh.getRange(row, 1, 1, S_COLS).getValues()[0], st = rowStats_(r);
  return {
    student: {sid: sid, name: String(r[1])},
    stats: {sessions: st.sessions, questions: st.questions, correct: st.correct},
    recent: st.recent.slice().reverse().map(e => Object.assign({sid: sid}, e)),
    wrong: st.wrong.slice(0, 60)
  };
}
function isPractice_(l) { return l.total > 0; }



/* ---------------- 서버가 잠깐 기억해 두는 목록 (지문·모둠·레이스) ---------------- */
function cached_(key, fn) {
  key = gk_(key);
  const c = CacheService.getScriptCache();
  try {
    const head = c.get(key);
    if (head) {
      const k = head.indexOf(':'), n = Number(head.slice(0, k));
      if (n === 0) return JSON.parse(head.slice(k + 1));
      const parts = []; for (let i = 0; i < n; i++) parts.push(c.get(key + '_' + i));
      if (parts.every(x => x !== null)) return JSON.parse(parts.join(''));
    }
  } catch (e) {}
  const v = fn();
  try {
    const str = JSON.stringify(v), size = 90000;
    if (str.length <= size) c.put(key, '0:' + str, CACHE_SEC);
    else if (str.length <= size * 9) {
      const n = Math.ceil(str.length / size);
      for (let i = 0; i < n; i++) c.put(key + '_' + i, str.slice(i * size, (i + 1) * size), CACHE_SEC);
      c.put(key, n + ':', CACHE_SEC);
    }
  } catch (e) {}
  return v;
}
function forget_(key) { try { CacheService.getScriptCache().remove(gk_(key)); } catch (e) {} }


/* ---------------- 학생당 한 줄: 연습 결과를 학생 줄에 합쳐 둠 ---------------- */
function parseJ_(v, d) { try { const x = JSON.parse(String(v || '')); return x == null ? d : x; } catch (e) { return d; } }
function rowStats_(r) {
  return {sessions: Number(r[6]) || 0, questions: Number(r[7]) || 0, correct: Number(r[8]) || 0, last: iso_(r[10]),
    recent: parseJ_(r[12], []), wrong: parseJ_(r[13], [])};
}
/* entries: [{at, pageTitle, mode, score, total, page, wrong: [en…]}] in time order */
function mergeStats_(cur, entries) {
  const wrong = {};
  (cur.wrong || []).forEach(w => { wrong[w.page + '|' + w.en] = w; });
  entries.forEach(e => {
    cur.sessions++; cur.questions += e.total; cur.correct += e.score;
    if (!cur.last || e.at > cur.last) cur.last = e.at;
    (e.wrong || []).forEach(en => { const k = e.page + '|' + en; const o = wrong[k] || (wrong[k] = {en: en, page: e.page, count: 0, last: ''}); o.count++; if (e.at > o.last) o.last = e.at; });
    cur.recent.push({at: e.at, pageTitle: e.pageTitle, mode: e.mode, score: e.score, total: e.total, page: e.page, wrong: (e.wrong || []).slice(0, 20)});
  });
  cur.recent.sort((a, b) => a.at < b.at ? -1 : 1);
  cur.recent = cur.recent.slice(-RECENT_KEEP);
  cur.wrong = Object.keys(wrong).map(k => wrong[k]).sort((a, b) => b.count - a.count || (b.last > a.last ? 1 : -1)).slice(0, WRONG_KEEP);
  return cur;
}
function statsCells_(st) {
  const top = st.wrong.slice(0, 10).map(w => `${w.en}(${w.count})`).join(', ');
  return [st.sessions, st.questions, st.correct, st.questions ? Math.round(st.correct / st.questions * 100) / 100 : '',
    st.last ? new Date(st.last) : '', asText_(top), JSON.stringify(st.recent), JSON.stringify(st.wrong)];
}
/* 연습 결과와 내 학습 데이터를 학생 줄에 한 번에 씀 (시트를 한 번 읽고, 이어진 칸을 한 번에 씀) */
function addPractice_(sid, entry, cliMy) {
  return withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
    if (!row) fail_('학생 정보를 찾을 수 없어요. 다시 로그인해 주세요.', 'auth');
    const r = sh.getRange(row, 1, 1, S_COLS).getValues()[0];
    const mm = myMerge_(r[MY_COL - 1], cliMy), out = {};
    if (entry) {
      const st = mergeStats_(rowStats_(r), [entry]);
      sh.getRange(row, 7, 1, 9).setValues([statsCells_(st).concat([mm.changed ? JSON.stringify(mm.my) : r[MY_COL - 1]])]);   // G~O열
      sh.getRange(row, 4).setValue(new Date());   // 최근 접속
      out.me = meCachePut_(sid, String(r[1]), st);
    } else if (mm.changed) {
      sh.getRange(row, MY_COL).setValue(JSON.stringify(mm.my));
    }
    if (cliMy) out.myPack = myPack_(sid, mm.my);
    return out;
  });
}
/* 예전 '기록' 탭(연습 한 번에 한 줄)을 학생 줄에 합치고 탭 이름을 '지난 기록'으로 바꿈 */
function migrateLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), old = ss.getSheetByName(SHEET_LOG);
  if (!old || old.getLastRow() < 2) return 0;
  const n = old.getLastRow();
  const rows = old.getRange(2, 1, n - 1, LOG_HEADERS.length).getValues();
  const by = {};
  rows.forEach(r => {
    const total = Number(r[6]) || 0; if (!total) return;   // 예전 '등록'·'로그인' 줄은 건너뜀
    const sid = String(r[1]).trim(); if (!sid) return;
    (by[sid] = by[sid] || []).push({at: iso_(r[0]), pageTitle: String(r[3]), mode: String(r[4]), score: Number(r[5]) || 0, total: total,
      page: String(r[9] || ''), wrong: String(r[8] || '').split(' ; ').filter(Boolean)});
  });
  withLock_(() => {
    const sh = sheet_(SHEET_STUDENTS), m = sh.getLastRow();
    if (m >= 2) {
      const all = sh.getRange(2, 1, m - 1, S_COLS).getValues();
      all.forEach((r, i) => {
        const sid = String(r[0]).trim(), list = by[sid]; if (!list) return;
        list.sort((a, b) => a.at < b.at ? -1 : 1);
        sh.getRange(i + 2, 7, 1, 8).setValues([statsCells_(mergeStats_(rowStats_(r), list))]);
      });
    }
    const done = ss.getSheetByName(SHEET_LOG_OLD);
    if (done) {   // '지난 기록'이 이미 있으면 그 아래에 붙이고 '기록' 탭은 지움
      done.getRange(done.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows.map(r => r.map(v => typeof v === 'string' ? asText_(v) : v)));
      ss.deleteSheet(old);
    } else old.setName(SHEET_LOG_OLD);
  });
  return rows.length;
}

/* ---------------- 지문 목록 (학생은 읽기만, 선생님만 등록·수정) ---------------- */
function readSets_() {
  const sh = sheet_(SHEET_SETS), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, SET_HEADERS.length).getValues().map((r, i) => {
    let words = []; try { words = JSON.parse(String(r[4] || '[]')); } catch (e) {}
    return {row: i + 2, id: String(r[0]).trim(), name: String(r[1]), order: Number(r[2]) || 0, words: words, updated: iso_(r[6])};
  }).filter(x => x.id).sort((a, b) => a.order - b.order || a.row - b.row);
}
/* 지문 꾸러미: 내용으로 만든 짧은 버전과 함께 보냄. 브라우저가 가진 버전과 같으면 단어는 보내지 않음 */
function setsBundle_() {
  return cached_('c_sets2', () => {
    const sets = readSets_().map(x => ({id: x.id, name: x.name, words: x.words, updated: x.updated}));
    const v = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(sets), Utilities.Charset.UTF_8)).slice(0, 12);
    return {v: v, sets: sets};
  });
}
function setsFor_(have) {
  const b = setsBundle_();
  return have && have === b.v ? {v: b.v, same: true} : {v: b.v, sets: b.sets};
}
function sets_(req) {
  return Object.assign({ok: true}, setsFor_(req && req.v));
}
function cleanWords_(list) {
  if (!Array.isArray(list)) fail_('단어 목록 형식이 올바르지 않아요.');
  const out = [];
  list.forEach(w => {
    if (!Array.isArray(w)) return;
    const en = String(w[0] || '').trim().slice(0, 80), ko = String(w[1] || '').trim().slice(0, 120), note = String(w[2] || '').trim().slice(0, 160);
    if (en && ko) out.push(note ? [en, ko, note] : [en, ko]);
  });
  if (out.length > 600) fail_('한 지문에 단어는 600개까지 넣을 수 있어요.');
  return out;
}
function adminSetSave_(req) {
  checkToken_(req.token, 'a');
  const name = String(req.name || '').trim().replace(/\s+/g, ' ');
  if (!name || name.length > 40) fail_('지문 이름을 1~40자로 써 주세요.');
  const words = cleanWords_(req.words || []);
  const json = JSON.stringify(words);
  if (json.length > 45000) fail_('단어가 너무 많아요. 지문을 둘로 나눠 주세요.');
  return withLock_(() => {
    const sh = sheet_(SHEET_SETS), now = new Date(), list = readSets_();
    const cur = req.id ? list.find(x => x.id === String(req.id)) : null;
    if (req.id && !cur) fail_('그 지문을 찾을 수 없어요. 새로고침해 주세요.');
    if (cur) {
      sh.getRange(cur.row, 2, 1, 4).setValues([[asText_(name), cur.order, words.length, json]]);
      sh.getRange(cur.row, 7).setValue(now);
      forget_('c_sets2');
      return {ok: true, id: cur.id, setsPack: setsFor_('')};
    }
    const id = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 7);
    const order = list.reduce((m, x) => Math.max(m, x.order), 0) + 1;
    sh.appendRow([id, asText_(name), order, words.length, json, now, now]);
    forget_('c_sets2');
    return {ok: true, id: id, setsPack: setsFor_('')};
  });
}
function adminSetDelete_(req) {
  checkToken_(req.token, 'a');
  return withLock_(() => {
    const cur = readSets_().find(x => x.id === String(req.id || ''));
    if (!cur) fail_('그 지문을 찾을 수 없어요.');
    sheet_(SHEET_SETS).deleteRow(cur.row);
    forget_('c_sets2');
    return {ok: true, setsPack: setsFor_('')};
  });
}
function adminSetOrder_(req) {
  checkToken_(req.token, 'a');
  const ids = (Array.isArray(req.ids) ? req.ids : []).map(String);
  return withLock_(() => {
    const sh = sheet_(SHEET_SETS);
    readSets_().forEach(x => { const k = ids.indexOf(x.id); sh.getRange(x.row, 3).setValue(k < 0 ? ids.length + x.order : k + 1); });
    forget_('c_sets2');
    return {ok: true, setsPack: setsFor_('')};
  });
}

/* 사진 인식 점검: Apps Script 편집기에서 이 함수를 골라 실행하면 필요한 권한을 묻고, 실행 로그에 결과가 나와요 */
function checkOcr() {
  if (typeof Drive === 'undefined') throw new Error("Drive API가 추가되지 않았어요. 왼쪽 '서비스' 옆 + → Drive API → 추가를 해 주세요.");
  const blob = Utilities.newBlob('Word check 단어 확인', 'text/plain', 'check.txt');
  let id = '';
  try {
    const meta = {mimeType: MimeType.GOOGLE_DOCS};
    if (Drive.Files.create) { meta.name = '단어장 점검 (임시)'; id = Drive.Files.create(meta, blob).id; }
    else { meta.title = '단어장 점검 (임시)'; id = Drive.Files.insert(meta, blob, {convert: true}).id; }
    const text = DocumentApp.openById(id).getBody().getText();
    console.log('점검 완료: 드라이브와 문서 권한이 모두 있어요. 읽은 글자 → ' + text.trim());
    return '점검 완료';
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
}

/* 단어표 사진 → 글자: 선생님 구글 드라이브의 무료 글자 인식(OCR)을 쓰고, 임시 문서는 바로 휴지통으로 */
function adminOcr_(req) {
  checkToken_(req.token, 'a');
  const b64 = String(req.data || '');
  if (!b64) fail_('사진이 없어요.');
  if (b64.length > 14000000) fail_('사진이 너무 커요. 10MB보다 작은 사진으로 올려 주세요.');
  if (typeof Drive === 'undefined') fail_("사진 인식을 쓰려면 Apps Script 편집기 왼쪽 '서비스' 옆 + → Drive API → 추가를 한 뒤, 새 버전으로 다시 배포해 주세요.", 'no_drive');
  const mime = /^image\/(png|jpeg|gif|bmp|webp)$/.test(String(req.mime)) ? String(req.mime) : 'image/jpeg';
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, 'word-photo');
  let id = '';
  try {
    const meta = {mimeType: MimeType.GOOGLE_DOCS};
    if (Drive.Files.create) { meta.name = '단어장 사진 인식 (임시)'; id = Drive.Files.create(meta, blob, {ocrLanguage: 'ko'}).id; }
    else { meta.title = '단어장 사진 인식 (임시)'; id = Drive.Files.insert(meta, blob, {ocr: true, ocrLanguage: 'ko'}).id; }
    const text = DocumentApp.openById(id).getBody().getText();
    return {ok: true, text: text};
  } catch (err) {
    if (err && err.userMessage) throw err;
    const why = String(err && err.message || err);
    console.error('adminOcr: ' + why);
    if (/permission|권한|authoriz|scope|access/i.test(why)) fail_("구글 드라이브 권한이 아직 없어요. Apps Script에서 checkOcr 함수를 한 번 실행해 권한을 허용한 뒤, 새 버전으로 다시 배포해 주세요. (" + why.slice(0, 160) + ')', 'ocr_auth');
    fail_('구글 글자 인식에서 오류가 났어요: ' + why.slice(0, 200) + ' — Apps Script에서 checkOcr 함수를 실행해 점검해 주세요.', 'ocr_error');
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
}


/* ---------------- 반·모둠 (선생님이 등록, 모둠 입장 화면에 나옴) ---------------- */
function readTeams_() {
  const sh = sheet_(SHEET_TEAMS), n = sh.getLastRow(); if (n < 2) return [];
  const rows = sh.getRange(2, 1, n - 1, 3).getValues().map(r => ({cls: String(r[0]).trim(), team: String(r[1]).trim(), order: Number(r[2]) || 0})).filter(r => r.cls && r.team);
  rows.sort((a, b) => a.order - b.order);
  const classes = [];
  rows.forEach(r => { let c = classes.find(x => x.name === r.cls); if (!c) classes.push(c = {name: r.cls, teams: []}); if (c.teams.indexOf(r.team) < 0) c.teams.push(r.team); });
  return classes;
}
function teams_() { return {ok: true, classes: cached_('c_teams', readTeams_)}; }
function adminTeamsSave_(req) {
  checkToken_(req.token, 'a');
  const list = Array.isArray(req.classes) ? req.classes : [];
  const rows = []; let order = 0;
  list.forEach(c => {
    const cls = String(c && c.name || '').trim().slice(0, 20); if (!cls) return;
    (Array.isArray(c.teams) ? c.teams : []).forEach(t => { const team = String(t || '').trim().slice(0, 20); if (team) rows.push([asText_(cls), asText_(team), ++order]); });
  });
  if (rows.length > 300) fail_('모둠은 300개까지 등록할 수 있어요.');
  return withLock_(() => {
    const sh = sheet_(SHEET_TEAMS), n = sh.getLastRow();
    if (n > 1) sh.deleteRows(2, n - 1);
    if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
    forget_('c_teams');
    return {ok: true, classes: readTeams_()};
  });
}
function teamLogin_(req) {
  const cls = String(req.cls || '').trim(), team = String(req.team || '').trim();
  const c = cached_('c_teams', readTeams_).find(x => x.name === cls);
  if (!c || c.teams.indexOf(team) < 0) fail_('등록된 반·모둠이 아니에요. 선생님께 확인해 주세요.');
  const token = makeToken_('t', cls + '\n' + team, 10);
  const first = races_({token: token, v: req.v});   // 레이스 목록과 지문을 같이 보내 입장 뒤 한 번 더 기다리지 않게
  return {ok: true, token: token, team: {cls: cls, team: team}, races: first.races, setsPack: first.setsPack};
}
function teamOf_(token) {
  const id = checkToken_(token, 't'), k = id.indexOf('\n');
  return {cls: id.slice(0, k), team: id.slice(k + 1)};
}

/* ---------------- 레이스 (선생님이 지문을 묶어 만듦) ---------------- */
function readRaces_() {
  const sh = sheet_(SHEET_RACES), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, RACE_HEADERS.length).getValues().map((r, i) => {
    const j = v => { try { return JSON.parse(String(v || '')); } catch (e) { return null; } };
    return {row: i + 2, id: String(r[0]).trim(), name: String(r[1]), classes: j(r[2]) || [], sources: j(r[3]) || [], settings: j(r[4]) || {}, status: String(r[5] || '준비'), created: iso_(r[6]), updated: iso_(r[7])};
  }).filter(x => x.id);
}
const RACE_STATUS = ['준비', '진행', '마감'];
function cleanRace_(req) {
  const name = String(req.name || '').trim().replace(/\s+/g, ' ');
  if (!name || name.length > 40) fail_('레이스 이름을 1~40자로 써 주세요.');
  const sources = (Array.isArray(req.sources) ? req.sources : []).map(String).filter(s => /^[A-Za-z0-9-]{1,30}$/.test(s)).slice(0, 40);
  if (!sources.length) fail_('레이스에 쓸 지문을 하나 이상 골라 주세요.');
  let classes = (Array.isArray(req.classes) ? req.classes : []).map(c => String(c).trim()).filter(Boolean).slice(0, 40);
  if (classes.indexOf('*') >= 0) classes = ['*'];   // 전체 반 (나중에 추가한 반도 포함)
  if (!classes.length) fail_('레이스에 참여할 반을 하나 이상 골라 주세요.');
  const s = req.settings || {};
  const mode = ['time', 'match', 'spell'].indexOf(s.mode) >= 0 ? s.mode : 'time';
  const settings = {
    mode: mode,
    count: Math.max(0, Math.min(300, Math.floor(Number(s.count) || 0))),
    penalty: Math.max(0, Math.min(30, Math.floor(Number(s.penalty) || 0))),
    rel: !!s.rel,
    live: !!s.live,
    showRank: !!s.showRank
  };
  return {name: name, sources: sources, classes: classes, settings: settings};
}
function adminRaceSave_(req) {
  checkToken_(req.token, 'a');
  const r = cleanRace_(req);
  return withLock_(() => {
    const sh = sheet_(SHEET_RACES), now = new Date();
    const cur = req.id ? readRaces_().find(x => x.id === String(req.id)) : null;
    if (req.id && !cur) fail_('그 레이스를 찾을 수 없어요. 새로고침해 주세요.');
    const vals = [asText_(r.name), JSON.stringify(r.classes), JSON.stringify(r.sources), JSON.stringify(r.settings)];
    forget_('c_races');
    if (cur) { sh.getRange(cur.row, 2, 1, 4).setValues([vals]); sh.getRange(cur.row, 8).setValue(now); return {ok: true, id: cur.id}; }
    const id = 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 7);
    sh.appendRow([id].concat(vals, ['준비', now, now]));
    return {ok: true, id: id};
  });
}
function adminRaceStatus_(req) {
  checkToken_(req.token, 'a');
  const st = String(req.status || '');
  if (RACE_STATUS.indexOf(st) < 0) fail_('알 수 없는 상태예요.');
  return withLock_(() => {
    const cur = readRaces_().find(x => x.id === String(req.id || ''));
    if (!cur) fail_('그 레이스를 찾을 수 없어요.');
    sheet_(SHEET_RACES).getRange(cur.row, 6).setValue(st);
    forget_('c_races');
    return {ok: true};
  });
}
function adminRaceDelete_(req) {
  checkToken_(req.token, 'a');
  return withLock_(() => {
    const cur = readRaces_().find(x => x.id === String(req.id || ''));
    if (!cur) fail_('그 레이스를 찾을 수 없어요.');
    sheet_(SHEET_RACES).deleteRow(cur.row);
    forget_('c_races');
    return {ok: true};
  });
}
function adminRaces_(req) {
  checkToken_(req.token, 'a');
  const logs = readRaceLogs_();
  return {ok: true, setsPack: setsFor_(req.v), archive: archiveInfo_(), classes: readTeams_(), races: readRaces_().map(x => ({id: x.id, name: x.name, classes: x.classes, sources: x.sources, settings: x.settings, status: x.status,
    done: logs.filter(l => l.race === x.id && l.status === '완료').length, running: logs.filter(l => l.race === x.id && l.status === '진행 중').length}))};
}

/* ---------------- 레이스 기록 ---------------- */
function readRaceLogs_() {
  const sh = sheet_(SHEET_RACELOG), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, RACELOG_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, started: r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || 0, race: String(r[1]).trim(), raceName: String(r[2]), cls: String(r[3]).trim(), team: String(r[4]).trim(),
    status: String(r[5]), done: Number(r[6]) || 0, total: Number(r[7]) || 0, seconds: r[8] === '' ? null : Number(r[8]), penalty: Number(r[9]) || 0, wrong: Number(r[10]) || 0,
    wrongWords: String(r[11] || ''), updated: iso_(r[12])
  })).filter(l => l.race);
}
const raceHasClass_ = (race, cls) => race.classes.indexOf('*') >= 0 || race.classes.indexOf(cls) >= 0;
function openRaceFor_(raceId, who) {
  const race = racesCached_().find(x => x.id === String(raceId || ''));
  if (!race || !raceHasClass_(race, who.cls)) fail_('참여할 수 없는 레이스예요.');
  return race;
}
function racesCached_() { return cached_('c_races', () => readRaces_().map(r => ({id: r.id, name: r.name, classes: r.classes, sources: r.sources, settings: r.settings, status: r.status}))); }
function races_(req) {
  const who = teamOf_(req.token);
  const logs = readRaceLogs_().filter(l => l.cls === who.cls && l.team === who.team);
  const list = racesCached_().filter(r => raceHasClass_(r, who.cls) && r.status !== '준비').map(r => {
    const mine = logs.find(l => l.race === r.id);
    return {id: r.id, name: r.name, sources: r.sources, settings: r.settings, status: r.status,
      mine: mine ? {status: mine.status, seconds: mine.seconds, done: mine.done, total: mine.total, wrong: mine.wrong} : null};
  });
  return {ok: true, team: who, races: list, setsPack: setsFor_(req.v)};
}
function raceStart_(req) {
  const who = teamOf_(req.token);
  return withLock_(() => {
    const race = openRaceFor_(req.race, who);
    const mine = readRaceLogs_().find(l => l.race === race.id && l.cls === who.cls && l.team === who.team);
    if (mine && mine.status === '완료') fail_('이 레이스는 이미 끝냈어요. 다시 하려면 선생님께 말씀드리세요.', 'race_done');
    if (mine) {   // 새로고침·화면 꺼짐 뒤 이어 하기
      return {ok: true, resumed: true, elapsed: (Date.now() - mine.started) / 1000, done: mine.done, wrong: mine.wrong, penalty: mine.penalty};
    }
    if (race.status !== '진행') fail_(race.status === '마감' ? '마감된 레이스예요.' : '아직 열리지 않은 레이스예요.');
    const total = Math.max(0, Math.floor(Number(req.total) || 0));
    const now = new Date();
    sheet_(SHEET_RACELOG).appendRow([now, race.id, asText_(race.name), asText_(who.cls), asText_(who.team), '진행 중', 0, total, '', 0, 0, '', now]);
    return {ok: true, resumed: false, elapsed: 0, done: 0, wrong: 0, penalty: 0};
  });
}
function raceUpdate_(req, finish) {
  const who = teamOf_(req.token);
  return withLock_(() => {
    const mine = readRaceLogs_().find(l => l.race === String(req.race || '') && l.cls === who.cls && l.team === who.team);
    if (!mine) fail_('레이스 기록을 찾을 수 없어요. 처음부터 다시 시작해 주세요.');
    if (mine.status === '완료') return {ok: true, already: true};
    const done = Math.max(0, Math.floor(Number(req.done) || 0)), wrong = Math.max(0, Math.floor(Number(req.wrong) || 0));
    const penalty = Math.max(0, Number(req.penalty) || 0);
    const sh = sheet_(SHEET_RACELOG);
    sh.getRange(mine.row, 7).setValue(done);
    sh.getRange(mine.row, 10, 1, 2).setValues([[penalty, wrong]]);
    sh.getRange(mine.row, 13).setValue(new Date());
    if (finish) {
      // 시간은 서버가 잰다: 시작부터 지금까지 + 벌점
      const seconds = Math.round(((Date.now() - mine.started) / 1000 + penalty) * 10) / 10;
      const words = (Array.isArray(req.wrongWords) ? req.wrongWords : []).slice(0, 100).map(w => String(w).slice(0, 60)).join(' ; ');
      sh.getRange(mine.row, 6).setValue('완료');
      sh.getRange(mine.row, 9).setValue(seconds);
      sh.getRange(mine.row, 12).setValue(asText_(words));
      const same = readRaceLogs_().filter(l => l.race === mine.race && l.cls === who.cls && l.status === '완료' && l.seconds !== null);
      const rank = same.filter(l => l.seconds < seconds).length + 1;
      return {ok: true, seconds: seconds, rank: rank, of: same.length};
    }
    return {ok: true};
  });
}
function raceProgress_(req) { return raceUpdate_(req, false); }
function raceFinish_(req) { return raceUpdate_(req, true); }
function adminBoard_(req) {
  checkToken_(req.token, 'a');
  const race = readRaces_().find(x => x.id === String(req.race || ''));
  if (!race) fail_('그 레이스를 찾을 수 없어요.');
  const cls = String(req.cls || '');
  const rows = readRaceLogs_().filter(l => l.race === race.id && (!cls || l.cls === cls)).map(l => ({
    cls: l.cls, team: l.team, status: l.status, done: l.done, total: l.total, seconds: l.seconds, penalty: l.penalty, wrong: l.wrong,
    wrongWords: l.wrongWords ? l.wrongWords.split(' ; ') : [], elapsed: l.status === '진행 중' ? (Date.now() - l.started) / 1000 + l.penalty : null
  }));
  return {ok: true, race: {id: race.id, name: race.name, classes: race.classes, sources: race.sources, settings: race.settings, status: race.status}, rows: rows, teams: readTeams_()};
}
function adminRaceReset_(req) {
  checkToken_(req.token, 'a');
  return withLock_(() => {
    const mine = readRaceLogs_().find(l => l.race === String(req.race || '') && l.cls === String(req.cls || '') && l.team === String(req.team || ''));
    if (!mine) fail_('지울 기록이 없어요.');
    sheet_(SHEET_RACELOG).deleteRow(mine.row);
    return {ok: true};
  });
}


/* ---------------- 지난 레이스 기록 옮기기: N일 지난 줄을 '지난 레이스 기록' 탭으로 ---------------- */
const SHEET_LOG_OLD = '지난 기록';
const SHEET_RACELOG_OLD = '지난 레이스 기록';
function archiveInfo_() {
  const props = PropertiesService.getScriptProperties();
  const rows = name => { const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); return sh ? Math.max(0, sh.getLastRow() - 1) : 0; };
  return {days: Number(props.getProperty('ARCHIVE_DAYS') || 0), auto: props.getProperty('ARCHIVE_AUTO') === '1', last: props.getProperty('ARCHIVE_LAST') || '',
    raceRows: rows(SHEET_RACELOG), oldRows: rows(SHEET_RACELOG_OLD)};
}
function moveOld_(name, oldName, headers, cutoff, keepRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = sheet_(name), n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, 1, n - 1, headers.length).getValues();
  const when = v => v instanceof Date ? v.getTime() : (Date.parse(v) || 0);
  const old = [], keep = [];
  vals.forEach(r => (when(r[0]) && when(r[0]) < cutoff && !(keepRow && keepRow(r)) ? old : keep).push(r));
  if (!old.length) return 0;
  const fix = r => r.map(v => typeof v === 'string' ? asText_(v) : v);
  const osh = ss.getSheetByName(oldName) || ensureSheet_(ss, oldName, headers);
  osh.getRange(osh.getLastRow() + 1, 1, old.length, headers.length).setValues(old.map(fix));   // 먼저 복사하고
  sh.deleteRows(2, n - 1);                                                                       // 원래 탭을 비운 뒤
  if (keep.length) sh.getRange(2, 1, keep.length, headers.length).setValues(keep.map(fix));      // 남길 줄을 다시 씀
  return old.length;
}
function archiveNow_(days) {
  if (!(days > 0)) return {race: 0};
  const cutoff = Date.now() - days * 864e5;
  return withLock_(() => {
    const r = {race: moveOld_(SHEET_RACELOG, SHEET_RACELOG_OLD, RACELOG_HEADERS, cutoff, row => String(row[5]) === '진행 중')};
    PropertiesService.getScriptProperties().setProperty('ARCHIVE_LAST', new Date().toISOString());
    return r;
  });
}
/* 매일 새벽 자동으로 실행 (선생님 화면에서 켜면 트리거가 생김) */
function archiveDaily() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ARCHIVE_AUTO') !== '1') return;
  archiveNow_(Number(props.getProperty('ARCHIVE_DAYS') || 0));
}
function setArchiveTrigger_(on) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'archiveDaily').forEach(t => ScriptApp.deleteTrigger(t));
  if (on) ScriptApp.newTrigger('archiveDaily').timeBased().everyDays(1).atHour(3).create();
}
function adminArchive_(req) {
  checkToken_(req.token, 'a');
  const props = PropertiesService.getScriptProperties();
  const days = Math.max(0, Math.min(3650, Math.floor(Number(req.days) || 0)));
  if (req.save) {
    if (req.auto && !days) fail_('며칠 지난 기록을 옮길지 정해 주세요.');
    props.setProperty('ARCHIVE_DAYS', String(days));
    props.setProperty('ARCHIVE_AUTO', req.auto ? '1' : '0');
    try { setArchiveTrigger_(!!req.auto); }
    catch (e) { fail_("자동 실행을 켜려면 Apps Script에서 setup을 한 번 실행해 권한을 허용해 주세요. (" + String(e.message || e).slice(0, 120) + ')', 'trigger'); }
  }
  const moved = req.now ? archiveNow_(days) : null;
  return {ok: true, moved: moved, archive: archiveInfo_()};
}

/* ---------------- 시트 읽기·쓰기 ---------------- */
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = {[SHEET_STUDENTS]: STUDENT_HEADERS, [SHEET_SETS]: SET_HEADERS, [SHEET_TEAMS]: TEAM_HEADERS, [SHEET_RACES]: RACE_HEADERS, [SHEET_RACELOG]: RACELOG_HEADERS, [SHEET_LOG_OLD]: LOG_HEADERS}[name] || LOG_HEADERS;
  return ss.getSheetByName(name) || ensureSheet_(ss, name, headers);
}
function findStudentRow_(sh, sid) {
  const n = sh.getLastRow(); if (n < 2) return 0;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === sid) return i + 2;
  return 0;
}
function studentInfo_(sid) {
  const sh = sheet_(SHEET_STUDENTS), row = findStudentRow_(sh, sid);
  if (!row) fail_('학생 정보를 찾을 수 없어요. 다시 로그인해 주세요.', 'auth');
  return {sid: sid, name: String(sh.getRange(row, 2).getValue())};
}
function readStudents_() {
  const sh = sheet_(SHEET_STUDENTS), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 6).getValues().map(r => ({sid: String(r[0]).trim(), name: String(r[1]), joined: r[2], lastSeen: r[3], hasPw: !!r[5]})).filter(s => s.sid);
}
// 시트가 학번 앞자리 0을 지우거나 '='로 시작하는 글을 수식으로 바꾸지 않도록 글자로 고정
function asText_(v) { v = String(v); return /^[0-9=+\-@]/.test(v) ? "'" + v : v; }

/* ---------------- 보안 도구 ---------------- */
const newSalt_ = () => 'v2:' + Utilities.getUuid();
function hashPw_(pw, salt) {
  let h = String(salt) + ':' + pw;
  const rounds = /^v2:/.test(String(salt)) ? HASH_ROUNDS : HASH_ROUNDS_OLD;
  for (let i = 0; i < rounds; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + pw, Utilities.Charset.UTF_8));
  return h;
}
function sign_(s) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) fail_('서버 설치가 끝나지 않았어요. 선생님께 알려 주세요.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret)).replace(/=+$/, '');
}
function makeToken_(kind, id, hours) {
  const body = [kind, Utilities.base64EncodeWebSafe(id, Utilities.Charset.UTF_8).replace(/=+$/, ''), Date.now() + hours * 3600 * 1000].join('.');
  return body + '.' + sign_(body);
}
function checkToken_(token, kind) {
  const p = String(token || '').split('.');
  if (p.length !== 4 || p[0] !== kind) fail_('다시 로그인해 주세요.', 'auth');
  const body = p.slice(0, 3).join('.');
  if (sign_(body) !== p[3] || Number(p[2]) < Date.now()) fail_('로그인이 만료되었어요. 다시 로그인해 주세요.', 'auth');
  let id = p[1]; while (id.length % 4) id += '=';
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(id)).getDataAsString('UTF-8');
}
function cleanSid_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!/^[0-9A-Za-z-]{2,20}$/.test(s)) fail_('학번은 숫자(또는 영문)로 2~20자로 써 주세요.');
  return s;
}
function cleanName_(v) {
  const s = String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
  if (s.length < 1 || s.length > 20 || /^[=+\-@]/.test(s)) fail_('이름을 1~20자로 써 주세요.');
  return s;
}
function cleanPin_(v) {   // 새로 정하는 학생 비밀번호: 숫자 5자리 (통장·카드 4자리, 휴대폰 4·6자리와 겹치지 않게)
  const s = String(v == null ? '' : v).trim();
  if (!/^\d{5}$/.test(s)) fail_('비밀번호는 숫자 5자리로 정해 주세요.');
  return s;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (lock.hasLock()) return fn();
  // 오래 기다려도 차례가 안 오면 아무것도 바꾸지 않고 'busy'로 알려 줌 (학생 화면이 잠시 뒤 다시 보냄)
  if (!lock.tryLock(20000)) fail_('지금 쓰는 사람이 많아요. 잠시 후 다시 해 주세요.', 'busy');
  try { return fn(); } finally { lock.releaseLock(); }
}
function fail_(msg, code) { const e = new Error(msg); e.userMessage = msg; e.code = code; throw e; }
function iso_(d) { return d instanceof Date ? d.toISOString() : (d ? String(d) : ''); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
