// AOB9 · 9팀 팀 서버 (리그전 + 건의사항·오류 신고(자동 개선) + MD's PICK + OTP 요청 + 보험사 연락처 팀 공유) · 버전 2026-11-13 — 구글 Apps Script, 관리자(강남혁) 구글 계정에서 한 번 배포
// 저장하는 것: 리그(별명·주간 숫자), 건의사항(이름·제목·내용·상태·답변). 고객 정보는 받지 않습니다.
// 👥 팀원(2026-11-13): 팀원에게 보일 이름·마지막 접속 시각(숨기면 저장 안 함)·한마디·공개 열쇠 + 교환 게시판 글(지역·나이대 같은 요약만) + 팀원끼리 보내는 '잠긴 글'.
//   잠긴 글 = 보내는 사람 기기에서 받는 사람의 공개 열쇠로 잠근 것 → 이 서버와 관리자는 내용을 열 수 없음. 받으면 지우고, 안 받은 글은 14일 뒤 지움.
// OTP가 '사용 중'인 사람만 글을 올릴 수 있습니다 (access.json 확인). 건의사항 검토는 관리자 열쇠가 있어야 합니다.
const ACCESS = 'https://raw.githubusercontent.com/sicrics27-gif/sicrics27-gif.github.io/main/mh-analyzer/access.json'; // 2026-11-04: 코드 저장소 비공개 → 공개 배포 저장소의 입장 목록
function allowed(c) {
  if (!c || !/^[0-9a-f]{24}$/.test(c)) return false;
  const cache = CacheService.getScriptCache();
  let j = cache.get('acc');
  if (!j) { j = UrlFetchApp.fetch(ACCESS + '?t=' + Date.now(), { muteHttpExceptions: true }).getContentText(); cache.put('acc', j, 300); }
  try { const a = JSON.parse(j); const u = (a.users || {})[c]; return !!(u && u.s === 1); } catch (e) { return false; }
}
// 팀원 본인 확인 (보험사 연락처처럼 남의 휴대폰 번호가 든 자료용): OTP 사용 중 + 그 OTP에 등록된 구글 계정으로 로그인한 사람만
// 앱이 보낸 구글 접근 토큰을 구글(tokeninfo)에 물어 이메일을 확인하고, access.json 의 e(이메일 해시)와 맞는지 봄. 토큰·이메일은 저장하지 않음(10분 캐시만)
function accUsers() {
  const cache = CacheService.getScriptCache();
  let j = cache.get('acc');
  if (!j) { j = UrlFetchApp.fetch(ACCESS + '?t=' + Date.now(), { muteHttpExceptions: true }).getContentText(); cache.put('acc', j, 300); }
  try { return JSON.parse(j).users || {}; } catch (e) { return {}; }
}
function hex24(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 24); }
function member(c, g) {
  if (!allowed(c) || !g || String(g).length > 4000) return false;
  const cache = CacheService.getScriptCache(), h = 'tk_' + hex24('tk:' + g);
  let em = cache.get(h);
  if (!em) {
    const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(g), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return false;
    let t = {}; try { t = JSON.parse(r.getContentText()); } catch (e) { return false; }
    if (!t.email || String(t.email_verified) !== 'true') return false;
    em = String(t.email).toLowerCase(); cache.put(h, em, 600);
  }
  const u = accUsers()[c];
  return !!(u && u.s === 1 && u.e === hex24('mhs:' + em));
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function props() { return PropertiesService.getScriptProperties(); }
function isAdmin(k) { const a = props().getProperty('ADMIN_KEY'); return !!(a && k && k === a); }

// ---------- 팀 서버를 새 프로젝트로 다시 만들었을 때 (주소가 바뀜) ----------
// 예전 서버가 쓰던 'AOB 건의사항' 시트(건의 · MD픽 · 보험사 연락처 · 회사 일정 …)가 내 드라이브에 있으면 그 시트를 이어서 씀 → 자료가 비지 않음.
// 그 사이 새 시트에 들어온 줄은 예전 시트로 옮겨 담음. 프로젝트마다 한 번만 확인(스크립트 속성 ADOPT). 바꾸기 전 시트 번호는 SUG_SHEET_WAS 에 남김.
function adoptOld() {
  const p = props(); if (p.getProperty('ADOPT')) return;
  const lk = LockService.getScriptLock(); if (!lk.tryLock(10000)) return;
  try {
    if (p.getProperty('ADOPT')) return;
    const cur = p.getProperty('SUG_SHEET') || '', size = ss => ss.getSheets().reduce((n, sh) => n + sh.getLastRow(), 0);
    let best = null, bestN = -1, n = 0; const it = DriveApp.getFilesByName('AOB 건의사항');
    while (it.hasNext() && n++ < 15) {
      const f = it.next(); if (f.isTrashed()) continue;
      let ss = null; try { ss = SpreadsheetApp.openById(f.getId()); } catch (e) { ss = null; }
      if (!ss) continue; const k = size(ss); if (k > bestN) { best = ss; bestN = k; }
    }
    if (best && best.getId() !== cur) {
      if (cur) p.setProperty('SUG_SHEET_WAS', cur);
      p.setProperty('SUG_SHEET', best.getId());
      if (cur) { try { adoptRows(SpreadsheetApp.openById(cur), best); } catch (e) { p.setProperty('ADOPT_ERR', String(e).slice(0, 200)); } }
    }
    p.setProperty('ADOPT', String(Date.now()));
  } catch (e) { try { p.setProperty('ADOPT', 'err ' + String(e).slice(0, 150)); } catch (e2) { } }
  finally { lk.releaseLock(); }
}
function adoptRows(from, to) {
  from.getSheets().forEach(sh => {
    const v = sh.getDataRange().getValues(); if (v.length < 2) return;
    const name = sh.getName(); let t = to.getSheetByName(name);
    if (!t) { t = to.insertSheet(name); t.appendRow(v[0]); t.setFrozenRows(1); }
    const have = {}; t.getDataRange().getValues().forEach(r => { have[String(r[0])] = 1; });
    v.slice(1).forEach(r => { const id = String(r[0]); if (!id || have[id]) return; t.appendRow(r); have[id] = 1; });
  });
}
// ---------- 건의사항: 관리자 구글 드라이브의 시트 'AOB 건의사항'에 저장 ----------
const COLS = ['id', 'at', 'c', 'n', 'title', 'desc', 'st', 'reply', 'rat', 'likes', 'kind', 'img', 'diag', 'fix', 'pr', 'fsum', 'ft'];
// ---------- 오류·개선 신고 → GitHub 자동 개선 (2026-10-31) ----------
// 캡처 사진은 관리자 구글 드라이브 'AOB 오류 캡처' 폴더에만 저장(공개 안 함). 신고가 들어오면 GitHub 에 알림 → Claude 가 고친 안(PR)을 만듦
const REPO = 'sicrics27-gif/mh-analyzer';
function shotFolder() {
  const p = props(); let id = p.getProperty('SHOT_DIR'), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); } catch (e) { f = null; } }
  if (!f) { f = DriveApp.createFolder('AOB 오류 캡처'); p.setProperty('SHOT_DIR', f.getId()); }
  return f;
}
function ghDispatch(type, payload) {
  const t = props().getProperty('GH_TOKEN'); if (!t) return false;
  const r = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/dispatches', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + t, Accept: 'application/vnd.github+json' }, payload: JSON.stringify({ event_type: type, client_payload: payload }) });
  return r.getResponseCode() < 300;
}
function bugOne(id) { return rows(sheet()).find(x => String(x.id) === String(id)); }
function sheet() {
  const p = props(); let id = p.getProperty('SUG_SHEET'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('AOB 건의사항');
    p.setProperty('SUG_SHEET', ss.getId());
    const sh = ss.getSheets()[0]; sh.setName('건의'); sh.appendRow(COLS); sh.setFrozenRows(1);
  }
  return ss.getSheetByName('건의') || ss.getSheets()[0];
}
function rows(sh) {
  const v = sh.getDataRange().getValues(); v.shift();
  return v.map((r, i) => { const o = { row: i + 2 }; COLS.forEach((k, j) => { o[k] = r[j]; }); return o; }).filter(o => o.id);
}
function sugList(me, admin) {
  const sh = sheet();
  return rows(sh).map(o => {
    const likes = String(o.likes || '').split(',').filter(Boolean);
    return { id: String(o.id), at: +o.at, n: String(o.n), title: String(o.title), desc: String(o.desc), st: String(o.st || 'new'), reply: String(o.reply || ''), rat: +o.rat || 0,
      likes: likes.length, liked: !!(me && likes.indexOf(me) >= 0), mine: !!(me && o.c === me),
      kind: String(o.kind || 'idea'), shot: !!o.img, fix: String(o.fix || ''), pr: String(o.pr || ''), fsum: String(o.fsum || '') };
  }).sort((a, b) => b.at - a.at);
}

function doPost(e) { try { return doPost_(e); } catch (err) { return out({ ok: false, error: '팀 서버 오류: ' + err }); } }
function doPost_(e) {
  const b = JSON.parse(e.postData.contents || '{}');
  adoptOld();
  if (b.a === 'mt_sync') return mtSync(b);   // 접속 표시는 자주 오므로 줄 세우지 않음 (자기 칸만 고침)
  const lock = LockService.getScriptLock(); lock.waitLock(8000);
  try {
    if (/^mt_(send|ack|recall|post|unpost)$/.test(b.a || '')) return mtBox(b);
    // 관리자 열쇠 받기 (처음 한 번만 · 그다음부터는 아무도 못 받음)
    if (b.a === 'claim') {
      if (props().getProperty('ADMIN_KEY')) return out({ ok: false, error: '이미 관리자가 정해져 있습니다' });
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      const k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      props().setProperty('ADMIN_KEY', k); sheet();
      return out({ ok: true, key: k });
    }
    if (b.a === 'sug_add') {
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      const title = String(b.title || '').trim().slice(0, 60), desc = String(b.desc || '').trim().slice(0, 1500);
      if (!title) return out({ ok: false, error: '제목을 적어 주세요' });
      const kind = ['bug', 'imp', 'idea'].indexOf(b.kind) >= 0 ? b.kind : 'idea';
      const sh = sheet(), mine = rows(sh).filter(o => o.c === b.c && Date.now() - +o.at < 3600e3).length;
      if (mine >= 10) return out({ ok: false, error: '한 시간에 10개까지 올릴 수 있습니다' });
      const id = Utilities.getUuid().slice(0, 8);
      let img = '';
      // 2026-11-04: 캡처 사진은 받지 않음 (고객 정보가 관리자 드라이브로 오지 않게)
      if (false && b.img && String(b.img).length < 3000000) { try { const f = shotFolder().createFile(Utilities.newBlob(Utilities.base64Decode(String(b.img)), 'image/jpeg', id + '.jpg')); img = f.getId(); } catch (e) { img = ''; } }
      const ft = Utilities.getUuid().replace(/-/g, '');
      const auto = kind !== 'idea' && !!props().getProperty('GH_TOKEN');
      sh.appendRow([id, Date.now(), b.c, String(b.n || '팀원').slice(0, 10), title, desc, 'new', '', '', '', kind, img, String(b.diag || '').slice(0, 6000), auto ? 'queued' : '', '', '', ft]);
      if (auto) { const ok = ghDispatch('aob-bug', { id: id, t: ft, u: ScriptApp.getService().getUrl() }); if (!ok) { const o = bugOne(id); sheet().getRange(o.row, COLS.indexOf('fix') + 1).setValue(''); } }
      return out({ ok: true, id: id, auto: auto });
    }
    if (b.a === 'sug_like') {
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      const sh = sheet(), o = rows(sh).find(x => String(x.id) === String(b.id)); if (!o) return out({ ok: false, error: '없는 글' });
      const likes = String(o.likes || '').split(',').filter(Boolean), i = likes.indexOf(b.c);
      if (i >= 0) likes.splice(i, 1); else likes.push(b.c);
      sh.getRange(o.row, COLS.indexOf('likes') + 1).setValue(likes.join(','));
      return out({ ok: true });
    }
    if (b.a === 'sug_rev') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만 검토할 수 있습니다' });
      const sh = sheet(), o = rows(sh).find(x => String(x.id) === String(b.id)); if (!o) return out({ ok: false, error: '없는 글' });
      const st = ['new', 'review', 'yes', 'done', 'no'].indexOf(b.st) >= 0 ? b.st : o.st;
      sh.getRange(o.row, COLS.indexOf('st') + 1, 1, 3).setValues([[st, String(b.reply || '').slice(0, 1000), Date.now()]]);
      return out({ ok: true });
    }
    if (b.a === 'sug_del') {
      const sh = sheet(), o = rows(sh).find(x => String(x.id) === String(b.id)); if (!o) return out({ ok: false, error: '없는 글' });
      if (!(isAdmin(b.k) || (allowed(b.c) && o.c === b.c && (o.st || 'new') === 'new'))) return out({ ok: false, error: '지울 수 없습니다' });
      sh.deleteRow(o.row);
      return out({ ok: true });
    }
    // 관리자: GitHub 열쇠 맡기기 (오류 신고가 오면 자동 개선을 부르는 데만 씀)
    if (b.a === 'gh_set') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const g = String(b.gh || '').trim(); if (!/^(gh[pousr]_|github_pat_)[\w]+$/.test(g)) return out({ ok: false, error: '열쇠 형식이 아닙니다' });
      props().setProperty('GH_TOKEN', g); return out({ ok: true });
    }
    // 자동 개선 결과 (GitHub 워크플로가 보냄 · 신고마다 한 번 쓰는 열쇠 ft 로 확인)
    if (b.a === 'bug_done') {
      const o = bugOne(b.id); if (!o || !o.ft || String(o.ft) !== String(b.t)) return out({ ok: false, error: '열쇠가 맞지 않습니다' });
      const st = ['ready', 'nochange', 'review', 'done', 'fail', 'working'].indexOf(b.st) >= 0 ? b.st : 'fail';
      const sh = sheet(); sh.getRange(o.row, COLS.indexOf('fix') + 1, 1, 3).setValues([[st, String(b.pr || o.pr || ''), String(b.sum || o.fsum || '').slice(0, 1500)]]);
      if (st === 'done') sh.getRange(o.row, COLS.indexOf('st') + 1, 1, 3).setValues([['done', String(o.reply || '') || '자동 개선으로 반영했습니다', Date.now()]]);
      return out({ ok: true });
    }
    // 관리자: 고친 안 반영 → GitHub 에 합치기·배포 요청
    if (b.a === 'bug_ok') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const o = bugOne(b.id); if (!o || String(o.fix) !== 'ready' || !o.pr) return out({ ok: false, error: '반영할 고친 안이 없습니다' });
      if (!ghDispatch('aob-merge', { id: String(o.id), t: String(o.ft), u: ScriptApp.getService().getUrl(), pr: String(o.pr) })) return out({ ok: false, error: 'GitHub 열쇠가 없거나 만료됐습니다' });
      sheet().getRange(o.row, COLS.indexOf('fix') + 1).setValue('merging'); return out({ ok: true });
    }
    // 관리자: 다시 고쳐 보기
    if (b.a === 'bug_retry') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const o = bugOne(b.id); if (!o) return out({ ok: false, error: '없는 글' });
      const ft = Utilities.getUuid().replace(/-/g, ''), sh = sheet();
      sh.getRange(o.row, COLS.indexOf('fix') + 1, 1, 4).setValues([['queued', '', '', ft]]);
      if (!ghDispatch('aob-bug', { id: String(o.id), t: ft, u: ScriptApp.getService().getUrl(), more: String(b.more || '').slice(0, 800) })) { sh.getRange(o.row, COLS.indexOf('fix') + 1).setValue(''); return out({ ok: false, error: 'GitHub 열쇠가 없거나 만료됐습니다' }); }
      return out({ ok: true });
    }
    // ---------- OTP 요청 (아직 OTP가 없는 동료 → 관리자) ----------
    if (b.a === 'otp_req') {
      const n = String(b.n || '').trim().slice(0, 12), m = String(b.m || '').trim().toLowerCase().slice(0, 80);
      if (!n || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m)) return out({ ok: false, error: '이름과 구글 계정(이메일)을 정확히 넣어 주세요' });
      const sh = otpSheet(), all = otpRows(sh), day = Date.now() - 864e5;
      if (all.filter(o => +o.at > day).length >= 40) return out({ ok: false, error: '오늘 요청이 너무 많습니다. 관리자에게 직접 연락하세요' });
      const same = all.find(o => o.m === m && o.st === 'new'); if (same) return out({ ok: true, rid: String(same.rid), again: true });
      const rid = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
      sh.appendRow([rid, Date.now(), n, m, String(b.dev || '').slice(0, 20), String(b.note || '').slice(0, 200), 'new', '', '']);
      return out({ ok: true, rid: rid });
    }
    if (b.a === 'otp_done' || b.a === 'otp_no') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const sh = otpSheet(), o = otpRows(sh).find(x => String(x.rid) === String(b.rid)); if (!o) return out({ ok: false, error: '없는 요청' });
      sh.getRange(o.row, OTC.indexOf('st') + 1, 1, 3).setValues([[b.a === 'otp_done' ? 'done' : 'no', b.a === 'otp_done' ? String(b.code || '').slice(0, 30) : '', Date.now()]]);
      return out({ ok: true });
    }
    // ---------- MD's PICK (추천 상품 노트) ----------
    if (b.a === 'md_add' || b.a === 'md_edit') {
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      const sh = mdSheet(), o = { cat: String(b.cat || '기타').slice(0, 20), ins: String(b.ins || '').trim().slice(0, 20), prod: String(b.prod || '').trim().slice(0, 40), why: String(b.why || '').slice(0, 500), tip: String(b.tip || '').slice(0, 300) };
      if (!o.ins || !o.prod) return out({ ok: false, error: '보험사와 상품을 적어 주세요' });
      if (b.a === 'md_add') { const id = Utilities.getUuid().slice(0, 8); sh.appendRow([id, Date.now(), b.c, String(b.n || '팀원').slice(0, 10), o.cat, o.ins, o.prod, o.why, o.tip, '', '']); return out({ ok: true, id: id }); }
      const r = mdRows(sh).find(x => String(x.id) === String(b.id)); if (!r) return out({ ok: false, error: '없는 글' });
      if (!(r.c === b.c || isAdmin(b.k))) return out({ ok: false, error: '본인 글만 고칠 수 있습니다' });
      sh.getRange(r.row, MDC.indexOf('cat') + 1, 1, 5).setValues([[o.cat, o.ins, o.prod, o.why, o.tip]]); return out({ ok: true });
    }
    if (b.a === 'md_del') {
      const sh = mdSheet(), r = mdRows(sh).find(x => String(x.id) === String(b.id)); if (!r) return out({ ok: false, error: '없는 글' });
      if (!(isAdmin(b.k) || (allowed(b.c) && r.c === b.c))) return out({ ok: false, error: '지울 수 없습니다' });
      sh.deleteRow(r.row); return out({ ok: true });
    }
    if (b.a === 'md_like') {
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      const sh = mdSheet(), r = mdRows(sh).find(x => String(x.id) === String(b.id)); if (!r) return out({ ok: false, error: '없는 글' });
      const likes = String(r.likes || '').split(',').filter(Boolean), i = likes.indexOf(b.c); if (i >= 0) likes.splice(i, 1); else likes.push(b.c);
      sh.getRange(r.row, MDC.indexOf('likes') + 1).setValue(likes.join(',')); return out({ ok: true });
    }
    if (b.a === 'md_pin') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const sh = mdSheet(), r = mdRows(sh).find(x => String(x.id) === String(b.id)); if (!r) return out({ ok: false, error: '없는 글' });
      sh.getRange(r.row, MDC.indexOf('pin') + 1).setValue(r.pin ? '' : '1'); return out({ ok: true });
    }
    // ---------- 📇 보험사 연락처 팀 공유 (2026-11-04) ----------
    // 한 사람이 카톡 연락처 글을 붙여 넣으면 보험사별로 합쳐 저장 → 팀원 앱이 열릴 때 자동으로 받음. 관리자 구글 시트 'AOB 건의사항'의 '보험사연락처' 탭에만 저장(공개 안 함)
    if (b.a === 'cd_put' || b.a === 'cd_get') {
      if (!member(b.c, b.g)) return out({ ok: false, error: 'member', msg: 'OTP와 같은 구글 계정으로 로그인한 팀원만 받을 수 있습니다' });
      const sh = cdSheet();
      if (b.a === 'cd_put') {
        const cos = Array.isArray(b.cos) ? b.cos.slice(0, 80) : []; let n = 0;
        const all = cdRows(sh), by = String(b.n || '팀원').slice(0, 10);
        cos.forEach(co => {
          if (!co || !co.key) return; const key = String(co.key).slice(0, 30), inc = cdClean(co, b.date, b.at);
          const r = all.find(x => String(x.key) === key);
          if (!r) { sh.appendRow([key, inc.at, inc.date, by, JSON.stringify(inc)]); n++; return; }
          let old = {}; try { old = JSON.parse(r.json); } catch (e) { old = {}; }
          const m = cdMerge(old, inc); if (!m) return;
          sh.getRange(r.row, 1, 1, 5).setValues([[key, m.at, m.date, by, JSON.stringify(m)]]); n++;
        });
        return out({ ok: true, n: n });
      }
      const rows = cdRows(sh), cos = []; let at = 0;
      rows.forEach(r => { try { const o = JSON.parse(r.json); cos.push(o); at = Math.max(at, +r.at || 0); } catch (e) { } });
      return out({ ok: true, at: at, cos: cos });
    }
    if (b.a === 'sched_put' || b.a === 'sched_get') {
      if (!member(b.c, b.g)) return out({ ok: false, error: 'member', msg: 'OTP와 같은 구글 계정으로 로그인한 팀원만 쓸 수 있습니다' });
      const sh = scSheet(), rows = scRows(sh);
      if (b.a === 'sched_put') {
        const ym = String(b.ym || ''); if (!/^\d{4}-\d{2}$/.test(ym)) return out({ ok: false, error: 'ym' });
        const r = rows.find(x => String(x.ym) === ym); let old = []; if (r) { try { old = JSON.parse(r.json) || []; } catch (e) { old = []; } }
        const custLike = t => { t = String(t || ''); if (/1차|2차|3차|상담|클로징|db|인포|계약|만남|증권|청구/i.test(t) && !/교육|초청|석회|세미나|강의/.test(t)) return true; return /^[가-힣]{2,3}\s+[가-힣]{2,4}\s+\S/.test(t) && !/교육|초청|석회|미팅|워킹|위클리|마감|회의|공유|세미나|시상|강의/.test(t); }; // 고객 약속 같은 줄은 받지 않음 (2026-11-05)
        old = old.filter(x => !custLike(x && x.x));
        const m = scMerge(old, (Array.isArray(b.items) ? b.items.slice(0, 300) : []).filter(x => !custLike(x && x.x)));
        if (!m.add) return out({ ok: true, add: 0 });
        const vals = [ym, Date.now(), String(b.n || '팀원').slice(0, 10), JSON.stringify(m.items)];
        if (r) sh.getRange(r.row, 1, 1, 4).setValues([vals]); else sh.appendRow(vals);
        return out({ ok: true, add: m.add });
      }
      const sched = {}; rows.sort((a, b2) => String(a.ym) < String(b2.ym) ? -1 : 1).slice(-3).forEach(r => { try { sched[String(r.ym)] = { at: +r.at, items: JSON.parse(r.json) }; } catch (e) { } });
      return out({ ok: true, sched: sched });
    }
    // ---------- 🔎 상품 정보 (2026-11-07) ----------
    // 팀원 누구나 원수사 안내 문자를 올림(pism_put) → 관리자 앱이 꺼내 NHC 장부로(pism_take) → 대표님이 확정한 값만 팀 공개(pi_put / pi_get)
    if (b.a === 'pism_put') {
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      let t = String(b.text || '').slice(0, 3000).replace(/01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}/g, '(전화번호)').replace(/\d{6}[-\s]?[1-4]\d{6}/g, '(주민번호)');
      if (t.trim().length < 10) return out({ ok: false, error: '글이 너무 짧습니다' });
      const sh = piSheet('상품문자', ['id', 'at', 'n', 'from', 'text']);
      sh.appendRow([Utilities.getUuid().slice(0, 8), Date.now(), String(b.n || '팀원').slice(0, 10), String(b.from || '').slice(0, 40), t]);
      return out({ ok: true });
    }
    if (b.a === 'pism_take') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const sh = piSheet('상품문자', ['id', 'at', 'n', 'from', 'text']), v = sh.getDataRange().getValues(); v.shift();
      const rows = v.filter(r => r[0]).map(r => ({ id: r[0], at: r[1], n: r[2], from: r[3], text: r[4] }));
      if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      return out({ ok: true, rows: rows });
    }
    if (b.a === 'pi_put') {
      if (!isAdmin(b.k)) return out({ ok: false, error: '관리자만' });
      const items = Array.isArray(b.items) ? b.items.slice(0, 2000) : [];
      const sh = piSheet('상품정보', ['at', 'json']);
      if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      const at = Date.now(), chunk = [];
      for (let i = 0; i < items.length; i += 40) chunk.push([at, JSON.stringify(items.slice(i, i + 40))]);
      if (chunk.length) sh.getRange(2, 1, chunk.length, 2).setValues(chunk);
      return out({ ok: true, n: items.length });
    }
    if (b.a === 'pi_get') {
      if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
      const sh = piSheet('상품정보', ['at', 'json']), v = sh.getDataRange().getValues(); v.shift();
      let items = [], at = 0; v.forEach(r => { try { items = items.concat(JSON.parse(r[1])); at = Math.max(at, +r[0] || 0); } catch (e) { } });
      return out({ ok: true, at: at, items: items });
    }
    // ---------- 리그전 (주간 점수) ----------
    if (!allowed(b.c)) return out({ ok: false, error: 'OTP 사용자가 아닙니다' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(b.week || '')) return out({ ok: false, error: 'week' });
    const p = props(), key = 'w_' + b.week, t = JSON.parse(p.getProperty(key) || '{}');
    const n = (x, m) => Math.max(0, Math.min(+x || 0, m));
    t[b.c] = { n: String(b.n || '플래너').slice(0, 10), xp: n(b.xp, 20000), call: n(b.call, 2000), meet: n(b.meet, 500), msg: n(b.msg, 3000), lv: n(b.lv, 200), tot: n(b.tot, 10000000), st: n(b.st, 3650), ti: String(b.ti || '').slice(0, 16), at: Date.now() };
    p.setProperty(key, JSON.stringify(t));
    const all = p.getKeys().filter(k => k.indexOf('w_') === 0).sort(); while (all.length > 12) p.deleteProperty(all.shift());
    return out({ ok: true });
  } finally { lock.releaseLock(); }
}
function piSheet(name, cols) {
  const ss = SpreadsheetApp.openById(props().getProperty('SUG_SHEET') || sheet().getParent().getId());
  let sh = ss.getSheetByName(name); if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}
// ---------- 👥 팀원: 접속 상태 · 잠긴 글(고객 보내기·교환) · 교환 게시판 (2026-11-13) ----------
// 팀원 번호 = OTP 코드를 한 번 더 섞은 값(다른 팀원에게 OTP 코드가 보이지 않게). 접속 표시는 스크립트 속성 'mt_<번호>' 한 칸씩, 잠긴 글·게시판은 시트.
const MTB = ['id', 'at', 'to', 'from', 'fn', 'kind', 'meta', 'ct', 'ref'], MTP = ['id', 'at', 'by', 'n', 'meta'];
const MT_NO = { ok: false, error: 'member', msg: 'OTP와 같은 구글 계정으로 로그인한 팀원만 쓸 수 있습니다' };
function mtActive() {   // 지금 OTP가 살아 있는 팀원 번호 목록 (10분 캐시)
  const cache = CacheService.getScriptCache(); let s = cache.get('mtact');
  if (!s) { const u = accUsers(); s = Object.keys(u).filter(c => u[c] && u[c].s === 1).map(c => hex24('mt:' + c)).join(','); if (s) cache.put('mtact', s, 600); }
  return s || '';
}
function mtScrub(t) { return String(t == null ? '' : t).replace(/01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}/g, '(전화번호)').replace(/\d{6}[-\s]?[1-4]\d{6}/g, '(주민번호)').replace(/[\r\n\t]+/g, ' '); }
// 시트에 글자로만 넣기: 앞에 ' 를 붙여 숫자·날짜·수식으로 바뀌지 않게 (=, +, - 로 시작하는 글이 수식이 되는 것도 막음). 읽을 때는 ' 를 뗌
function mtT(s) { return "'" + String(s == null ? '' : s); }
function mtU(s) { return String(s == null ? '' : s).replace(/^'/, ''); }
function mtRows(sh, cols) { const v = sh.getDataRange().getValues(); v.shift(); return v.map((r, i) => { const o = { row: i + 2 }; cols.forEach((k, j) => { o[k] = k === 'at' ? (+r[j] || 0) : mtU(r[j]); }); return o; }).filter(o => o.id); }
function mtDel(sh, list) { list.slice().sort((a, b) => b.row - a.row).forEach(o => sh.deleteRow(o.row)); }
function mtBump() { const p = props(), cur = +p.getProperty('mtv') || 0; p.setProperty('mtv', String(Math.max(Date.now(), cur + 1))); }   // 받은함·게시판이 바뀔 때마다 올라가는 번호 (앱이 같은 번호면 다시 안 읽음)
function mtId(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 13); }
// 내 접속 시각을 남기고 팀원 목록(+ 바뀐 게 있으면 내 받은함·게시판)을 받음
function mtSync(b) {
  if (!member(b.c, b.g)) return out(MT_NO);
  const me = hex24('mt:' + b.c), now = Date.now(), P = props().getProperties();
  let old = {}; try { old = JSON.parse(P['mt_' + me] || '{}') || {}; } catch (e) { old = {}; }
  const rec = { n: mtScrub(b.n || old.n || '팀원').slice(0, 12), at: b.hide ? 0 : now, pub: String(b.pub || old.pub || '').slice(0, 200), hide: b.hide ? 1 : 0, say: b.say === undefined || b.say === null ? String(old.say || '') : mtScrub(b.say).slice(0, 40) };
  props().setProperty('mt_' + me, JSON.stringify(rec));
  const act = mtActive(), rows = [];
  Object.keys(P).forEach(k => {
    if (k.indexOf('mt_') !== 0 || k.length !== 27) return; const id = k.slice(3); if (id === me || act.indexOf(id) < 0) return;
    try { const o = JSON.parse(P[k]); rows.push({ id: id, n: String(o.n || '팀원'), at: o.hide ? 0 : (+o.at || 0), hide: o.hide ? 1 : 0, pub: String(o.pub || ''), say: String(o.say || '') }); } catch (e) { }
  });
  rows.push({ id: me, me: 1, n: rec.n, at: now, hide: rec.hide, pub: rec.pub, say: rec.say });
  const v = String(P.mtv || '0');
  if (String(b.v || '') === v) return out({ ok: true, mt: 2, me: me, now: now, v: v, same: 1, rows: rows });
  const all = mtRows(piSheet('팀원함', MTB), MTB), lim = now - 14 * 864e5;
  const box = all.filter(o => o.to === me && o.at >= lim).map(o => ({ id: o.id, at: o.at, from: o.from, fn: o.fn, kind: o.kind, meta: o.meta, ct: o.ct, ref: o.ref }));
  const sent = all.filter(o => o.from === me && o.at >= lim).map(o => o.id);
  const posts = mtRows(piSheet('교환글', MTP), MTP).filter(o => o.at >= now - 30 * 864e5 && (o.by === me || act.indexOf(o.by) >= 0)).map(o => ({ id: o.id, at: o.at, by: o.by, n: o.n, meta: o.meta, mine: o.by === me ? 1 : 0 })).sort((x, y) => y.at - x.at).slice(0, 200);
  return out({ ok: true, mt: 2, me: me, now: now, v: v, rows: rows, box: box, sent: sent, posts: posts });
}
// 잠긴 글 보내기(mt_send) · 받은 글 지우기(mt_ack) · 보낸 글 되찾기(mt_recall) · 교환 게시판 올리기/내리기(mt_post / mt_unpost) — 한 번에 하나씩(잠금 안에서)
function mtBox(b) {
  if (!member(b.c, b.g)) return out(MT_NO);
  const me = hex24('mt:' + b.c), now = Date.now();
  if (b.a === 'mt_post' || b.a === 'mt_unpost') {
    const sh = piSheet('교환글', MTP), all = mtRows(sh, MTP);
    if (b.a === 'mt_unpost') { const hit = all.filter(o => o.id === String(b.id) && o.by === me); mtDel(sh, hit); if (hit.length) mtBump(); return out({ ok: true, n: hit.length }); }
    const meta = mtScrub(b.meta).trim().slice(0, 160); if (meta.length < 2) return out({ ok: false, error: '내용이 없습니다' });
    const old = all.filter(o => o.at < now - 30 * 864e5);
    if (all.filter(o => o.by === me && o.at >= now - 30 * 864e5).length >= 10) return out({ ok: false, error: '내놓은 글은 10개까지입니다. 먼저 내린 뒤 올려 주세요' });
    if (old.length) mtDel(sh, old);
    const id = mtId('p'); sh.appendRow([mtT(id), now, mtT(me), mtT(mtScrub(b.n || '팀원').slice(0, 12)), mtT(meta)]);
    mtBump(); return out({ ok: true, id: id });
  }
  const sh = piSheet('팀원함', MTB), all = mtRows(sh, MTB);
  if (b.a === 'mt_ack') { const ids = (Array.isArray(b.ids) ? b.ids : []).slice(0, 100).map(String), hit = all.filter(o => o.to === me && ids.indexOf(o.id) >= 0); mtDel(sh, hit); if (hit.length) mtBump(); return out({ ok: true, n: hit.length }); }
  if (b.a === 'mt_recall') { const hit = all.filter(o => o.from === me && o.id === String(b.id)); mtDel(sh, hit); if (hit.length) mtBump(); return out({ ok: true, n: hit.length }); }
  const to = String(b.to || ''), kind = String(b.kind || ''), ct = String(b.ct || '');
  if (!/^[0-9a-f]{24}$/.test(to) || to === me || !/^[a-z]{2,12}$/.test(kind)) return out({ ok: false, error: '받는 사람을 다시 골라 주세요' });
  if (mtActive().indexOf(to) < 0 || !props().getProperty('mt_' + to)) return out({ ok: false, error: '받는 팀원이 아직 팀원 화면을 시작하지 않았습니다' });
  if (ct.length > 30000 || !/^[A-Za-z0-9_.\-]*$/.test(ct)) return out({ ok: false, error: '내용을 보낼 수 없습니다' });
  if (all.filter(o => o.from === me && o.at > now - 864e5).length >= 80) return out({ ok: false, error: '오늘 보낸 글이 너무 많습니다' });
  const old = all.filter(o => o.at < now - 14 * 864e5); if (old.length) mtDel(sh, old);
  const id = mtId('m');
  sh.appendRow([mtT(id), now, mtT(to), mtT(me), mtT(mtScrub(b.n || '팀원').slice(0, 12)), mtT(kind), mtT(mtScrub(b.meta).slice(0, 200)), mtT(ct), mtT(String(b.ref || '').replace(/[^a-z0-9]/g, '').slice(0, 14))]);
  mtBump(); return out({ ok: true, id: id });
}
// OTP 요청 시트
const OTC = ['rid', 'at', 'n', 'm', 'dev', 'note', 'st', 'code', 'dt'];
function otpSheet() {
  const ss = SpreadsheetApp.openById(props().getProperty('SUG_SHEET') || sheet().getParent().getId());
  let sh = ss.getSheetByName('OTP요청'); if (!sh) { sh = ss.insertSheet('OTP요청'); sh.appendRow(OTC); sh.setFrozenRows(1); }
  return sh;
}
function otpRows(sh) { const v = sh.getDataRange().getValues(); v.shift(); return v.map((r, i) => { const o = { row: i + 2 }; OTC.forEach((k, j) => { o[k] = r[j]; }); return o; }).filter(o => o.rid); }
// MD's PICK 시트
const MDC = ['id', 'at', 'c', 'n', 'cat', 'ins', 'prod', 'why', 'tip', 'likes', 'pin'];
function mdSheet() {
  const ss = SpreadsheetApp.openById(props().getProperty('SUG_SHEET') || sheet().getParent().getId());
  let sh = ss.getSheetByName('MD픽'); if (!sh) { sh = ss.insertSheet('MD픽'); sh.appendRow(MDC); sh.setFrozenRows(1); }
  return sh;
}
function mdRows(sh) { const v = sh.getDataRange().getValues(); v.shift(); return v.map((r, i) => { const o = { row: i + 2 }; MDC.forEach((k, j) => { o[k] = r[j]; }); return o; }).filter(o => o.id); }
// 보험사 연락처 시트 (보험사 한 곳 = 한 줄)
const CDC = ['key', 'at', 'date', 'by', 'json'];
function cdSheet() {
  const ss = SpreadsheetApp.openById(props().getProperty('SUG_SHEET') || sheet().getParent().getId());
  let sh = ss.getSheetByName('보험사연락처'); if (!sh) { sh = ss.insertSheet('보험사연락처'); sh.appendRow(CDC); sh.setFrozenRows(1); }
  return sh;
}
function cdRows(sh) { const v = sh.getDataRange().getValues(); v.shift(); return v.map((r, i) => { const o = { row: i + 2 }; CDC.forEach((k, j) => { o[k] = r[j]; }); return o; }).filter(o => o.key); }
function cdClean(co, date, at) {
  const s = (x, n) => String(x == null ? '' : x).slice(0, n);
  const ppl = (Array.isArray(co.people) ? co.people : []).slice(0, 60).map(p => ({ nm: s(p.nm, 20), role: s(p.role, 12), ph: s(p.ph, 20), note: s(p.note, 60) }));
  const svc = (Array.isArray(co.svc) ? co.svc : []).slice(0, 40).map(x => ({ l: s(x.l, 40), v: s(x.v, 200), phs: (x.phs || []).slice(0, 6).map(p => s(p, 20)), mails: (x.mails || []).slice(0, 4).map(m => s(m, 60)) }));
  const links = (Array.isArray(co.links) ? co.links : []).slice(0, 20).map(l => ({ l: s(l.l, 40), u: s(l.u, 300) })).filter(l => /^https?:\/\//.test(l.u));
  const notes = (Array.isArray(co.notes) ? co.notes : []).slice(0, 30).map(x => s(x, 200));
  return { n: s(co.n, 30), key: s(co.key, 30), k: co.k === '생명' ? '생명' : '손해', people: ppl, svc: svc, links: links, notes: notes, date: s(co.date || date, 10), at: +(co.at || at) || Date.now() };
}
// 합치기 (2026-11-04 대표님 결정): 팀원 누가 올려도 '추가만' — 기존보다 줄어든 내용은 반영 안 함, 오타 정도 차이는 덮어쓰지 않음.
//  같은 사람(이름+직함 또는 같은 번호)은 더 새 기준일일 때만 번호·메모를 바꾸고, 새 사람·새 창구·새 링크만 더함
function near(a, b) { a = String(a || '').replace(/\s/g, ''); b = String(b || '').replace(/\s/g, ''); if (a === b) return true; if (!a || !b || Math.abs(a.length - b.length) > 1) return false; let d = 0; for (let i = 0, j = 0; i < a.length || j < b.length;) { if (a[i] === b[j]) { i++; j++; continue; } d++; if (d > 1) return false; if (a.length > b.length) i++; else if (b.length > a.length) j++; else { i++; j++; } } return true; }
function cdMerge(old, inc) {
  const newer = String(inc.date || '') >= String(old.date || ''), ph = x => String(x || '').replace(/\D/g, '');
  const ppl = (old.people || []).map(p => Object.assign({}, p)); let ch = false;
  inc.people.forEach(p => {
    const o = ppl.find(q => (q.role === p.role && near(q.nm, p.nm)) || (ph(p.ph) && ph(q.ph) === ph(p.ph)));
    if (!o) { ppl.push(p); ch = true; return; }
    if (newer && ph(p.ph) && ph(p.ph) !== ph(o.ph)) { o.ph = p.ph; ch = true; }
    if (newer && p.note && String(p.note).length > String(o.note || '').length) { o.note = p.note; ch = true; }
  });
  const sk = x => String(x.l || '').replace(/\s/g, ''), svc = (old.svc || []).slice();
  inc.svc.forEach(x => { if (!svc.some(y => near(sk(y), sk(x)) && (y.phs || []).join(',') === (x.phs || []).join(','))) { svc.push(x); ch = true; } });
  const links = (old.links || []).slice(); inc.links.forEach(l => { if (!links.some(y => y.u === l.u)) { links.push(l); ch = true; } });
  const notes = (old.notes || []).slice(); (inc.notes || []).forEach(x => { if (!notes.some(y => near(y, x))) { notes.push(x); ch = true; } });
  if (!ch) return null;
  return Object.assign({}, old, { n: old.n || inc.n, people: ppl.slice(0, 80), svc: svc.slice(0, 60), links: links.slice(0, 30), notes: notes.slice(0, 40), date: newer ? inc.date : old.date, at: Math.max(+old.at || 0, +inc.at || 0) });
}
// ---------- 👥 회사 일정 팀 공유 (2026-11-04): 시트 '회사일정' (달 하나 = 한 줄). 팀원 누구나 올림, 추가만 ----------
const SCC = ['ym', 'at', 'by', 'json'];
function scSheet() {
  const ss = SpreadsheetApp.openById(props().getProperty('SUG_SHEET') || sheet().getParent().getId());
  let sh = ss.getSheetByName('회사일정'); if (!sh) { sh = ss.insertSheet('회사일정'); sh.appendRow(SCC); sh.setFrozenRows(1); }
  return sh;
}
function scRows(sh) { const v = sh.getDataRange().getValues(); v.shift(); return v.map((r, i) => { const o = { row: i + 2 }; SCC.forEach((k, j) => { o[k] = r[j]; }); return o; }).filter(o => o.ym); }
function scMerge(old, inc) {
  const out = old.slice(); let add = 0;
  inc.forEach(x => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(x.d || '')) || !String(x.x || '').trim()) return;
    const it = { d: String(x.d), t: String(x.t || '').replace(/\D/g, '').slice(0, 4), x: String(x.x).trim().slice(0, 60), big: x.big ? 1 : 0 };
    if (out.some(y => y.d === it.d && (y.t || '') === it.t && !!y.big === !!it.big && near(y.x, it.x))) return; // 같은 날·같은 시각에 오타 정도 차이면 그대로
    out.push(it); add++;
  });
  return { items: out.slice(0, 400), add: add };
}
function mdList(me) {
  return mdRows(mdSheet()).map(o => { const likes = String(o.likes || '').split(',').filter(Boolean); return { id: String(o.id), at: +o.at, by: String(o.n), cat: String(o.cat), ins: String(o.ins), prod: String(o.prod), why: String(o.why || ''), tip: String(o.tip || ''), likes: likes.length, liked: !!(me && likes.indexOf(me) >= 0), mine: !!(me && o.c === me), pin: !!o.pin }; });
}
function doGet(e) { try { return doGet_(e); } catch (err) { return out({ ok: false, error: '팀 서버 오류: ' + err }); } }
function doGet_(e) {
  const prm = Object.assign({}, e.parameter || {});
  adoptOld();
  // 'c' 는 구글이 예약한 이름이라 주소에 쓰면 '파일을 열 수 없습니다' 가 남 → 앱은 uid 로 보냄
  if (prm.uid) prm.c = prm.uid;
  if (prm.a === 'mt') return out({ ok: true, mt: 2 });   // 👥 팀원 기능이 있는 서버인지 앱이 확인
  if (prm.a === 'otpreq') { if (!isAdmin(prm.k)) return out({ ok: false, error: '관리자만' }); return out({ ok: true, rows: otpRows(otpSheet()).filter(o => o.st === 'new').map(o => ({ rid: String(o.rid), at: +o.at, n: String(o.n), m: String(o.m), dev: String(o.dev || ''), note: String(o.note || '') })) }); }
  if (prm.a === 'otpstat') { const o = otpRows(otpSheet()).find(x => String(x.rid) === String(prm.rid) && x.m === String(prm.m || '').toLowerCase()); if (!o) return out({ ok: false }); return out({ ok: true, st: String(o.st), code: o.st === 'done' ? String(o.code || '') : '' }); }
  // 자동 개선 워크플로가 신고 내용·사진을 받아 감 (신고마다 한 번 쓰는 열쇠)
  if (prm.a === 'bug_get') {
    const o = bugOne(prm.id); if (!o || !o.ft || String(o.ft) !== String(prm.t)) return out({ ok: false, error: '열쇠가 맞지 않습니다' });
    let img = ''; if (o.img) { try { img = Utilities.base64Encode(DriveApp.getFileById(String(o.img)).getBlob().getBytes()); } catch (e) { img = ''; } }
    return out({ ok: true, id: String(o.id), kind: String(o.kind), title: String(o.title), desc: String(o.desc), diag: String(o.diag || ''), reply: String(o.reply || ''), img: img });
  }
  // 관리자: 캡처 사진 보기
  if (prm.a === 'bug_img') {
    if (!isAdmin(prm.k)) return out({ ok: false, error: '관리자만' });
    const o = bugOne(prm.id); if (!o || !o.img) return out({ ok: false });
    try { return out({ ok: true, img: Utilities.base64Encode(DriveApp.getFileById(String(o.img)).getBlob().getBytes()) }); } catch (e) { return out({ ok: false }); }
  }
  if (prm.a === 'md') { const me = allowed(prm.c) ? prm.c : ''; return out({ ok: true, rows: mdList(me) }); }
  if (prm.a === 'sug') {
    const me = allowed(prm.c) ? prm.c : '';
    return out({ ok: true, admin: !props().getProperty('ADMIN_KEY') ? 'none' : 'set', bug: 2, gh: !!props().getProperty('GH_TOKEN'), rows: sugList(me) });
  }
  const week = (prm.week || ''), me = prm.c || '';
  const t = JSON.parse(props().getProperty('w_' + week) || '{}');
  const rowsL = Object.keys(t).map(k => Object.assign({ me: k === me }, t[k]));
  return out({ ok: true, week: week, rows: rowsL, sug: true });
}

// ▶ 처음 한 번: 편집기 위 함수 칸에서 '권한설정' 을 고르고 ▶ 실행 → 권한 검토 → 허용 (시트·드라이브·인터넷 접속을 한 번에 허락)
function 권한설정() {
  UrlFetchApp.fetch(ACCESS + '?t=' + Date.now(), { muteHttpExceptions: true });
  CacheService.getScriptCache().put('ping', '1', 60);
  sheet(); otpSheet(); mdSheet(); cdSheet(); shotFolder();
  Logger.log('완료: 팀 서버 권한이 모두 허락됐습니다. 이제 배포 → 배포 관리 → 연필 → 새 버전 → 배포 를 해 주세요.');
}
