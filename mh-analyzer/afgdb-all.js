// AOB9 · 회사 DB(link.afg.kr) "AOB로 복사" 북마크 본체 (2026-11-19 대표님: "고객마다 누르는 거 말고 한 번에")
// - 목록 화면에서 누르면: 내 DB 목록의 모든 페이지 → 고객마다 상세 화면을 차례로 열어(대표님이 직접 누르는 것과 같은 길)
//   이름·연락처·생년월일·성별·지역·인수상태·배분일·특이사항을 모아 한 번에 복사한다.
//   가운데 번호가 가려진 고객(아직 한 번도 통화하지 않은 고객)은 번호 칸을 한 번 눌러 보고, 그래도 가려져 있으면 가려진 그대로 담는다.
// - 상세 화면에서 누르면: 예전처럼 그 고객만 담는다.
// 모은 글은 이 브라우저 안에서 클립보드로만 간다 — 어디로도 보내지 않는다.
(async function () {
  if (window.__aobRun) { alert('이미 담는 중입니다. 오른쪽 위 상자를 보세요.'); return; }
  var W = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var PH = /^01[016789][-*\d]{8,11}$/;
  function phoneEl() { return [].slice.call(document.querySelectorAll('input')).find(function (e) { return PH.test(String(e.value || '').replace(/\s/g, '')); }); }
  function nameVal() { var e = [].slice.call(document.querySelectorAll('input')).find(function (x) { return /^[가-힣]{2,5}$/.test(String(x.value || '').trim()); }); return e ? e.value.trim() : ''; }
  function grab() {
    var L = [];
    document.querySelectorAll('input,textarea,select').forEach(function (e) {
      var t = (e.type || '').toLowerCase(); if (/hidden|file|button|submit|password|search/.test(t)) return;
      var v;
      if (t === 'checkbox' || t === 'radio') { if (!e.checked) return; var lb = (e.labels && e.labels[0] && e.labels[0].innerText) || (e.parentElement && e.parentElement.innerText) || e.value || ''; v = '[선택] ' + lb.trim().split('\n')[0]; }
      else if (e.tagName === 'SELECT') v = e.selectedIndex >= 0 ? e.options[e.selectedIndex].text : '';
      else v = e.value;
      v = String(v || '').trim(); if (!v) return;
      var lab = '';
      for (var p = e.parentElement, i = 0; p && i < 4 && !lab; p = p.parentElement, i++) { var tx = (p.innerText || '').split('\n')[0].trim(); if (tx && tx !== v && tx.length < 16 && !/^\[/.test(tx)) lab = tx; }
      L.push(/^\[선택\]/.test(v) ? v : (lab ? lab + ': ' : '') + v);
    });
    // 2026-11-19: 녹취는 회사 사이트에서 듣도록 DB 번호·메뉴·녹취 개수만 같이 담는다 (녹음 파일·임시 주소는 담지 않음)
    var all = document.body.innerText || '', id = (location.pathname.match(/\/db\/(\d+)/) || [])[1] || '', mn = new URLSearchParams(location.search).get('menu') || '';
    var rc = (all.match(/녹취파일\s*목록\s*\((\d+)개\)/) || [])[1] || '';
    var meta = id ? 'DB번호: ' + id + ' 메뉴: ' + mn + ' 녹취: ' + rc + '\n' : '';
    return { k: L.slice(0, 3).join('|'), v: '[AFG-DB]\n' + L.join('\n') + '\n[PAGE]\n' + meta + all.slice(0, 3000) };
  }
  function copy(txt, done) {
    // 오래 담은 뒤에는 창이 '누른 직후'가 아니라서 복사가 대답 없이 멈출 수 있다 → 2초 안에 안 되면 [📋 AOB9로 복사] 단추로
    var fired = false, ok = function () { if (!fired) { fired = true; done(true); } }, no = function () { if (!fired) { fired = true; done(false); } };
    setTimeout(no, 2000);
    try { if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(txt).then(ok, no); return; } } catch (e) {}
    try { var ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); var r = document.execCommand('copy'); ta.remove(); r ? ok() : no(); } catch (e) { no(); }
  }
  var rowsOf = function () { return document.querySelectorAll('tbody tr'); };
  // ── 🔁 인수상태 동기화 (2026-11-20 대표님: "앱의 고객상태와 회사 사이트 고객상태를 자동으로 · 최근에 바꾼 쪽으로")
  //   AOB9에서 [🔁 상태 복사] → 클립보드의 [AOB-ST] 목록(DB번호·이름·상태·앱에서 바꾼 시각)을 읽어,
  //   고객 상세 화면을 열 때 회사 사이트의 '인수상태 변경시점'이 앱보다 오래됐으면 그 상태 칸을 눌러 바꾼다(바꾸기 전에 목록을 보여 주고 묻는다).
  //   목록은 이 브라우저 안에서만 쓰고 어디로도 보내지 않는다.
  var PEND = null, stRes = [];
  async function readPend() {
    var t = '';
    try { t = await navigator.clipboard.readText(); } catch (e) { return null; }
    var m = /\[AOB-ST\](\{[\s\S]*\})/.exec(t || ''); if (!m) return null;
    try { var o = JSON.parse(m[1]), P = {}; (o.L || []).forEach(function (x) { if (x && x.id && x.st) P[String(x.id)] = x; }); return Object.keys(P).length ? P : null; } catch (e) { return null; }
  }
  function stBoxes() {   // 인수상태 칸들 (글자가 상태 이름인 체크·라디오)
    var NAMES = ['부재', 'TA거절', '상담중', '가입불가(AS신청)', 'AS재신청', 'AS승인', '상담완료', '청약완료', '전산완료'];
    return [].slice.call(document.querySelectorAll('input[type=checkbox],input[type=radio]')).map(function (e) {
      var lb = ((e.labels && e.labels[0] && e.labels[0].innerText) || (e.parentElement && e.parentElement.innerText) || (e.parentElement && e.parentElement.parentElement && e.parentElement.parentElement.innerText) || '').trim().split('\n')[0].trim();
      return { e: e, t: lb };
    }).filter(function (x) { return NAMES.indexOf(x.t) >= 0; });
  }
  function siteSt() { var b = stBoxes().find(function (x) { return x.e.checked; }); return b ? b.t : ''; }
  function siteAt() { return ((document.body.innerText || '').match(/인수상태\s*변경시점\s*(\d{4}-\d{2}-\d{2}(?:\s\d{2}:\d{2})?)/) || [])[1] || ''; }
  async function applySt() {
    if (!PEND) return;
    var id = (location.pathname.match(/\/db\/(\d+)/) || [])[1]; if (!id || !PEND[id] || PEND[id].done) return;
    var want = PEND[id], cur = siteSt(), at = siteAt(), r = { n: want.n, st: want.st, from: cur };
    PEND[id].done = 1;
    if (cur === want.st) { r.k = 'same'; stRes.push(r); return; }
    if (at && want.at && at >= want.at) { r.k = 'site'; r.at = at; stRes.push(r); return; }   // 회사 사이트가 더 최근 → 회사 쪽이 맞음 (앱이 가져갈 때 맞춰짐)
    var box = stBoxes().find(function (x) { return x.t === want.st; });
    if (!box) { r.k = 'nobox'; stRes.push(r); return; }
    var oc = window.confirm, oa = window.alert; window.confirm = function () { return true; }; window.alert = function () {};
    try {
      (box.e.labels && box.e.labels[0] ? box.e.labels[0] : box.e).click(); await W(500);
      if (!box.e.checked) { box.e.click(); await W(400); }
      // 저장 단추가 따로 있으면 (상태 칸 가까이에서) 누른다
      var sv = null;
      for (var p2 = box.e.parentElement, k2 = 0; p2 && k2 < 8 && !sv; p2 = p2.parentElement, k2++) sv = [].slice.call(p2.querySelectorAll('button')).find(function (b) { return /^(저장|수정|변경|상태\s*변경|적용)/.test((b.innerText || '').trim()); });
      if (!sv) sv = [].slice.call(document.querySelectorAll('button')).find(function (b) { var tx = (b.innerText || '').trim(); return /^(저장|수정하기|변경하기|인수상태\s*변경)$/.test(tx); });
      if (sv) { sv.click(); await W(1300); }
      for (var k3 = 0; k3 < 10 && siteSt() !== want.st; k3++) await W(200);
      r.k = siteSt() === want.st ? (sv || siteAt() !== at ? 'ok' : 'okq') : 'fail';
    } catch (e) { r.k = 'fail'; }
    window.confirm = oc; window.alert = oa;
    stRes.push(r);
  }
  function stSummary() {
    if (!PEND) return '';
    var c = function (k) { return stRes.filter(function (x) { return x.k === k; }); }, miss = Object.keys(PEND).filter(function (id) { return !PEND[id].done; }).map(function (id) { return PEND[id].n; });
    var ok = c('ok').concat(c('okq')), bad = c('fail').concat(c('nobox'));
    return '<div style="margin-top:8px;padding:8px 10px;border-radius:10px;background:#FFF6E8">🔁 <b>인수상태 맞추기</b><br>' +
      (ok.length ? '✅ 회사 사이트를 바꿈 ' + ok.length + '명' + (c('okq').length ? ' (저장 단추 없이 바뀜 ' + c('okq').length + '명 — 회사 사이트에서 한 번 확인해 주세요)' : '') + '<br>' : '') +
      (c('same').length ? '· 이미 같음 ' + c('same').length + '명<br>' : '') +
      (c('site').length ? '· 회사 사이트가 더 최근이라 그대로 ' + c('site').length + '명 (AOB9에 붙여넣으면 앱이 회사 상태로 바뀝니다)<br>' : '') +
      (bad.length ? '⚠ 못 바꿈 ' + bad.length + '명: ' + bad.map(function (x) { return x.n + '(' + x.st + ')'; }).join(', ') + '<br>' : '') +
      (miss.length ? '· 목록에서 못 찾음 ' + miss.length + '명: ' + miss.slice(0, 6).join(', ') + (miss.length > 6 ? ' …' : '') : '') + '</div>';
  }
  function askPend(say2) {
    return new Promise(function (res) {
      var L = Object.keys(PEND).map(function (id) { return PEND[id]; });
      say2('AOB9에서 바꾼 인수상태 <b>' + L.length + '명</b>을 회사 사이트에도 맞출까요?<div style="max-height:160px;overflow:auto;margin:6px 0;font-size:13px">' + L.map(function (x) { return '· ' + String(x.n || '').replace(/[<>&]/g, '') + ' → <b>' + String(x.st).replace(/[<>&]/g, '') + '</b>'; }).join('<br>') + '</div><span style="font-size:12px;color:#5b6472">회사 사이트에서 그보다 나중에 바꾼 고객은 건드리지 않습니다.</span><br><button id="aobStY" style="margin-top:6px;font-weight:800">네, 맞추고 담기</button> <button id="aobStN">아니요, 담기만</button>');
      document.getElementById('aobStY').onclick = function () { res(true); };
      document.getElementById('aobStN').onclick = function () { res(false); };
    });
  }
  var onList = function () { return rowsOf().length > 0 && !/\/db\/\d+/.test(location.pathname); };

  // ── 상세 화면: 그 고객만 (예전 방식)
  if (/\/db\/\d+/.test(location.pathname) || (!onList() && !/\/db\/?$/.test(location.pathname))) {
    PEND = await readPend();
    var did = (location.pathname.match(/\/db\/(\d+)/) || [])[1];
    if (PEND && did && PEND[did] && confirm('AOB9에서 바꾼 인수상태(' + PEND[did].st + ')로 회사 사이트도 바꿀까요?\n(회사 사이트에서 더 나중에 바꿨으면 건드리지 않습니다)')) { await applySt(); var rr = stRes[0]; if (rr) alert(rr.k === 'ok' || rr.k === 'okq' ? '인수상태를 ' + rr.st + '(으)로 바꿨습니다' : rr.k === 'same' ? '이미 같은 상태입니다' : rr.k === 'site' ? '회사 사이트에서 더 나중에 바꿔서 그대로 둡니다' : '인수상태를 바꾸지 못했습니다 · 직접 바꿔 주세요'); }
    var g = grab(), K = 'oc9_db_bag', bag;
    try { bag = JSON.parse(localStorage.getItem(K) || 'null'); } catch (e) { bag = null; }
    if (!bag || Date.now() - bag.t > 12 * 3600e3) bag = { t: Date.now(), b: [] };
    var dup = bag.b.some(function (x) { return x.k === g.k; }); if (!dup) bag.b.push(g); bag.t = Date.now();
    try { localStorage.setItem(K, JSON.stringify(bag)); } catch (e) {}
    var out1 = bag.b.map(function (x) { return x.v; }).join('\n[END]\n');
    copy(out1, function (ok) { if (ok) alert((dup ? '이미 담은 고객입니다.\n' : '담았습니다.\n') + '지금까지 ' + bag.b.length + '명 복사됨\n\nAOB9 → 회사 DB 고객 가져오기 → 붙여넣기\n(목록 화면에서 누르면 한 번에 전부 담깁니다)'); else prompt('아래 글을 전체 복사하세요', out1); });
    return;
  }

  // ── 목록 화면: 전부
  window.__aobRun = 1;
  var box = document.createElement('div');
  box.style.cssText = 'position:fixed;top:14px;right:14px;z-index:2147483647;width:300px;background:#fff;color:#172033;border:2px solid #2F6BFF;border-radius:14px;box-shadow:0 10px 30px rgba(0,0,0,.25);padding:14px 16px;font:14px/1.6 system-ui,"Malgun Gothic",sans-serif';
  document.body.appendChild(box);
  var stop = false;
  var say = function (h) { box.innerHTML = '<b style="font-size:15px">AOB로 복사 · 한 번에 담기</b><div style="margin-top:6px">' + h + '</div>'; };
  say('준비 중…');
  PEND = await readPend();
  if (PEND && !(await askPend(say))) PEND = null;
  var pageBtns = function () { return [].slice.call(document.querySelectorAll('nav button, .MuiPagination-ul button')); };
  var curPage = function () { var b = pageBtns().find(function (x) { return x.getAttribute('aria-current'); }); return b ? +b.innerText.trim() || 1 : 1; };
  var maxPage = function () { return pageBtns().map(function (x) { return +x.innerText.trim() || 0; }).reduce(function (a, b) { return Math.max(a, b); }, 1); };
  async function waitList() { for (var k = 0; k < 60 && !onList(); k++) await W(150); await W(350); }
  async function goPage(p) {
    for (var t = 0; t < 3 && curPage() !== p; t++) {
      var b = pageBtns().find(function (x) { return +x.innerText.trim() === p; }) || pageBtns().find(function (x) { return /next/i.test(x.getAttribute('aria-label') || ''); });
      if (!b) break; b.click(); await W(900);
    }
  }
  var out = [], seen = {}, nMask = 0, nFail = 0, P = 1, total = 0, prevName = '', menuNow = '';
  // 2026-11-19 대표님: "30명이 전부가 아닐텐데" — 보장리스트Ⅰ·Ⅱ·실버·달러… 왼쪽 메뉴의 모든 '리스트'를 차례로 돈다
  var menuBtns = function () { return [].slice.call(document.querySelectorAll('button')).filter(function (b) { return b.getBoundingClientRect().left < 300 && /리스트/.test(b.innerText || '') && !/쿠폰/.test(b.innerText || ''); }); };
  var menus = menuBtns().map(function (b) { return b.innerText.trim(); }).filter(function (x, i, a) { return a.indexOf(x) === i; });
  if (!menus.length) menus = [''];
  // 상세 → 목록은 '뒤로 가기' 대신 왼쪽 메뉴를 다시 눌러 돌아온다 (뒤로 가기는 가끔 페이지를 통째로 다시 열어 담던 것이 날아감)
  async function backTo(p) { var b = menuNow && menuBtns().find(function (x) { return x.innerText.trim() === menuNow; }); if (b) b.click(); else history.back(); await W(600); await waitList(); if (curPage() !== p) { await goPage(p); await waitList(); } }
  // 중간에 화면이 다시 열려도 이어서 담도록 이 탭에만 잠깐 기억 (탭을 닫으면 사라짐)
  try { var SV = JSON.parse(sessionStorage.getItem('aob_all') || 'null'); if (SV && Date.now() - SV.t < 3 * 3600e3 && SV.out && SV.out.length) { out = SV.out; SV.out.forEach(function (v) { seen[v.split('\n').slice(1, 4).join('|')] = 1; }); } } catch (e) {}
  var save = function () { try { sessionStorage.setItem('aob_all', JSON.stringify({ t: Date.now(), out: out })); } catch (e) {} };
  var realRows = function () { return [].slice.call(rowsOf()).filter(function (r) { return r.children.length >= 6; }).length; };
  for (var mi = 0; mi < menus.length && !stop; mi++) {
   menuNow = menus[mi];
   if (menuNow) { var mb = menuBtns().find(function (b) { return b.innerText.trim() === menuNow; }); if (!mb) continue; mb.click(); await W(2200); await waitList(); }
   if (!realRows()) continue;
   P = maxPage(); prevName = '';
  for (var p = 1; p <= P && !stop; p++) {
    await goPage(p); await waitList();
    var n = rowsOf().length;
    for (var i = 0; i < n && !stop; i++) {
      if (!onList()) { await backTo(p); }
      if (curPage() !== p) { await goPage(p); await waitList(); }
      var tr = rowsOf()[i]; if (!tr) break;
      total++;
      if (realRows() === 0) break;
      say((menuNow ? menuNow + ' · ' : '') + p + ' / ' + P + ' 페이지 · ' + (i + 1) + '번째 고객 여는 중<br>지금까지 <b>' + out.length + '명</b> 담음' + (nMask ? ' (번호 가려짐 ' + nMask + ')' : '') + '<br><button id="aobStop" style="margin-top:6px">그만 담고 지금까지 복사</button>');
      window.__aobStat = { menu: mi + 1, menus: menus.length, p: p, P: P, i: i, got: out.length, mask: nMask, fail: nFail };
      var sb = document.getElementById('aobStop'); if (sb) sb.onclick = function () { stop = true; };
      (tr.children[3] || tr.children[1] || tr).click();
      for (var k = 0; k < 60 && !/\/db\/\d+/.test(location.pathname); k++) await W(150);
      var ph = null, nm = '';
      for (k = 0; k < 40; k++) { ph = phoneEl(); nm = nameVal(); if (ph && nm && nm !== prevName) break; await W(150); }
      if (!ph || !nm) { nFail++; await backTo(p); continue; }
      // 가운데 번호가 가려져 있으면 잠깐 기다렸다가, 번호 칸을 눌러 본다 (대표님이 누르는 것과 같은 동작)
      for (k = 0; k < 8 && /\*/.test(ph.value); k++) { await W(250); ph = phoneEl() || ph; }
      if (/\*/.test(ph.value)) {
        [ph.parentElement, ph.parentElement && ph.parentElement.parentElement, ph].forEach(function (el) { try { el && el.click(); } catch (e) {} });
        for (k = 0; k < 10 && /\*/.test((phoneEl() || ph).value); k++) await W(200);
        ph = phoneEl() || ph;
      }
      if (/\*/.test(ph.value)) nMask++;
      prevName = nm;
      if (PEND) await applySt();
      var gg = grab(); if (!seen[gg.k]) { seen[gg.k] = 1; out.push(gg.v); save(); }
      await backTo(p);
    }
  }
  }
  // 📮 DB 신청하기 · DB 신청현황 표도 같이 (2026-11-20 대표님 요청 · 고객 정보 없음: 디비 이름·수량·날짜·설명뿐)
  var req = null;
  async function tableOf() {
    var hd = [].slice.call(document.querySelectorAll('thead th')).map(function (x) { return (x.innerText || '').trim(); });
    var rs = [].slice.call(rowsOf()).map(function (r) { return [].slice.call(r.children).map(function (td) { return (td.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 120); }); }).filter(function (r) { return r.length >= 3; });
    return { head: hd, rows: rs };
  }
  async function openSide(re) { var b = [].slice.call(document.querySelectorAll('button')).find(function (x) { return x.getBoundingClientRect().left < 300 && re.test((x.innerText || '').trim()); }); if (!b) return false; b.click(); await W(2500); await waitList(); return true; }
  if (!stop) {
    try {
      say('DB 신청하기 · 신청현황 표 담는 중…');
      req = { at: Date.now() };
      if (await openSide(/^DB\s*신청하기$/)) { var ap = await tableOf(); var note = [].slice.call(document.querySelectorAll('main *, body *')).map(function (x) { return x.children.length ? '' : (x.innerText || '').trim(); }).filter(function (t) { return /DB\s*지원\s*대상|지원여부/.test(t) && t.length < 80; }).slice(0, 2); ap.note = note; req.apply = ap; }
      if (await openSide(/^DB\s*신청현황$/)) {
        var hs = await tableOf(), PP = Math.min(maxPage(), 15);
        for (var pp = 2; pp <= PP && !stop; pp++) { await goPage(pp); await waitList(); hs.rows = hs.rows.concat((await tableOf()).rows); }
        req.hist = hs;
      }
      if (!req.apply && !req.hist) req = null;
    } catch (e) { req = null; }
  }
  window.__aobRun = 0;
  var txt = out.join('\n[END]\n') + (req ? '\n[END]\n[AFG-DBREQ]\n' + JSON.stringify(req) : '');
  window.__aobStat = { done: 1, got: out.length, mask: nMask, fail: nFail, menus: menus.length };
  var fin = function (ok) {
    say('<b>' + out.length + '명</b> 담았습니다' + (nMask ? '<br>· 번호가 가려진 고객 ' + nMask + '명(아직 통화 안 함)은 가려진 번호로 들어가고, 통화 뒤 다시 누르면 번호가 채워집니다' : '') + (nFail ? '<br>· 못 연 고객 ' + nFail + '명' : '') +
      (ok ? '<br><br>✅ 클립보드에 복사됨<br>AOB9 → 고객 탭 ＋ 추가 → 회사 DB 고객 가져오기 → <b>붙여넣기</b>' : '<br><br><button id="aobCopy" style="font-size:16px;font-weight:800;padding:10px 16px;border:0;border-radius:10px;background:#2F6BFF;color:#fff;cursor:pointer">📋 AOB9로 복사</button>') +
      stSummary() + '<br><button id="aobX" style="margin-top:8px">닫기</button>');
    var c = document.getElementById('aobCopy'); if (c) c.onclick = function () { copy(txt, function (k2) { if (k2) fin(true); else prompt('아래 글을 전체 복사하세요', txt); }); };
    document.getElementById('aobX').onclick = function () { box.remove(); };
  };
  if (!out.length) { say('담은 고객이 없습니다. 내 DB 목록 화면에서 다시 눌러 주세요.<br><button id="aobX">닫기</button>'); document.getElementById('aobX').onclick = function () { box.remove(); }; return; }
  // 끝 화면(복사 단추)을 먼저 띄우고, 그다음 저절로 복사를 시도한다 — 저절로 복사가 대답 없이 멈춰도 단추는 남는다
  fin(false);
  copy(txt, function (ok) { if (ok) fin(true); });
})();
