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
    return { k: L.slice(0, 3).join('|'), v: '[AFG-DB]\n' + L.join('\n') + '\n[PAGE]\n' + document.body.innerText.slice(0, 3000) };
  }
  function copy(txt, done) {
    // 오래 담은 뒤에는 창이 '누른 직후'가 아니라서 복사가 대답 없이 멈출 수 있다 → 2초 안에 안 되면 [📋 AOB9로 복사] 단추로
    var fired = false, ok = function () { if (!fired) { fired = true; done(true); } }, no = function () { if (!fired) { fired = true; done(false); } };
    setTimeout(no, 2000);
    try { if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(txt).then(ok, no); return; } } catch (e) {}
    try { var ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); var r = document.execCommand('copy'); ta.remove(); r ? ok() : no(); } catch (e) { no(); }
  }
  var rowsOf = function () { return document.querySelectorAll('tbody tr'); };
  var onList = function () { return rowsOf().length > 0 && !/\/db\/\d+/.test(location.pathname); };

  // ── 상세 화면: 그 고객만 (예전 방식)
  if (/\/db\/\d+/.test(location.pathname) || (!onList() && !/\/db\/?$/.test(location.pathname))) {
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
      var gg = grab(); if (!seen[gg.k]) { seen[gg.k] = 1; out.push(gg.v); save(); }
      await backTo(p);
    }
  }
  }
  window.__aobRun = 0;
  var txt = out.join('\n[END]\n');
  window.__aobStat = { done: 1, got: out.length, mask: nMask, fail: nFail, menus: menus.length };
  var fin = function (ok) {
    say('<b>' + out.length + '명</b> 담았습니다' + (nMask ? '<br>· 번호가 가려진 고객 ' + nMask + '명(아직 통화 안 함)은 가려진 번호로 들어가고, 통화 뒤 다시 누르면 번호가 채워집니다' : '') + (nFail ? '<br>· 못 연 고객 ' + nFail + '명' : '') +
      (ok ? '<br><br>✅ 클립보드에 복사됨<br>AOB9 → 고객 탭 ＋ 추가 → 회사 DB 고객 가져오기 → <b>붙여넣기</b>' : '<br><br><button id="aobCopy" style="font-size:16px;font-weight:800;padding:10px 16px;border:0;border-radius:10px;background:#2F6BFF;color:#fff;cursor:pointer">📋 AOB9로 복사</button>') +
      '<br><button id="aobX" style="margin-top:8px">닫기</button>');
    var c = document.getElementById('aobCopy'); if (c) c.onclick = function () { copy(txt, function (k2) { if (k2) fin(true); else prompt('아래 글을 전체 복사하세요', txt); }); };
    document.getElementById('aobX').onclick = function () { box.remove(); };
  };
  if (!out.length) { say('담은 고객이 없습니다. 내 DB 목록 화면에서 다시 눌러 주세요.<br><button id="aobX">닫기</button>'); document.getElementById('aobX').onclick = function () { box.remove(); }; return; }
  // 끝 화면(복사 단추)을 먼저 띄우고, 그다음 저절로 복사를 시도한다 — 저절로 복사가 대답 없이 멈춰도 단추는 남는다
  fin(false);
  copy(txt, function (ok) { if (ok) fin(true); });
})();
