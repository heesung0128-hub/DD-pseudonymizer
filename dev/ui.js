/*
 * 화면 코드. 보안 원칙 (이 파일 전체에 적용):
 *  1) 원본·대응표·비밀번호는 아래 `state` 객체(메모리)에만 둔다.
 *     localStorage / sessionStorage / IndexedDB / 쿠키 / Cache API 는 사용하지 않는다.
 *  2) 네트워크 API(fetch, XHR, WebSocket, sendBeacon, 외부 <script>/<img>)를 쓰지 않는다. (CSP 로도 차단됨)
 *  3) 파일 내용·셀 값을 HTML 에 넣을 때는 반드시 esc() 를 거친다. 인라인 이벤트 속성(onclick=)은 쓰지 않는다.
 *  4) 파일 저장은 Blob + <a download> 로만 한다. (서버로 보내지 않음)
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  var MODE_LABEL = { keep: '유지', token: '토큰화', 'delete': '삭제' };
  var PAGE = 100, MAX_PREVIEW = 200;

  function freshState() {
    return {
      file: null, sheetName: null, grid: null, merged: 0, headerRow: 0,
      workKey: null, work: null, hdr: null, rec: null, cfg: null, res: null,
      sel: { type: null }, viewRows: PAGE,
      keyId: null, keyIdSig: null,
      rs: { pendingKey: null, pendingKeyName: '', entries: null, restorer: null, keyId: null, input: null, result: null, sheetView: null }
    };
  }
  var state = freshState();

  /* ------------------------------------------------------------ 공통 도우미 */
  var toastTimer = null;
  function toast(msg, bad) {
    var t = $('toast'); t.textContent = msg; t.className = bad ? 'bad' : ''; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, bad ? 6000 : 2800);
  }
  function fmtSize(n) { return n > 1048576 ? (n / 1048576).toFixed(1) + 'MB' : Math.max(1, Math.round(n / 1024)) + 'KB'; }
  function today() { var d = new Date(); return d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0'); }

  function downloadBlob(blob, name) {
    var url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 15000);
  }
  function downloadText(text, name, mime, bom) { downloadBlob(new Blob([bom ? '﻿' : '', text], { type: mime + ';charset=utf-8' }), name); }

  function copyText(text, okMsg) {
    function fallback() {
      var ta = document.createElement('textarea'); ta.value = text; ta.className = 'offscreen';
      document.body.appendChild(ta); ta.select();
      var ok = false; try { ok = document.execCommand('copy'); } catch (e) { }
      ta.remove(); toast(ok ? okMsg : '복사하지 못했어요. 직접 선택해서 복사해 주세요.', !ok);
    }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(function () { toast(okMsg); }, fallback);
    else fallback();
  }

  function openModal(html, wire) {
    var ov = $('overlay'); ov.innerHTML = '<div class="modal" role="dialog" aria-modal="true">' + html + '</div>'; ov.hidden = false;
    if (wire) wire(ov.firstChild);
    var f = ov.querySelector('button.primary, button'); if (f) f.focus();
  }
  function closeModal() { var ov = $('overlay'); ov.hidden = true; ov.innerHTML = ''; }

  function makeDrop(zone, input, button, handler) {
    function pick() { input.click(); }
    button.addEventListener('click', function (e) { e.stopPropagation(); pick(); });
    zone.addEventListener('click', pick);
    zone.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } });
    ['dragenter', 'dragover'].forEach(function (ev) { zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add('over'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.remove('over'); }); });
    zone.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) handler(f); });
    input.addEventListener('change', function () { var f = input.files[0]; input.value = ''; if (f) handler(f); });
  }
  // 창 어디에 놓아도 브라우저가 파일을 열어 버리지 않게 막는다.
  ['dragover', 'drop'].forEach(function (ev) { window.addEventListener(ev, function (e) { e.preventDefault(); }); });

  /* ------------------------------------------------------------ 탭 */
  function gotoTab(n) {
    if (n === 2 && !state.file) return;
    if (n === 2) enterStep2();
    [1, 2, 3].forEach(function (i) {
      $('p' + i).hidden = i !== n;
      $('tab' + i).setAttribute('aria-selected', i === n ? 'true' : 'false');
    });
    window.scrollTo(0, 0);
  }
  function updateTabs() { $('tab2').setAttribute('aria-disabled', state.file ? 'false' : 'true'); }
  document.querySelectorAll('nav.tabs button').forEach(function (b) {
    b.addEventListener('click', function () { var n = +b.dataset.tab; if (n === 2 && !state.file) { toast('먼저 파일을 불러와 주세요.', true); return; } gotoTab(n); });
  });

  /* ------------------------------------------------------------ ① 불러오기 */
  makeDrop($('drop1'), $('file1'), $('pick1'), loadFile);

  async function loadFile(file) {
    try {
      if (file.size > 60 * 1048576) throw new Error('파일이 너무 커요 (60MB 초과).');
      var ext = (file.name.split('.').pop() || '').toLowerCase(), buf = await file.arrayBuffer(), f = { name: file.name, size: file.size };
      if (ext === 'csv' || ext === 'tsv') {
        var d = Core.decodeBytes(buf);
        f.kind = 'csv'; f.enc = d.enc; f.rows = Core.parseDelimited(d.text, Core.detectDelimiter(d.text)); f.sheets = ['CSV'];
      } else if (ext === 'xlsx' || ext === 'xls' || ext === 'xlsm') {
        f.kind = 'xlsx'; f.wb = Core.readWorkbook(buf); f.sheets = f.wb.SheetNames.slice();
        if (!f.sheets.length) throw new Error('시트가 없는 파일이에요.');
      } else throw new Error('.xlsx .xls .csv 파일만 불러올 수 있어요.');
      var keepRs = state.rs;
      state = freshState(); state.rs = keepRs;
      state.file = f;
      selectSheet(f.sheets[0]);
      updateTabs();
    } catch (e) { toast('불러오지 못했어요: ' + e.message, true); }
  }

  function selectSheet(name) {
    var f = state.file; state.sheetName = name;
    if (f.kind === 'csv') { state.grid = Core.normalizeGrid(f.rows); state.merged = 0; }
    else { var r = Core.sheetToGrid(f.wb.Sheets[name]); state.grid = r.grid; state.merged = r.merged; }
    state.headerRow = Core.guessHeader(state.grid);
    state.workKey = null; state.cfg = null;
    renderStep1();
  }

  function renderStep1() {
    var f = state.file, box = $('loaded'); box.hidden = !f; if (!f) { box.innerHTML = ''; return; }
    var g = state.grid, nShow = Math.min(g.length, 12), nC = g[0].length, h = '';
    h += '<div class="card"><h2>' + esc(f.name) + ' <span class="muted small">(' + fmtSize(f.size) + (f.enc ? ' · ' + esc(f.enc) : '') + ')</span></h2>';
    if (f.sheets.length > 1) {
      h += '<p>시트 선택: <select id="sheetSel">' + f.sheets.map(function (s) { return '<option' + (s === state.sheetName ? ' selected' : '') + ' value="' + esc(s) + '">' + esc(s) + '</option>'; }).join('') + '</select></p>';
    }
    h += '<p>표 크기: <b>' + g.length + '행 × ' + nC + '열</b>';
    if (state.merged) h += ' <span class="muted">· 병합 셀 ' + state.merged + '칸은 같은 값으로 채워서 읽었어요.</span>';
    h += '</p>';
    h += '<p><b>머리글(제목) 행</b>을 골라 주세요. 머리글 위쪽 행은 <u>가져오지도, 내보내지도 않아요</u>. &nbsp; 번호 입력: <input type="number" id="hdrNum" min="1" max="' + g.length + '" value="' + (state.headerRow + 1) + '" class="w5"></p>';
    h += '<div class="tablewrap"><table class="grid"><thead><tr><th class="rn">머리글</th><th class="rn">행</th>';
    for (var c = 0; c < nC; c++) h += '<th>' + Core.colLetter(c) + '</th>';
    h += '</tr></thead><tbody>';
    for (var r = 0; r < nShow; r++) {
      h += '<tr' + (r === state.headerRow ? ' class="hdrrow"' : '') + '><td class="rn"><input type="radio" name="hdr" value="' + r + '"' + (r === state.headerRow ? ' checked' : '') + ' aria-label="' + (r + 1) + '행을 머리글로"></td><td class="rn">' + (r + 1) + '</td>';
      for (c = 0; c < nC; c++) h += '<td' + (r < state.headerRow ? ' class="del"' : '') + '>' + esc(g[r][c]) + '</td>';
      h += '</tr>';
    }
    h += '</tbody></table></div>';
    if (g.length > nShow) h += '<p class="muted small">앞의 ' + nShow + '행만 보여요. 더 아래 행이 머리글이면 위 번호 칸에 입력하세요.</p>';
    var col = Core.findBracketCells(g);
    if (col.length) h += '<div class="banner bad"><b>토큰과 헷갈릴 수 있어요.</b> 원본에 <code>[[...]]</code> 모양의 글자가 ' + col.length + '칸에 이미 있어요 (예: ' + esc(col[0].text.slice(0, 40)) + '). 그대로 진행하면 복원할 때 섞일 수 있으니 원본에서 바꾸는 것을 권해요.</div>';
    h += '<div class="row mt12"><button class="btn primary" type="button" id="goStep2">다음: 가릴 곳 고르기 →</button></div></div>';
    box.innerHTML = h;
    var sel = $('sheetSel'); if (sel) sel.addEventListener('change', function () { selectSheet(sel.value); });
    box.querySelectorAll('input[name=hdr]').forEach(function (rb) { rb.addEventListener('change', function () { state.headerRow = +rb.value; state.workKey = null; renderStep1(); }); });
    $('hdrNum').addEventListener('change', function () { var v = Math.max(1, Math.min(g.length - 1, parseInt(this.value, 10) || 1)); state.headerRow = v - 1; state.workKey = null; renderStep1(); });
    $('goStep2').addEventListener('click', function () { gotoTab(2); });
  }

  /* ------------------------------------------------------------ ② 선택·내보내기 */
  function enterStep2() {
    var key = state.sheetName + '|' + state.headerRow;
    if (state.workKey === key && state.cfg) { refresh(); return; }
    var work = state.grid.slice(state.headerRow);
    if (work.length < 2) { toast('머리글 아래에 데이터가 없어요. 머리글 행을 다시 골라 주세요.', true); gotoTab(1); return; }
    state.work = work; state.workKey = key;
    var nC = work[0].length, cfg = Core.newConfig();
    state.hdr = []; state.rec = [];
    for (var c = 0; c < nC; c++) {
      var name = work[0][c].trim() || ('열 ' + Core.colLetter(c));
      state.hdr.push(name); state.rec.push(Core.isRecommended(name) && !!work[0][c].trim());
      cfg.colMode.push('keep'); cfg.colPrefix.push(Core.sanitizePrefix(name, '값')); cfg.colCross.push(Core.crossDefault(name));
    }
    state.cfg = cfg; state.sel = { type: null }; state.viewRows = PAGE; state.keyId = null; state.keyIdSig = null;
    $('manualWords').value = ''; $('manualPrefix').value = '기타'; $('keyMsg').innerHTML = '';
    refresh();
  }

  function cfgSig() {
    var c = state.cfg;
    return JSON.stringify([state.workKey, c.colMode, c.colPrefix, c.colCross, Array.from(c.cellOv), c.manualWords, c.manualPrefix, Array.from(c.crossExcluded)]);
  }

  function refresh() {
    state.res = Core.compute(state.work, state.cfg);
    renderRec(); renderColTable(); renderGrid(); renderPreview(); renderSelBar(); renderCross(); renderWarn(); renderExport();
  }

  /* ---- 추천 배너 */
  function renderRec() {
    var recs = [], c;
    for (c = 0; c < state.rec.length; c++) if (state.rec[c] && state.cfg.colMode[c] === 'keep') recs.push(c);
    var b = $('recBanner');
    if (!recs.length) { b.innerHTML = ''; return; }
    b.innerHTML = '<div class="banner info"><b>개인정보로 보이는 열을 찾았어요:</b> ' + recs.map(function (c) { return esc(state.hdr[c]); }).join(', ') +
      ' <button class="btn sm primary" type="button" id="applyRec">추천대로 토큰화</button></div>';
    $('applyRec').addEventListener('click', function () { recs.forEach(function (c) { setColMode([c], 'token'); }); refresh(); });
  }

  /* ---- 열 설정 표 */
  function renderColTable() {
    var cfg = state.cfg, res = state.res, nC = state.work[0].length, h = '<thead><tr><th>열</th><th>예시</th><th>처리</th><th>접두어</th><th>문장 속에서도 찾기</th></tr></thead><tbody>';
    for (var c = 0; c < nC; c++) {
      var m = cfg.colMode[c], samples = [];
      for (var r = 1; r < state.work.length && samples.length < 2; r++) if (state.work[r][c].trim()) samples.push(state.work[r][c].trim().slice(0, 14));
      var w = res.colWarn[c], wtxt = [];
      if (m !== 'delete') { if (w.rrn) wtxt.push('주민번호 모양 ' + w.rrn); if (w.phone) wtxt.push('전화번호 모양 ' + w.phone); if (w.mail) wtxt.push('이메일 모양 ' + w.mail); }
      h += '<tr><td><b>' + esc(state.hdr[c]) + '</b>' + (state.rec[c] ? '<span class="badge rec">추천: 가리기</span>' : '') + (wtxt.length ? '<span class="badge warn">⚠ ' + wtxt.join(' · ') + '칸</span>' : '') + '</td>' +
        '<td class="sample">' + esc(samples.join(' · ')) + '</td><td><div class="seg">' +
        ['keep', 'token', 'delete'].map(function (k) { return '<button type="button" class="m-' + k + (m === k ? ' on' : '') + '" data-act="mode" data-c="' + c + '" data-m="' + k + '">' + MODE_LABEL[k] + '</button>'; }).join('') + '</div></td>' +
        '<td>' + (m === 'token' ? '<input type="text" data-act="prefix" data-c="' + c + '" value="' + esc(cfg.colPrefix[c]) + '" aria-label="접두어">' : '<span class="muted">-</span>') + '</td>' +
        '<td>' + (m === 'token' ? '<label class="chk"><input type="checkbox" data-act="cross" data-c="' + c + '"' + (cfg.colCross[c] ? ' checked' : '') + '> 사용</label>' : '<span class="muted">-</span>') + '</td></tr>';
    }
    $('colTable').innerHTML = h + '</tbody>';
  }
  $('colTable').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-act=mode]'); if (!b) return;
    setColMode([+b.dataset.c], b.dataset.m); refresh();
  });
  $('colTable').addEventListener('change', function (e) {
    var t = e.target, c = +t.dataset.c;
    if (t.dataset.act === 'prefix') { state.cfg.colPrefix[c] = Core.sanitizePrefix(t.value, state.hdr[c]); refresh(); }
    else if (t.dataset.act === 'cross') { state.cfg.colCross[c] = t.checked; refresh(); }
  });

  function setColMode(cols, mode) {
    var cfg = state.cfg;
    cols.forEach(function (c) {
      cfg.colMode[c] = mode;
      Array.from(cfg.cellOv.keys()).forEach(function (k) { if (+k.split(',')[1] === c) cfg.cellOv.delete(k); });   // 열 단위 설정이 칸 단위 설정을 덮어씀
    });
  }

  /* ---- 원본 / 미리보기 표 */
  var gridEl = $('gridWrap'), prevEl = $('prevWrap');

  function renderGrid() {
    var w = state.work, res = state.res, cfg = state.cfg, nC = w[0].length, L = Math.min(w.length, state.viewRows + 1);
    var st = gridEl.scrollTop, sl = gridEl.scrollLeft;
    var h = '<table class="grid"><thead><tr><th class="rn"></th>';
    for (var c = 0; c < nC; c++) {
      var m = cfg.colMode[c];
      h += '<th data-c="' + c + '" class="m-' + m + '">' + esc(state.hdr[c]) + '<span class="chm">' + MODE_LABEL[m] + '</span>' + (state.rec[c] && m === 'keep' ? '<span class="chr">★추천</span>' : '') + '</th>';
    }
    h += '</tr></thead><tbody>';
    for (var r = 1; r < L; r++) {
      h += '<tr><th class="rn" data-r="' + r + '">' + r + '</th>';
      for (c = 0; c < nC; c++) {
        var k = res.kind[r * nC + c], cls = (k & Core.K_DEL) ? 'del' : (k & Core.K_TOKEN) ? 'tok' : (k & Core.K_CROSS) ? 'crs' : '';
        if (k & Core.W_ANY) cls += ' wrn';
        h += '<td data-r="' + r + '" data-c="' + c + '"' + (cls ? ' class="' + cls.trim() + '"' : '') + '>' + esc(w[r][c]) + '</td>';
      }
      h += '</tr>';
    }
    gridEl.innerHTML = h + '</tbody></table>';
    gridEl.scrollTop = st; gridEl.scrollLeft = sl;
    var shown = L - 1, total = w.length - 1;
    $('rowInfo').textContent = '데이터 ' + total + '행 중 ' + shown + '행을 보여줘요 (선택·처리는 전체 행에 적용돼요)';
    $('moreRows').hidden = shown >= total;
    applySel();
  }

  function renderPreview() {
    var res = state.res, nC = res.nC, keep = res.keepCols, L = Math.min(res.nR, state.viewRows + 1);
    var st = prevEl.scrollTop, sl = prevEl.scrollLeft;
    var h = '<table class="grid"><thead><tr><th class="rn"></th>' + keep.map(function (c) { return '<th>' + esc(state.hdr[c]) + '</th>'; }).join('') + '</tr></thead><tbody>';
    for (var r = 1; r < L; r++) {
      h += '<tr><th class="rn">' + r + '</th>';
      for (var j = 0; j < keep.length; j++) {
        var k = res.kind[r * nC + keep[j]], cls = (k & Core.K_TOKEN) ? 'tok' : (k & Core.K_CROSS) ? 'crs' : '';
        if (k & Core.W_ANY) cls += ' wrn';
        h += '<td' + (cls ? ' class="' + cls.trim() + '"' : '') + '>' + esc(res.out[r][j]) + '</td>';
      }
      h += '</tr>';
    }
    prevEl.innerHTML = keep.length ? h + '</tbody></table>' : '<p class="muted pad12">내보낼 열이 없어요.</p>';
    prevEl.scrollTop = st; prevEl.scrollLeft = sl;
  }
  $('moreRows').addEventListener('click', function () { state.viewRows += 200; renderGrid(); renderPreview(); });

  // 세로 스크롤 동기화
  var syncing = false;
  function syncScroll(from, to) { from.addEventListener('scroll', function () { if (syncing) return; syncing = true; to.scrollTop = from.scrollTop; syncing = false; }); }
  syncScroll(gridEl, prevEl); syncScroll(prevEl, gridEl);

  /* ---- 선택 (열 머리글 / 행 번호 / 칸 드래그) */
  var drag = null;
  function inSel(r, c) {
    var s = state.sel;
    if (s.type === 'col') return s.cols.has(c);
    if (s.type === 'row') return s.rows.has(r);
    if (s.type === 'cell') return r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2;
    return false;
  }
  function applySel() {
    var s = state.sel;
    gridEl.querySelectorAll('td[data-r]').forEach(function (td) { td.classList.toggle('sel', inSel(+td.dataset.r, +td.dataset.c)); });
    gridEl.querySelectorAll('thead th[data-c]').forEach(function (th) { th.classList.toggle('sel', s.type === 'col' && s.cols.has(+th.dataset.c)); });
    gridEl.querySelectorAll('tbody th[data-r]').forEach(function (th) { th.classList.toggle('sel', s.type === 'row' && s.rows.has(+th.dataset.r)); });
  }
  function pickMulti(type, key, idx, e) {
    var s = state.sel, set = key === 'cols' ? 'cols' : 'rows';
    if (e.shiftKey && s.type === type && s.anchor != null) {
      var a = Math.min(s.anchor, idx), b = Math.max(s.anchor, idx); s[set] = new Set(); for (var i = a; i <= b; i++) s[set].add(i);
    } else if (e.ctrlKey || e.metaKey) {
      if (s.type !== type) { s = state.sel = { type: type }; s[set] = new Set(); }
      if (s[set].has(idx)) s[set].delete(idx); else s[set].add(idx);
      s.anchor = idx; if (!s[set].size) state.sel = { type: null };
    } else { s = state.sel = { type: type, anchor: idx }; s[set] = new Set([idx]); }
  }
  gridEl.addEventListener('mousedown', function (e) {
    if (e.button !== 0) return;
    var t = e.target.closest('thead th[data-c], tbody th[data-r], td[data-r]'); if (!t) return;
    e.preventDefault();
    if (t.tagName === 'TH' && t.dataset.c !== undefined) pickMulti('col', 'cols', +t.dataset.c, e);
    else if (t.tagName === 'TH') pickMulti('row', 'rows', +t.dataset.r, e);
    else { var r = +t.dataset.r, c = +t.dataset.c; drag = { r: r, c: c }; state.sel = { type: 'cell', r1: r, r2: r, c1: c, c2: c }; }
    applySel(); renderSelBar();
  });
  gridEl.addEventListener('mouseover', function (e) {
    if (!drag) return;
    var t = e.target.closest('td[data-r]'); if (!t) return;
    var r = +t.dataset.r, c = +t.dataset.c;
    state.sel = { type: 'cell', r1: Math.min(drag.r, r), r2: Math.max(drag.r, r), c1: Math.min(drag.c, c), c2: Math.max(drag.c, c) };
    applySel(); renderSelBar();
  });
  document.addEventListener('mouseup', function () { drag = null; });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { if (!$('overlay').hidden) closeModal(); else if (state.sel.type) { state.sel = { type: null }; applySel(); renderSelBar(); } }
  });

  function selCells() {
    var s = state.sel, out = [], nC = state.work[0].length, r, c;
    if (s.type === 'row') s.rows.forEach(function (r0) { for (c = 0; c < nC; c++) out.push([r0, c]); });
    if (s.type === 'cell') for (r = s.r1; r <= s.r2; r++) for (c = s.c1; c <= s.c2; c++) out.push([r, c]);
    return out.filter(function (p) { return state.cfg.colMode[p[1]] !== 'delete'; });
  }

  function renderSelBar() {
    var s = state.sel, bar = $('selbar'), n = state.work.length - 1;
    if (!s.type) { bar.innerHTML = '<span class="muted">열 머리글 = 열 전체 · 행 번호 = 한 줄 · 칸을 끌면 범위 선택 (Ctrl/Shift 로 여러 개, Esc 로 해제)</span>'; return; }
    if (s.type === 'col') {
      var names = Array.from(s.cols).sort(function (a, b) { return a - b; }).map(function (c) { return state.hdr[c]; });
      bar.innerHTML = '<b>' + esc(names.join(', ')) + '</b> 열 전체(' + n + '칸) 선택됨 <span class="grow"></span>' +
        '<button class="btn sm primary" type="button" data-sa="col-token">토큰화</button><button class="btn sm" type="button" data-sa="col-keep">유지</button><button class="btn sm danger" type="button" data-sa="col-delete">열 삭제</button>';
    } else {
      var cells = selCells(), desc = s.type === 'row' ? Array.from(s.rows).sort(function (a, b) { return a - b; }).join(', ') + '번 행 선택됨' : '칸 범위 선택됨';
      bar.innerHTML = '<b>' + desc + '</b> (' + cells.length + '칸) <span class="grow"></span>' +
        '<button class="btn sm primary" type="button" data-sa="cell-token">이 칸들 토큰화</button><button class="btn sm" type="button" data-sa="cell-keep">이 칸들 유지</button><button class="btn sm" type="button" data-sa="cell-clear">열 설정 따르기</button>';
    }
  }
  $('selbar').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-sa]'); if (!b) return;
    var a = b.dataset.sa, s = state.sel, cfg = state.cfg;
    if (a.indexOf('col-') === 0) setColMode(Array.from(s.cols), a.slice(4));
    else {
      selCells().forEach(function (p) {
        var key = p[0] + ',' + p[1];
        if (a === 'cell-clear') cfg.cellOv.delete(key);
        else { var m = a === 'cell-token' ? 'token' : 'keep'; if (m === cfg.colMode[p[1]]) cfg.cellOv.delete(key); else cfg.cellOv.set(key, m); }
      });
    }
    refresh();
  });

  /* ---- 교차 치환 */
  $('manualWords').addEventListener('change', function () {
    state.cfg.manualWords = this.value.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean); refresh();
  });
  $('manualPrefix').addEventListener('change', function () { this.value = state.cfg.manualPrefix = Core.sanitizePrefix(this.value, '기타'); refresh(); });

  function renderCross() {
    var res = state.res, cands = res.candidates, tot = cands.reduce(function (a, s) { return a + s.count; }, 0), b = $('crossStatus');
    if (!cands.length) { b.innerHTML = '<p class="muted">문장 속에서 추가로 바뀌는 곳이 없어요. (열 표의 “문장 속에서도 찾기”를 켜면 찾아요)</p>'; return; }
    b.innerHTML = '<div class="row"><span>다른 열의 문장 속에서 <b>' + tot + '곳</b>(값 ' + cands.length + '개)이 함께 바뀌어요.</span> ' +
      (res.needConfirm ? '<span class="tag need">확인 필요</span>' : '<span class="tag done">확인함 ✓</span>') +
      '<button class="btn sm ' + (res.needConfirm ? 'primary' : '') + '" type="button" id="btnCross">바뀔 목록 보기 · 확인</button></div>';
    $('btnCross').addEventListener('click', openCrossDialog);
  }

  function openCrossDialog() {
    var res = state.res, cfg = state.cfg, cands = res.candidates;
    if (!cands.length) return;
    var h = '<h2>문장 속에서 함께 바뀔 목록</h2><p class="muted">긴 값부터 찾아서 바꿔요. 바꾸면 안 되는 것은 체크를 풀어 주세요. 짧은 값은 다른 낱말의 일부와 겹칠 수 있어요.</p>' +
      '<table class="xtable"><thead><tr><th>바꿈</th><th>원래 값</th><th>토큰</th><th>횟수</th><th>들어 있는 열 · 문장 예</th></tr></thead><tbody>';
    cands.forEach(function (s, i) {
      var cols = Array.from(s.cols).map(function (c) { return esc(state.hdr[c]); }).join(', ');
      var smp = s.samples[0] ? esc(s.samples[0].b) + '<mark class="h">' + esc(s.samples[0].v) + '</mark>' + esc(s.samples[0].a) : '';
      h += '<tr><td><input type="checkbox" data-i="' + i + '"' + (cfg.crossExcluded.has(s.value) ? '' : ' checked') + '></td><td>' + esc(s.value) + (s.value.length <= 2 ? ' <span class="badge warn">짧음</span>' : '') + (s.manual ? ' <span class="badge rec">직접 추가</span>' : '') +
        '</td><td class="mono">' + (s.token ? esc(s.token) : '<span class="muted">(확인 후 부여)</span>') + '</td><td>' + s.count + '</td><td>' + cols + '<div class="muted small">…' + smp + '…</div></td></tr>';
    });
    h += '</tbody></table><div class="btns"><button class="btn" type="button" id="xCancel">닫기</button><button class="btn primary" type="button" id="xOk">이 목록으로 확인</button></div>';
    openModal(h, function (m) {
      m.querySelector('#xCancel').addEventListener('click', closeModal);
      m.querySelector('#xOk').addEventListener('click', function () {
        cands.forEach(function (s, i) { if (m.querySelector('input[data-i="' + i + '"]').checked) cfg.crossExcluded.delete(s.value); else cfg.crossExcluded.add(s.value); });
        cfg.crossConfirmedSig = res.crossSig;
        closeModal(); refresh();
      });
    });
  }

  /* ---- 경고 */
  function renderWarn() {
    var res = state.res, cfg = state.cfg, items = [], c, nC = res.nC;
    if (res.collisions.length) items.push(['bad', '원본에 <code>[[...]]</code> 모양의 글자가 ' + res.collisions.length + '칸에 이미 있어요 (예: ' + esc(res.collisions[0].text.slice(0, 40)) + '). 복원할 때 토큰과 섞일 수 있으니 원본을 고치는 것을 권해요.']);
    var rrn = 0, ph = 0, ml = 0;
    for (c = 0; c < nC; c++) if (cfg.colMode[c] !== 'delete') { rrn += res.colWarn[c].rrn; ph += res.colWarn[c].phone; ml += res.colWarn[c].mail; }
    if (rrn) items.push(['bad', '결과 표에 <b>주민등록번호 모양</b>이 ' + rrn + '칸 남아 있어요. 해당 열을 토큰화하거나 삭제하세요. (표에서 노란 테두리)']);
    if (ph) items.push(['warn', '결과 표에 <b>전화번호 모양</b>이 ' + ph + '칸 남아 있어요. (노란 테두리)']);
    if (ml) items.push(['warn', '결과 표에 <b>이메일 모양</b>이 ' + ml + '칸 남아 있어요. (노란 테두리)']);
    var keepRec = [];
    for (c = 0; c < nC; c++) if (state.rec[c] && cfg.colMode[c] === 'keep') keepRec.push(state.hdr[c]);
    if (keepRec.length) items.push(['warn', '개인정보로 보이는 열이 <b>유지</b> 상태예요: ' + keepRec.map(esc).join(', ')]);
    if (!res.entries.length) items.push(['warn', '아직 토큰화한 칸이 없어요.']);
    items.push(['info', '<b>남는 열의 조합</b>(반·성별·점수·비고 등)만으로도 누구인지 짐작될 수 있어요. 필요 없는 열은 삭제해 주세요. 비고 속 이름은 <b>성까지 똑같이 쓴 경우만</b> 찾아요. 이름만 쓴 경우는 “직접 추가할 단어”에 적어 주세요.']);
    $('warnBody').innerHTML = items.map(function (i) { return '<div class="banner ' + i[0] + '">' + i[1] + '</div>'; }).join('');
  }

  /* ---- 내보내기 */
  function ensureKeyId() {
    var sig = state.res.entSig;
    if (!state.res.entries.length) return null;
    if (state.keyIdSig !== sig) { state.keyId = Core.newKeyId(); state.keyIdSig = sig; }
    return state.keyId;
  }

  function renderExport() {
    var res = state.res, v = Core.verify(state.work, res), h = '';
    h += '<p>내보낼 표: <b>' + (res.nR - 1) + '행 × ' + res.keepCols.length + '열</b> · 토큰 <b>' + res.entries.length + '개</b>' +
      (res.prefixCounts.length ? ' (' + res.prefixCounts.map(function (p) { return esc(p[0]) + ' ' + p[1]; }).join(', ') + ')' : '') + '</p>';
    if (res.needConfirm) h += '<div class="banner warn">문장 속에서 함께 바뀔 목록을 먼저 확인해야 내보낼 수 있어요. (위 3번의 “바뀔 목록 보기 · 확인”)</div>';
    h += v.bad === 0 ? '<div class="banner good small">✔ 자체 검증: 키로 되돌리면 원본과 <b>100% 일치</b>해요 (' + v.total + '칸 검사).</div>'
      : '<div class="banner bad small">⚠ 자체 검증: ' + v.bad + '칸이 원본과 달라져요. 원본에 <code>[[...]]</code> 모양의 글자가 있는지 확인하세요.</div>';
    $('exportStatus').innerHTML = h;
    $('instr').textContent = Core.buildInstruction(res);
    var dis = res.needConfirm || !res.keepCols.length;
    ['btnXlsx', 'btnCsv', 'btnCopyAll', 'btnCopyTsv', 'btnCopyMd', 'btnCopyInstr'].forEach(function (id) { $(id).disabled = dis; });
  }

  // 내보내기 직전 경고: 설정이 바뀔 때마다 한 번 확인받는다.
  function guardExport(action) {
    var res = state.res, cfg = state.cfg;
    if (res.needConfirm) { openCrossDialog(); return; }
    if (cfg.ackSig === cfgSig()) { action(); return; }
    var kept = res.keepCols.filter(function (c) { return cfg.colMode[c] !== 'token'; }).map(function (c) { return state.hdr[c]; });
    openModal('<h2>내보내기 전에 확인하세요</h2>' +
      (res.entries.length ? '' : '<div class="banner bad">토큰화한 칸이 없어요. 이대로 내보내면 원본과 같아요.</div>') +
      '<p>토큰화하지 않은 열의 <b>조합</b>으로 개인이 특정될 수 있어요.</p>' +
      '<p>지금 그대로 나가는 열: <b>' + (kept.length ? kept.map(esc).join(', ') : '없음') + '</b></p>' +
      '<ul class="clean"><li>반·성별·점수·특이사항처럼 남는 값이 적은 사람만 해당될 때 알아볼 수 있어요.</li><li>필요 없는 열은 삭제하고, 비고에 이름 외의 식별 정보(학교·가족 이야기 등)가 없는지 보세요.</li><li>AI 서비스 약관과 학교의 개인정보 지침을 확인하세요.</li></ul>' +
      '<div class="btns"><button class="btn" type="button" id="gCancel">돌아가기</button><button class="btn primary" type="button" id="gOk">확인했어요. 내보내기</button></div>',
      function (m) {
        m.querySelector('#gCancel').addEventListener('click', closeModal);
        m.querySelector('#gOk').addEventListener('click', function () { cfg.ackSig = cfgSig(); closeModal(); action(); });
      });
  }

  function baseName() { var id = ensureKeyId(); return '가명처리표' + (id ? '_' + id : '') + '_' + today(); }

  $('btnXlsx').addEventListener('click', function () {
    guardExport(function () {
      var u8 = new Uint8Array(Core.makeXlsxBlobParts(state.res.out));
      downloadBlob(new Blob([u8], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), baseName() + '.xlsx');
      toast('엑셀 파일을 저장했어요.');
    });
  });
  $('btnCsv').addEventListener('click', function () {
    guardExport(function () { downloadText(Core.toCSV(state.res.out), baseName() + '.csv', 'text/csv', true); toast('CSV 파일을 저장했어요.'); });
  });
  $('btnCopyTsv').addEventListener('click', function () { guardExport(function () { copyText(Core.toTSV(state.res.out), '표를 복사했어요. AI 채팅창에 붙여넣으세요.'); }); });
  $('btnCopyMd').addEventListener('click', function () { guardExport(function () { copyText(Core.toMarkdown(state.res.out), '마크다운 표를 복사했어요.'); }); });
  $('btnCopyAll').addEventListener('click', function () {
    guardExport(function () { copyText(Core.buildInstruction(state.res) + '\n\n' + Core.toTSV(state.res.out), '지시문과 표를 복사했어요. AI 채팅창에 붙여넣으세요.'); });
  });
  $('btnCopyInstr').addEventListener('click', function () { copyText(Core.buildInstruction(state.res), '지시문을 복사했어요.'); });

  $('pwShow').addEventListener('change', function () { $('pw1').type = $('pw2').type = this.checked ? 'text' : 'password'; });

  $('btnKey').addEventListener('click', async function () {
    var res = state.res, msg = $('keyMsg'), p1 = $('pw1').value, p2 = $('pw2').value;
    if (!res.entries.length) { msg.innerHTML = '<div class="banner warn">저장할 대응표가 없어요. 먼저 토큰화할 열을 고르세요.</div>'; return; }
    if (res.needConfirm) { openCrossDialog(); return; }
    if (p1.length < 8) { msg.innerHTML = '<div class="banner bad">비밀번호는 8자 이상으로 정해 주세요.</div>'; return; }
    if (p1 !== p2) { msg.innerHTML = '<div class="banner bad">두 비밀번호가 서로 달라요.</div>'; return; }
    if (!(window.crypto && crypto.subtle)) { msg.innerHTML = '<div class="banner bad">이 브라우저에서는 암호화 기능을 쓸 수 없어요. 최신 Chrome·Edge·Safari·Firefox 로 파일을 직접 열어 주세요.</div>'; return; }
    var btn = this; btn.disabled = true; msg.innerHTML = '<div class="banner info">암호화하는 중… (몇 초 걸려요)</div>';
    try {
      var id = ensureKeyId(), obj = await Core.encryptKey(res.entries, p1, id);
      var back = await Core.decryptKey(obj, p1);          // 저장 전에 열리는지 검증
      if (JSON.stringify(back.entries) !== JSON.stringify(res.entries)) throw new Error('검증 실패');
      downloadText(JSON.stringify(obj, null, 2), '키파일_' + id + '_' + today() + '.key.json', 'application/json', false);
      msg.innerHTML = '<div class="banner good"><b>키 파일을 저장했어요.</b> (키 ID <b>' + id + '</b> · 토큰 ' + res.entries.length + '개) 가명처리한 표를 내보낼 때와 <u>설정을 바꿨다면 키 파일도 다시 저장</u>해야 해요. 비밀번호는 키 파일과 따로 보관하세요.</div>';
      $('pw1').value = $('pw2').value = '';
    } catch (e) { msg.innerHTML = '<div class="banner bad">암호화에 실패했어요: ' + esc(e.message) + '</div>'; }
    btn.disabled = false;
  });

  /* ------------------------------------------------------------ ③ 복원 */
  makeDrop($('dropKey'), $('fileKey'), $('pickKey'), async function (file) {
    try {
      var obj = JSON.parse(await file.text());
      if (!obj || obj.format !== 'pseudonym-key') throw new Error('x');
      state.rs.pendingKey = obj; state.rs.pendingKeyName = file.name;
      $('keyFileInfo').innerHTML = '선택한 키 파일: <b>' + esc(file.name) + '</b> (키 ID ' + esc(String(obj.keyId)) + ')';
      $('openMsg').innerHTML = ''; $('pwOpen').focus();
    } catch (e) { state.rs.pendingKey = null; $('keyFileInfo').innerHTML = ''; $('openMsg').innerHTML = '<div class="banner bad">키 파일 형식이 아니에요. 가명처리 때 저장한 <code>.key.json</code> 파일을 선택하세요.</div>'; }
  });

  $('btnOpenKey').addEventListener('click', async function () {
    var rs = state.rs, msg = $('openMsg'), pw = $('pwOpen').value;
    if (!rs.pendingKey) { msg.innerHTML = '<div class="banner bad">먼저 키 파일을 선택하세요.</div>'; return; }
    if (!pw) { msg.innerHTML = '<div class="banner bad">비밀번호를 입력하세요.</div>'; return; }
    var btn = this; btn.disabled = true; msg.innerHTML = '<div class="banner info">키를 여는 중… (몇 초 걸려요)</div>';
    try {
      var r = await Core.decryptKey(rs.pendingKey, pw);
      rs.entries = r.entries; rs.keyId = r.keyId; rs.restorer = Core.makeRestorer(r.entries);
      var counts = {}; r.entries.forEach(function (e) { var p = e[0].slice(2, -2); p = p.slice(0, p.lastIndexOf('_')); counts[p] = (counts[p] || 0) + 1; });
      msg.innerHTML = '<div class="banner good"><b>키를 열었어요.</b> 키 ID <b>' + esc(r.keyId) + '</b> · 토큰 ' + r.entries.length + '개 (' + Object.keys(counts).map(function (p) { return esc(p) + ' ' + counts[p]; }).join(', ') + ')<br><span class="small">가명처리할 때 본 키 ID 와 같은지 확인하세요. 다른 키로 복원하면 엉뚱한 값이 들어갈 수 있어요.</span></div>';
      $('pwOpen').value = '';
    } catch (e) {
      rs.entries = null; rs.restorer = null;
      msg.innerHTML = '<div class="banner bad"><b>키를 열지 못했어요.</b> ' + esc(e.message) + '</div>';
    }
    btn.disabled = false;
  });

  makeDrop($('dropAi'), $('fileAi'), $('pickAi'), async function (file) {
    try {
      var ext = (file.name.split('.').pop() || '').toLowerCase(), buf = await file.arrayBuffer(), inp = { name: file.name, ext: ext };
      if (ext === 'xlsx' || ext === 'xls') { inp.kind = 'xlsx'; inp.buf = buf; Core.readWorkbook(buf); }
      else if (ext === 'csv' || ext === 'tsv') { var d = Core.decodeBytes(buf); inp.kind = 'csv'; inp.grid = Core.normalizeGrid(Core.parseDelimited(d.text, Core.detectDelimiter(d.text))); }
      else if (ext === 'txt' || ext === 'md' || ext === 'markdown') { inp.kind = 'text'; $('aiText').value = Core.decodeBytes(buf).text; state.rs.input = { kind: 'text', name: file.name, ext: ext === 'md' || ext === 'markdown' ? 'md' : 'txt' }; showAiInfo(); return; }
      else throw new Error('.xlsx .csv .txt .md 파일만 넣을 수 있어요.');
      $('aiText').value = ''; state.rs.input = inp; showAiInfo();
    } catch (e) { toast('파일을 읽지 못했어요: ' + e.message, true); }
  });
  $('aiText').addEventListener('input', function () {
    var i = state.rs.input;
    if (!i || i.kind !== 'text') state.rs.input = { kind: 'text', name: '', ext: 'txt' };   // 붙여넣기가 파일 입력을 대체
    showAiInfo();
  });
  function showAiInfo() {
    var i = state.rs.input, b = $('aiFileInfo');
    b.innerHTML = i && i.name ? '입력 파일: <b>' + esc(i.name) + '</b> (' + (i.kind === 'xlsx' ? '엑셀' : i.kind === 'csv' ? 'CSV' : '텍스트') + ') <button class="btn sm" type="button" id="aiClear">지우기</button>' : '';
    var cl = $('aiClear'); if (cl) cl.addEventListener('click', function () { state.rs.input = null; $('aiText').value = ''; showAiInfo(); });
  }

  $('btnRestore').addEventListener('click', function () {
    var rs = state.rs;
    if (!rs.restorer) { toast('먼저 키 파일을 열어 주세요.', true); return; }
    var inp = rs.input || (($('aiText').value) ? { kind: 'text', name: '', ext: 'txt' } : null);
    if (!inp) { toast('복원할 AI 결과를 붙여넣거나 파일로 넣어 주세요.', true); return; }
    var fix = $('fixVariants').checked, rep = Core.newReport(), res = { kind: inp.kind, ext: inp.ext, rep: rep };
    try {
      if (inp.kind === 'text') { res.parts = []; res.text = rs.restorer.restore($('aiText').value, rep, res.parts, fix); }
      else if (inp.kind === 'csv') { var r = Core.restoreGrid(inp.grid, rs.restorer, rep, fix); res.grid = r.grid; res.changed = r.changed; }
      else { res.wb = Core.readWorkbook(inp.buf); res.changedSet = Core.restoreWorkbook(res.wb, rs.restorer, rep, fix); res.sheet = res.wb.SheetNames[0]; }
    } catch (e) { toast('복원 중 오류: ' + e.message, true); return; }
    rs.result = res; renderResult();
  });

  function renderResult() {
    var res = state.rs.result, rep = res.rep, h = '', box = $('resultCard'); box.hidden = false;
    h += '<div class="banner ' + (rep.unknown.size || rep.variants.size ? 'warn' : 'good') + '"><b>복원한 토큰 ' + rep.replaced + '개</b> (종류 ' + rep.restoredSet.size + '가지)' +
      (rep.fixedVariants ? ' · 변형 복원 ' + rep.fixedVariants + '개 포함' : '') + '</div>';
    if (rep.unknown.size) {
      h += '<div class="banner bad"><b>키에 없는 토큰 ' + rep.unknown.size + '종</b> — 복원하지 못하고 그대로 두었어요. 다른 키 파일이거나 AI가 만들어 낸 토큰일 수 있어요.<ul>' +
        Array.from(rep.unknown.entries()).map(function (e) { return '<li class="mono">' + esc(e[0]) + ' × ' + e[1] + '</li>'; }).join('') + '</ul></div>';
    }
    if (rep.variants.size) {
      h += '<div class="banner warn"><b>변형 의심 토큰 ' + rep.variants.size + '종</b> — AI가 모양을 바꿨을 수 있어요.<ul>' +
        Array.from(rep.variants.entries()).map(function (e) {
          var v = e[1]; return '<li><span class="mono">' + esc(e[0]) + '</span> × ' + v.count + ' → ' + (v.fixed ? '<b>' + esc(v.suggest) + ' 로 복원했어요</b>' : v.suggest ? esc(v.suggest) + ' 로 보여요 (위 체크박스를 켜고 다시 복원하면 함께 바뀌어요)' : '키에 대응되는 토큰이 없어요');
        }).join('') + '</ul></div>';
    }
    $('report').innerHTML = h;
    var body = $('resultBody'), btns = $('resultBtns'); btns.innerHTML = '';
    function addBtn(label, fn, primary) { var b = document.createElement('button'); b.type = 'button'; b.className = 'btn' + (primary ? ' primary' : ''); b.textContent = label; b.addEventListener('click', fn); btns.appendChild(b); }
    if (res.kind === 'text') {
      body.innerHTML = '<pre class="out">' + res.parts.map(function (p) { return p.k === 't' ? esc(p.s) : '<span class="' + p.k + '">' + esc(p.s) + '</span>'; }).join('') + '</pre>' +
        '<div class="legend"><span><i class="lg-ok"></i>복원됨</span><span><i class="lg-unk"></i>키에 없는 토큰</span><span><i class="lg-var"></i>변형 의심</span></div>';
      addBtn('복사', function () { copyText(res.text, '복원 결과를 복사했어요.'); }, true);
      addBtn('저장 (.' + res.ext + ')', function () { downloadText(res.text, '복원결과_' + today() + '.' + res.ext, 'text/plain', false); toast('저장했어요.'); });
    } else if (res.kind === 'csv') {
      body.innerHTML = tablePreview(res.grid, function (r, c) { return res.changed.has(r + ',' + c); });
      addBtn('표 복사 (TSV)', function () { copyText(Core.toTSV(res.grid), '복원한 표를 복사했어요.'); }, true);
      addBtn('CSV로 저장', function () { downloadText(Core.toCSV(res.grid), '복원결과_' + today() + '.csv', 'text/csv', true); toast('저장했어요.'); });
    } else {
      var names = res.wb.SheetNames, g = Core.sheetToGrid(res.wb.Sheets[res.sheet]).grid;
      body.innerHTML = (names.length > 1 ? '<p>시트: <select id="rsSheet">' + names.map(function (n) { return '<option' + (n === res.sheet ? ' selected' : '') + ' value="' + esc(n) + '">' + esc(n) + '</option>'; }).join('') + '</select></p>' : '') +
        tablePreview(g, function (r, c) { return res.changedSet.has(res.sheet + '!' + XLSX.utils.encode_cell({ r: r, c: c })); }, true);
      var sel = $('rsSheet'); if (sel) sel.addEventListener('change', function () { res.sheet = sel.value; renderResult(); });
      addBtn('현재 시트 복사 (TSV)', function () { copyText(Core.toTSV(g), '복원한 표를 복사했어요.'); }, true);
      addBtn('엑셀(.xlsx)로 저장', function () {
        downloadBlob(new Blob([new Uint8Array(XLSX.write(res.wb, { bookType: 'xlsx', type: 'array' }))], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), '복원결과_' + today() + '.xlsx'); toast('저장했어요.');
      });
    }
  }
  // xlsx 미리보기는 시트 좌표 오프셋이 달라 변경 칸 강조를 생략한다(복원 수는 보고서에 표시).
  function tablePreview(g, isChanged, xlsx) {
    var L = Math.min(g.length, MAX_PREVIEW), nC = g[0] ? g[0].length : 0, h = '<div class="gridwrap"><table class="grid"><tbody>';
    for (var r = 0; r < L; r++) {
      h += '<tr><th class="rn">' + (r + 1) + '</th>';
      for (var c = 0; c < nC; c++) h += '<td' + (!xlsx && isChanged(r, c) ? ' class="tok"' : '') + '>' + esc(g[r][c]) + '</td>';
      h += '</tr>';
    }
    return h + '</tbody></table></div>' + (g.length > L ? '<p class="muted small">앞의 ' + L + '행만 보여줘요. 복사·저장에는 전체가 들어가요.</p>' : '') + (xlsx ? '' : '<p class="muted small">파란 칸 = 복원된 칸</p>');
  }

  /* ------------------------------------------------------------ 모두 지우기 */
  $('btnClear').addEventListener('click', function () {
    openModal('<h2>모두 지울까요?</h2><p>불러온 파일, 대응표, 복원한 내용, 입력한 비밀번호를 이 창에서 모두 지워요. (저장한 파일은 그대로예요.)</p><p class="muted small">복사해 둔 내용(클립보드)은 이 앱이 지울 수 없어요.</p>' +
      '<div class="btns"><button class="btn" type="button" id="cNo">취소</button><button class="btn danger" type="button" id="cYes">모두 지우기</button></div>', function (m) {
        m.querySelector('#cNo').addEventListener('click', closeModal);
        m.querySelector('#cYes').addEventListener('click', function () { closeModal(); clearAll(); toast('모두 지웠어요.'); });
      });
  });
  function clearAll() {
    state = freshState();                 // 데이터를 가진 모든 참조를 끊어 가비지 컬렉션 대상으로 만든다
    document.querySelectorAll('input[type=password], input[type=file], textarea').forEach(function (el) { el.value = ''; });
    $('manualPrefix').value = '기타'; $('fixVariants').checked = false; $('pwShow').checked = false; $('pw1').type = $('pw2').type = 'password';
    ['loaded', 'recBanner', 'colTable', 'gridWrap', 'prevWrap', 'selbar', 'crossStatus', 'warnBody', 'exportStatus', 'instr', 'keyMsg', 'keyFileInfo', 'openMsg', 'aiFileInfo', 'report', 'resultBody', 'resultBtns'].forEach(function (id) { $(id).innerHTML = ''; });
    $('loaded').hidden = true; $('resultCard').hidden = true;
    updateTabs(); gotoTab(1);
  }

  // 작업 내용은 메모리에만 있으므로, 창을 닫기 전에 한 번 알려 준다.
  window.addEventListener('beforeunload', function (e) { if (state.file || state.rs.entries) { e.preventDefault(); e.returnValue = ''; } });

  /* ------------------------------------------------------------ 시작 */
  if (!(window.crypto && crypto.subtle)) toast('이 브라우저에서는 암호화(키 파일) 기능을 쓸 수 없어요. 최신 브라우저를 사용해 주세요.', true);
  $('bootfail').hidden = true; $('app').hidden = false;
  showAiInfo();
})();
