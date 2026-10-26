'use strict';
/*
 * Core — 화면(DOM)과 무관한 순수 로직.
 *  - 파일 읽기(표시값 그대로) / 토큰화 / 교차 치환 / 키 암호화·복호화 / 복원
 *  - 네트워크·저장소(localStorage·IndexedDB·쿠키)를 쓰는 코드는 이 파일 어디에도 없다.
 *  - 모든 데이터는 호출자가 들고 있는 메모리 객체에만 존재한다.
 */
var Core = (function () {

  var K_KEEP = 1, K_TOKEN = 2, K_CROSS = 4, K_DEL = 8, W_PHONE = 16, W_RRN = 32, W_MAIL = 64;
  var W_ANY = W_PHONE | W_RRN | W_MAIL;

  // 토큰 형식: [[접두어_001]] (일련번호 3자리 이상, 0 채움)
  var TOKEN_EXACT_SRC = '\\[\\[([^\\[\\]\\n]+)_(\\d{3,})\\]\\]';
  var BRACKET_LIKE_RE = /\[\[[^\[\]\n]*\]\]/;

  function X() { return (typeof XLSX !== 'undefined') ? XLSX : globalThis.XLSX; }

  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function pad3(n) { return String(n).padStart(3, '0'); }

  function colLetter(i) {
    var s = ''; i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }

  // 접두어에서 토큰 구조를 깨는 문자([ ] 줄바꿈)를 없애고 공백은 _ 로 바꾼다.
  function sanitizePrefix(s, fallback) {
    var t = String(s == null ? '' : s).replace(/[\[\]\r\n\t]/g, ' ').trim().replace(/\s+/g, '_');
    return t || fallback || '값';
  }

  /* ---------------------------------------------------------------- 파일 읽기 */

  // 인코딩 판별: BOM → UTF-8(엄격) → 실패하면 한국어 엑셀 CSV에 흔한 EUC-KR(CP949)
  function decodeBytes(buf) {
    var u8 = new Uint8Array(buf);
    if (u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) return { text: new TextDecoder('utf-8').decode(u8.subarray(3)), enc: 'UTF-8' };
    if (u8[0] === 0xFF && u8[1] === 0xFE) return { text: new TextDecoder('utf-16le').decode(u8.subarray(2)), enc: 'UTF-16' };
    if (u8[0] === 0xFE && u8[1] === 0xFF) return { text: new TextDecoder('utf-16be').decode(u8.subarray(2)), enc: 'UTF-16' };
    try { return { text: new TextDecoder('utf-8', { fatal: true }).decode(u8), enc: 'UTF-8' }; }
    catch (e) { return { text: new TextDecoder('euc-kr').decode(u8), enc: 'EUC-KR(CP949)' }; }
  }

  function detectDelimiter(text) {
    var line = text.slice(0, 4000).split(/\r?\n/)[0] || '';
    var c = (line.match(/,/g) || []).length, t = (line.match(/\t/g) || []).length, s = (line.match(/;/g) || []).length;
    if (t > c && t >= s) return '\t';
    if (s > c) return ';';
    return ',';
  }

  // CSV를 직접 파싱한다. SheetJS의 CSV 파서는 "0101" 을 숫자 101 로 바꿔 버리므로 쓰지 않는다.
  function parseDelimited(text, delim) {
    var rows = [], row = [], f = '', i = 0, n = text.length, q = false;
    while (i < n) {
      var ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i += 2; continue; } q = false; i++; continue; }
        f += ch; i++; continue;
      }
      if (ch === '"' && f === '') { q = true; i++; continue; }
      if (ch === delim) { row.push(f); f = ''; i++; continue; }
      if (ch === '\r') { if (text[i + 1] === '\n') i++; row.push(f); f = ''; rows.push(row); row = []; i++; continue; }
      if (ch === '\n') { row.push(f); f = ''; rows.push(row); row = []; i++; continue; }
      f += ch; i++;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }

  function isBlank(s) { return String(s).trim() === ''; }

  function normalizeGrid(rows) {
    var nC = 0, i, j;
    for (i = 0; i < rows.length; i++) nC = Math.max(nC, rows[i].length);
    var g = rows.map(function (r) { var o = new Array(nC); for (var c = 0; c < nC; c++) o[c] = r[c] == null ? '' : String(r[c]); return o; });
    var lastR = -1, lastC = -1;
    for (i = 0; i < g.length; i++) for (j = 0; j < nC; j++) if (!isBlank(g[i][j])) { if (i > lastR) lastR = i; if (j > lastC) lastC = j; }
    if (lastR < 0) return [['']];
    g = g.slice(0, lastR + 1).map(function (r) { return r.slice(0, lastC + 1); });
    return g;
  }

  function readWorkbook(buf) {
    // dateNF: 날짜 서식 14(지역별 서식)를 고정 형식으로 보여 줌. 셀의 표시 문자열(w)을 그대로 쓰기 위해 cellText 유지.
    return X().read(buf, { type: 'array', cellDates: false, dateNF: 'yyyy-mm-dd', cellText: true });
  }

  function cellStr(cell) {
    if (!cell) return '';
    if (cell.w != null) return String(cell.w);
    if (cell.v != null) return String(cell.v);
    return '';
  }

  // 시트 → 문자열 표. 셀에 "보이는 값"(w)을 그대로 쓰므로 학번 0101, 날짜 서식이 유지된다.
  // 병합 셀은 왼쪽 위 값을 병합 범위 전체에 복사해 행마다 값이 있도록 한다.
  function sheetToGrid(ws) {
    var U = X().utils;
    var keys = Object.keys(ws || {});
    var minR = Infinity, minC = Infinity, maxR = -1, maxC = -1, k, a;
    for (k = 0; k < keys.length; k++) {
      if (keys[k][0] === '!') continue;
      if (isBlank(cellStr(ws[keys[k]]))) continue;
      a = U.decode_cell(keys[k]);
      if (a.r < minR) minR = a.r; if (a.r > maxR) maxR = a.r;
      if (a.c < minC) minC = a.c; if (a.c > maxC) maxC = a.c;
    }
    if (maxR < 0) return { grid: [['']], merged: 0 };
    var nR = maxR - minR + 1, nC = maxC - minC + 1;
    if (nR * nC > 4000000) throw new Error('표가 너무 커요. 필요한 범위만 남겨 다시 저장해 주세요.');
    var g = [], r, c;
    for (r = 0; r < nR; r++) { var row = new Array(nC); for (c = 0; c < nC; c++) row[c] = cellStr(ws[U.encode_cell({ r: minR + r, c: minC + c })]); g.push(row); }
    var merged = 0;
    (ws['!merges'] || []).forEach(function (m) {
      var br = m.s.r - minR, bc = m.s.c - minC;
      if (br < 0 || bc < 0 || br >= nR || bc >= nC) return;
      var base = g[br][bc]; if (isBlank(base)) return;
      for (var rr = m.s.r; rr <= m.e.r; rr++) for (var cc = m.s.c; cc <= m.e.c; cc++) {
        var y = rr - minR, x = cc - minC;
        if (y < 0 || x < 0 || y >= nR || x >= nC || (y === br && x === bc)) continue;
        g[y][x] = base; merged++;
      }
    });
    return { grid: g, merged: merged };
  }

  function guessHeader(grid) {
    var lim = Math.min(grid.length, 10), cnt = [], max = 0, r;
    for (r = 0; r < lim; r++) { cnt[r] = grid[r].filter(function (v) { return !isBlank(v); }).length; if (cnt[r] > max) max = cnt[r]; }
    for (r = 0; r < lim; r++) if (cnt[r] >= Math.max(2, Math.ceil(max * 0.6))) return r;
    return 0;
  }

  // 원본에 이미 [[...]] 모양의 문자열이 있는 칸 (토큰과 헷갈릴 수 있음)
  function findBracketCells(grid) {
    var out = [];
    for (var r = 0; r < grid.length; r++) for (var c = 0; c < grid[r].length; c++)
      if (BRACKET_LIKE_RE.test(grid[r][c])) out.push({ r: r, c: c, text: grid[r][c] });
    return out;
  }

  /* ---------------------------------------------------------------- 자동 추천 / 패턴 경고 */

  var REC_RE = /이름|성명|학번|번호|연락처|전화|휴대|핸드폰|주소|생년월일|생일|주민|이메일|메일|보호자|학부모|phone|name|email|address/i;
  var CROSS_RE = /이름|성명|보호자|학부모|연락처|전화|휴대|핸드폰|주소|메일|생년월일|생일|주민|학번|name|email|address/i;
  function isRecommended(header) { return REC_RE.test(header); }
  function crossDefault(header) { return CROSS_RE.test(header); }

  var PHONE_RE = /(?<![0-9])(?:01[016789]|02|0[3-6][1-5]|070|050\d?|080)[-.\s]?\d{3,4}[-.\s]?\d{4}(?![0-9])/;
  var RRN_RE = /(?<![0-9])\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])[-\s]?[1-8]\d{6}(?![0-9])/;
  var MAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+/;

  /* ---------------------------------------------------------------- 토큰화 */

  function newConfig() {
    return {
      colMode: [], colPrefix: [], colCross: [],
      cellOv: new Map(),          // "r,c" → 'token' | 'keep' (열 설정보다 우선하는 칸 단위 설정)
      manualWords: [], manualPrefix: '기타',
      crossExcluded: new Set(),   // 교차 치환 확인창에서 제외한 값
      crossConfirmedSig: null,    // 사용자가 확인한 교차 치환 목록의 서명
      ackSig: null                // 내보내기 전 경고를 확인한 설정의 서명
    };
  }

  function effMode(cfg, r, c) {
    var m = cfg.colMode[c];
    if (m === 'delete') return 'delete';
    return cfg.cellOv.get(r + ',' + c) || m;
  }

  function crossEligible(v) { return v.length >= 2 && !(/^\d+$/.test(v) && v.length < 4); }

  var ALNUM = /[0-9A-Za-z]/;
  function boundaryOK(str, i, v) {
    if (ALNUM.test(v[0]) && i > 0 && ALNUM.test(str[i - 1])) return false;
    var e = i + v.length;
    if (ALNUM.test(v[v.length - 1]) && e < str.length && ALNUM.test(str[e])) return false;
    return true;
  }

  // 긴 값부터 일치시키는 치환. 첫 글자별 버킷(긴 값 우선)으로 수천 개 값도 빠르게 처리한다.
  function scanReplace(str, buckets, hit) {
    var out = '', last = 0, i = 0, n = str.length, any = false;
    while (i < n) {
      var b = buckets.get(str[i]), m = null;
      if (b) for (var k = 0; k < b.length; k++) {
        var v = b[k].value;
        if (str.startsWith(v, i) && boundaryOK(str, i, v)) { m = b[k]; break; }
      }
      if (m) { out += str.slice(last, i) + hit(m, i); i += m.value.length; last = i; any = true; }
      else i++;
    }
    return any ? out + str.slice(last) : str;
  }

  function fnv(str, h) { h = h || 2166136261; for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h; }

  /*
   * work: 문자열 표(0행 = 머리글). 반환:
   *  text/kind(칸별 결과와 상태 플래그), out(내보낼 표, 삭제 열 제외), entries([토큰, 원래값]),
   *  candidates(교차 치환 대상), needConfirm, colWarn, collisions
   */
  function compute(work, cfg) {
    var nR = work.length, nC = nR ? work[0].length : 0, r, c, v;
    var text = new Array(nR), kind = new Uint8Array(nR * nC);
    for (r = 0; r < nR; r++) text[r] = work[r].slice();

    var counters = new Map(), byKey = new Map(), entries = [], crossVals = new Map();
    function tokenFor(prefix, val) {
      var key = prefix + '\u0000' + val, t = byKey.get(key);       // 같은 접두어·같은 값 → 항상 같은 토큰
      if (!t) {
        var n = (counters.get(prefix) || 0) + 1; counters.set(prefix, n);
        t = '[[' + prefix + '_' + pad3(n) + ']]';
        byKey.set(key, t); entries.push([t, val]);
      }
      return t;
    }

    // 머리글 행
    for (c = 0; c < nC; c++) kind[c] = cfg.colMode[c] === 'delete' ? K_DEL : K_KEEP;

    // 1단계: 선택한 칸을 토큰으로
    for (r = 1; r < nR; r++) for (c = 0; c < nC; c++) {
      var idx = r * nC + c, mode = effMode(cfg, r, c);
      if (mode === 'delete') { kind[idx] = K_DEL; continue; }
      if (mode === 'token') {
        v = work[r][c];
        var core = v.trim();
        if (core === '') { kind[idx] = K_KEEP; continue; }
        var lead = v.slice(0, v.indexOf(core)), trail = v.slice(lead.length + core.length);   // 앞뒤 공백은 그대로 보존
        var t = tokenFor(cfg.colPrefix[c] || '값', core);
        text[r][c] = lead + t + trail; kind[idx] = K_TOKEN;
        if (cfg.colCross[c] && crossEligible(core) && !crossVals.has(core)) crossVals.set(core, t);
      } else kind[idx] = K_KEEP;
    }

    // 2단계: 교차 치환 — 토큰화한 값(및 직접 추가한 단어)이 '유지' 칸의 문장 속에 있으면 같은 토큰으로
    var sources = [];
    crossVals.forEach(function (tok, val) { sources.push({ value: val, token: tok, manual: false, count: 0, cols: new Set(), samples: [] }); });
    var seen = new Set(crossVals.keys());
    cfg.manualWords.forEach(function (w) {
      w = w.trim();
      if (w.length < 2 || seen.has(w)) return;
      seen.add(w); sources.push({ value: w, token: null, manual: true, count: 0, cols: new Set(), samples: [] });
    });
    sources.sort(function (a, b) { return b.value.length - a.value.length; });   // 긴 값 먼저
    var buckets = new Map();
    sources.forEach(function (s) { var ch = s.value[0], arr = buckets.get(ch); if (!arr) buckets.set(ch, arr = []); arr.push(s); });
    var manualPrefix = cfg.manualPrefix || '기타';

    if (sources.length) {
      for (r = 1; r < nR; r++) for (c = 0; c < nC; c++) {
        if (kind[r * nC + c] !== K_KEEP) continue;
        var s0 = text[r][c];
        if (!s0) continue;
        var accepted = false, cc = c;
        var res = scanReplace(s0, buckets, function (src, pos) {
          src.count++; src.cols.add(cc);
          if (src.samples.length < 3) src.samples.push({ b: s0.slice(Math.max(0, pos - 8), pos), v: src.value, a: s0.slice(pos + src.value.length, pos + src.value.length + 8) });
          if (cfg.crossExcluded.has(src.value)) return src.value;
          if (!src.token) src.token = tokenFor(manualPrefix, src.value);
          accepted = true; return src.token;
        });
        if (accepted) { text[r][c] = res; kind[r * nC + c] = K_CROSS; }
      }
    }
    var candidates = sources.filter(function (s) { return s.count > 0; });
    var crossSig = candidates.map(function (s) { return s.value + '\u0001' + s.count; }).join('\u0002');
    var needConfirm = candidates.length > 0 && cfg.crossConfirmedSig !== crossSig;

    // 3단계: 결과에 남은 전화번호·주민등록번호·이메일 모양
    var colWarn = [];
    for (c = 0; c < nC; c++) colWarn.push({ phone: 0, rrn: 0, mail: 0 });
    for (r = 1; r < nR; r++) for (c = 0; c < nC; c++) {
      var ix = r * nC + c;
      if (kind[ix] & K_DEL) continue;
      var tx = text[r][c];
      if (tx.length < 6) continue;
      if (RRN_RE.test(tx)) { kind[ix] |= W_RRN; colWarn[c].rrn++; }
      if (PHONE_RE.test(tx)) { kind[ix] |= W_PHONE; colWarn[c].phone++; }
      if (MAIL_RE.test(tx)) { kind[ix] |= W_MAIL; colWarn[c].mail++; }
    }

    // 충돌: 삭제하지 않을 칸 중 원본에 이미 [[...]] 모양이 있는 곳
    var collisions = [];
    for (r = 0; r < nR; r++) for (c = 0; c < nC; c++) {
      if (cfg.colMode[c] === 'delete') continue;
      if (BRACKET_LIKE_RE.test(work[r][c])) collisions.push({ r: r, c: c, text: work[r][c] });
    }

    var keepCols = [];
    for (c = 0; c < nC; c++) if (cfg.colMode[c] !== 'delete') keepCols.push(c);
    var out = text.map(function (row) { return keepCols.map(function (cc2) { return row[cc2]; }); });

    var entSig = 0;
    entries.forEach(function (e) { entSig = fnv(e[0] + '\u0001' + e[1], entSig || undefined); });
    var prefixCounts = [];
    counters.forEach(function (n, p) { prefixCounts.push([p, n]); });

    return {
      nR: nR, nC: nC, text: text, kind: kind, keepCols: keepCols, out: out,
      entries: entries, entSig: entries.length + ':' + entSig, prefixCounts: prefixCounts,
      candidates: candidates, crossSig: crossSig, needConfirm: needConfirm,
      colWarn: colWarn, collisions: collisions
    };
  }

  // 자체 검증: 결과 표를 키로 되돌렸을 때 원본과 한 칸도 다르지 않은지 확인
  function verify(work, res) {
    var map = new Map(res.entries), re = new RegExp(TOKEN_EXACT_SRC, 'g'), bad = 0, total = 0;
    for (var r = 0; r < res.nR; r++) for (var c = 0; c < res.nC; c++) {
      if (res.kind[r * res.nC + c] & K_DEL) continue;
      total++;
      var back = res.text[r][c].replace(re, function (m) { var x = map.get(m); return x === undefined ? m : x; });
      if (back !== work[r][c]) bad++;
    }
    return { total: total, bad: bad };
  }

  /* ---------------------------------------------------------------- 내보내기용 문자열 */

  function tsvCell(s) { return /[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
  function toTSV(out) { return out.map(function (r) { return r.map(tsvCell).join('\t'); }).join('\n'); }
  function csvCell(s) { return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
  function toCSV(out) { return out.map(function (r) { return r.map(csvCell).join(','); }).join('\r\n'); }
  function mdCell(s) { return s.replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>'); }
  function toMarkdown(out) {
    if (!out.length) return '';
    var line = function (r) { return '| ' + r.map(mdCell).join(' | ') + ' |'; };
    var sep = '| ' + out[0].map(function () { return '---'; }).join(' | ') + ' |';
    return [line(out[0]), sep].concat(out.slice(1).map(line)).join('\n');
  }

  function buildInstruction(res) {
    var ex = res.prefixCounts.map(function (p) { return '[[' + p[0] + '_NNN]]'; });
    return [
      '[AI에게 드리는 지시]',
      '이 자료에는 개인정보를 가리기 위한 "토큰"이 들어 있습니다. 예: ' + (res.entries.length ? res.entries[0][0] : '[[이름_001]]'),
      '- [[...]] 형태의 토큰은 철자·괄호·번호를 바꾸지 말고 그대로 쓰세요.',
      '- 토큰의 실제 값(이름·번호 등)을 추측하거나 만들어 내지 마세요.',
      '- 토큰을 번역하거나 줄이거나 합치지 말고, 새 토큰도 만들지 마세요.',
      ex.length ? '- 이 자료에 쓰인 토큰 종류: ' + ex.join(', ') : ''
    ].filter(Boolean).join('\n');
  }

  function textWidth(s) { var w = 0; for (var i = 0; i < s.length; i++) w += s.charCodeAt(i) > 0x2e80 ? 2 : 1; return w; }

  function makeXlsxBlobParts(out, sheetName) {
    var U = X().utils, ws = U.aoa_to_sheet(out);       // 문자열은 문자열 셀로 저장되어 0101 이 유지됨
    var widths = [];
    for (var c = 0; c < (out[0] || []).length; c++) {
      var w = 6; for (var r = 0; r < Math.min(out.length, 200); r++) w = Math.max(w, textWidth(out[r][c] || ''));
      widths.push({ wch: Math.min(w + 2, 50) });
    }
    ws['!cols'] = widths;
    var wb = U.book_new(); U.book_append_sheet(wb, ws, sheetName || '가명처리');
    return X().write(wb, { bookType: 'xlsx', type: 'array' });
  }

  /* ---------------------------------------------------------------- 키 파일 암호화 */
  // PBKDF2(SHA-256, 60만 회) 로 비밀번호에서 AES-256 키를 만들고 AES-GCM 으로 암호화한다.
  // salt(16B)·IV(12B)는 저장할 때마다 crypto.getRandomValues 로 새로 만든다.
  var KDF_ITER = 600000, KDF_MIN = 600000, KDF_MAX = 20000000;

  function b64(u8) { var s = '', n = 0x8000; for (var i = 0; i < u8.length; i += n) s += String.fromCharCode.apply(null, u8.subarray(i, i + n)); return btoa(s); }
  function unb64(s) { var bin = atob(s), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }

  async function deriveKey(pw, salt, iter, usage) {
    var base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw.normalize('NFC')), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, usage);
  }
  // 헤더 정보를 AAD 로 묶어, 헤더를 바꿔치기하면 복호화가 실패하도록 한다.
  function aad(keyId, iter) { return new TextEncoder().encode('pseudonym-key/1/' + keyId + '/' + iter); }

  function newKeyId() { var u = crypto.getRandomValues(new Uint8Array(3)); return Array.prototype.map.call(u, function (x) { return x.toString(16).padStart(2, '0'); }).join('').toUpperCase(); }

  async function encryptKey(entries, password, keyId) {
    var salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
    var key = await deriveKey(password, salt, KDF_ITER, ['encrypt']);
    var pt = new TextEncoder().encode(JSON.stringify({ v: 1, entries: entries }));
    var ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv, additionalData: aad(keyId, KDF_ITER), tagLength: 128 }, key, pt);
    return {
      format: 'pseudonym-key', version: 1,
      note: '암호화된 키 파일입니다. 비밀번호를 잃어버리면 복원할 수 없습니다.',
      keyId: keyId,
      kdf: { name: 'PBKDF2', hash: 'SHA-256', iterations: KDF_ITER, salt: b64(salt) },
      cipher: { name: 'AES-GCM', keyLength: 256, iv: b64(iv) },
      ciphertext: b64(new Uint8Array(ct))
    };
  }

  function KeyError(code, msg) { var e = new Error(msg); e.code = code; return e; }

  async function decryptKey(obj, password) {
    if (!obj || obj.format !== 'pseudonym-key' || obj.version !== 1 || !obj.kdf || !obj.cipher || !obj.ciphertext || typeof obj.keyId !== 'string')
      throw KeyError('format', '키 파일 형식이 아니에요. 가명처리 때 저장한 .key.json 파일을 선택해 주세요.');
    var iter = obj.kdf.iterations;
    if (obj.kdf.name !== 'PBKDF2' || obj.kdf.hash !== 'SHA-256' || obj.cipher.name !== 'AES-GCM' || !(iter >= KDF_MIN && iter <= KDF_MAX))
      throw KeyError('format', '지원하지 않는 키 파일입니다.');
    var salt, iv, ct;
    try { salt = unb64(obj.kdf.salt); iv = unb64(obj.cipher.iv); ct = unb64(obj.ciphertext); }
    catch (e) { throw KeyError('format', '키 파일이 손상되었어요.'); }
    var key = await deriveKey(password, salt, iter, ['decrypt']), pt;
    try { pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: iv, additionalData: aad(obj.keyId, iter), tagLength: 128 }, key, ct); }
    catch (e) { throw KeyError('password', '비밀번호가 맞지 않거나 키 파일이 변경·손상되었어요. 복원하지 못했습니다.'); }
    var data = JSON.parse(new TextDecoder().decode(pt));
    if (!data || !Array.isArray(data.entries)) throw KeyError('format', '키 파일 내용이 올바르지 않아요.');
    return { entries: data.entries, keyId: obj.keyId };
  }

  /* ---------------------------------------------------------------- 복원 */

  var OPEN = '(?:\\[\\[|\\[|［［|［|【|〔|「|『|\\{\\{|<<|《)';
  var CLOSE = '(?:\\]\\]|\\]|］］|］|】|〕|」|』|\\}\\}|>>|》)';
  var SEP = '(?:\\s*\\\\?[_＿\\-‐‑–—]\\s*|\\s+)?';

  function newReport() { return { replaced: 0, restoredSet: new Set(), unknown: new Map(), variants: new Map(), fixedVariants: 0 }; }

  /*
   * 키의 토큰을 원래 값으로 바꾸는 복원기.
   *  - 정확한 [[접두어_NNN]] 은 키에서 찾아 치환, 키에 없으면 그대로 두고 '키에 없는 토큰'으로 보고.
   *  - 키에 있는 접두어와 비슷한 변형([[이름_01]], [이름_001], \_ 이스케이프 등)은 '변형 의심'으로 보고.
   *    fixVariants 를 켠 경우에 한해, 키의 토큰으로 정확히 대응되는 변형만 함께 복원한다.
   */
  function makeRestorer(entries) {
    var map = new Map(entries), prefSet = new Set();
    entries.forEach(function (e) {
      var inner = e[0].slice(2, -2), k = inner.lastIndexOf('_');
      if (k > 0) prefSet.add(inner.slice(0, k));
    });
    var prefs = Array.from(prefSet).sort(function (a, b) { return b.length - a.length; }).map(escRe);
    var src = TOKEN_EXACT_SRC;
    if (prefs.length) src += '|' + OPEN + '\\s*(' + prefs.join('|') + ')' + SEP + '(\\d{1,7})\\s*' + CLOSE;
    var re = new RegExp(src, 'g');

    function restore(str, rep, parts, fixVariants) {
      var out = '', last = 0, mm;
      re.lastIndex = 0;
      while ((mm = re.exec(str)) !== null) {
        if (mm[0].length === 0) { re.lastIndex++; continue; }
        var piece = str.slice(last, mm.index), r, k, vPref = null, vNum = null;
        if (mm[1] !== undefined) {
          var val = map.get(mm[0]);
          if (val !== undefined) { r = val; k = 'ok'; rep.replaced++; rep.restoredSet.add(mm[0]); }
          else if (mm[1].slice(-1) === '\\' && prefSet.has(mm[1].slice(0, -1))) { vPref = mm[1].slice(0, -1); vNum = mm[2]; }   // [[이름\_001]] (마크다운 이스케이프)
          else { r = mm[0]; k = 'unk'; rep.unknown.set(mm[0], (rep.unknown.get(mm[0]) || 0) + 1); }
        } else { vPref = mm[3]; vNum = mm[4]; }
        if (vPref !== null) {
          var sug = '[[' + vPref + '_' + pad3(parseInt(vNum, 10)) + ']]', has = map.has(sug);
          var vr = rep.variants.get(mm[0]);
          if (!vr) rep.variants.set(mm[0], vr = { count: 0, suggest: has ? sug : null, fixed: false });
          vr.count++;
          if (fixVariants && has) { r = map.get(sug); k = 'fix'; rep.replaced++; rep.fixedVariants++; vr.fixed = true; rep.restoredSet.add(sug); }
          else { r = mm[0]; k = 'var'; }
        }
        out += piece + r;
        if (parts) { if (piece) parts.push({ k: 't', s: piece }); parts.push({ k: k, s: r }); }
        last = mm.index + mm[0].length;
      }
      var tail = str.slice(last);
      out += tail;
      if (parts && tail) parts.push({ k: 't', s: tail });
      return out;
    }
    return { restore: restore, size: map.size };
  }

  function restoreGrid(grid, restorer, rep, fix) {
    var changed = new Set();
    var g = grid.map(function (row, r) {
      return row.map(function (v, c) {
        var nv = v ? restorer.restore(v, rep, null, fix) : v;
        if (nv !== v) changed.add(r + ',' + c);
        return nv;
      });
    });
    return { grid: g, changed: changed };
  }

  // 엑셀 통합 문서의 문자열 셀을 직접 고쳐 서식·시트 구조를 최대한 유지한다.
  function restoreWorkbook(wb, restorer, rep, fix) {
    var changed = new Set();
    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      Object.keys(ws).forEach(function (k) {
        if (k[0] === '!') return;
        var cell = ws[k];
        if (cell && cell.t === 's' && typeof cell.v === 'string') {
          var nv = restorer.restore(cell.v, rep, null, fix);
          if (nv !== cell.v) { cell.v = nv; delete cell.w; delete cell.h; delete cell.r; changed.add(name + '!' + k); }
        }
      });
    });
    return changed;
  }

  var api = {
    K_KEEP: K_KEEP, K_TOKEN: K_TOKEN, K_CROSS: K_CROSS, K_DEL: K_DEL, W_PHONE: W_PHONE, W_RRN: W_RRN, W_MAIL: W_MAIL, W_ANY: W_ANY,
    KDF_ITER: KDF_ITER,
    colLetter: colLetter, sanitizePrefix: sanitizePrefix,
    decodeBytes: decodeBytes, detectDelimiter: detectDelimiter, parseDelimited: parseDelimited, normalizeGrid: normalizeGrid,
    readWorkbook: readWorkbook, sheetToGrid: sheetToGrid, guessHeader: guessHeader, findBracketCells: findBracketCells,
    isRecommended: isRecommended, crossDefault: crossDefault,
    newConfig: newConfig, compute: compute, verify: verify,
    toTSV: toTSV, toCSV: toCSV, toMarkdown: toMarkdown, buildInstruction: buildInstruction, makeXlsxBlobParts: makeXlsxBlobParts,
    newKeyId: newKeyId, encryptKey: encryptKey, decryptKey: decryptKey,
    newReport: newReport, makeRestorer: makeRestorer, restoreGrid: restoreGrid, restoreWorkbook: restoreWorkbook
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
