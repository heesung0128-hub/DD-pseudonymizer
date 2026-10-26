// node dev/test-core.js — 가상 데이터만 사용하는 Core 로직 테스트
const assert = require('assert');
globalThis.XLSX = require('./xlsx.full.min.js');
const Core = require('./core.js');

let pass = 0;
function ok(name, cond, extra) { if (!cond) { console.error('FAIL', name, extra || ''); process.exitCode = 1; } else { pass++; console.log('ok  ', name); } }

// ---- 가상 명렬표 30명 (실제 사람 아님, 연락처는 존재하지 않는 번호 대역)
const sur = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임'];
const giv = ['하늘', '서준', '지우', '도윤', '서연', '예준', '하윤', '주원', '지호', '수아', '민재', '채원', '시우', '유나', '건우'];
const rows = [['이름', '학번', '연락처', '비고', '점수']];
const names = [];
for (let i = 0; i < 30; i++) names.push(sur[i % 10] + giv[(i * 7 + 3) % 15] + (i >= 15 ? '' : ''));
// 중복 이름 방지
const uniq = new Set(names); ok('가상 이름 30개 고유', uniq.size === 30, uniq.size);
for (let i = 0; i < 30; i++) {
  const n = names[i], friend = names[(i + 1) % 30];
  rows.push([n, '01' + String(i + 1).padStart(2, '0'), '010-5550-' + String(1000 + i), `${n}은(는) 모둠 활동에서 ${friend}와(과) 협력함. 담임 확인.`, String(60 + i)]);
}
// 앞뒤 공백이 있는 이름, 연락처가 비고 속에 있는 경우
rows[2][0] = ' ' + rows[2][0] + ' ';
rows[3][3] += ' 보호자 연락처 ' + rows[3][2] + ' 로 회신.';

// ---- xlsx 로 저장했다가 다시 읽기: 학번 '0101' 유지
{
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '명렬표');
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const wb2 = Core.readWorkbook(buf);
  const { grid } = Core.sheetToGrid(wb2.Sheets['명렬표']);
  ok('xlsx 학번 0101 유지', grid[1][1] === '0101', grid[1][1]);
  ok('xlsx 전체 일치', JSON.stringify(grid) === JSON.stringify(rows));
}
// ---- csv: 0101 유지 + EUC-KR 판별
{
  const csv = rows.map(r => r.map(v => /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v).join(',')).join('\r\n');
  const d = Core.decodeBytes(new TextEncoder().encode(csv).buffer);
  const g = Core.normalizeGrid(Core.parseDelimited(d.text, Core.detectDelimiter(d.text)));
  ok('csv 학번 0101 유지', g[1][1] === '0101');
  ok('csv 전체 일치', JSON.stringify(g) === JSON.stringify(rows));
  const bad = new Uint8Array([0xC0, 0xCC, 0xB8, 0xA7]); // "이름" (CP949)
  ok('EUC-KR 판별', Core.decodeBytes(bad.buffer).text === '이름', Core.decodeBytes(bad.buffer).enc);
}
// ---- 병합 셀 / 날짜 서식
{
  const ws = XLSX.utils.aoa_to_sheet([['반', '이름'], ['1반', '가나다'], ['', '라마바']]);
  ws['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'S');
  const wb2 = Core.readWorkbook(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  const r = Core.sheetToGrid(wb2.Sheets.S);
  ok('병합 셀 채움', r.grid[2][0] === '1반' && r.merged === 1, JSON.stringify(r));
}

// ---- 설정: 이름·학번·연락처 토큰화(이름·연락처·학번은 교차 치환)
function makeCfg(work) {
  const cfg = Core.newConfig();
  work[0].forEach((h, c) => {
    cfg.colMode[c] = ['이름', '학번', '연락처'].includes(h) ? 'token' : 'keep';
    cfg.colPrefix[c] = Core.sanitizePrefix(h);
    cfg.colCross[c] = Core.crossDefault(h);
  });
  return cfg;
}
const work = rows;
const cfg = makeCfg(work);
let res = Core.compute(work, cfg);
ok('토큰 형식', /^\[\[이름_001\]\]$/.test(res.out[1][0]), res.out[1][0]);
ok('앞뒤 공백 보존', res.out[2][0] === ' [[이름_002]] ', JSON.stringify(res.out[2][0]));
ok('같은 값=같은 토큰(비고 속 친구 이름)', res.out[1][3].includes(res.out[2][0].trim()), res.out[1][3]);
ok('교차 치환 후보 존재 + 확인 필요', res.candidates.length > 0 && res.needConfirm);
ok('이름이 비고에 남지 않음', !names.some(n => res.out.slice(1).some(r => r[3].includes(n))));
ok('비고 속 연락처도 치환', !res.out[3][3].includes('010-5550-'), res.out[3][3]);
ok('점수 열 유지', res.out[5][4] === work[5][4]);
const v = Core.verify(work, res);
ok('자체 검증 0건 불일치', v.bad === 0, JSON.stringify(v));
cfg.crossConfirmedSig = res.crossSig;
res = Core.compute(work, cfg);
ok('확인 후 needConfirm=false', !res.needConfirm);

// 긴 값 우선: "김하윤" vs "하윤" 같은 포함 관계
{
  const w = [['이름', '비고'], ['김하윤', '김하윤과 하윤'], ['하윤', 'x']];
  const c = makeCfg(w);
  const r = Core.compute(w, c);
  ok('긴 값 우선(김하윤 → 이름_001, 하윤 → 이름_002)', r.out[1][1].startsWith(r.out[1][0] + '과 ') && r.out[1][1].endsWith(r.out[2][0]), r.out[1][1] + ' | ' + r.out[1][0] + ' | ' + r.out[2][0]);
  ok('포함 관계에서도 100% 복원', Core.verify(w, r).bad === 0);
}
// 영숫자 경계: 학번 "0101" 이 "10101" 안에서 치환되지 않음
{
  const w = [['학번', '비고'], ['0101', '코드 10101 과 0101 학생']];
  const c = makeCfg(w); const r = Core.compute(w, c);
  ok('숫자 경계', r.out[1][1].startsWith('코드 10101 과 ') && r.out[1][1].includes(r.out[1][0]), r.out[1][1]);
}
// 직접 추가 단어 + 제외
{
  const w = [['이름', '비고'], ['김하늘', '하늘이는 성실함. 서울초 출신']];
  const c = makeCfg(w); c.manualWords = ['하늘이', '서울초'];
  let r = Core.compute(w, c);
  ok('추가 단어 치환', r.out[1][1].startsWith('[[기타_001]]는') && r.out[1][1].includes('[[기타_002]]'), r.out[1][1]);
  c.crossExcluded.add('서울초'); r = Core.compute(w, c);
  ok('제외한 값은 치환 안 함', r.out[1][1].includes('서울초') && Core.verify(w, r).bad === 0, r.out[1][1]);
}
// 패턴 경고 + 충돌
{
  const w = [['이름', '비고'], ['가', '연락 010-1234-5678 / 900101-1234567 / a@b.kr / [[이름_001]]']];
  const c = makeCfg(w); const r = Core.compute(w, c);
  ok('전화/주민/메일 경고', (r.kind[3] & Core.W_PHONE) && (r.kind[3] & Core.W_RRN) && (r.kind[3] & Core.W_MAIL));
  ok('[[...]] 충돌 경고', r.collisions.length === 1);
  ok('충돌 시 자체 검증이 불일치 감지', Core.verify(w, r).bad === 1);
}
// 열 삭제 / 칸 단위 설정
{
  const w = [['a', 'b'], ['x', 'y'], ['x', 'z']];
  const c = Core.newConfig(); w[0].forEach((h, i) => { c.colMode[i] = 'keep'; c.colPrefix[i] = h; c.colCross[i] = false; });
  c.colMode[1] = 'delete'; c.cellOv.set('2,0', 'token');
  const r = Core.compute(w, c);
  ok('열 삭제', r.out[0].length === 1 && r.out[0][0] === 'a');
  ok('칸 단위 토큰화', r.out[1][0] === 'x' && r.out[2][0] === '[[a_001]]');
}
// 4자리 이상 일련번호
{
  const w = [['n'], ...Array.from({ length: 1001 }, (_, i) => ['v' + i])];
  const c = Core.newConfig(); c.colMode[0] = 'token'; c.colPrefix[0] = 'n'; c.colCross[0] = false;
  const r = Core.compute(w, c);
  ok('일련번호 3자리/4자리', r.out[1][0] === '[[n_001]]' && r.out[1000][0] === '[[n_1000]]' && r.out[999][0] === '[[n_999]]', r.out[1000][0]);
}

// ---- 암호화
(async () => {
  const entries = res.entries;
  const t0 = Date.now();
  const key = await Core.encryptKey(entries, '테스트-비밀번호-1234', 'ABC123');
  console.log('PBKDF2 ms', Date.now() - t0, 'iter', key.kdf.iterations);
  ok('반복 횟수 ≥ 600000', key.kdf.iterations >= 600000);
  ok('평문이 키 파일에 없음', !JSON.stringify(key).includes(names[0]) && !JSON.stringify(key).includes('이름_001'));
  const key2 = await Core.encryptKey(entries, '테스트-비밀번호-1234', 'ABC123');
  ok('salt·IV 무작위', key.kdf.salt !== key2.kdf.salt && key.cipher.iv !== key2.cipher.iv && key.ciphertext !== key2.ciphertext);
  const dec = await Core.decryptKey(JSON.parse(JSON.stringify(key)), '테스트-비밀번호-1234');
  ok('복호화 항목 일치', JSON.stringify(dec.entries) === JSON.stringify(entries) && dec.keyId === 'ABC123');
  let err = null; try { await Core.decryptKey(key, '틀린비밀번호12'); } catch (e) { err = e; }
  ok('틀린 비밀번호 → password 오류', err && err.code === 'password', err && err.message);
  const tamper = JSON.parse(JSON.stringify(key)); tamper.keyId = 'ZZZ999'; err = null;
  try { await Core.decryptKey(tamper, '테스트-비밀번호-1234'); } catch (e) { err = e; }
  ok('헤더 변조 감지(AAD)', err && err.code === 'password');
  err = null; try { await Core.decryptKey({ hello: 1 }, 'x'); } catch (e) { err = e; }
  ok('다른 JSON → format 오류', err && err.code === 'format');
  const weak = JSON.parse(JSON.stringify(key)); weak.kdf.iterations = 1000; err = null;
  try { await Core.decryptKey(weak, '테스트-비밀번호-1234'); } catch (e) { err = e; }
  ok('낮은 반복 횟수 거부', err && err.code === 'format');

  // ---- 복원: 예시 AI 답변
  const T = n => res.out[n][0], S = n => res.out[n][1];
  const e1 = res.out[1][0], e2 = res.out[2][0].trim(), id1 = res.out[1][1], ph1 = res.out[1][2];
  const aiAnswer = `1. ${e1}(${id1}) 학생은 ${ph1} 로 연락 가능하며, 모둠 활동에서 ${e2}와(과) 협력함.\n- ${e2} 학생 칭찬 문구 초안: "${e2}은(는) …"\n**${e1}** 의 학번은 ${id1}.`;
  const expected = `1. ${names[0]}(0101) 학생은 010-5550-1000 로 연락 가능하며, 모둠 활동에서 ${names[1]}와(과) 협력함.\n- ${names[1]} 학생 칭찬 문구 초안: "${names[1]}은(는) …"\n**${names[0]}** 의 학번은 0101.`;
  const restorer = Core.makeRestorer(dec.entries);
  const rep = Core.newReport();
  const parts = [];
  const got = restorer.restore(aiAnswer, rep, parts, false);
  ok('AI 답변 복원 = 원본 100% 일치', got === expected, got);
  ok('복원 토큰 수', rep.replaced === 8, rep.replaced);
  ok('parts 이어붙이면 결과와 같음', parts.map(p => p.s).join('') === got);

  // 표 전체 왕복: 가명처리한 표를 복원하면 원본과 한 칸도 다르지 않음
  const rep2 = Core.newReport();
  const back = Core.restoreGrid(res.out, restorer, rep2, false).grid;
  const orig = work.map(r => r.slice());
  ok('표 전체 왕복 100% 일치', JSON.stringify(back) === JSON.stringify(orig));

  // 변형 / 미등록 토큰
  const rep3 = Core.newReport();
  const msg = `${e1} / [[이름_01]] / [이름_001] / [[이름\\_001]] / 【이름 1】 / [[이름_999]] / [[없는열_001]]`;
  const noFix = restorer.restore(msg, rep3, null, false);
  ok('변형 의심 4종 탐지', rep3.variants.size === 4, Array.from(rep3.variants.keys()).join(' | '));
  ok('키에 없는 토큰 2종', rep3.unknown.size === 2, Array.from(rep3.unknown.keys()).join(' | '));
  ok('기본은 변형을 복원하지 않음', noFix.includes('[[이름_01]]') && noFix.includes('[이름_001]'));
  const rep4 = Core.newReport();
  const fixed = restorer.restore(msg, rep4, null, true);
  ok('옵션 켜면 대응되는 변형도 복원', !fixed.includes('[[이름_01]]') && rep4.fixedVariants === 4, fixed);

  // xlsx 복원
  const ws = XLSX.utils.aoa_to_sheet(res.out); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'AI');
  const wb3 = Core.readWorkbook(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  const rep5 = Core.newReport();
  const ch = Core.restoreWorkbook(wb3, restorer, rep5, false);
  const back2 = Core.sheetToGrid(wb3.Sheets.AI).grid;
  ok('xlsx 복원 왕복 100% 일치(0101 포함)', JSON.stringify(back2) === JSON.stringify(orig) && back2[1][1] === '0101' && ch.size > 0);

  console.log(`\n${pass} checks passed` + (process.exitCode ? ' — WITH FAILURES' : ''));
})();
