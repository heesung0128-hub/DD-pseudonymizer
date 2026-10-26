// node dev/build.js — dev/ 의 조각들을 합쳐 CSP 해시가 들어간 단일 ../index.html 을 만든다.
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const d = __dirname;
// HTML 파서는 줄바꿈을 LF 로 정규화한 뒤 해시를 계산하므로, 처음부터 LF 로 통일한다.
const rd = f => fs.readFileSync(path.join(d, f), 'utf8').replace(/^﻿/, '').replace(/\r\n?/g, '\n');
const sha = s => "'sha256-" + crypto.createHash('sha256').update(s, 'utf8').digest('base64') + "'";

let lib = rd('xlsx.full.min.js');
if (/<\/script/i.test(lib)) lib = lib.replace(/<\/script/gi, '<\\/script');
const css = rd('style.css');
const app = rd('core.js') + '\n' + rd('ui.js');
if (/<\/script/i.test(app)) throw new Error('app code contains </script');

const csp = [
  "default-src 'none'",
  'script-src ' + sha(lib) + ' ' + sha(app),   // 이 파일 안의 두 스크립트 블록만 실행 허용
  'style-src ' + sha(css),
  'img-src data:',
  "connect-src 'none'", "font-src 'none'", "media-src 'none'", "object-src 'none'",
  "frame-src 'none'", "child-src 'none'", "worker-src 'none'", "manifest-src 'none'",
  "base-uri 'none'", "form-action 'none'"
].join('; ');

// 자리표시자를 한 번에 치환(본문 안에 '$&' 등이 있어도 안전하도록 함수형 replace 사용)
const marks = { __CSP__: csp, __CSS__: css, __SHEETJS__: lib, __APP__: app };
const html = rd('template.html').replace(/__(CSP|CSS|SHEETJS|APP)__/g, m => marks[m]);
if (/__(CSP|CSS|SHEETJS|APP)__/.test(html.replace(lib, '').replace(app, ''))) throw new Error('unreplaced placeholder');
fs.writeFileSync(path.join(d, '..', 'index.html'), html, 'utf8');
console.log('index.html', Buffer.byteLength(html) + ' bytes');
console.log(csp.replace(/sha256-[^']{8}[^']*/g, 'sha256-…'));
