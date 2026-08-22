// Catches calls to functions that no longer exist.
//
// `node --check` only validates syntax, so deleting a function while leaving
// its call sites behind passes cleanly and then throws at runtime — usually
// inside a try/catch, where it surfaces as a misleading "could not load"
// message rather than an error anyone can act on.
//
//   node tests/check-refs.js sriramstudio_admin.html

const fs = require('fs');

const file = process.argv[2];
if (!file) { console.error('usage: node tests/check-refs.js <file.html>'); process.exit(2); }

const html = fs.readFileSync(file, 'utf8');
const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]);
if (!blocks.length) { console.error('no inline script blocks found'); process.exit(2); }

// Strip comments and string literals so their contents are not mistaken for code.
let src = blocks.join('\n;\n')
  .replace(/\/\*[\s\S]*?\*\//g, ' ')
  .replace(/(^|[^:])\/\/[^\n]*/g, '$1 ')
  .replace(/`(?:\\[\s\S]|[^`\\])*`/g, '``')
  .replace(/'(?:\\[\s\S]|[^'\\\n])*'/g, "''")
  .replace(/"(?:\\[\s\S]|[^"\\\n])*"/g, '""');

const defined = new Set();
for (const m of src.matchAll(/\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g)) defined.add(m[1]);
for (const m of src.matchAll(/\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=/g)) defined.add(m[1]);
for (const m of src.matchAll(/\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*[,;\n]/g)) defined.add(m[1]);
for (const m of src.matchAll(/\bfunction\s*\(([^)]*)\)/g)) {
  m[1].split(',').forEach(p => { const n = p.trim().split('=')[0].trim(); if (n) defined.add(n); });
}

const KEYWORDS = new Set(['if','for','while','switch','catch','function','return','typeof','new',
  'do','else','await','yield','delete','void','in','of','case','throw','instanceof']);

const GLOBALS = new Set(['console','document','window','alert','confirm','setTimeout','clearTimeout',
  'setInterval','clearInterval','fetch','Promise','JSON','Object','Array','String','Number','Boolean',
  'Math','Date','RegExp','Error','Map','Set','parseInt','parseFloat','isNaN','isFinite','encodeURI',
  'encodeURIComponent','decodeURI','decodeURIComponent','localStorage','sessionStorage','navigator',
  'location','URL','URLSearchParams','Intl','Symbol','BigInt','requestAnimationFrame','structuredClone',
  'queueMicrotask','btoa','atob','print','open','close','blur','focus','scroll','scrollTo','matchMedia']);

const missing = new Map();
for (const m of src.matchAll(/(^|[^.\w$])([A-Za-z_$][\w$]*)\s*\(/gm)) {
  const name = m[2];
  if (KEYWORDS.has(name) || GLOBALS.has(name) || defined.has(name)) continue;
  missing.set(name, (missing.get(name) || 0) + 1);
}

if (!missing.size) {
  console.log('PASS  every called function is defined  (' + defined.size + ' definitions seen)');
  process.exit(0);
}
console.log('FAIL  called but never defined:');
[...missing.entries()].sort((a, b) => b[1] - a[1])
  .forEach(([n, c]) => console.log('  ' + n + '()  - ' + c + ' call site(s)'));
process.exit(1);
