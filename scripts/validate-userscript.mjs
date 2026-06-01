import fs from 'node:fs';

const code = fs.readFileSync('douban-neodb-export.user.js', 'utf8');
const expectedSheets = [
  '看过',
  '在看',
  '想看',
  '听过',
  '在听',
  '想听',
  '读过',
  '在读',
  '想读',
  '玩过',
  '在玩',
  '想玩',
  '看过的舞台剧',
  '想看的舞台剧',
];

const checks = [
  ['Tampermonkey header', code.includes('// ==UserScript==')],
  ['NeoDB-compatible sheet count', (code.match(/\['(?:done|doing|mark)',/g) || []).length === 14],
  ['No unsupported drama progress sheet', !code.includes("['doing', '在看的舞台剧']")],
  ['Excel download', code.includes('XLSX.writeFile')],
  ['Referer header', code.includes('Referer:')],
  ['No Chrome extension API', !code.includes('chrome.')],
  ['Empty sheets retained', !code.includes('if (group.records.length === 0)')],
  ...expectedSheets.map(sheet => [`Sheet: ${sheet}`, code.includes(`'${sheet}'`)]),
];

for (const [name, passed] of checks) {
  console.log(`${passed ? 'PASS' : 'FAIL'} ${name}`);
}

if (checks.some(([, passed]) => !passed)) {
  process.exit(1);
}
