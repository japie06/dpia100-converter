// Checks that the demo CSV files in tests/demo show every message of the converter. Run by hand after
// make-demo.ps1 (not part of the test pipeline): node tests/demo/check-demo.mjs
// The codes come from the core of dpia100.html: every err(...), warn(...), grouped warning and fatal code.
// The demo files go through that same core; a code that no file triggers is listed as missing.
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { loadCore, pagePath } from '../web/core.mjs';

const dir = fileURLToPath(new URL('./', import.meta.url));
const core = loadCore();
const html = readFileSync(pagePath, 'utf8');
const source = /<script id="dpia100-core">([\s\S]*?)<\/script>/.exec(html)[1];

const all = new Set();
for (const m of source.matchAll(/err\(rowNumber, '([^']+)', '(\w+)'/g)) all.add(`error ${m[1]}: ${m[2]}`);
for (const m of source.matchAll(/warn\('(\w+)'/g)) all.add(`warning ${m[1]}`);
for (const m of source.matchAll(/r\.warnings\.(?:push|unshift)\(\{ code: '(\w+)'/g)) all.add(`warning ${m[1]}`);
for (const m of source.matchAll(/r\.fatal = \{ code: '(\w+)'/g)) all.add(`file ${m[1]}`);
all.add('report commaReported');
all.add('report formatConversions');

// Not reachable from a demo file: the reason is printed with the result.
const exempt = new Map([
  ['file validation', 'the row errors of 1-fouten.csv'],
  ['file noRecords', 'not reachable from a file that has lines'],
  ['error Mutatiewaarde: unicode', 'a lone surrogate; a file decoder replaces it']
]);
const table = /STANDARD_RUBRIEKEN_TABLE = String\.raw`([\s\S]*?)`/.exec(source)[1];
if (!/^\S+ +(PS|DV) [TU] /m.test(table)) exempt.set('error Mutatiewaarde: notTime', 'no standard time rubriek (T, U) in the list');

const seen = new Set();
const problems = [];
const files = readdirSync(dir).filter(f => f.endsWith('.csv')).sort();
for (const file of files) {
  const r = JSON.parse(JSON.stringify(core.analyze(new Uint8Array(readFileSync(dir + file)), file)));
  if (r.fatal) seen.add(`file ${r.fatal.code}`);
  for (const e of r.errors) seen.add(`error ${e.field}: ${e.code}`);
  for (const w of r.warnings) seen.add(`warning ${w.code}`);
  if (r.commaReported) seen.add('report commaReported');
  if (r.formatConversions.length > 0) seen.add('report formatConversions');
  const summary = r.fatal ? `${r.fatal.code}${r.errors.length ? ` (${r.errors.length} errors)` : ''}` : `ok, ${r.warnings.length} warning(s)`;
  console.log(`${file.padEnd(34)} ${summary}`);

  if (file === '1-fouten.csv') {
    // Row 2 is the good row; every other row shows exactly one error, so the remark matches the message.
    const perRow = new Map();
    for (const e of r.errors) perRow.set(e.row, (perRow.get(e.row) || 0) + 1);
    for (const [row, n] of perRow) {
      if (row === 2) problems.push('1-fouten.csv: row 2 should have no error');
      else if (n !== 1) problems.push(`1-fouten.csv: row ${row} has ${n} errors, expected 1`);
    }
    for (const rec of r.records) if (rec.row !== 2 && !perRow.has(rec.row)) problems.push(`1-fouten.csv: row ${rec.row} has no error`);
  } else if (/^\d-waarschuwingen/.test(file) && r.fatal) {
    problems.push(`${file}: expected no errors, got ${r.fatal.code}`);
  }
}

const missing = [...all].filter(c => !seen.has(c) && !exempt.has(c)).sort();
console.log(`\n${all.size} messages, ${seen.size} shown by the demo files, ${exempt.size} not reachable:`);
for (const [code, why] of exempt) console.log(`  - ${code}: ${why}`);
for (const p of problems) console.log(`PROBLEM: ${p}`);
for (const c of missing) console.log(`MISSING: ${c}`);
if (missing.length || problems.length) process.exit(1);
console.log('All messages covered.');
