// Loads the core script of dpia100.html into Node for the unit tests. The text between
// <script id="dpia100-core"> and </script> is run unchanged in its own context, so the tested code is the
// shipped code (as with tests/web/harness.js in the PowerShell runner).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

export const repoRoot = fileURLToPath(new URL('../../', import.meta.url));
export const pagePath = fileURLToPath(new URL('../../dpia100.html', import.meta.url));
export const casePath = name => fileURLToPath(new URL(`../cases/${name}`, import.meta.url));

// extraRubrieken: lines added to the standard rubrieken table, for test rubrieken the real list lacks
// (T and U time rubrieken), as tests/cases/*/extra-rubrieken.txt does in the PowerShell runner.
export function loadCore({ extraRubrieken = [] } = {}) {
  const html = readFileSync(pagePath, 'utf8').replace(/\r\n?/g, '\n');
  const match = /<script id="dpia100-core">([\s\S]*?)<\/script>/.exec(html);
  if (!match) throw new Error('No <script id="dpia100-core"> in dpia100.html');
  const marker = 'const STANDARD_RUBRIEKEN_TABLE = String.raw`\n';
  if (!match[1].includes(marker)) throw new Error('No STANDARD_RUBRIEKEN_TABLE in the core of dpia100.html');
  const code = match[1].replace(marker, marker + extraRubrieken.map(l => `${l}\n`).join(''));
  return new vm.Script(`${code}\nDPIA100;`, { filename: 'dpia100-core.js' }).runInNewContext({ TextDecoder, TextEncoder });
}

// Analyze a CSV given as text (encoded as UTF-8). The result comes from another realm, so it is turned
// into plain objects for deepStrictEqual.
export function analyzeText(core, text, fileName = 'test.csv') {
  return JSON.parse(JSON.stringify(core.analyze(new Uint8Array(Buffer.from(text, 'utf8')), fileName)));
}

// A CSV with the standard header and the given rows (arrays of 7 cells), joined with the delimiter.
export const HEADER = ['Rubriek', 'Pers. Nr', 'Dienstv', 'S of V', 'Waarde', 'Ingangsdatum', 'Productiesoort'];
export function csv(rows, delim = ';', header = HEADER) {
  return [header, ...rows].map(r => (Array.isArray(r) ? r.join(delim) : r)).join('\r\n') + '\r\n';
}

// An inzetregels CSV (decision 21): the HR Core columns, and a row from the sandbox settings with overrides.
export const INZET_HEADER = ['Pers. Nr', 'Dienstverband', 'OE sleutel', 'Functie', 'Datum vanaf', 'Datum t/m', 'Uren', 'Soort inzet'];
export const inzetCsv = (rows, delim = ';') => csv(rows, delim, INZET_HEADER);
export const inzetRow = (extra = {}) => [extra.persNr ?? '10000', extra.dv ?? '1', extra.oe ?? '314', extra.functie ?? '020',
  extra.vanaf ?? '1-2-2027', extra.tm ?? '', extra.uren ?? '', extra.soort ?? ''];
