// Loads the core script of dpia100.html into Node for the unit tests. The text between
// <script id="dpia100-core"> and </script> is run unchanged in its own context, so the tested code is the
// shipped code (as with tests/web/harness.js in the PowerShell runner).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

export const repoRoot = fileURLToPath(new URL('../../', import.meta.url));
export const pagePath = fileURLToPath(new URL('../../dpia100.html', import.meta.url));
export const casePath = name => fileURLToPath(new URL(`../cases/${name}`, import.meta.url));

export function loadCore() {
  const html = readFileSync(pagePath, 'utf8').replace(/\r\n?/g, '\n');
  const match = /<script id="dpia100-core">([\s\S]*?)<\/script>/.exec(html);
  if (!match) throw new Error('No <script id="dpia100-core"> in dpia100.html');
  return new vm.Script(`${match[1]}\nDPIA100;`, { filename: 'dpia100-core.js' }).runInNewContext({ TextDecoder, TextEncoder });
}

// Analyze a CSV given as text (encoded as UTF-8). The result comes from another realm, so it is turned
// into plain objects for deepStrictEqual.
export function analyzeText(core, text, fileName = 'test.csv') {
  return JSON.parse(JSON.stringify(core.analyze(new Uint8Array(Buffer.from(text, 'utf8')), fileName)));
}

// A CSV with the standard header and the given rows (arrays of 7 cells), joined with the delimiter.
export const HEADER = ['Rubriek', 'Pers. Nr', 'Dienstv', 'S of V', 'Waarde', 'Ingangsdatum', 'Productiesoort'];
export function csv(rows, delim = ';') {
  return [HEADER, ...rows].map(r => (Array.isArray(r) ? r.join(delim) : r)).join('\r\n') + '\r\n';
}
