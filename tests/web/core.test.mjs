// Unit tests for the core of dpia100.html (node:test, no dependencies): npm run test:unit.
// They pin the rules function by function with readable failures. Whole-file parity with dpia100.ps1 is
// tested by tests/run-tests.ps1 (every case on PS 5.1, PS 7 and Edge) and tests/web/differential.ps1.
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { loadCore, analyzeText, csv } from './core.mjs';

const core = loadCore();
const row = (waarde, extra = {}) => [extra.rubriek ?? 'E00101', extra.persNr ?? '123', extra.dv ?? '1', extra.sv ?? 'V', waarde, extra.datum ?? '1-11-2026', 'NOR'];

describe('dates (PS: Test-And-ParseDate)', () => {
  // Results probed on PowerShell 5.1 and 7.
  const cases = {
    '1 11 2026': '01112026', '1  11 2026': '01112026', '1-11-2026': '01112026', '1/11/2026': '01112026',
    '1.11.2026': '01112026', '1.11-2026': '01112026', '1 - 11 - 2026': '01112026', '2026-11-01': '01112026',
    '01112026': '01112026', ' 1-11-2026': '01112026', '1 11 2026': '01112026', '1\u008511 2026': '01112026',
    '29-02-2028': '29022028', '29-02-2000': '29022000', '1-11-0001': '01110001', '31-12-9999': '31129999',
    '2026-1-01': null, '2026-11-1': null, '1112026': null, '011126': null, '1-11-26': null, '31-02-2026': null,
    '29-02-1900': null, '1-11-0000': null, '0-11-2026': null, '2026/11/01': null, '１-11-2026': null,
    '1﻿11 2026': null, '00112026': null, '1-11-2026-': null, '': null
  };
  for (const [input, expected] of Object.entries(cases)) {
    test(JSON.stringify(input), () => assert.equal(core.parseDate(input), expected));
  }
});

describe('delimiter (PS: Detect-CsvDelimiter)', () => {
  test('most frequent wins', () => assert.equal(analyzeText(core, csv([row('x')], ',')).delimiter, ','));
  test('tie between ; and , gives ;', () => {
    const text = 'Rubriek;Pers. Nr;Dienstv;S of V;Waarde (a, b, c, d, e, f);Ingangsdatum;Productiesoort\r\nE00101;123;1;S;x;1-11-2026;NOR\r\n';
    assert.equal(analyzeText(core, text).delimiter, ';');
  });
  test('tie between tab and , gives tab', () => {
    const text = 'Rubriek\tPers. Nr\tDienstv\tS of V\tWaarde (a, b, c, d, e, f)\tIngangsdatum\tProductiesoort\r\nE00101\t123\t1\tS\tx\t1-11-2026\tNOR\r\n';
    assert.equal(analyzeText(core, text).delimiter, '\t');
  });
});

describe('CSV records (PS: ConvertFrom-Csv quirks)', () => {
  test('a first row starting with # is validated, not dropped', () => {
    const r = analyzeText(core, csv([row('x', { rubriek: '#N/B' }), row('y')]));
    assert.deepEqual(r.errors.map(e => [e.row, e.field, e.code]), [[2, 'Rubriek', 'chars']]);
  });
  test('blank and whitespace-only lines are not rows; ;;; rows are', () => {
    const r = analyzeText(core, csv([row('a'), '', '   ', ';;;;;;', row('b')]));
    assert.deepEqual(r.records.map(x => x.row), [2, 4]);
    assert.equal(r.skippedEmptyRows, 1);
  });
  test('spaces before an opening quote, text after a closing quote', () => {
    const r = analyzeText(core, csv([['E00101', '123', '1', 'S', '  "a;b"', '"1-11"-2026', 'NOR']]));
    assert.equal(r.records[0].waarde, 'a;b');
    assert.equal(r.records[0].ingang, '01112026');
  });
  test('values beyond the header are dropped', () => {
    const r = analyzeText(core, csv([[...row('x'), 'extra', 'more']]));
    assert.equal(r.fatal, null);
    assert.equal(r.records.length, 1);
  });
  test('a quoted line break is an error in Waarde', () => {
    const r = analyzeText(core, csv([['E00101', '123', '1', 'S', '"regel1\nregel2"', '1-11-2026', 'NOR']]));
    assert.deepEqual(r.errors.map(e => e.code), ['control']);
  });
});

describe('.NET string semantics', () => {
  test('NBSP and U+0085 are trimmed, U+FEFF is not', () => {
    assert.equal(analyzeText(core, csv([row('x', { persNr: ' 123\u0085' })])).fatal, null);
    assert.deepEqual(analyzeText(core, csv([row('x', { persNr: '123﻿' })])).errors.map(e => e.code), ['digits']);
  });
  test('ß is not uppercased to SS', () => {
    assert.deepEqual(analyzeText(core, csv([row('x', { rubriek: 'e0010ß' })])).errors.map(e => e.code), ['chars']);
  });
  test('leading zeros are stripped from the keys (decision 4)', () => {
    const rec = analyzeText(core, csv([row('x', { persNr: '000123', dv: '001' })])).records[0];
    assert.deepEqual([rec.persNr, rec.dienstv], ['123', '1']);
  });
});

describe('amounts by delimiter (decision 8)', () => {
  // [value, ';' result, ',' result, tab result]; result = [written value, kind] with kind
  // 'reported' (comma conversion listed without confirmation), a warning code, or null.
  const table = [
    ['12,50', ['12.50', 'reported'], ['12.50', 'commaChanged'], ['12.50', 'commaChanged']],
    ['-12,50', ['-12.50', 'reported'], ['-12.50', 'commaChanged'], ['-12.50', 'commaChanged']],
    ['1,234', ['1.234', 'reported'], ['1,234', 'thousandEn'], ['1.234', 'commaAsDecimal']],
    ['3.250', ['3.250', 'dotDecimal'], ['3.250', null], ['3.250', 'dotDecimal']],
    ['1.234,56', ['1.234,56', 'separator'], ['1.234,56', 'separator'], ['1.234,56', 'separator']],
    ['1.234.567', ['1.234.567', 'separator'], ['1.234.567', 'separator'], ['1.234.567', 'separator']],
    ['12.50', ['12.50', null], ['12.50', null], ['12.50', null]],
    ['1.11.2026', ['1.11.2026', null], ['1.11.2026', null], ['1.11.2026', null]]
  ];
  const delims = [';', ',', '\t'];
  for (const [value, ...expected] of table) {
    delims.forEach((d, i) => {
      test(`${value} in a ${JSON.stringify(d)} file`, () => {
        const cell = d === ',' && value.includes(',') ? `"${value}"` : value;
        const r = analyzeText(core, csv([row(cell)], d));
        const kind = r.commaReported ? 'reported' : (r.warnings[0] ? r.warnings[0].code : null);
        assert.deepEqual([r.records[0].waarde, kind], expected[i]);
      });
    });
  }
  test('Waarde over 50 characters is truncated with a warning (decision 3)', () => {
    const r = analyzeText(core, csv([row('x'.repeat(60))]));
    assert.equal(r.records[0].waarde.length, 50);
    assert.equal(r.warnings[0].code, 'truncated');
  });
});

describe('Windows-1252 (decisions 1 and 7)', () => {
  test('€ and accents are one byte; C1 controls of the undefined bytes round-trip', () => {
    assert.deepEqual([...core.encodeOutput(['€é\u0081'])], [0x80, 0xE9, 0x81, 13, 10]);
  });
  test('U+0080 and characters outside 1252 are rejected', () => {
    assert.throws(() => core.encodeOutput(['\u0080']));
    assert.equal(core.unencodableChars('ĳ ok 😀'), 'ĳ (U+0133), \uD83D (U+D83D), \uDE00 (U+DE00)');
  });
  test('decomposed accents are composed first', () => {
    assert.equal(analyzeText(core, csv([row('coördinator')])).records[0].waarde, 'coördinator');
  });
});

describe('record layout (field-layout.md)', () => {
  test('reproduces the verified Excel-tool example byte for byte', () => {
    const r = analyzeText(core, csv([['E00101', '10', '1', 'S', 'Test', '1 11 2024', 'NOR'], ['E00101', '123450', '1', 'S', 'Test', '5 11 2024', 'NOR']]));
    const lines = core.buildLines(r.records, { procescode: 'CMT', creatiedatum: '07102026', gebruiker: 'TSTUSER' });
    assert.deepEqual([...lines], [
      'CMTE00101    10   1' + ' '.repeat(40) + 'S   07102026         TSTUSERTest' + ' '.repeat(46) + '   01112024' + ' '.repeat(30) + 'NOR      ',
      'CMTE00101123450   1' + ' '.repeat(40) + 'S   07102026         TSTUSERTest' + ' '.repeat(46) + '   05112024' + ' '.repeat(30) + 'NOR      '
    ]);
  });
  test('a PS key has no dienstverband part', () => {
    const r = analyzeText(core, csv([row('x', { dv: '' })]));
    const [line] = core.buildLines(r.records, { procescode: 'CMT', creatiedatum: '07102026', gebruiker: 'U' });
    assert.equal(line.substring(9, 59), '   123'.padEnd(50));
  });
});

describe('prompt values and file name', () => {
  test('procescode: empty is CMT, uppercased, max 3', () => {
    assert.equal(core.checkProcescode('').value, 'CMT');
    assert.equal(core.checkProcescode(' con ').value, 'CON');
    assert.equal(core.checkProcescode('TOOLONG').error, 'procescode');
  });
  test('creatiedatum: same formats as the CSV', () => {
    assert.equal(core.checkCreatiedatum('2026-10-07').value, '07102026');
    assert.equal(core.checkCreatiedatum('31 02 2026').error, 'creatiedatum');
  });
  test('creatiegebruiker: mandatory, max 16, Windows-1252', () => {
    assert.equal(core.checkGebruiker('  ').error, 'gebruikerEmpty');
    assert.equal(core.checkGebruiker('ABCDEFGHIJKLMNOPQ').error, 'gebruikerTooLong');
    assert.equal(core.checkGebruiker('Łukasz').error, 'gebruikerChars');
    assert.equal(core.checkGebruiker('ABCDEFGHIJKLMNOP').value, 'ABCDEFGHIJKLMNOP');
  });
  test('output name DPIA100_<csv name>.txt (decisions 12 and 13)', () => {
    assert.equal(core.outputFileName('mutaties oktober.csv'), 'DPIA100_mutaties oktober.txt');
    assert.equal(core.outputFileName('a.b.csv'), 'DPIA100_a.b.txt');
    assert.equal(core.isValidOutputName('dpia100.txt'), true);
    assert.equal(core.isValidOutputName('mutaties.txt'), false);
  });
});
