// Unit tests for the core of dpia100.html (node:test, no dependencies): npm run test:unit.
// They pin the rules function by function with readable failures. Whole-file parity with dpia100.ps1 is
// tested by tests/run-tests.ps1 (every case on PS 5.1, PS 7 and Edge) and tests/web/differential.ps1.
import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { loadCore, analyzeText, csv } from './core.mjs';

const core = loadCore();
const row = (waarde, extra = {}) => [extra.rubriek ?? 'E00101', extra.persNr ?? '123', extra.dv ?? '1', extra.sv ?? 'V', waarde, extra.datum ?? '1-11-2026', extra.prod ?? 'NOR'];

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
    // A dotted date is not an amount, but gets the date warning (decision 16).
    ['1.11.2026', ['1.11.2026', 'dateShape'], ['1.11.2026', 'dateShape'], ['1.11.2026', 'dateShape']]
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

describe('empty values, dates and times (decision 16)', () => {
  const kinds = value => analyzeText(core, csv([row(value)])).warnings.map(w => w.code);
  test('an empty Waarde is written, with a warning', () => {
    const r = analyzeText(core, csv([row(''), row('x'), row('  ', { persNr: '7' })]));
    assert.deepEqual(r.records.map(x => x.waarde), ['', 'x', '']);
    assert.deepEqual(r.warnings, [{ code: 'emptyValue', rows: [{ row: 2, persNr: '123' }, { row: 4, persNr: '7' }] }]);
  });
  test('dates not as dd-mm-yyyy get a warning and are written as typed', () => {
    for (const v of ['1-11-2026', '1/11/26', '01/11/2026', '2026-11-01', '1 11 2026', '1.11.2026']) {
      assert.deepEqual(kinds(v), ['dateShape'], v);
    }
    const r = analyzeText(core, csv([row('1-11-2026')]));
    assert.equal(r.records[0].waarde, '1-11-2026');
    assert.deepEqual(r.warnings[0].values, [{ row: 2, value: '1-11-2026' }]);
  });
  test('times not as hh:mm or hh:mm:ss get a warning', () => {
    for (const v of ['8:30', '8:30:00', '08:3:00']) assert.deepEqual(kinds(v), v === '08:3:00' ? [] : ['timeShape'], v);
  });
  test('the Excel-tool shapes, hh:mm:ss (type U) and number-like values pass', () => {
    for (const v of ['01-11-2026', '08:30', '08:30:00', '01112026', '8.30', '1-11', '1-11/2026', '123:45', 'x']) {
      assert.deepEqual(kinds(v), [], v);
    }
  });
  test('one warning per kind, after the row warnings', () => {
    const r = analyzeText(core, csv([row('8:30'), row('1-11-2026'), row(''), row('x'.repeat(60))]));
    assert.deepEqual(r.warnings.map(w => w.code), ['truncated', 'emptyValue', 'dateShape', 'timeShape']);
  });
});

describe('rubriek and productiesoort', () => {
  const codes = extra => analyzeText(core, csv([row('x', extra)])).errors.map(e => e.code);
  test('rubriek has 1 to 6 characters (decision 17)', () => {
    assert.deepEqual(codes({ rubriek: 'E0010' }), []);
    assert.deepEqual(codes({ rubriek: 'M' }), []);
    assert.deepEqual(codes({ rubriek: 'E001011' }), ['tooLong']);
    assert.deepEqual(codes({ rubriek: ' e00101 ' }), []);
    assert.deepEqual(codes({ rubriek: '#N/B' }), ['chars']);
  });
  test('a shorter code is left-aligned in positions 4-9 (issue #25)', () => {
    const r = analyzeText(core, csv([row('x', { rubriek: 'E0010' })]));
    const [line] = core.buildLines(r.records, { procescode: 'CMT', creatiedatum: '07102026', gebruiker: 'U' });
    assert.equal(line.substring(3, 9), 'E0010 ');
  });
  test('productiesoort NOR or COR, empty is NOR', () => {
    assert.deepEqual(codes({ prod: 'cor' }), []);
    assert.equal(analyzeText(core, csv([row('x', { prod: ' cor ' })])).records[0].productie, 'COR');
    assert.equal(analyzeText(core, csv([row('x', { prod: '' })])).records[0].productie, 'NOR');
    assert.deepEqual(codes({ prod: 'JWP' }), ['jwp']);
    assert.deepEqual(codes({ prod: 'XYZ' }), ['invalid']);
  });
});

describe('standard rubrieken (decisions 17-20)', () => {
  // Rubrieken used below, from the embedded list: P00001 DV N 6,2 stam+var; P00027 DV N stam only;
  // P00063 DV N variabel only; P00100 DV blocked; P00301 PS S 50; P00303 PS S 6; P00320 PS N 9,0;
  // P00321 PS D; P00332 DV S 3; P00633 DV blank Blok ind mut; SETKB DV S 2.
  const analyze = rows => analyzeText(core, csv(rows));
  const errs = rows => analyze(rows).errors.map(e => [e.field, e.code]);
  const value = (rubriek, v, extra = {}) => analyze([row(v, { rubriek, ...extra })]);

  test('the list is parsed: count, date, attributes, unknown codes', () => {
    assert.ok(core.standardRubriekCount > 4000, `${core.standardRubriekCount} rubrieken`);
    assert.match(core.STANDARD_RUBRIEKEN_DATE, /^\d{2}-\d{2}-\d{4}$/);
    assert.deepEqual(JSON.parse(JSON.stringify(core.standardRubriek('P00001'))),
      { code: 'P00001', gegsoort: 'DV', type: 'N', length: 6, decimals: 2, stam: true, variabel: true, blocked: false });
    assert.equal(core.standardRubriek('SETKB').length, 2);
    assert.equal(core.standardRubriek('P00633').blocked, false);
    assert.equal(core.standardRubriek('P00100').blocked, true);
    assert.equal(core.standardRubriek('E00101'), null);
  });
  test('an unknown code (eigen rubriek) gets no message and keeps the old checks', () => {
    const r = value('E00101', '12,50', { dv: '' });
    assert.deepEqual([r.errors, r.warnings, r.records[0].waarde, r.commaReported], [[], [], '12.50', true]);
  });
  test('the code is found after trim and uppercase', () => {
    assert.deepEqual(errs([row('1', { rubriek: ' p00001 ', dv: '' })]), [['Dienstverband', 'dvRequired']]);
  });

  test('dienstverband must match the gegevenssoort (decision 18)', () => {
    assert.deepEqual(errs([row('1', { rubriek: 'P00001', dv: '' })]), [['Dienstverband', 'dvRequired']]);
    assert.deepEqual(errs([row('1', { rubriek: 'P00001', dv: '0' })]), [['Dienstverband', 'dvRequired']]);
    assert.deepEqual(errs([row('x', { rubriek: 'P00301', dv: '1', sv: 'S' })]), [['Dienstverband', 'psOnly']]);
    assert.deepEqual(errs([row('x', { rubriek: 'P00301', dv: '', sv: 'S' })]), []);
  });
  test('a broken dienstverband gets one error, not two', () => {
    assert.deepEqual(errs([row('1', { rubriek: 'P00001', dv: '1a' })]), [['Dienstverband', 'digits']]);
    assert.deepEqual(errs([row('x', { rubriek: 'P00301', dv: '1234', sv: 'S' })]), [['Dienstverband', 'tooLong']]);
  });
  test('S of V must be allowed by the rubriek (decision 18)', () => {
    assert.deepEqual(errs([row('1', { rubriek: 'P00063', sv: 'S' })]), [['S of V', 'variabelOnly']]);
    assert.deepEqual(errs([row('1', { rubriek: 'P00027', sv: 'V' })]), [['S of V', 'stamOnly']]);
    assert.deepEqual(errs([row('1', { rubriek: 'P00001', sv: 'S' }), row('1', { rubriek: 'P00001', sv: 'V' })]), []);
    assert.equal(analyze([row('1', { rubriek: 'P00063', sv: 'S' })]).errors[0].rubriek, 'P00063');
  });

  test('Blok ind mut J is one warning listing the rows (decision 19)', () => {
    const r = analyze([row('1', { rubriek: 'P00100' }), row('1', { rubriek: 'P00633' }), row('', { rubriek: 'P00100' })]);
    assert.deepEqual(r.warnings, [
      { code: 'blocked', rows: [{ row: 2, rubriek: 'P00100' }, { row: 4, rubriek: 'P00100' }] },
      { code: 'emptyValue', rows: [{ row: 4, persNr: '123' }] }
    ]);
  });

  test('numeric: decision 8 applies, then must be a number that fits (decision 20)', () => {
    const r = value('P00001', '12,50');
    assert.deepEqual([r.records[0].waarde, r.commaReported, r.errors], ['12.50', true, []]);
    for (const v of ['-7', '007', '12.340', '999.99', '0.5', '']) assert.deepEqual(value('P00001', v).errors, [], v);
    assert.equal(value('P00001', '3.250').warnings[0].code, 'dotDecimal');
    for (const v of ['abc', 'NULL', '1.234,56', '+5', '1-11-2026', '- 7.5']) {
      assert.deepEqual(value('P00001', v).errors.map(e => e.code), ['notNumber'], v);
    }
    assert.deepEqual(analyzeText(core, csv([row('"1,234"', { rubriek: 'P00001' })], ',')).errors.map(e => e.code), ['notNumber']);
  });
  test('numeric: decimals and digits of the rubriek', () => {
    const e = v => value('P00001', v).errors.map(x => [x.code, x.kept ?? x.max]);
    assert.deepEqual(e('12.345'), [['decimals', '12.34']]);
    assert.deepEqual(e('-12.345'), [['decimals', '-12.34']]);
    assert.deepEqual(e('1234'), [['digits', 3]]);
    assert.deepEqual(e('0001234'), [['digits', 3]]);
    assert.deepEqual(value('P00320', '12.5', { dv: '', sv: 'S' }).errors.map(x => [x.code, x.decimals, x.kept]), [['decimals', 0, '12']]);
    assert.deepEqual(value('P00320', '123456789', { dv: '', sv: 'S' }).errors, []);
  });
  test('date: rewritten as dd-mm-yyyy and reported; not a date is an error', () => {
    const r = analyze(['1-11-2026', '2026-11-01', '01112026', '01-11-2026', '1.11.2026'].map(v => row(v, { rubriek: 'P00321', dv: '', sv: 'S' })));
    assert.deepEqual(r.records.map(x => x.waarde), Array(5).fill('01-11-2026'));
    assert.deepEqual(r.formatConversions.map(c => c.row), [2, 3, 4, 6]);
    assert.deepEqual(r.warnings, []);
    for (const v of ['1/11/26', '46327', 'NULL', '31-02-2026']) {
      assert.deepEqual(value('P00321', v, { dv: '', sv: 'S' }).errors.map(e => e.code), ['notDate'], v);
    }
  });
  test('text: as typed, no amount, date or time handling; length of the rubriek', () => {
    const r = analyze(['12,50', '1-11-2026', '8:30', '1.234,56'].map(v => row(v, { rubriek: 'P00301', dv: '', sv: 'S' })));
    assert.deepEqual([r.records.map(x => x.waarde), r.warnings, r.commaConversions], [['12,50', '1-11-2026', '8:30', '1.234,56'], [], []]);
    assert.deepEqual(value('P00303', 'ABCDEFG', { dv: '', sv: 'S' }).errors.map(e => [e.code, e.length]), [['textTooLong', 6]]);
    assert.deepEqual(value('P00303', 'ABCDEF', { dv: '', sv: 'S' }).errors, []);
    assert.equal(value('P00301', 'x'.repeat(60), { dv: '', sv: 'S' }).warnings[0].code, 'truncated');
  });

  describe('time rubrieken T (hh:mm) and U (hh:mm:ss), added for the test', () => {
    const timeCore = loadCore({ extraRubrieken: ['ZZT001 DV T  5 - SV N', 'ZZU001 DV U  8 - SV N'] });
    const t = (rubriek, v) => analyzeText(timeCore, csv([row(v, { rubriek })]));
    test('rewritten to the shape of the type and reported', () => {
      for (const [rub, v, out] of [['ZZT001', '8:30', '08:30'], ['ZZT001', '08:30:00', '08:30'], ['ZZT001', '23:59', '23:59'],
        ['ZZU001', '8:30', '08:30:00'], ['ZZU001', '08:30:15', '08:30:15']]) {
        const r = t(rub, v);
        assert.deepEqual([r.errors, r.records[0].waarde, r.formatConversions.length], [[], out, v === out ? 0 : 1], `${rub} ${v}`);
      }
    });
    test('not a time, or seconds in a T rubriek, is an error', () => {
      for (const [rub, v] of [['ZZT001', '08:30:15'], ['ZZT001', '24:00'], ['ZZT001', '8.30'], ['ZZU001', 'abc'], ['ZZU001', '12:60:00']]) {
        assert.deepEqual(t(rub, v).errors.map(e => [e.code, e.type]), [['notTime', rub === 'ZZT001' ? 'T' : 'U']], `${rub} ${v}`);
      }
    });
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
