/* Test driver for the core of dpia100.html, run in headless Edge by tests/web/WebTest.ps1.
   The core script is copied byte for byte from dpia100.html into the same page, so the tested code
   is the shipped code.

   For every case it does what dpia100.ps1 does on the console: analyze the CSV, answer the prompts
   from the case's answers (an invalid answer is followed by the next one, as with piped input), ask
   the Y confirmation for warnings, and build the file. It returns the exit code (0/1), the output
   bytes and an English transcript with the same lines as the PowerShell console, so the case files
   expected-output.txt and unexpected-output.txt apply unchanged. The English texts below are copies
   of the messages in dpia100.ps1; keep them in sync. */
(() => {
  'use strict';

  const ISSUES = {
    'Rubriek/empty': () => 'Field is empty',
    'Rubriek/tooLong': () => 'Exceeds max length of 6 chars',
    'Rubriek/chars': () => 'Contains invalid characters (alphanumeric only)',
    'Pers. Nr/empty': () => 'Field is empty',
    'Pers. Nr/digits': () => 'Must contain digits only (no decimals, spaces or letters)',
    'Pers. Nr/zero': () => 'Cannot be zero',
    'Pers. Nr/tooLong': () => 'Exceeds max length of 6 digits',
    'Dienstverband/digits': () => 'Must contain digits only (no decimals, spaces or letters)',
    'Dienstverband/tooLong': () => 'Exceeds max length of 3 digits',
    'S of V/invalid': () => "Must strictly be 'S' (Stam) or 'V' (Variabel)",
    'Ingangsdatum/invalid': () => 'Invalid or unparseable date',
    'Productiesoort/jwp': () => 'JWP is not supported by this script (requires a verrekenperiode). Use NOR',
    'Productiesoort/invalid': () => 'Must be one of: NOR',
    'Mutatiewaarde/control': () => 'Contains line break, tab or control character',
    'Mutatiewaarde/unicode': () => 'Contains invalid Unicode characters',
    'Mutatiewaarde/encoding': e => `Characters not available in Windows-1252: ${e.chars}`
  };

  const label = w => `Row ${w.row} (Pers. Nr: ${w.persNr})`;
  const WARNINGS = {
    commaChanged: w => commaMessage(w.conversions),
    thousandEn: w => `${label(w)}: Waarde '${w.value}' has a thousand separator (English notation in a comma-separated file). Written as typed; a numeric rubriek rejects it at processing (PI_2229). Write ${w.whole}${w.fraction}.`,
    commaAsDecimal: w => `${label(w)}: Waarde '${w.value}' was written as '${w.converted}' (comma read as decimal). If the comma is a thousand separator, write ${w.whole}${w.fraction}.`,
    dotDecimal: w => `${label(w)}: Waarde '${w.value}' is read by Beaufort as ${w.asComma} in a numeric rubriek. If the dot is a thousand separator, write it without the dot.`,
    separator: w => `${label(w)}: Waarde '${w.value}' looks like an amount with a thousand separator. A numeric rubriek rejects it at processing (PI_2229); write it without, e.g. 1234.56.`,
    truncated: w => `${label(w)}: Waarde was truncated from ${w.length} to 50 chars:\n    Original : '${w.original}'\n    Truncated: '${w.truncated}'`
  };

  function commaMessage(conversions) {
    const shown = conversions.slice(0, 5).map(c => `row ${c.row} '${c.from}' -> '${c.to}'`).join(', ');
    const more = conversions.length > 5 ? `, and ${conversions.length - 5} more` : '';
    return `Comma changed to dot in ${conversions.length} Waarde value(s), because a numeric rubriek in Beaufort only accepts a dot (a text rubriek gets the dot too): ${shown}${more}.`;
  }

  const PROMPT_WARNINGS = {
    procescode: () => 'Procescode must be 1 to 3 alphanumeric characters (e.g. CMT, CON).',
    creatiedatum: r => `Invalid date '${r.input}'. Please use 'dd mm yyyy' or 'dd-mm-yyyy'.`,
    gebruikerEmpty: () => 'Creatiegebruiker cannot be empty.',
    gebruikerTooLong: () => 'Creatiegebruiker cannot exceed 16 characters.',
    gebruikerChars: r => `Creatiegebruiker contains characters that do not exist in Windows-1252: ${r.chars}`
  };

  const fromBase64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  const toBase64 = bytes => {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  };

  function errorValue(e) {
    return e.code === 'control' ? e.value.replace(/[\x00-\x1F\x7F]+/g, ' <line break/tab> ') : e.value;
  }

  function runCase(c) {
    const out = [];
    const result = (exitCode, bytes) => ({
      name: c.name, exitCode, output: bytes ? toBase64(bytes) : null, transcript: out.join('\n'),
      suggestedName: DPIA100.outputFileName(c.fileName)
    });
    const fail = msg => { out.push(`[ERROR] ${msg}`); return result(1, null); };

    // PS checks an explicit output name before anything else (decision 12); the page checks the name
    // the user chooses in the save dialog.
    if (c.outputName && !DPIA100.isValidOutputName(c.outputName)) {
      return fail(`The output file name '${c.outputName}' must start with 'DPIA100' (for example dpia100.txt), otherwise Beaufort does not find it.`);
    }

    const r = DPIA100.analyze(fromBase64(c.input), c.fileName);
    const f = r.fatal;
    if (f && f.code === 'read') return fail(`Could not read '${c.fileName}': ${f.message}`);
    out.push(`Input encoding    : ${r.encodingName}`);
    if (f && f.code === 'empty') return fail(`The file '${c.fileName}' is completely empty. Aborting.`);
    if (f && f.code === 'headerOnly') return fail(`The file '${c.fileName}' contains only headers and NO data rows. Aborting.`);
    out.push(`Detected delimiter: '${r.delimiter}'`);
    if (f && f.code === 'noRecords') return fail(`No data records could be parsed from '${c.fileName}'. Aborting.`);
    out.push('Column mapping:');
    for (const m of r.mapping) out.push(`  ${m.field.padEnd(15)} <- ${m.column ? `'${m.column}'` : '(not found)'}`);
    if (f && f.code === 'missingHeaders') {
      return fail(`The following mandatory column headers are missing from the input file: ${f.fields.join(', ')}. Aborting.`);
    }
    if (f && f.code === 'doubleMapped') {
      const details = f.groups.map(g => `'${g.column}' -> ${g.fields.join(' and ')}`).join('; ');
      return fail(`A CSV column was matched to more than one field: ${details}. Rename the header(s). Aborting.`);
    }
    if (f && f.code === 'validation') {
      out.push(`[FATAL ERROR] ${r.errors.length} validation error(s) found. Process aborted!`);
      out.push('No export file was created. Please fix the issues below in your CSV:');
      out.push('Row Column Value Issue');
      for (const e of r.errors) out.push(`${e.row} ${e.field} ${errorValue(e)} ${ISSUES[`${e.field}/${e.code}`](e)}`);
      return result(1, null);
    }
    if (f && f.code === 'allEmpty') return fail(`No data rows found (all ${f.count} row(s) were empty). Aborting.`);
    if (f) return fail(`Unknown fatal code ${f.code}`);

    out.push(`[OK] ${r.records.length} row(s) checked, no errors found.`);
    if (r.commaReported) out.push(commaMessage(r.commaConversions));
    if (r.warnings.length > 0) {
      out.push(`[WARNINGS] ${r.warnings.length} warning(s). The values are written as shown:`);
      for (const w of r.warnings) out.push(`  - ${WARNINGS[w.code](w)}`);
    }

    // Prompts. Each answer read is echoed on its own line, as the console does with piped input.
    const answers = c.answers.slice();
    const next = () => {
      if (answers.length === 0) return null;
      const a = answers.shift();
      out.push(a);
      return a;
    };
    const ask = (check, allowEof) => {
      for (;;) {
        const a = next();
        if (a === null && !allowEof) return null;
        const res = check(a === null ? '' : a);
        if (!res.error) return res.value;
        out.push(`WARNING: ${PROMPT_WARNINGS[res.error](res)}`);
        if (a === null) return null;
      }
    };
    // At the end of input Read-Host returns nothing: procescode and creatiedatum take their default.
    const procescode = ask(DPIA100.checkProcescode, true);
    const creatiedatum = ask(a => DPIA100.checkCreatiedatum(a, c.today ? new Date(c.today) : undefined), true);
    const gebruiker = ask(DPIA100.checkGebruiker, false);
    if (gebruiker === null) {
      out.push('[UNEXPECTED ERROR] No answer for Creatiegebruiker.');
      return result(1, null);
    }
    if (r.warnings.length > 0) {
      const a = next();
      if (!/^(y|yes)$/i.test(DPIA100.trim(a === null ? '' : a))) return fail('Cancelled. No file was created.');
    }

    const bytes = DPIA100.encodeOutput(DPIA100.buildLines(r.records, { procescode, creatiedatum, gebruiker }));
    out.push('[SUCCESS] File successfully created:');
    out.push(c.outputName || DPIA100.outputFileName(c.fileName));
    out.push(`Records parsed : ${r.records.length}`);
    if (r.skippedEmptyRows > 0) out.push(`Empty rows skipped: ${r.skippedEmptyRows}`);
    if (r.warnings.length > 0) out.push(`Written with ${r.warnings.length} confirmed warning(s), see above.`);
    return result(0, bytes);
  }

  // Entry point: called with the case list, logs the results for WebTest.ps1 to pick up.
  window.runHarness = cases => {
    let payload;
    try {
      payload = {
        version: DPIA100.VERSION,
        encodable: DPIA100.encodableCodePoints(),
        cases: cases.map(c => {
          try { return runCase(c); } catch (e) { return { name: c.name, exitCode: -1, output: null, transcript: `HARNESS EXCEPTION: ${e.stack || e}` }; }
        })
      };
    } catch (e) {
      payload = { error: String(e.stack || e) };
    }
    console.log('DPIA100-RESULT:' + toBase64(new TextEncoder().encode(JSON.stringify(payload))));
  };
})();
