# CLAUDE.md: DPIA100 converter

Notes for an AI agent or developer working on this project. All paths are relative to the project
root. The user guide is `README.md` (Dutch), the file format is in `field-layout.md`, and history is
in `docs/CHANGELOG.md`.

## What this is

`dpia100.ps1` converts a CSV of HR mutations into `dpia100.txt`, the fixed-width import file for
Visma/Raet **Beaufort** (function 100254, *Import externe indirecte mutaties*). It replaces the official
Excel tool's "Maak DPIA100" button with stricter validation. Users are non-technical colleagues at
a Dutch municipality who drag a CSV onto `dpia100.bat`.

## Files

| Path | Purpose |
|---|---|
| `dpia100.ps1` | The script, in the root. All logic is in this one file. |
| `dpia100.bat` | Launcher, in the root next to the script; contents are also in `README.md`. |
| `README.md`, `field-layout.md` | User guide (Dutch) and file format reference, in the root. |
| `CLAUDE.md` | These notes, in the root. |
| `docs/CHANGELOG.md` | Change history. |
| `tests/run-tests.ps1` | Regression suite: every case in `tests/cases` on PowerShell 5.1 and 7. |
| `tests/beaufort/` | Probe files for what Beaufort accepts (`make-probes.ps1`), results in `findings.md`. |
| `docs/DPIA100 record indeling.pdf` | Beaufort help (2024): record layout. Primary spec. Local only, not in the repo (copyrighted). |
| `docs/Microsoft Word - Handboek DPIA100 via Excel.doc.pdf` | Raet handbook (2005): Excel tool and Beaufort setup. Local only, not in the repo (copyrighted). |

## Hard constraints (do not break)

- **Output bytes.** Windows-1252, exactly 187 bytes per record, CRLF line endings. Output must stay
  byte-identical to the Excel tool for the same input (see the example in `field-layout.md`).
- **Alignment.** pers_nr is right-aligned in 6 positions, dv right-aligned in 3, and creatiegebruiker
  right-aligned in 16. Both dates are written as 3 spaces + `ddMMyyyy`. This was confirmed against
  Excel tool output, not taken from the docs.
- **Fail fast, report everything.** Any validation error aborts with no output file. All errors are
  collected into one table first. The only non-fatal cases are Waarde warnings: truncation over 50
  characters and amounts Beaufort may misread (decisions 3 and 8). With warnings, the file is only
  written after the user types `Y`; anything else (also end of input) cancels like an error.
- **Nothing changes silently.** Do not "clean" input into a different valid-looking value (for
  example stripping non-digits, or writing `?` for characters that can't be encoded). Reject instead.
  The comma -> dot conversion of amounts is the one deliberate change, and it is always listed: as a
  report line in a `;` file (Dutch notation), otherwise as a warning.
- **Windows PowerShell 5.1 compatible.** `dpia100.bat` calls `powershell.exe`, which is always 5.1.
  Also keep it working on PowerShell 7.
- **Script file encoding.** Save `dpia100.ps1` as UTF-8 **with BOM** and CRLF. Without a BOM, 5.1
  reads the file as ANSI.
- **Keep the .bat minimal.** It only launches the script. Put logic and defaults in the `.ps1`.
- **File name starts with `DPIA100`.** Beaufort only finds import files with that prefix (case does not
  matter). The script rejects any other `-OutputTxtPath` name before the stale-file logic is armed, so
  an unrelated file is never renamed to `.old` (decision 12). Probe files follow the same rule.

## Script structure

Top to bottom:
1. Comment-based help: description, layout table, `.KEY_ARCHITECTURAL_DECISIONS` (numbered),
   `.NOT_IN_CURRENT_SPEC` (A-D).
2. `param`: `-InputCsvPath` (empty means a file picker), `-OutputTxtPath` (default `.\dpia100.txt`),
   `-NoPause`.
3. Exit handling:
   - `Exit-Script` is the single exit point. It renames a stale output file to `.old` when the code
     is not 0, and pauses unless `-NoPause` is passed.
   - `Stop-WithError` prints a message and calls `Exit-Script`.
   - A `trap` catches unexpected terminating errors.
4. Constrained Language Mode check, strict Windows-1252 encoder, `Get-UnencodableChars`,
   `Read-CsvFile` (raw bytes plus encoding detection).
5. Input path resolution and the file picker. The output path is resolved early, so every failure
   can retire a stale file.
6. Helpers:
   - `Detect-CsvDelimiter`
   - `Split-CsvHeader` (a custom parser for the header record)
   - `Get-UniqueHeaderNames`
   - `Test-And-ParseDate`
   - `Find-PropertyByPattern` (patterns of 2 characters or fewer must match the whole header)
7. Prompts: procescode, creatiedatum, creatiegebruiker.
8. CSV parsing: the header via `Split-CsvHeader`, the data rows via `ConvertFrom-Csv -Header`.
9. Header mapping. The mapping is printed; the script aborts on a missing field or on one column
   mapped to two fields.
10. Per-row validation into `$validationErrors` and `$preparedRecords`. Completely empty rows are
    skipped.
11. Abort on errors, build the 187-character lines with a length check, then write with the strict
    encoder.

## Conventions

- Comments use `# DECISION:`, `# ASSUMPTION:`, `# NOTE:` and `# TODO:`. Header decisions are numbered
  and referenced from the code, for example "(decision 8)" or "(.NOT_IN_CURRENT_SPEC B)". Keep them
  in sync when changing behaviour.
- Field labels in messages: `Rubriek`, `Pers. Nr`, `Dienstverband`, `S of V`, `Mutatiewaarde`,
  `Ingangsdatum`, `Productiesoort`.
- Validation errors are `[PSCustomObject]@{ Row; Column; Value; Issue }`. Row 1 is the header.
- Update `docs/CHANGELOG.md` with every behavioural change, and run the test suite.
- Keep `dpia100.ps1` ending in CRLF; a missing LF at the end makes git treat it as binary.
- Don't edit `.ps1` files with Git Bash `sed -i`: it silently converts CRLF to LF. Use the Edit tool or
  PowerShell.
- Curly quotes (`’ ‘ “ ”`) are string delimiters in PowerShell. Put text containing them in a
  single-quoted here-string (`@'...'@`), and save scripts with non-ASCII literals as UTF-8 with BOM.

## Scope decisions by the owner

- **In scope:** PS and DV only, productiesoort NOR only, verrekenperiode always blank.
- **No rubriek lookup list.** Rubrieken are created freely in Beaufort, so a list would go out of
  date. PS or DV is derived from whether Dienstverband is filled. Beaufort rejects a DV rubriek with
  a PS key at import (PI_1650), but in probe round 1 a PS rubriek with a DV key was **accepted**; see
  probe P04-7 in `tests/beaufort/findings.md`.
- **Header matching stays flexible.** Users rename headers, and the data validation acts as a safety
  net. The printed mapping covers Mutatiewaarde, which can't be validated.
- **Row numbers are not mapped to physical CSV lines.** Users check the CSV or the Excel source.
- **Waarde is free.** An empty Waarde is allowed, and no format is forced (dates, times, text are
  passed through as typed). Beaufort validates the value against the rubriek.
- **Amounts: warn, don't block (decision 8).** Probes P08 showed that Beaufort reads the dot as the
  decimal point (`3.250` is silently 3,25) and rejects a comma in a numeric rubriek at processing
  (PI_2229), not at import. The script can't tell a numeric rubriek from a text rubriek, so it converts
  `12,50` to `12.50` and warns about thousand-separator shapes, but never rejects them. The delimiter
  hints at the notation: `;` = Dutch Excel (decimal comma), `,` = English Excel (decimal dot), tab = no
  hint. The full table is in decision 8; cases `ok-amounts*` cover each delimiter.
- **The official PDF docs stay leading.** Beaufort accepts more at import than the docs allow
  (lowercase codes, other alignments, leading zeros, tabs, 186/188-byte lines; see
  `tests/beaufort/findings.md`). Don't relax the script because of that.
- **Duplicate rows are up to the user.** Beaufort creates both; the script does not check.
- **Beaufort is the safety net for rubriek-level errors.** It rejects only the records with errors and
  reports them clearly, including S/V mismatches. Don't duplicate those checks.
- **Candidates for later** are listed in `.NOT_IN_CURRENT_SPEC`: other object types, verrekenperiode
  plus JWP, and the rubriek lookup.

## Testing

- **Regression suite.** `pwsh tests\run-tests.ps1` (optionally `-Filter err-*`). Runs each folder in
  `tests/cases` on `powershell.exe` (5.1) and `pwsh`. A case has `input.csv` plus either
  `expected.txt` (exact output bytes) or no `expected.txt` (must fail: exit 1, no output, stale file
  renamed to `.old`). Optional `expected-output.txt` (console lines that must appear), `answers.txt`
  (prompt answers; default `CMT` / `07 10 2026` / `TSTUSER` / `Y`, the last one confirms warnings) and `output-name.txt` (output file name;
  default `DPIA100_<case>.<engine>.txt`). Run it after every script change.
- **Fixtures are raw bytes** (`.gitattributes`: `-text`). Write new ones with
  `[IO.File]::WriteAllBytes`, not an editor that may change encoding or line endings. Build
  `expected.txt` from the layout in `field-layout.md`, not by copying the script's own output.
- **Gold reference.** `ok-reference-excel-tool` holds the two Excel-tool records from
  `field-layout.md`. Add more real Excel-tool output there when available.
- **Special characters.** `ok-chars-<encoding>` runs the same Dutch character set (accents, `€`, curly
  quotes, dashes, ASCII punctuation) through ANSI, UTF-8, UTF-8 BOM, UTF-16 LE and BE input.
  `ok-chars-decomposed` covers combining accents; `err-chars-outside-1252` covers `ĳ`, `Ł`, `ş`, emoji,
  etc. The runner captures console output in the OEM code page, so match error text on ASCII parts
  such as `(U+0133)`.
- **Prompt answers are ASCII only** in tests: piping non-ASCII to a child PowerShell depends on console
  code pages.
- **Beaufort behaviour.** `tests/beaufort/make-probes.ps1` writes raw probe files (bypassing the
  script) into `tests/beaufort/probes/` (git-ignored): `DPIA100_probes.txt` with everything that can
  share one file, and `DPIA100_P07a-lf.txt`. Keep it to as few files as possible, because each file is
  a manual import for the owner. Beaufort reads line by line, so line-level variations (length,
  blank line) can share the main file; only whole-file variations (line endings, encoding) need their
  own file. The owner imports them and saves the reports in `tests/beaufort/Findings <date>/`
  (git-ignored: server paths and names). Results go into `tests/beaufort/findings.md`.
- **Headless runs.** Pipe the prompt answers in (add `Y` to confirm warnings) and pass `-NoPause`
  (works on 5.1 and 7): `printf 'CMT\n07 10 2026\nTSTUSER\nY\n' | pwsh -NoProfile -File dpia100.ps1 -InputCsvPath t.csv -OutputTxtPath out.txt -NoPause`
- **pwsh on Linux.** It writes LF line endings, so convert to CRLF before comparing.
  `DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1` breaks NFC normalization, so test the decomposed `ö` case
  with ICU available.
- **Real Windows.** The suite covers 5.1 via `powershell.exe -File`, not the `.bat` itself. Untested:
  the machine with a GPO-enforced `RemoteSigned` policy. Under RemoteSigned the
  script runs if it is local and unblocked; a downloaded copy needs `Unblock-File`.

## Working with the owner

- Read the current files before changing anything; the owner edits the script between sessions.
- When asked for a review or report, give findings only, not code. Implement only what the owner
  selects.
- Be concise. Ask before decisions that change scope or behaviour.
