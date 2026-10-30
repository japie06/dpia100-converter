# CLAUDE.md: DPIA100 converter

Notes for an AI agent or developer working on this project. All paths are relative to the project
root. The user guide is `README.md` (Dutch), the file format is in `field-layout.md`, and history is
in `docs/CHANGELOG.md`.

## What this is

`dpia100.ps1` converts a CSV of HR mutations into `DPIA100_<csv name>.txt` (next to the CSV), the fixed-width import file for
Visma/Raet **Beaufort** (function 100254, *Import externe indirecte mutaties*). It replaces the official
Excel tool's "Maak DPIA100" button with stricter validation. Users are non-technical colleagues at
a Dutch municipality who drag a CSV onto `dpia100.bat`.

`dpia100.html` is the same converter as a web page (decision 15): one self-contained file, Dutch UI,
everything local in the browser, opened from the release zip or GitHub Pages. It is a line-by-line
JavaScript port of the script, not a wrapper: a browser cannot run PowerShell. The two share the spec
(`field-layout.md`, the numbered decisions), the fixtures in `tests/cases`, and a differential test.

## Files

| Path | Purpose |
|---|---|
| `dpia100.ps1` | The script, in the root. All logic is in this one file. |
| `dpia100.bat` | Launcher, in the root next to the script; contents are also in `README.md`. |
| `dpia100.html` | The web version, in the root: CSP, style, core script (logic), UI script (Dutch), all inline. |
| `docs/privacy-en-beveiliging.md` | Dutch explanation of the web version's privacy guardrails, for IT and privacy officers. |
| `tests/web/WebTest.ps1` | Helpers to run pages in headless Edge, the CSP policy and the static page checks. |
| `tests/web/harness.js` | Runs the cases on the core of `dpia100.html` and produces the PowerShell console lines. |
| `package.json`, `playwright.config.mjs` | Developer test tooling for the page (`npm test`): node:test and Playwright on the installed Edge. Not needed by users. |
| `tests/web/core.test.mjs`, `tests/web/core.mjs` | Unit tests of the page's core (node:test) and the loader that takes the core from the page. |
| `tests/web/ui.spec.mjs` | Playwright UI tests: file input and download first, the open and save dialogs stubbed. |
| `tests/web/update-csp.ps1` | Rewrites the CSP hashes in `dpia100.html` after a change to an inline block. |
| `tests/web/differential.ps1` | Random CSV files through the script and the page; both must agree. |
| `README.md`, `field-layout.md` | User guide (Dutch) and file format reference, in the root. |
| `voorbeeld.csv` | Example input (Windows-1252, `;`, CRLF) shipped in the release zip and shown in `README.md`. Byte-identical to `tests/cases/ok-voorbeeld/input.csv`; change both together. |
| `CLAUDE.md` | These notes, in the root. |
| `docs/CHANGELOG.md` | Change history, one `## vX.Y.Z (date)` section per release. |
| `LICENSE` | MIT. |
| `.github/workflows/ci.yml` | Runs the tests on every push/PR; a `v*` tag publishes a release zip and deploys the page to GitHub Pages. |
| `tests/run-tests.ps1` | Regression suite: every case in `tests/cases` on PowerShell 5.1, 7 and the page in Edge, plus page checks. |
| `tests/beaufort/` | Probe files for what Beaufort accepts (`make-probes.ps1`), results in `findings.md`. |
| `docs/DPIA100 record indeling.pdf` | Beaufort help (2024): record layout. Primary spec. Local only, not in the repo (copyrighted). |
| `docs/Microsoft Word - Handboek DPIA100 via Excel.doc.pdf` | Raet handbook (2005): Excel tool and Beaufort setup. Local only, not in the repo (copyrighted). |
| `docs/DPIA100-Excel-tool/` | Source of the official Excel tool (`DPIA100.xlsm` V 06.12.6, VBA exports). Local only, git-ignored: copyrighted, contains the sheet password and the municipality's rubriek list. Summary in `field-layout.md` (*Excel tool*). |

## Hard constraints (do not break)

- **Output bytes.** Windows-1252, exactly 187 bytes per record, CRLF line endings. Output must stay
  byte-identical to the Excel tool for the same input (see the example in `field-layout.md`).
- **Alignment.** pers_nr is right-aligned in 6 positions, dv right-aligned in 3, and creatiegebruiker
  right-aligned in 16. Both dates are written as 3 spaces + `ddMMyyyy`. This was confirmed against
  Excel tool output, not taken from the docs.
- **Fail fast, report everything.** Any validation error aborts with no output file. All errors are
  collected into one table first. The only non-fatal cases are Waarde warnings: truncation over 50
  characters, amounts Beaufort may misread, empty values (a delete in Beaufort) and dates or times not
  in the Excel-tool shape (decisions 3, 8 and 16). With warnings, the file is only
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
- **Two front ends, one set of rules (decision 15).** Every rule change goes into `dpia100.ps1` **and**
  the core script of `dpia100.html`, with the same decision references. The English console texts are
  mirrored in `tests/web/harness.js` (so `expected-output.txt` checks the page too); the page itself
  shows Dutch texts. Run the suite and `tests/web/differential.ps1` after any change to either.
- **The page stays self-contained and offline.** No external resources, libraries, fonts, build step or
  minification. Its scripts may not use network, navigation or code-from-string APIs (`fetch`,
  `XMLHttpRequest`, `location`, `open()`, `innerHTML`, `eval`, ...); `page-security` enforces the list in
  `tests/web/WebTest.ps1`. Build DOM with `textContent`/`createElement` only (Trusted Types is on).
- **CSP hashes.** The CSP allows only the inline blocks with the hashes it lists. After changing the
  style, core or UI block run `pwsh tests\web\update-csp.ps1`; otherwise the browser refuses the block.
  Hashes are over LF-normalized text, so line endings do not matter. Keep the three blocks (ids
  `dpia100-style`, `dpia100-core`, `dpia100-ui`) and no others; no inline event handlers or `style=`.
- **The core has no DOM access.** It takes bytes and returns plain objects with codes; the UI and the
  harness turn codes into Dutch and English text. Emulate .NET where it differs from JS (whitespace set
  of `Trim`/`\s`, `ToUpper` per character, `\d` = any Unicode digit); the comments name each case.
- **The page never touches existing files.** It writes only when the user saves; in the save dialog it
  refuses names not starting with `DPIA100`; after a download it warns about `(1).txt`. Never add
  storage of CSV content; `localStorage` holds only procescode and user name, and only when ticked.

## Script structure

Top to bottom:
1. Comment-based help: description, layout table, `.KEY_ARCHITECTURAL_DECISIONS` (numbered),
   `.NOT_IN_CURRENT_SPEC` (A-D).
2. `param`: `-InputCsvPath` (empty means a file picker), `-OutputTxtPath` (empty means
   `DPIA100_<csv name>.txt` next to the input CSV, decision 13), `-NoPause`.
3. Exit handling:
   - `Exit-Script` is the single exit point. It renames a stale output file to `.old` when the code
     is not 0, and pauses unless `-NoPause` is passed.
   - `Stop-WithError` prints a message and calls `Exit-Script`.
   - A `trap` catches unexpected terminating errors.
4. Constrained Language Mode check, strict Windows-1252 encoder, `Get-UnencodableChars`,
   `Read-CsvFile` (raw bytes plus encoding detection).
5. Input path resolution and the file picker. The output path is derived from the input when not
   given, and resolved early, so every failure can retire a stale file.
6. Helpers:
   - `Detect-CsvDelimiter`
   - `Split-CsvHeader` (a custom parser for the header record)
   - `Get-UniqueHeaderNames`
   - `Test-And-ParseDate`
   - `Find-PropertyByPattern` (patterns of 2 characters or fewer must match the whole header)
7. CSV parsing: the header via `Split-CsvHeader`, the data rows via `ConvertFrom-Csv -Header`, with a
   line break prefixed so a first row starting with `#` is not taken for a `#TYPE` line.
8. Header mapping. The mapping is printed; the script aborts on a missing field or on one column
   mapped to two fields.
9. Per-row validation into `$validationErrors` and `$preparedRecords`. Completely empty rows are
   skipped. Empty, date-like and time-like values are collected for one grouped warning each
   (`Join-Examples`). Abort on errors, otherwise print `[OK]` and list the warnings.
10. Prompts: procescode, creatiedatum, creatiegebruiker. They come after the validation, so a broken
    CSV never asks anything (decision 14). Then the `Y` confirmation, only with warnings.
11. Build the 187-character lines with a length check, then write with the strict encoder.

## Web page structure (`dpia100.html`)

1. Comment, then `<meta charset>` and directly after it the CSP meta tag (before anything can load).
2. `<style id="dpia100-style">`: light and dark theme via CSS variables, works down to phone width.
3. Static markup: header, step 1 (drop zone, file bar), step 2 (check), step 3 (form), privacy section,
   footer. A `#startup-error` banner shows until the UI script removes it (JavaScript off or blocked).
4. `<script id="dpia100-core">`: `const DPIA100`, the port. `analyze(bytes, fileName)` covers everything up
   to the row validation (decision 14) and returns `{ fatal, errors, warnings, commaConversions,
   records, mapping, ... }`. Also `checkProcescode`, `checkCreatiedatum`, `checkGebruiker`, `buildLines`,
   `encodeOutput`, `outputFileName`, `isValidOutputName`. Each function names its PowerShell counterpart.
   The data rows are parsed by a port of PowerShell's `ImportCsvHelper` (`CsvCommands.cs`), quirks
   included; a lone CR is data, as in 5.1.
5. `<script id="dpia100-ui">`: Dutch texts per code, rendering, file choice (open dialog with a handle
   for *Opnieuw controleren*, plain file input otherwise), save (save dialog, otherwise Blob download),
   remembered settings. Logs `DPIA100 converter <version> gestart` (the smoke test waits for it).

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

- **In scope:** PS and DV only, productiesoort NOR or COR, verrekenperiode always blank. Rubriek exactly
  6 characters.
- **No rubriek lookup list.** Rubrieken are created freely in Beaufort, so a list would go out of
  date. PS or DV is derived from whether Dienstverband is filled. Beaufort rejects a DV rubriek with
  a PS key at import (PI_1650), but in probe round 1 a PS rubriek with a DV key was **accepted**; see
  probe P04-7 in `tests/beaufort/findings.md`.
- **Header matching stays flexible.** Users rename headers, and the data validation acts as a safety
  net. The printed mapping covers Mutatiewaarde, which can't be validated.
- **Row numbers are not mapped to physical CSV lines.** Users check the CSV or the Excel source.
- **Waarde is free.** No format is forced (dates, times, text are passed through as typed). Beaufort
  validates the value against the rubriek. An empty Waarde is allowed but means a deliberate delete of the
  current value, so it gets a warning; dates not as `dd-mm-yyyy` and times not as `hh:mm` (what the Excel
  tool wrote) get a warning too (decision 16). Each kind is one warning listing the rows.
- **The Excel tool source is a reference, not the truth.** It is from 2005 and has bugs (encoding by save
  path, quoting, Dutch-locale formulas); see `field-layout.md`. Use it to understand layout choices and
  unbuilt features (BG, LV, SD, IZ widths; verrekenperiode), then probe in Beaufort.
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

- **Regression suite.** `pwsh tests\run-tests.ps1` (optionally `-Filter err-*`, `-Engine edge`). Runs
  each folder in `tests/cases` on `powershell.exe` (5.1), `pwsh` and `edge` (the page). A case has `input.csv` plus either
  `expected.txt` (exact output bytes) or no `expected.txt` (must fail: exit 1, no output, stale file
  renamed to `.old`). Optional `expected-output.txt` (console lines that must appear, **in that
  order**), `unexpected-output.txt` (whole lines that must not appear), `answers.txt`
  (prompt answers; default `CMT` / `07 10 2026` / `TSTUSER` / `Y`, the last one confirms warnings) and `output-name.txt` (output file name;
  default `DPIA100_<case>.<engine>.txt`). An empty `default-output` marker runs the case without
  `-OutputTxtPath` on a copy `out\<case>.<engine>.csv`, so the script's default name lands on that same
  output path. Run it after every script change.
- **Order of operations (decision 14).** Prompt text is not captured when stdin is piped, but each answer
  read is echoed on its own line. The runner uses that: a failure case without `answers.txt` must not
  show `TSTUSER` or `[OK]`. The `ok-order-*` cases expect the echo, so they fail if it ever disappears
  and the check would pass silently. GitHub Actions also runs it
  on `windows-latest` for every push and pull request.
- **The `edge` engine.** All cases run in one headless Edge page: the core block copied from
  `dpia100.html` plus `tests/web/harness.js`, which answers the prompts like piped input and returns the
  exit code, output bytes and the English console lines. Path-only console lines (`Input file`,
  `Output file`, the path after `[SUCCESS]`, `Encoding`) are not checked for `edge`, and there is no
  `.old` check: the page never touches files. `msedge.exe` hands off to a detached process, so stdout is
  lost; pages report through `console.log`, which `--enable-logging` writes to `chrome_debug.log` in a
  throwaway profile. Never let a test page start a real download: it hangs headless Edge.
- **Page checks** (with `edge`, `-Filter page-*` for only these): `page-security` (CSP equals the
  canonical policy, no external resources, forbidden-API scan), `page-version` (equals `$ScriptVersion`),
  `page-1252-table` (the page's encoder equals .NET's strict 1252), `page-smoke-crlf`/`-lf` (the real
  page starts under its CSP without violations).
- **Unit and UI tests of the page (`npm test`, Node.js LTS).** The roles are separate: `run-tests.ps1` is
  the authoritative parity runner; `npm test` runs the fast unit tests and the UI interactions.
  - `npm run test:unit`: `node:test`, no dependencies. `tests/web/core.test.mjs` tests the core
    function by function (dates, delimiter, CSV quirks, .NET semantics, the decision 8 table per
    delimiter, 1252, layout, prompts). `tests/web/core.mjs` loads the core block from the page in a `vm`.
  - `npm run test:ui`: Playwright Test, `tests/web/ui.spec.mjs`, on the **installed Edge**
    (`channel: 'msedge'` in `playwright.config.mjs`). Never run `npx playwright install`; set
    `PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1` for `npm ci`. The main route is primary: dialogs switched off,
    a real file chooser and a real Blob download, with exact bytes, warnings confirmation, errors,
    *Onthoud* across a reload, form checks and drag and drop. The open and save dialogs are tested
    separately with stubs (save, re-check, refused name, changed-file guard). Every test fails on a CSP
    violation or script error. Failure traces and screenshots go to `tests/out/playwright`.
  - No linters for now (ESLint, html-validate): `page-security` and the CSP hashes cover the risks.
  - Keep `@playwright/test` pinned exactly in `package.json` and commit `package-lock.json`.
- **Differential test.** `pwsh tests\web\differential.ps1 [-Count N] [-Seed S] [-Engine powershell]`
  generates random CSVs (encodings, delimiters, quoting quirks, broken values) into `tests/out/diff`,
  runs each through the script and all through the page, and compares exit codes, bytes, key console
  lines and error positions. CI runs 40 files with the run number as seed; rerun a failure with that
  seed. It captures PowerShell output as UTF-8 (`-Command` with `[Console]::OutputEncoding`).
- **Visual check of the page.** Headless Edge can screenshot a state: inject a script that loads a
  fixture through the file input into a copy of the page (add its hash with `Set-PagePolicy`) and run
  `msedge --headless --screenshot=x.png --window-size=1280,2000 --virtual-time-budget=5000 <file URL>`.
  The window cannot be narrower than about 500 px; add `--blink-settings=preferredColorScheme=1` for
  light mode or `--force-dark-mode` for dark.
- **Manual checks in a real browser** (not covered headless): drag and drop from Explorer, the real save
  dialog (overwrite prompt, a name without DPIA100), *Opnieuw controleren*
  after saving the CSV in Excel while it is still open, the page from a network share, and Firefox
  (download path).
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
  `make-probes.ps1 -Web` (round 4) tests the web app instead. It writes one probe CSV of valid input,
  `web-probes.csv`, and lets the page's core make `DPIA100_web-probes.txt` through the `edge` harness.
  The owner can recreate that file in the real page and compare. Every W row gets its own ingangsdatum, so
  processing does not overwrite values and each stored value can be checked. Round 4 found that Beaufort
  reported the bytes 0x80-0x9F and 0xFF (`€`, curly quotes, dashes) as missing in the processing report.
  That is **unverified**: the Beaufort GUI accepts these characters when typed. Retest in #20; the
  script keeps accepting them (owner).
- **Headless runs.** Pipe the prompt answers in (add `Y` to confirm warnings) and pass `-NoPause`
  (works on 5.1 and 7): `printf 'CMT\n07 10 2026\nTSTUSER\nY\n' | pwsh -NoProfile -File dpia100.ps1 -InputCsvPath t.csv -OutputTxtPath out.txt -NoPause`
- **pwsh on Linux.** It writes LF line endings, so convert to CRLF before comparing.
  `DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1` breaks NFC normalization, so test the decomposed `ö` case
  with ICU available.
- **Real Windows.** The suite covers 5.1 via `powershell.exe -File`, not the `.bat` itself. Tested by
  the owner on a municipality PC (#4, 2026-10-29): `RemoteSigned` at CurrentUser/LocalMachine, no
  GPO (MachinePolicy/UserPolicy `Undefined`), FullLanguage. The `.bat` (drag-and-drop and file picker)
  and *Run with PowerShell* run even with blocked files from a downloaded zip, because the Process-scope
  Bypass wins. Only `.\dpia100.ps1` from a PowerShell window fails ("not digitally signed") until
  unblocked. Output next to a CSV in a OneDrive/SharePoint folder works. A GPO-enforced policy was not
  seen; if it ever is, the `.bat` Bypass is overridden and unblocking becomes required.

## Releasing

- Open work is tracked in GitHub Issues (labels `beaufort-probe`, `enhancement`, `later`, `docs`, `bug`),
  not in a markdown checklist.
- To release: bump `$ScriptVersion` in `dpia100.ps1` and `VERSION` in the core of `dpia100.html` (then
  `update-csp.ps1`), add a `## vX.Y.Z (date)` section to `docs/CHANGELOG.md`, merge to main, then
  `git tag vX.Y.Z` and push the tag. CI runs the tests, checks that the tag matches both versions,
  publishes a release with the zip (html, ps1, bat, voorbeeld.csv, README, LICENSE) and that changelog
  section as notes, and deploys `dpia100.html` to GitHub Pages as `index.html`. Pages must be enabled
  once (Settings > Pages > Source: GitHub Actions). Version scheme: patch = fix, minor = new behaviour,
  major = output format change.

## Working with the owner

- Read the current files before changing anything; the owner edits the script between sessions.
- When asked for a review or report, give findings only, not code. Implement only what the owner
  selects.
- Be concise. Ask before decisions that change scope or behaviour.
