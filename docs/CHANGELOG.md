# Changelog

All notable changes to `dpia100.ps1`, `dpia100.bat` and `dpia100.html`. Newest first. Format based on
[Keep a Changelog](https://keepachangelog.com/). Releases are numbered from v1.0.0 (tag `vX.Y.Z`,
zip on the GitHub Releases page); older entries are dated.

## v1.2.0 (2026-10-30)

### Added
- **Web version `dpia100.html`** (decision 15): one self-contained page that does what the script does,
  without a console and without installing anything. It works offline from the release zip or a network
  share, and is published on GitHub Pages (<https://japie06.github.io/dpia100-converter/>) for every
  release. The interface is Dutch:
  - Step 1: drop or choose the CSV.
  - Step 2: the check, with the column mapping, every error with a fix hint, warnings, and a preview of
    the records.
  - Step 3: procescode, creatiedatum and creatiegebruiker. With warnings, a checkbox replaces typing `Y`.
  - In Edge and Chrome, *Opnieuw controleren* re-reads the CSV after it was fixed in Excel.
  - The file is saved through a save dialog where available, or as a download otherwise. After a
    download the page warns that Edge may name a second copy `… (1).txt`.
  - *Onthoud* stores the procescode and user name in the browser; nothing from the CSV is ever stored.
  - Same rules and byte-identical output as the script. The page cannot rename an old output file
    (decision 10); it writes nothing until the user saves, and refuses a file name that does not start
    with `DPIA100` (decision 12).
- Privacy guardrails for the page:
  - A Content Security Policy that allows no network access and only the page's own hashed inline blocks.
  - Trusted Types.
  - No external resources.
  - A test (`page-security`) that rejects network, navigation and code-from-string APIs in its scripts.
  - Explained for IT and privacy officers in `docs/privacy-en-beveiliging.md`.
- Tests: `tests/run-tests.ps1` runs every case also on the page, in headless Edge (`-Engine edge`), plus
  checks of the CSP, the version, the Windows-1252 table and a smoke test of the real page (CRLF and LF).
  `tests/web/differential.ps1` compares the script and the page on random CSV files.
  `tests/web/update-csp.ps1` rewrites the CSP hashes after a change.
- Developer test tooling for the page, with Node.js (#19): `npm test` runs unit tests of the core
  (`node:test`) and UI tests with Playwright in the installed Edge, with the file input and Blob download
  as the primary route and the open and save dialogs stubbed. CI runs it in a separate `web` job.
- The *Privacy en veiligheid* section of the page is collapsed by default; the link in the header opens it.
- Beaufort probe round 4 for the web app: `tests/beaufort/make-probes.ps1 -Web` writes one probe CSV and
  lets the page make one import file from it (`DPIA100_web-probes.txt`).
- CI runs the web tests and the differential test, checks the page version on a release tag, adds
  `dpia100.html` to the zip and deploys it to GitHub Pages.

### Fixed
- A first data row starting with `#` (for example an Excel `#N/B` in the first column) was dropped
  without a message: `ConvertFrom-Csv` read it as a `#TYPE` line. It is now validated like any row. Test
  `err-hash-first-row`.
- When the header had as many of two delimiters (for example 6 `;` and 6 `,`), PowerShell 7 chose one at
  random on each run. Ties now go in the fixed order `;`, tab, `,` (what Windows PowerShell 5.1 already
  did). Test `ok-delimiter-tie`.

## v1.1.0 (2026-10-29)

### Added
- `voorbeeld.csv`: an example input file, included in the release zip and shown in the README. Test case
  `ok-voorbeeld` keeps it valid.

### Changed
- The CSV is checked before the three questions (decision 14, #11). On errors the script stops right
  away, without asking procescode, creatiedatum and creatiegebruiker. Without errors it prints
  `[OK] N row(s) checked, no errors found.`, lists any warnings, asks the questions, and then asks the
  `Y` confirmation for warnings (`Create the file anyway, with the N warning(s) above?`).
- Test runner: `expected-output.txt` lines must now appear in order, a new `unexpected-output.txt` lists
  lines that must not appear, and a failure case without `answers.txt` must stop before the prompts.
  New cases `ok-order-no-warnings` and `ok-order-warnings`; the `err-warnings-*` cases check the order.
- The output file is written next to the input CSV and named after it: `mutaties oktober.csv` gives
  `DPIA100_mutaties oktober.txt` in the same folder (decision 13, #10). It used to be `dpia100.txt` in
  the working directory, which after a drag-and-drop onto the `.bat` was an unpredictable folder. Users
  who looked for `dpia100.txt` in a fixed place now find the file next to their CSV. An explicit
  `-OutputTxtPath` works as before. The stale-file protection (`.old`) applies to the new name.
- The input and output paths are printed up front (`Input file`, `Output file`), before the encoding and
  delimiter lines. The `[SUCCESS]` line puts the path on its own line, so a long path stays readable.
  Test cases `ok-output-default` and `err-output-default`; the test runner has a new
  `default-output` marker to run a case without `-OutputTxtPath`.
- The prompts are spelled as correct Dutch: `Creatiedatum` and `Creatiegebruiker` (were `Creatie datum`
  and `Creatie gebruiker`), also in the related warnings. Header matching is unchanged: a CSV header
  `Ingangs datum` or `Mutatie waarde` is still recognised.
- README rewritten for less technical users: an *Installeren* section (download, unblock, unzip), an
  example of the console window and of `voorbeeld.csv`, an explanation of *Column mapping*, a simpler
  section on amounts, and Dutch spelling fixes.

## v1.0.0 (2026-10-28)

First numbered release.

### Added
- The version is shown in the start banner (`=== Beaufort DPIA100 Generator v1.0.0 ===`).
- MIT license.
- GitHub Actions: the regression suite runs on PowerShell 5.1 and 7 for every push and pull request.
  A `v*` tag publishes a GitHub Release with a zip of `dpia100.ps1`, `dpia100.bat`, `README.md` and
  `LICENSE`.
- README: download from Releases, unblocking a downloaded zip, and a note that the project is not
  affiliated with Visma/Raet.

### Changed
- Republished as the public repository `dpia100-converter` with a fresh history. Personal data was
  removed from examples and tests, and the Visma/Raet PDFs are no longer included.

## 2026-10-28 (amount warnings)

### Added
- The CSV delimiter decides the number notation (decision 8). In a `;` file (Dutch Excel) the comma is
  the decimal sign: `12,50` and `1,234` become `12.50` and `1.234` and are only reported, without a
  confirmation prompt. In a `,` file (English Excel) the dot is the decimal sign: `3.250` passes without
  a warning, and `1,234` is written as typed with a thousand-separator warning. A tab file keeps
  warning in both directions. Tests: `ok-amounts`, `ok-amounts-delimiter-comma`,
  `ok-amounts-delimiter-comma-decimals`, `ok-amounts-delimiter-tab`.

### Changed
- Amounts that Beaufort may read differently are no longer rejected but give a warning, because the
  rubriek may be a text rubriek (decision 8, based on Beaufort probes P08): `3.250`, `12.345`,
  `1.234,56`, `1,234.56`, `1.234.567` are written as typed; `1,234` becomes `1.234` with a warning.
- The comma -> dot conversion (`12,50` -> `12.50`) is kept and now always listed (see Added).
- With warnings (including truncation), the script asks `Create the file anyway?` before writing. Only
  `Y` continues; anything else, also end of input, cancels without a file (decision 3).

### Fixed
- A date with dots in Waarde (`1.11.2026`) was rejected as an amount with a thousand separator.

## 2026-10-28

### Fixed
- Ingangsdatum and the creatiedatum prompt now accept `2026-11-01` and `01112026`, as documented.
  `Test-And-ParseDate` passed an untyped array to `TryParseExact`, so PowerShell picked the
  single-format overload and no format ever matched. Only the regex fallback (day-month-year with
  separators) worked.
- `dpia100.ps1` ended in a lone CR without LF, which made git treat it as a binary file.

### Added
- The output file name must start with `DPIA100` (any case); otherwise Beaufort does not find it.
  Another `-OutputTxtPath` name stops the script before any prompt, and an existing file with that name is
  left alone (not renamed to `.old`). Decision 12.
- Git repository with `.gitattributes` (test fixtures stored as raw bytes) and `.gitignore`.
- Regression suite `tests/run-tests.ps1` with 43 cases in `tests/cases`, run on PowerShell 5.1 and 7.
  It covers Dutch special characters (accents, `€`, curly quotes, dashes) in every input encoding,
  letters typed with combining accents, and characters outside Windows-1252 (`ĳ`, `Ł`, `ş`, emoji).
- Beaufort probe generator `tests/beaufort/make-probes.ps1` (special characters, file format, amounts) and
  results sheet `tests/beaufort/findings.md`. All probes that can share a file go into one file,
  `DPIA100_probes.txt`; only the LF line-ending probe is a separate file. Round 1 results are recorded.
- Project documentation: `README.md` (Dutch user guide), `CLAUDE.md` (agent/developer notes),
  `docs/CHANGELOG.md`, `field-layout.md`.

### Changed
- `CHANGELOG.md` moved to `docs/`.
- The Waarde TODO about date values is replaced by an owner decision: no forced format on Waarde.
- Amount handling (comma -> dot, thousand-separator rejection, decision 8) is marked as unconfirmed
  (TODO) until Beaufort probes P08 are done.

## 2026-10-21 (documentation in script)

### Added
- `.NOT_IN_CURRENT_SPEC` section in the script header covering:
  - other object types (BG, IZ, IP, LV, SD, ZG, ZT);
  - verrekenperiode and JWP;
  - the rubriek lookup with per-type Waarde formatting;
  - the undocumented positions 149-178.

### Changed
- Comments on the NOR-only decision, the PS/DV assumption, the object_id build and the verrekenperiode
  now say these are limits of the current spec that may be lifted later.
- Positions 149-178 are no longer called "externe referentie"; neither PDF describes them.

## 2026-10-21

### Added
- Rejects Waarde amounts with a thousand separator (`3.250`, `1.234,56`, `1,234.56`). Beaufort would
  read `3.250` as 3.25.
- Prints the column mapping (field ← CSV header) on every run.
- Aborts if one CSV column is matched to more than one field.
- After any failed run, an existing output file is renamed to `<name>.old` so a stale file cannot be
  imported. A successful run reports when it replaced an existing file.
- Stops early with a clear message when PowerShell runs in Constrained Language Mode (AppLocker/WDAC).
- `TODO` for date values in Waarde (passed through as typed).

### Changed
- The script now parses the header row itself. Empty or duplicate header names (for example a header
  ending in `;;`) are made unique, so Windows PowerShell 5.1 and PowerShell 7 behave the same.
- Header patterns of 1-2 characters (`Dv`) must match the whole header. Before, `Advies` was matched
  as Dienstverband.
- Productiesoort: only `NOR` is accepted. `JWP` is rejected because the verrekenperiode is not
  implemented.
- Field labels are consistent in all messages: Rubriek, Pers. Nr, Dienstverband, S of V, Mutatiewaarde,
  Ingangsdatum, Productiesoort.
- The script file is saved as UTF-8 with BOM and CRLF, so Windows PowerShell 5.1 reads it correctly.
- Extra header patterns: `Persoonsnummer`, `Dienstverband`, `Dv`, `Mutatiewaarde`, `Productiesoort`,
  `Productiecode`.

### Removed
- Unused `-Encoding` parameter. Output is always Windows-1252.

## 2026-10-07

### Added
- Reads the CSV as raw bytes with encoding detection (UTF-8/UTF-16 BOM, strict UTF-8, otherwise
  Windows-1252), so Windows PowerShell 5.1 and PowerShell 7 read the same file the same way.
- A strict Windows-1252 encoder: characters outside Windows-1252 are reported instead of written as `?`.
- Waarde is normalized to NFC; line breaks, tabs and control characters are rejected.
- Completely empty rows are skipped and counted.
- The console window stays open on exit (`Press ENTER to close`) unless `-NoPause` is passed.
- Catch-all for unexpected errors with a readable message.

### Changed
- Pers. Nr and Dienstverband must be digits only. Before, non-digits were stripped silently
  (`1234.0` became `12340`). A Dienstverband of all zeros counts as empty (PS record).
- The file picker now appears when no input file is given. The input path no longer has a default.
- The output path works for both relative and absolute `-OutputTxtPath`; the full path is printed.
- The input file is read with `-LiteralPath` and shared read access, so it can stay open in Excel.
- The validation error table has a fixed width, so it also prints with redirected output.
- `dpia100.bat` passes `-InputCsvPath "%~1" -NoPause` to avoid a double pause prompt.

### Verified
- Output for a 5-record sample (pers_nr lengths 2-6) is byte-identical to the official Excel tool.

## Initial version

- CSV to fixed-width `dpia100.txt` (187 characters, Windows-1252) for PS and DV mutations.
- Delimiter detection (`;`, `,`, tab), column-order independent header matching.
- Prompts for procescode (default `CMT`), creatiedatum (default today) and creatiegebruiker.
- Validation of all rows with a combined error report; abort without output on any error.
- Waarde truncated to 50 characters with a warning; Dutch decimal comma converted to a dot.
