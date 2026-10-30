# Beaufort import probes: findings

What Beaufort (100254 import, 100251/100252) does with special characters, amounts and file-format
variations. The answers decide whether the script's checks and output format are too strict, too
loose, or right.

Already known (owner, not probed): Beaufort rejects only the records with errors and reports them;
an S/V mismatch is rejected; an empty Waarde is allowed; the import file name must start with `DPIA100`.

## Owner decisions on the findings
- **The official PDF docs stay leading, not Beaufort's leniency.** Beaufort accepts a lot at import
  (lowercase codes, other alignments, leading zeros, tabs, 186/188-byte lines). The script keeps
  following the docs and the Excel-tool layout; these findings are recorded only for reference.
- **Duplicate records:** Beaufort creates both. No logic in the script; it is up to the user.
- **Line endings:** CR-only loses records silently, LF works. Recorded; the script always writes CRLF.
- **Numeric probes use a stam rubriek** (TST004): variabel rubrieken cannot be created in the test
  environment, and stam/variabel should not matter for the import of the value.

## How to run a round
1. In the test environment you need: one test person with a dienstverband that is valid on the
   ingangsdatum, two text rubrieken (type S, length 50): `$RubTxt` (PS, stam) and `$RubDv` (DV, stam),
   and one numeric rubriek `$RubNum` (DV, stam, 2 decimals) for P08.
2. Fill in the settings at the top of `make-probes.ps1` and run `pwsh tests\beaufort\make-probes.ps1`.
   To repeat only some probes: `-Only P04-5,P04-6,P08` (writes only those records).
3. Import the two files from `probes\` (the names already start with `DPIA100`):
   - `DPIA100_probes.txt`: everything else (about 200 records)
   - `DPIA100_P07a-lf.txt`: LF line endings, must be a file of its own
4. Save the import reports in a dated folder here, e.g. `Findings 20261029\` (git-ignored: they
   contain server paths and names).
5. Records missing from the report were created: check them in 100251 (is every character and key
   stored as sent?). `probes\index.txt` lists what each probe ID contains.
6. Fill in the **Result** column, then delete the test mutations in 100251.

P below is the pers_nr of the test person (not published).

## Round 1 (2026-10-28, acceptance): mostly invalid
Reports: `Findings 20261028/` (not in git). Settings: test person P, dv 1. The roles of TST002 and
TST003 were swapped in the settings (TST002 is DV, TST003 is PS), and the probes were 17 separate files.

**What happened:**
- DV rubriek TST002 with PS key `P`: every record rejected with PI_1650 *object_id niet correct*.
  This is the expected PS/DV mismatch rejection, and it is why almost all probes failed.
- PS rubriek TST003 with DV key (P04-4..6): all three **created**. A PS rubriek with a filled
  Dienstverband is **not** rejected at import. (P04-5/6 had broken keys because of a generator bug, see
  round 2, so only P04-4 counts.)
- DV rubriek P02748 with DV key `P 1` (P08): every record rejected with PI_2250. Round 2 showed
  that dv 1 is fine, so the problem is P02748 itself.

**Valid results from round 1** (these happen before the rubriek check):
- The file is parsed exactly as `field-layout.md` describes. Beaufort's report labels pos 149-178 as
  **externe ref.** (.NOT_IN_CURRENT_SPEC D).
- P07b CR-only line endings: read as **1** data line instead of 2. **The second record is lost
  without any error.**
- P07d UTF-8 file: confirms decision 1. After `ö` every field shifts one position
  (ingangsdatum `0110202`, prod. soort `NO`), so the record fails with PI_2227. Only 1 of 2 lines was
  counted (unexplained). Not repeated in round 2.
- P07e UTF-8 BOM: the BOM shifts the first record by 3 bytes, so it fails with PI_1537 (proces
  onbekend). Not repeated in round 2.
- P07i blank line: counted as a record of its own and rejected with PI_1537 (all fields empty). The
  other records are read normally.
- P07f/P07g (186/188 bytes): both lines were parsed with every field in the right place, so Beaufort
  reads by line, not in fixed 189-byte blocks. This is why round 2 can put line-level probes in one file.

## Round 2 (2026-10-28, acceptance): valid
Reports: `Findings 20261028/DPIA100_probes.txt` and `DPIA100_P07a-lf.txt`. Settings: test person P,
dv 1, TST003 (PS), TST002 (DV), P02748, creatiedatum 01012020, ingangsdatum 01102020.

`DPIA100_probes.txt`: 199 lines read, **183 created, 16 rejected**. `DPIA100_P07a-lf.txt`: 2 of 2 created.
The 16 rejections, all explained:
- P04-5, P04-6 (PI_1650): **invalid probes.** A generator bug (PowerShell variable names ignore case,
  so `$dv` overwrote `$Dv`) repeated the pers_nr in the dv part of the key. Fixed; repeat in round 3.
- P04-8 (PI_1650): DV rubriek with a PS key, rejected as expected.
- P08-1..12 (PI_2250 *object_id P 1 is ongeldig*): P04-4 used exactly the same key on DV rubriek
  TST002 and was created, so dv 1 is valid. **P02748 is not a plain DV rubriek** (or not usable for this
  person). The amount probes are still untested: pick another numeric DV rubriek.
- The blank line (PI_1537), as in round 1.

**Created means "accepted at import", not "stored correctly".** Whether characters, dates and keys
were stored as sent has to be checked in 100251 (see the last column).

## Round 3 (2026-10-28, acceptance): P04-5, P04-6, P08
Generated with `-Only P04-5,P04-6,P08`. Settings as round 2, with `$RubNum` = TST004 (DV, stam, numeric,
2 decimals). Report: `Findings 20261028/Only P04-5,P04-6,P08.txt`. **Import (100254): 14 of 14 created.
Processing (100252): 10 processed, 4 rejected.**
- P04-5/6: the DV key is accepted left-aligned with one space and with leading zeros in dv, and processed.
- P08: the import does **not** validate a numeric Waarde. Processing does:
  - Rejected (PI_2229 *Ongeldige waarde voor numerieke rubriek*): every value with a comma: `12,50`,
    `-12,50`, `1.234,56`, `1,234.56`. The report shows **one** PI_2229 message for the 4 rejections, so
    they are easy to miss.
  - Processed: `12.50`, `-12.50`, `3.250`, `3250`, `12.345`, `0012.50`, `12.5`, `12`. `3.250` and
    `12.345` (3 decimals on a 2-decimal rubriek) are processed without any message, so Beaufort reads
    the dot as a decimal point: `3.250` becomes 3,25, not 3250.
  - The stored value per probe can no longer be checked: all records are on the same rubriek, key and
    ingangsdatum, so each processed value overwrote the previous one. A repeat run with extended logging
    gave the same result. Entered manually, `12,345` is rounded to 2 decimals; assumed the same for
    import.
  - For a future amount probe: give each value its own ingangsdatum, or check in 100251 before 100252.

## Round 4 (web app): files made by dpia100.html
A happy-flow file made with the web page imported fine in the sandbox (owner, 2026-10-30). Round 4
checks the web app more widely. The page is tested byte-identical to the script, so this is an
acceptance run of the whole chain: CSV → page → Beaufort.

How to run:
1. Same settings at the top of `make-probes.ps1` as round 3. Rubrieken `$RubTxt` (PS),
   `$RubDv` (DV) and `$RubNum` (DV, numeric, 2 decimals), all stam.
2. Run `pwsh tests\beaufort\make-probes.ps1 -Web`. It writes `probes\web-probes.csv` (the input) and lets
   the page's core make `probes\DPIA100_web-probes.txt` (one file, about 150 records). The script prints
   the SHA-256 of that file.
3. Optional, to cover the real interface: open `web-probes.csv` in `dpia100.html` with the same
   procescode, creatiedatum and user, confirm the warning, and compare the saved file with
   `fc /b` (or its SHA-256). It must be identical.
4. Import `DPIA100_web-probes.txt` (100254), check in 100251, and process W05 (100252). `probes\index-web.txt`
   lists each W-ID with its CSV row.

### Round 4 run 1 (2026-10-30)
Reports: `Findings 20261030/` (not in git). Real test person; settings as round 3.

- **`DPIA100_web-probes.txt`: 151 of 151 created at import and 151 of 151 processed (100252), none
  rejected.** All keys (W01, including leading zeros and dv `0`), codes (W02), date formats (W04) and the
  converted amounts (W05: no PI_2229) were accepted end to end. This confirms the web app's output.
- **New: Beaufort drops the characters 0x80-0x9F and 0xFF** (`€ ’ “ ” – — … Š ™ ÿ`, ...) without any
  message. Seen in the processing report of the raw P02/P03 probes (see the Results table): `a€b`
  became `ab`, `d’Ancona` became `dAncona`; 0xA0-0xFE arrive unchanged. The web probes were all
  processed, but in run 1 their W rows shared a rubriek, key and date, so each value overwrote the
  previous one and the stored characters cannot be checked. Issue **#20**.
- **New: extra decimals are not rounded up.** In the report `12.345` became `12,34` (2-decimal
  rubriek): cut off, or rounded on a binary value, but not `12,35`. The report also shows how Beaufort
  read each amount: `3.250` → `3,25`, `3250` → `3.250,00`, `0012.50` → `12,50`, `12` → `12,00`.
- Report details, not data: the *Uitgebreid foutverslag* shows at most 46 characters of a Waarde (so
  P03-5 and P03-6 cannot be checked there), shows leading spaces trimmed (P03-7), a tab as a space
  (P07j), the key zero-padded (`PS 000123`), and only one of two identical records (P07k).
- A comma amount on a non-existing person gives PI_2229, not PI_0802: the value check comes before the
  object check.
- Not a finding: `DPIA100_probes.txt` and `DPIA100_P07a-lf.txt` in this run were old files with the
  placeholder person `123`, so processing failed with PI_0802 (*bovenliggend object niet aanwezig*).

**For run 2** the W rows each have their own ingangsdatum (`$Ingang`, then one day later per row, up to
about five months), so after processing every value stays in the history and can be checked. The test
person must be in service on those dates.

| ID | Probe (CSV input → page) | Import (run 1) | Check stored value (run 2) | Processing (run 1) |
|---|---|---|---|---|
| W01 | Keys: PS, DV, leading zeros, dv `0` (PS key), spaces around keys | created | Right person and dienstverband? | processed |
| W02 | Lowercase rubriek, `STAM` / `s`, productiesoort empty / `nor` | created | | processed |
| W03 | Every Windows-1252 character 0x80-0xFF typed in UTF-8; decomposed accent; curly quotes; `;` and `"` in a value | created | 0x80-0x9F and 0xFF lost? (#20) | processed |
| W04 | Ingangsdatum in 6 input formats | created | Right dates? | processed |
| W05 | Amounts `12,50` `-12,50` `3250,5` `0,25` `12.50` `3250` (numeric rubriek) | created | 12,50 / -12,50 / 3250,50 / 0,25 / 12,50 / 3250,00? | processed, no PI_2229 |
| W06 | Exactly 50 characters; 60 cut to 50 (warning); empty Waarde | created | Complete / cut / empty? | processed |

## Results

| ID | Probe | Import (round 2) | Check in 100251 | Consequence for script |
|---|---|---|---|---|
| P01 | ASCII punctuation, one character per record | all 33 created | Are `" ' ; \ < > &` shown as sent? | None if stored correctly |
| P02 | Every Windows-1252 character 0x80-0xFF | all 123 created | **Round 4 report: 0x80-0x9F and 0xFF dropped, 0xA0-0xFE kept** | Decision 7 allows all of Windows-1252: to be decided in #20 |
| P02-0xA0 | Non-breaking space | created | Round 4 report: kept | |
| P03-1..4 | Realistic Dutch text | created | Round 4 report: accents kept; `’ “ ” € –` dropped (P03-3, P03-4) | Decision 1, 7; #20 |
| P03-5 | Accented character in the last Waarde position | created | Is the `é` at the end still there? | |
| P03-6 | Waarde exactly 50 characters | created | Complete? | Decision 3 |
| P03-7 | Leading spaces | created | Round 4 report: trimmed | None: script trims Waarde |
| P03-8 | Mixed case | created | | |
| P04-1 | PS key right-aligned (Excel tool) | created | | |
| P04-2 | PS key left-aligned | **created** | Same person? | None: script keeps the Excel-tool layout (owner: docs leading) |
| P04-3 | PS key with leading zeros `0P` (zero-padded to 6) | **created** | Same person, or a different one? | None: script strips them, as the docs say (decision 4) |
| P04-4 | DV key `P   1` (Excel tool) | created | | |
| P04-5 | DV key `P 1` | **created** (round 3) | Same dienstverband? | None: script keeps the Excel-tool layout |
| P04-6 | DV key `P 001` | **created** (round 3) | Same dienstverband? | None: script strips leading zeros (decision 4) |
| P04-7 | PS rubriek with DV key | **created** (again) | What happens to the dv part? Does 100252 process it? | Script derives PS/DV from Dienstverband: a filled Dienstverband on a PS rubriek is not stopped at import |
| P04-8 | DV rubriek with PS key | PI_1650 | | Expected |
| P05-1 | Procescode lowercase `cmt` | **created** | Linked to proces CMT? | None: script keeps uppercasing (docs) |
| P05-2 | Rubriek lowercase | **created** | Linked to the right rubriek? | idem |
| P05-3 | Creatiegebruiker left-aligned | **created** | Correct user? | None: script keeps the Excel-tool layout |
| P06-1 | Ingangsdatum `dd-mm-eejj` | created | Date correct? | |
| P06-2 | Ingangsdatum `ddmmjj` (`011020`) | created | **2020 or 1920?** | |
| P06-3 | Ingangsdatum left-aligned | created | Date correct? | |
| P06-4 | Creatiedatum `dd-mm-eejj` | created | Date correct? | |
| P08-1..12 | Amounts on TST004 | all 12 **created** (round 3) | 100252: commas rejected (PI_2229, one message for 4 records); `3.250` and `12.345` processed silently; round 4 report: `12.345` → `12,34` (not rounded up) | Owner: warn, don't block (decision 8). Comma -> dot kept, listed as a warning; thousand-separator shapes get a warning |
| P07f | Line of 186 bytes | **created** | | None: script keeps exactly 187 (docs) |
| P07g | Line of 188 bytes | **created** | | idem |
| P07h | Trailing spaces removed (181 bytes) | **created** | | Editors that trim spaces do no harm |
| P07j | Tab inside Waarde | **created** | Round 4 report: shown as a space | None: script keeps rejecting control characters |
| P07k | Same record twice | **both created** | | None: up to the user (owner) |
| P07i | Blank line | PI_1537, rest ok | | Harmless, but shows up as an error |
| P07c | No line break after the last line | **created** | | Last record is read |
| P07a | LF line endings (own file) | **2 of 2 created** | | LF works |
| P07b | CR-only line endings (round 1) | second record lost silently | | Script always writes CRLF |
| P07d | UTF-8 file with `ö` (round 1) | fields shift, PI_2227 | | Confirms decision 1 |
| P07e | UTF-8 BOM (round 1) | first record shifts, PI_1537 | | Script writes no BOM |
