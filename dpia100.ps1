<#
.SYNOPSIS
    Generates a fixed-width Beaufort DPIA100 import file from an input CSV file.
.DESCRIPTION
    - Automatically detects delimiter (;, ,, \t).
    - Fully independent of column order.
    - Strictly validates all fields and aborts if errors are found.
    - Truncates [Waarde] with a warning if it exceeds 50 characters (showing old & new values).
    - Warns about amounts Beaufort may read differently than intended (comma, thousand separator), about
      empty values (a delete in Beaufort) and about dates and times not written as dd-mm-yyyy / hh:mm.
      With warnings, the file is only written after the user types Y.
    - Complies with Visma / Raet Beaufort specifications (187-character fixed record layout).
    - Reads the CSV with automatic encoding detection (UTF-8 / UTF-16 / Windows-1252), so it behaves
      the same on Windows PowerShell 5.1 and PowerShell 7+.
    - Skips completely empty rows (e.g. trailing ';;;;;;' lines exported by Excel).
    - Keeps the console open on errors (unless -NoPause is passed by the .bat launcher).
    - Exports in Windows-1252 (ANSI) so characters like 'ö' do not cause column shifts.
    - Prints which CSV column was used for each field, so a wrong match is visible.
    - On any failure, an older output file is renamed to *.old so it cannot be imported by mistake.
    - Writes the output next to the input CSV as DPIA100_<csv name>.txt, unless -OutputTxtPath is given.
    - Productiesoort NOR and COR are supported (JWP needs a verrekenperiode, which is not implemented).
    - Knows the standard PS/DV rubrieken of Beaufort (decision 17): for those it checks the dienstverband,
      S of V and the value against the rubriek, and warns about a rubriek blocked for indirect mutation.
      Other codes (eigen rubrieken) are allowed and checked as before.
    - Scope: PS and DV mutations only. See .NOT_IN_CURRENT_SPEC for Beaufort features left out.

.BEAUFORT_LAYOUT_SPECIFICATION
    Total record length: Exactly 187 characters / bytes.
    -----------------------------------------------------------------------------------------
    Pos   1 -   3 (Len  3) : Procescode (Uppercase alphanumeric, e.g. CMT, CON)
    Pos   4 -   9 (Len  6) : Rubriekcode (1-6 uppercase alphanumeric, left-aligned, e.g. E00101)
    Pos  10 -  59 (Len 50) : Object_id (pers_nr [6] + space [1] + dv_vlgnr [3], padded to 50)
    Pos  60 -  60 (Len  1) : Indicator stam/variabel ('S' or 'V')
    Pos  61 -  71 (Len 11) : Creatiedatum (3 leading spaces + ddMMyyyy)
    Pos  72 -  87 (Len 16) : Creatiegebruiker (Right-aligned username, padded with spaces)
    Pos  88 - 137 (Len 50) : Mutatiewaarde (Left-aligned, space padded to 50 chars)
    Pos 138 - 148 (Len 11) : Ingangsdatum (3 leading spaces + ddMMyyyy)
    Pos 149 - 178 (Len 30) : Not a field: always 30 spaces (the Excel tool's padding, see .NOT_IN_CURRENT_SPEC D)
    Pos 179 - 181 (Len  3) : Productiesoort (NOR or COR in this script; Beaufort also knows JWP)
    Pos 182 - 187 (Len  6) : Verrekenperiode (Always 6 spaces in this script; Beaufort format ppjjjj)
    -----------------------------------------------------------------------------------------

.KEY_ARCHITECTURAL_DECISIONS
    1. ENCODING: Must strictly be Windows-1252 (ANSI).
       Beaufort HR Core (PPIA081 / Uniface) reads fixed single-byte character columns.
       In UTF-8, accented characters like 'ö' (0xC3 0xB6) occupy 2 bytes, shifting all subsequent
       columns by 1 position (truncating the date and turning 'NOR' into 'NO', triggering PI_2227).
       In Windows-1252, 'ö' is 1 byte (0xF6), preserving exact column alignment.
    2. ERROR HANDLING: Fail-fast with batch error collection.
       All rows are inspected and all validation errors across the entire CSV are compiled into a
       summary table before aborting, preventing partial imports and tedious single-row trial-and-error.
    3. WARNINGS INSTEAD OF ERRORS:
       These cases are non-fatal: a Waarde longer than 50 characters is truncated (showing the original and
       truncated values; this prevents the #VALUE! error seen in Excel formulas), amounts Beaufort may
       read differently (decision 8), empty values, dates and times (decision 16), and a standard rubriek
       blocked for indirect mutation (decision 19). All warnings are shown before the prompts, and after the
       prompts the file is only written if the user types Y. Anything else, including end of input, cancels
       like an error (decision 14). Changes that are certain (a decimal comma in a ';' file, a date or time
       in a standard date or time rubriek, decision 20) are only reported.
    4. NO LEADING ZEROS IN OBJECT_ID:
       Beaufort documentation explicitly specifies: "Er mogen geen voorloopnullen in staan".
       Leading zeros are stripped from pers_nr and dv_vlgnr before right-aligning them.
    5. STRICT NUMERIC KEYS:
       Pers. Nr and Dienstv must consist of digits only. Values like '1234.0' (Excel float export) or
       '12abc' are rejected instead of being silently "cleaned" into a different, valid-looking number.
    6. INPUT ENCODING DETECTION:
       Import-Csv / Get-Content default to ANSI on Windows PowerShell 5.1 but UTF-8 on PowerShell 7+.
       The CSV is therefore read as raw bytes: BOM -> that encoding; otherwise strict UTF-8; if the bytes
       are not valid UTF-8 -> Windows-1252 (a plain Excel "CSV" export).
    7. NO SILENT CHARACTER LOSS:
       The output encoder is strict. A character that does not exist in Windows-1252 is reported as a
       validation error instead of being written as '?'. Waarde is normalized to composed form (NFC) and
       may not contain line breaks, tabs or other control characters.
    8. AMOUNTS: WARN, DON'T BLOCK (owner; Beaufort behaviour from probes P08, tests/beaufort/findings.md):
       Applies to a rubriek that is not in the standard list (an eigen rubriek, decision 17) and to a
       standard numeric rubriek. For an eigen rubriek '1.234,56' may be a valid text value: the script
       cannot tell a numeric from a text rubriek, so number-like values are never rejected. A standard
       numeric rubriek also gets the errors of decision 20; a standard text rubriek is written as typed,
       without this table. Beaufort reads the
       dot as decimal point; a comma in a numeric rubriek is accepted at import (100254) but rejected at
       processing (100252, PI_2229).
       The delimiter hints at the notation: Dutch Excel saves CSV with ';' and a decimal comma, English
       Excel with ',' and a decimal dot. A tab-separated file gives no hint.
                                 ';' file (Dutch)       ',' file (English)       tab file
       '12,50', '-12,50'         -> '12.50', reported   -> '12.50', warning      -> '12.50', warning
       '1,234'                   -> '1.234', reported   as typed, warning (1234) -> '1.234', warning
       '3.250', '12.345'         as typed, warning      as typed, no warning     as typed, warning
       '1.234,56', '1,234.56',   as typed, warning      as typed, warning        as typed, warning
       '1.234.567'
       Comma conversions are listed together (one line or one warning). 'Reported' means printed without
       asking for confirmation. Beaufort processes '3.250' silently as 3,25, which is right in an English
       file and probably a formatted 3250 in a Dutch one. Dates with dots ('1.11.2026') are not amounts.
       Extra decimals are not rounded up by Beaufort: '12.345' became 12,34 in a 2-decimal rubriek
       (probe round 4, tests/beaufort/findings.md).
    9. HEADER ROW PARSED BY THE SCRIPT:
       Excel often adds empty header cells (header ending in ';;'). Windows PowerShell 5.1 aborts on
       duplicate/empty column names, PowerShell 7 renames them. The header row is parsed here and empty
       or duplicate names are made unique, so both versions behave the same.
       Patterns of 1-2 characters (e.g. 'Dv') must match the whole header; longer patterns match anywhere
       in the header. The same CSV column may not be mapped to two fields.
    10. STALE OUTPUT PROTECTION:
       If the run fails, an existing output file (from an earlier run) is renamed to '<name>.old' so an
       outdated import file cannot be imported by mistake.
    11. PRODUCTIESOORT NOR OR COR (current spec; JWP may be added later, see .NOT_IN_CURRENT_SPEC):
       An empty productiesoort becomes NOR. COR is valid in Beaufort (owner; also named by the Excel tool).
       JWP (herrekenen over een jaarwerkperiode) requires a verrekenperiode (ppjjjj) and a matching
       ingangsdatum. The verrekenperiode is not implemented, so JWP is rejected.
    12. OUTPUT FILE NAME STARTS WITH 'DPIA100':
       Beaufort only finds import files whose name starts with 'DPIA100' (owner). The check ignores case:
       a lowercase 'dpia100.txt' imports fine. Any other -OutputTxtPath name is rejected.
    13. OUTPUT LOCATION AND NAME:
       Without -OutputTxtPath, the file is written next to the input CSV as 'DPIA100_<csv name>.txt'
       (owner, issue #10). The working directory is no use: with a CSV dragged onto the .bat, Explorer
       sets it to an unpredictable folder (user profile, System32). The name always passes decision 12.
       An explicit -OutputTxtPath is used as given; a relative path resolves against the working directory.
    14. VALIDATE BEFORE THE PROMPTS:
       The CSV is read, mapped and fully validated before procescode, creatiedatum and creatiegebruiker
       are asked (owner, issue #11). On errors the script stops without asking anything. Without errors it
       prints '[OK] N row(s) checked', lists any warnings, asks the three prompts, and only then asks the
       Y confirmation for warnings. The prompt values are only needed to build the records.
    15. WEB VERSION (dpia100.html, owner):
       The browser page implements the same rules in JavaScript for users who should not need a console.
       Every rule change goes into both files; tests/run-tests.ps1 runs every case on PowerShell 5.1,
       PowerShell 7 and the page (headless Edge), and tests/web/differential.ps1 compares both on random
       input. By design the page cannot rename an old output file (decision 10); it writes nothing until
       the user saves, and refuses a file name that does not start with DPIA100 (decision 12).
    16. EMPTY VALUES, DATES AND TIMES: WARN (owner, after reviewing the Excel tool source):
       - An empty Waarde deletes the current value in Beaufort. The Excel tool only allowed it through the
         keyword NULL (a plain empty cell was an error for numeric, date and time rubrieken). The script
         writes it, with a warning listing the rows. The text NULL is written as typed (owner).
       - The Excel tool wrote a date in a date rubriek as dd-mm-yyyy and a time as hh:mm. Beaufort also has
         time rubrieken of type U, hh:mm:ss (owner, confirmed in Beaufort). For a rubriek that is not in the
         standard list, a value that looks like a date in another shape (1-11-2026, 1/11/26, 2026-11-01,
         1.11.2026) or a time in neither shape (8:30, 8:30:00) is written as typed, with a warning. The
         script cannot tell a date rubriek of an eigen rubriek from a text rubriek, so it never converts.
         8 digits (01112026) and '8.30' are not recognized: they may be numbers. A standard rubriek has a
         known type instead (decision 20): no date or time warnings there.
       Each kind is one warning listing the first rows, like the comma conversions.
    17. STANDARD RUBRIEKEN (owner, docs/rubrieken-export.md):
       The standard PS/DV rubrieken (Fix = J) of the Beaufort export 'Overzicht rubrieken' (function 210006)
       are embedded below, generated by tools/update-standard-rubrieken.ps1 together with the copy in
       dpia100.html. A customer cannot change the gegevenssoort, type, length, decimals, stam/variabel or
       Blok ind mut of a standard rubriek (Beaufort help), so these are the same for every municipality on
       that Beaufort version. Min/max and referentie values can be changed and are left out.
       A code that is not in the list is allowed without any message: it is an eigen rubriek (or a standard
       rubriek newer than the list) and is checked as before. Eigen rubrieken are issue #24.
       Rubriek codes are 1-6 letters and digits (the export has standard SETKB and eigen codes such as
       E0010). A shorter code is left-aligned in positions 4-9, as v1.2.0 did (owner); the Excel tool
       right-aligned it and Beaufort was never probed with one (issue #25).
    18. DIENSTVERBAND AND S OF V OF A STANDARD RUBRIEK (owner):
       A DV rubriek needs a dienstverband; empty or 0 gives a PS key, which Beaufort rejects at import
       (PI_1650). A PS rubriek may not have one: Beaufort accepts a DV key there at import (probe P04-7, #2),
       with an unknown effect. S on a variabel-only or V on a stam-only rubriek is rejected by Beaufort.
       All are validation errors. Without the list the key is derived from the Dienstverband column alone.
    19. BLOK IND MUT (owner):
       A standard rubriek with Blok ind mut = J is blocked for indirect mutation, so Beaufort does not import
       it through DPIA100. One warning lists the rows (decision 3); the file is still written after Y. A blank
       flag in the export counts as not blocked.
    20. WAARDE BY TYPE OF A STANDARD RUBRIEK (owner):
       The checks of decision 7 and the empty-value warning (decision 16) apply to every rubriek. Then:
       - N (numeric): decision 8 as for an eigen rubriek. Errors for what Beaufort rejects or changes: not a
         number after the comma conversion (text, NULL, 1.234,56, an English 1,234), more significant
         decimals than the rubriek has (Beaufort cuts 12.345 to 12.34, probe round 4), more digits before the
         decimal point than the length allows (length includes the decimal point and the decimals; confirmed
         in the Beaufort GUI, tests/beaufort/findings.md).
       - D (date): must be a date as accepted for Ingangsdatum; written as dd-mm-yyyy, as the Excel tool did.
         The rewrite is reported, not a warning. Not a date (1/11/26, 46327, text) is an error.
       - T (time hh:mm) and U (time hh:mm:ss): must be a time; written in the type's shape (8:30 -> 08:30),
         reported. Not a time, or seconds a T rubriek cannot hold, is an error.
       - Any other type (S text, H, I): written as typed. No comma conversion and no amount, date or time
         warnings. Longer than the rubriek allows is an error; over 50 characters keeps the truncation
         warning (decision 3).

.NOT_IN_CURRENT_SPEC
    Supported by Beaufort / the official Excel tool, not by this script. Candidates for a later version.
    A. OTHER OBJECT TYPES (gegevenssoorten). Object_id parts, space separated, no leading zeros:
         PS pers_nr                          (supported)
         DV pers_nr dv_vlgnr                 (supported)
         BG pers_nr dv_vlgnr bg_vlgnr        IZ pers_nr dv_vlgnr oe_oper_sl operfunc_kd ingang_dt
         IP reg_sal_kd inpas_nr (no pers_nr) LV pers_nr dv_vlgnr lv_rglnr
         SD pers_nr dv_vlgnr sd_rglnr        ZG pers_nr dv_vlgnr zg_nr
         ZT pers_nr dv_vlgnr zg_nr zt_nr
       Max lengths: bg_vlgnr 1-2, lv_rglnr 2-3, sd_rglnr 2, oe_oper_sl 10, operfunc_kd 7 (alfanumeriek),
       reg_sal_kd 2, inpas_nr 4, zg_nr 10, zt_nr 10 (the two PDFs differ for bg/lv). ingang_dt = Beaufort date.
       The Excel tool supported PS, DV, BG, LV, SD and IZ (not IP, ZG, ZT). Its source right-aligns each
       part in a fixed width that includes the separating space: pers_nr 6, dv 4 (space + 3), bg 2 (space
       + 1), lv 3 (space + 2), sd 3 (space + 2), IZ: oe_oper_sl 11 (space + 10), operfunc_kd 8 (space + 7),
       ingang_dt 11 (the mutation's ingangsdatum in the Windows short-date format; probably wrong, the PDF
       asks for the inzet's own date). Those widths follow the 2005 handbook, not the 2024 PDF (bg 2,
       lv 3), and were never verified in Beaufort: probe them before adding.
    B. VERREKENPERIODE (pos 182-187) AND JWP:
       Format ppjjjj (e.g. 132003). Periods 13-15 are only applied if the ingangsdatum is in December,
       otherwise Beaufort recalculates. For recalculation over a jaarwerkperiode the ingangsdatum must equal
       the start date of the last loonperiode. The Excel tool has a 'Verreken periode' column: exactly 6
       characters, right-aligned, written as typed without a format check.
    C. EIGEN RUBRIEKEN AND THE REST OF THE RUBRIEK DATA:
       The standard rubrieken are known (decisions 17-20). Not known: eigen rubrieken (ideas in issue #24),
       min/max and bereik, and the values of referentietabellen (customer settings, not in the list).
       The Excel tool's tab 'Rubriek' (rubriek_kd, gegsrt_id, type) flagged unknown rubrieken and checked
       type N as numeric (default), rewrote D as dd-mm-yyyy and T as hh:mm; H = hoofdletters had no effect
       since 2005. In N, D and T an empty cell was an error and the keyword NULL wrote an empty value. This
       script writes an empty value with a warning and NULL as typed (decision 16).
    D. POSITIONS 149-178:
       Not described in either PDF: the documented fields end at Ingangsdatum (pos 148) and Productiesoort
       starts at 179. Not a field: the Excel tool source right-aligns productiesoort in 33 positions, so
       these are its leading spaces. Filled with 30 spaces, which matches the Excel tool byte for byte.
#>

[CmdletBinding()]
param (
    # DECISION: No default path. If omitted (or empty), a file-open dialog is shown.
    [string]$InputCsvPath = $null,
    # Output is always written as Windows-1252 (see architectural decision 1).
    # DECISION: Empty means DPIA100_<csv name>.txt next to the input CSV (decision 13).
    [string]$OutputTxtPath = $null,
    # The .bat launcher has its own 'pause', so it passes -NoPause to avoid a double prompt.
    [switch]$NoPause
)

# NOTE: Bump together with a new section in docs/CHANGELOG.md and a git tag vX.Y.Z (see CLAUDE.md).
$ScriptVersion = '1.2.0'


# -------------------------------------------------------------------------
# Exit / Error Handling
# -------------------------------------------------------------------------

# DECISION: Every exit path goes through Exit-Script so the console window stays open when the
# script is started directly (double-click / right-click Run). The .bat launcher passes -NoPause.
function Exit-Script([int]$Code = 0) {
    # DECISION: On any failure, rename an existing output file from an earlier run (decision 10).
    if ($Code -ne 0 -and $script:fullOutputPath -and (Test-Path -LiteralPath $script:fullOutputPath -PathType Leaf)) {
        $oldFile = Get-Item -LiteralPath $script:fullOutputPath
        $oldStamp = $oldFile.LastWriteTime.ToString("dd-MM-yyyy HH:mm")
        $oldPath = "$($script:fullOutputPath).old"
        try {
            Move-Item -LiteralPath $script:fullOutputPath -Destination $oldPath -Force -ErrorAction Stop
            Write-Host "`nAn older output file (from $oldStamp) was renamed to '$oldPath'" -ForegroundColor Yellow
            Write-Host "so it cannot be imported by mistake." -ForegroundColor Yellow
        } catch {
            Write-Host "`n[WARNING] An older output file (from $oldStamp) still exists and could not be renamed:" -ForegroundColor Yellow
            Write-Host "  $($script:fullOutputPath)" -ForegroundColor Yellow
            Write-Host "  It is NOT the result of this run. Do NOT import it." -ForegroundColor Yellow
        }
    }
    if (-not $NoPause) {
        Write-Host ""
        Read-Host "Press ENTER to close" | Out-Null
    }
    exit $Code
}

function Stop-WithError([string]$Message) {
    Write-Host "`n[ERROR] $Message" -ForegroundColor Red
    Exit-Script 1
}

# DECISION: Catch-all for unexpected terminating errors (locked files, unreadable CSV, ...)
# so the user always sees a readable message instead of a window that disappears.
trap {
    Write-Host "`n[UNEXPECTED ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Exit-Script 1
}

# DECISION: AppLocker / WDAC can force Constrained Language Mode, which blocks the .NET calls below
# (encoding, file stream). Fail early with a clear message instead of a cryptic error halfway.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Stop-WithError "PowerShell runs in '$($ExecutionContext.SessionState.LanguageMode)' mode on this machine (AppLocker/WDAC policy). This script needs FullLanguage mode."
}

# DECISION: Strict Windows-1252 encoder. Characters that do not exist in Windows-1252 (Cyrillic, Greek,
# emoji, ...) would otherwise be replaced by '?' silently. With exception fallbacks they can be reported.
$strictAnsi = [System.Text.Encoding]::GetEncoding(1252, [System.Text.EncoderExceptionFallback]::new(), [System.Text.DecoderExceptionFallback]::new())

# Returns a readable list of characters in $text that cannot be represented in Windows-1252
# (empty string when everything is encodable).
function Get-UnencodableChars([string]$text) {
    $bad = [System.Collections.Generic.List[string]]::new()
    foreach ($ch in $text.ToCharArray()) {
        try {
            [void]$strictAnsi.GetBytes([string]$ch)
        } catch {
            $bad.Add(('{0} (U+{1:X4})' -f [string]$ch, [int]$ch))
        }
    }
    return (($bad | Select-Object -Unique) -join ', ')
}

# DECISION: Reads the CSV as raw bytes and detects the encoding explicitly, because Import-Csv /
# Get-Content behave differently on PowerShell 5.1 (ANSI) and 7+ (UTF-8).
# FileShare.ReadWrite allows reading a CSV that is still open in Excel.
function Read-CsvFile([string]$filePath) {
    $fs = New-Object System.IO.FileStream($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $ms = New-Object System.IO.MemoryStream
        $fs.CopyTo($ms)
        $bytes = $ms.ToArray()
    } finally {
        $fs.Dispose()
    }

    $enc = $null
    $offset = 0
    $name = ""
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $enc = New-Object System.Text.UTF8Encoding($false, $true); $offset = 3; $name = "UTF-8 (BOM)"
    } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $enc = [System.Text.Encoding]::Unicode; $offset = 2; $name = "UTF-16 LE"
    } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $enc = [System.Text.Encoding]::BigEndianUnicode; $offset = 2; $name = "UTF-16 BE"
    }

    if ($enc) {
        $text = $enc.GetString($bytes, $offset, $bytes.Length - $offset)
    } else {
        try {
            $text = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes)
            $name = "UTF-8 (no BOM)"
        } catch {
            $text = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
            $name = "Windows-1252"
        }
    }
    return [PSCustomObject]@{ Text = $text; EncodingName = $name }
}

# No path given (e.g. double-clicked directly) -> let the user browse to any CSV
# DECISION: GUI fallback ensures ease of use for non-technical users launching via shortcut.
if (-not $InputCsvPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dialog.Title  = "Select a CSV file to process"
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Warning "No file selected. Exiting."
        Exit-Script 0
    }
    $InputCsvPath = $dialog.FileName
}

# -LiteralPath: file names containing [ ] must not be treated as wildcards.
if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
    Stop-WithError "Input file not found: $InputCsvPath"
}
$InputCsvPath = (Resolve-Path -LiteralPath $InputCsvPath).ProviderPath

# DECISION: Default output next to the input CSV, named after it (decision 13).
if (-not $OutputTxtPath) {
    $OutputTxtPath = Join-Path (Split-Path -Parent $InputCsvPath) ("DPIA100_" + [System.IO.Path]::GetFileNameWithoutExtension($InputCsvPath) + ".txt")
}

# DECISION: Output path is resolved up front, so every later failure can retire a stale output file.
# Relative paths are resolved against the PowerShell working directory; absolute paths are used as-is
# (Join-Path would produce a broken path for those).
$basePath = if ([System.IO.Path]::IsPathRooted($OutputTxtPath)) { $OutputTxtPath } else { Join-Path (Get-Location).ProviderPath $OutputTxtPath }
$resolvedOutputPath = [System.IO.Path]::GetFullPath($basePath)

# DECISION: Beaufort only picks up import files whose name starts with 'DPIA100' (decision 12).
# Checked before $script:fullOutputPath is set, so a failure here never renames an unrelated file to .old.
$outputName = [System.IO.Path]::GetFileName($resolvedOutputPath)
if (-not $outputName.StartsWith('DPIA100', [System.StringComparison]::OrdinalIgnoreCase)) {
    Stop-WithError "The output file name '$outputName' must start with 'DPIA100' (for example dpia100.txt), otherwise Beaufort does not find it."
}
$script:fullOutputPath = $resolvedOutputPath

# -------------------------------------------------------------------------
# Helper Functions
# -------------------------------------------------------------------------

# DECISION: Inspects the first non-empty line of the file and calculates the frequency
# of common delimiters (;, ,, \t). This eliminates dependency on Regional Settings / Excel locale.
# DECISION: On a tie the first candidate in the order ';', tab, ',' wins. This is the order Windows
# PowerShell 5.1 happened to use; PowerShell 7 picked at random (hashtable order changes per run).
function Detect-CsvDelimiter([string]$text) {
    $firstLine = $text -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        throw "The CSV file is empty."
    }
    $firstLine = $firstLine.Trim([char]0xFEFF).Trim()

    $best = ';' # Default fallback
    $bestCount = 0
    foreach ($c in @(';', "`t", ',')) {
        $count = ($firstLine -split [regex]::Escape($c)).Count - 1
        if ($count -gt $bestCount) {
            $best = $c
            $bestCount = $count
        }
    }
    return $best
}

# DECISION: Parses only the header record (decision 9). Handles quoted names, "" inside quotes and
# line breaks inside quoted names (e.g. "Productie<newline>soort"). Returns the raw field names and the
# remaining text (the data rows), which is then parsed with ConvertFrom-Csv -Header.
function Split-CsvHeader([string]$text, [string]$delimiter) {
    $delim  = [char]$delimiter
    $fields = [System.Collections.Generic.List[string]]::new()
    $sb     = [System.Text.StringBuilder]::new()
    $i = 0
    $inQuotes = $false

    # Skip a BOM and blank lines before the header
    while ($i -lt $text.Length -and ($text[$i] -eq [char]0xFEFF -or $text[$i] -eq [char]13 -or $text[$i] -eq [char]10)) { $i++ }

    while ($i -lt $text.Length) {
        $c = $text[$i]
        if ($inQuotes) {
            if ($c -eq [char]'"') {
                if ($i + 1 -lt $text.Length -and $text[$i + 1] -eq [char]'"') {
                    [void]$sb.Append('"'); $i += 2; continue
                }
                $inQuotes = $false; $i++; continue
            }
            [void]$sb.Append($c); $i++; continue
        }
        if ($c -eq [char]'"') { $inQuotes = $true; $i++; continue }
        if ($c -eq $delim) { $fields.Add($sb.ToString()); [void]$sb.Clear(); $i++; continue }
        if ($c -eq [char]13 -or $c -eq [char]10) {
            if ($c -eq [char]13 -and $i + 1 -lt $text.Length -and $text[$i + 1] -eq [char]10) { $i += 2 } else { $i++ }
            break
        }
        [void]$sb.Append($c); $i++
    }
    $fields.Add($sb.ToString())

    return [PSCustomObject]@{ Fields = $fields.ToArray(); Rest = $text.Substring($i) }
}

# Makes header names usable as property names: line breaks -> space, empty -> placeholder,
# duplicates (case-insensitive) -> suffixed with (2), (3), ...
function Get-UniqueHeaderNames([string[]]$rawNames) {
    $seen   = @{}
    $result = [System.Collections.Generic.List[string]]::new()
    $empty  = 0
    foreach ($raw in $rawNames) {
        $name = ($raw -replace '[\r\n]+', ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            $empty++
            $name = "(empty column $empty)"
        }
        $base = $name
        $n = 2
        while ($seen.ContainsKey($name)) { $name = "$base ($n)"; $n++ }
        $seen[$name] = $true
        $result.Add($name)
    }
    return ,$result.ToArray()
}

# DECISION: Beaufort requires dates in 8-digit format (ddMMyyyy).
# Supports Dutch space-separated formats (e.g. '1 11 2024'), dashes, slashes, and ISO.
# Rejects impossible calendar dates (e.g. 31 02 2026).
function Test-And-ParseDate([string]$rawDate, [ref]$parsedDate) {
    if ([string]::IsNullOrWhiteSpace($rawDate)) { return $false }

    $clean = $rawDate.Trim()
    # NOTE: [string[]] is required. An untyped array makes PowerShell pick the single-format overload of
    # TryParseExact (the array is joined into one format string) and nothing would ever match.
    [string[]]$formats = @(
        "d M yyyy", "dd MM yyyy", "d MM yyyy", "dd M yyyy",
        "d-M-yyyy", "dd-MM-yyyy", "d/M/yyyy", "dd/MM/yyyy",
        "yyyy-MM-dd", "ddMMyyyy"
    )

    $dt = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($clean, $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        $parsedDate.Value = $dt
        return $true
    }

    # Fallback regex parsing (day, month, year)
    if ($clean -match '^(\d{1,2})[\s\.\-\/]+(\d{1,2})[\s\.\-\/]+(\d{4})$') {
        try {
            $parsedDate.Value = [DateTime]::new([int]$Matches[3], [int]$Matches[2], [int]$Matches[1])
            return $true
        } catch {
            return $false
        }
    }

    return $false
}

# DECISION: Resolves headers flexibly by keyword pattern.
# Strips BOM (\uFEFF), quotes, newlines, and leading/trailing whitespace.
# Makes the script 100% column-order agnostic and resilient against multiline headers like "Productie\nsoort".
# DECISION: Patterns of 1-2 characters (e.g. 'Dv') must equal the whole header; otherwise they would
# also match unrelated headers such as 'Advies'. Longer patterns match anywhere in the header.
function Find-PropertyByPattern([string[]]$headerNames, [string[]]$patterns) {
    foreach ($name in $headerNames) {
        $cleanName = ($name -replace '[\r\n"]', ' ').Trim()
        foreach ($p in $patterns) {
            $isMatch = if ($p.Length -le 2) { $cleanName -eq $p } else { $cleanName -like "*$p*" }
            if ($isMatch) {
                return $name
            }
        }
    }
    return $null
}

# -------------------------------------------------------------------------
# Standard Rubrieken (decision 17)
# -------------------------------------------------------------------------

# One line per standard PS/DV rubriek: code, gegevenssoort, type, length, decimals ('-' = blank),
# stam/variabel allowed ('S-', '-V', 'SV') and Blok ind mut ('J', 'N', '-' = blank, not blocked).
# The same text is in the core of dpia100.html (checked by the page-rubrieken test).
# BEGIN STANDARD RUBRIEKEN: generated by tools/update-standard-rubrieken.ps1 from the Beaufort export of 30-10-2026, do not edit (decision 17)
$StandardRubriekenDate = '30-10-2026'
$StandardRubriekenTable = @'
P00001 DV N  6 2 SV N
P00002 DV N  9 2 SV N
P00003 DV N  9 2 SV N
P00004 DV N  9 2 SV N
P00005 DV N  9 2 SV N
P00006 DV N  9 2 SV N
P00007 DV N  9 2 SV N
P00008 DV N  9 2 SV N
P00009 DV N  6 2 SV -
P00010 DV N  9 2 SV N
P00011 DV N  9 2 SV N
P00012 DV N  6 2 SV N
P00013 DV N  9 2 SV N
P00014 DV N  6 2 SV N
P00015 DV N  9 2 SV N
P00016 DV N  9 2 SV N
P00017 DV N  9 2 SV N
P00018 DV N  9 2 SV N
P00019 DV N  6 2 SV N
P00020 DV N  6 2 SV N
P00021 DV N  6 2 SV N
P00022 DV N  9 2 SV N
P00023 DV N  9 2 SV N
P00024 DV N  6 2 SV N
P00025 DV N  9 2 SV N
P00026 DV N  9 2 SV N
P00027 DV N  6 2 S- N
P00028 DV N  6 2 S- N
P00029 DV N  6 2 SV N
P00030 DV N  9 2 SV N
P00031 DV N  6 2 SV N
P00032 DV N  6 2 SV N
P00033 DV N  6 2 SV N
P00034 DV N  6 2 SV N
P00035 DV N  9 2 SV N
P00036 DV N  9 2 SV N
P00037 DV N  6 2 SV N
P00038 DV N  6 2 SV N
P00039 DV N  6 2 SV N
P00040 DV N  6 2 SV N
P00041 DV N  9 2 SV N
P00042 DV N  5 2 SV N
P00043 DV N  9 2 SV N
P00044 DV N  6 2 SV N
P00045 DV N  6 2 SV N
P00046 DV N  6 2 SV N
P00047 DV N  6 2 SV N
P00048 DV N  6 2 SV N
P00049 DV N  6 2 SV N
P00050 DV N  6 2 SV N
P00051 DV N  9 2 SV N
P00052 DV N  6 2 SV N
P00053 DV N  6 2 SV N
P00054 DV N  6 2 SV N
P00055 DV N  9 2 SV N
P00056 DV N  9 2 SV N
P00057 DV N  6 2 SV N
P00058 DV N  6 2 SV N
P00059 DV N  6 2 SV N
P00060 DV N  6 2 SV N
P00061 DV N  6 2 SV N
P00062 DV N  6 2 SV N
P00063 DV N  9 2 -V N
P00064 DV N  9 2 -V N
P00065 DV N  9 2 -V N
P00066 DV N  9 2 -V N
P00067 DV N  9 2 SV N
P00068 DV N  9 2 SV N
P00069 DV N  9 2 -V N
P00070 DV N  9 2 -V N
P00071 DV N  9 2 SV N
P00072 DV N  9 2 SV N
P00075 DV N  9 2 SV N
P00077 DV N  9 2 SV N
P00078 DV N  9 2 SV N
P00080 DV N  9 2 SV N
P00081 DV N  6 2 SV N
P00082 DV N  9 2 SV N
P00083 DV N  9 2 SV N
P00084 DV N  9 2 SV N
P00085 DV N  9 2 SV N
P00086 DV N  9 2 SV N
P00087 DV N  5 2 SV N
P00088 DV N  9 2 SV N
P00089 DV N  9 2 SV N
P00090 DV N  6 2 SV N
P00091 DV N  6 2 SV N
P00092 DV N  9 2 SV N
P00093 DV N  9 2 -V N
P00094 DV N  9 2 SV N
P00095 DV N  9 2 SV N
P00096 DV N  9 2 -V N
P00099 DV N  7 2 SV N
P00100 DV N  9 2 SV J
P00101 DV N  6 2 SV N
P00102 DV N  9 2 -V N
P00103 DV N  9 2 SV N
P00104 DV N  9 2 SV N
P00105 DV N  9 2 SV N
P00106 DV N  9 2 SV N
P00108 DV N  9 2 SV J
P00109 DV N  6 2 SV J
P00110 DV N  9 2 SV N
P00111 DV N  6 2 SV N
P00112 DV N  9 2 SV N
P00113 DV N  6 2 SV N
P00114 DV N  6 2 SV N
P00115 DV N  6 2 SV N
P00116 DV N  9 2 SV N
P00117 DV N  9 2 SV N
P00119 DV N  9 2 SV N
P00120 DV N  9 2 SV N
P00121 DV N  9 2 SV N
P00122 DV N  9 2 SV N
P00123 DV N  9 2 SV N
P00124 DV N  9 2 SV N
P00125 DV N  9 2 SV N
P00126 DV N  9 2 SV N
P00127 DV N  9 2 SV N
P00128 DV N  9 2 SV N
P00129 DV N  9 2 SV N
P00130 DV N  9 2 SV N
P00131 DV N  9 2 SV N
P00132 DV N  9 2 SV N
P00133 DV N  9 2 SV N
P00134 DV N  9 2 SV N
P00135 DV N  9 2 SV N
P00136 DV N  9 2 SV N
P00137 DV N  9 2 SV N
P00138 DV N  9 2 SV N
P00139 DV N  9 2 SV N
P00140 DV N  9 2 SV N
P00141 DV N  9 2 SV N
P00142 DV N  9 2 SV N
P00143 DV N  9 2 SV N
P00145 DV N  9 2 SV N
P00146 DV N  9 2 SV N
P00147 DV N  9 2 SV N
P00148 DV N  9 2 SV N
P00149 DV N  9 2 SV N
P00150 DV N  7 3 SV N
P00151 DV N  9 2 SV N
P00152 DV N  9 2 SV N
P00153 DV N  9 2 SV N
P00154 DV N  9 2 SV N
P00155 DV N  9 2 SV N
P00156 DV N  9 2 SV N
P00157 DV N 10 2 SV N
P00158 DV N  9 2 SV N
P00159 DV N  9 2 S- N
P00161 DV N  6 2 SV N
P00162 DV N  9 2 SV J
P00163 DV N  6 2 SV N
P00167 DV N 10 2 SV N
P00169 DV N 10 2 SV N
P00170 DV N  1 0 S- N
P00171 DV N  9 2 SV J
P00173 DV N  9 2 -V N
P00174 DV N  9 2 -V N
P00175 DV N  9 2 SV N
P00177 DV N  9 2 SV N
P00180 DV N  9 2 SV N
P00181 DV N  9 2 SV N
P00183 DV N  9 2 SV N
P00184 DV N  9 2 SV N
P00186 DV N  9 2 SV N
P00187 DV N  9 2 SV N
P00188 DV N  9 2 SV N
P00189 DV N  1 0 S- J
P00190 DV N  9 2 SV N
P00191 DV N  9 2 SV N
P00192 DV N  9 2 SV N
P00193 DV N  3 0 -V N
P00194 DV N  9 2 SV N
P00195 DV N  9 2 SV N
P00196 DV N  9 2 -V N
P00197 DV N  9 2 SV N
P00198 DV N  9 2 SV N
P00199 DV N  9 2 SV N
P00200 DV N  9 2 SV N
P00201 DV N  9 2 SV N
P00202 DV N  9 2 SV N
P00203 DV N  9 2 SV N
P00204 DV N  5 2 SV J
P00205 DV N  9 2 SV J
P00206 DV N  9 2 SV N
P00207 DV N  9 2 SV N
P00208 DV N  9 2 SV N
P00209 DV N  9 2 SV N
P00210 DV N  9 2 SV N
P00211 DV N  9 2 SV N
P00212 DV N  9 2 SV N
P00214 DV N  9 2 SV N
P00215 DV N  9 2 SV N
P00216 DV N  6 2 SV N
P00217 DV N  6 2 SV N
P00218 DV N  9 2 SV N
P00219 DV N  9 2 SV N
P00220 DV N  9 2 SV N
P00221 DV N  9 2 SV N
P00222 DV N  9 2 SV N
P00223 DV N  9 2 SV N
P00224 DV N  9 2 SV N
P00225 DV N  9 2 SV N
P00226 DV N  9 2 SV N
P00227 DV N  9 2 SV N
P00228 DV N  9 2 SV N
P00229 DV N  9 2 SV N
P00230 DV N  9 2 SV N
P00231 DV N  9 2 SV N
P00232 DV N  9 2 SV N
P00233 DV N  9 2 SV N
P00234 DV N  9 2 SV N
P00235 DV N  9 2 SV N
P00236 DV N  9 2 SV N
P00237 DV N  9 2 SV N
P00238 DV N  9 2 SV N
P00239 DV N  9 2 SV N
P00240 DV N  9 2 SV N
P00241 DV N  7 2 SV N
P00242 DV N  9 2 SV N
P00243 DV N  9 2 SV N
P00244 DV N  9 2 SV N
P00245 DV N  9 2 SV N
P00246 DV N  9 2 SV N
P00247 DV N  9 2 SV N
P00248 DV N  9 2 SV N
P00249 DV N  9 2 SV N
P00250 DV N  9 2 SV N
P00251 DV N  9 2 SV N
P00252 DV N  9 2 SV N
P00253 DV N  9 2 SV N
P00254 DV N  9 2 SV N
P00255 DV N  9 2 SV N
P00256 DV N  9 2 SV N
P00257 DV N  9 2 SV N
P00258 DV N  9 2 SV N
P00259 DV N  9 2 SV N
P00260 DV N  9 2 SV N
P00261 DV N  9 2 SV N
P00262 DV N  9 2 SV N
P00263 DV N  9 2 SV N
P00264 DV N  9 2 SV N
P00265 DV N  9 2 SV N
P00266 DV N  9 2 SV N
P00267 DV N  9 2 SV N
P00268 DV N  9 2 SV N
P00269 DV N  9 2 SV N
P00270 DV N  9 2 SV N
P00271 DV N  9 2 SV N
P00272 DV N  9 2 SV N
P00273 DV N  9 2 SV N
P00274 DV N  9 2 SV N
P00275 DV N  9 2 SV N
P00276 DV N  9 2 SV N
P00277 DV N  9 2 SV N
P00278 DV N  9 2 SV N
P00279 DV N  9 2 SV N
P00280 DV N  9 2 SV J
P00281 DV N  9 2 SV J
P00282 DV N  9 2 SV J
P00284 DV N  9 2 SV N
P00285 DV N  9 2 SV N
P00286 DV N  9 2 SV N
P00287 DV N  9 2 SV N
P00288 DV N  9 2 SV N
P00289 DV N  9 2 SV N
P00290 DV N  9 2 SV N
P00291 DV N  9 2 SV N
P00292 DV N  9 2 SV N
P00293 DV N  9 2 SV N
P00294 DV N  9 2 SV N
P00295 DV N  9 2 -V N
P00296 DV N  9 2 SV N
P00297 DV N  9 2 SV N
P00298 DV N  9 2 SV N
P00301 PS S 50 - S- N
P00302 PS S 10 - S- N
P00303 PS S  6 - S- N
P00304 PS S  1 - S- N
P00305 PS S 15 - S- N
P00306 PS S  2 - S- N
P00308 PS S 50 - S- N
P00311 DV S  4 - S- N
P00312 DV N  9 2 SV N
P00313 PS S 20 - S- N
P00314 DV S  3 - S- N
P00315 DV S  3 - S- N
P00317 DV N  3 - S- N
P00318 DV N  4 0 S- N
P00320 PS N  9 0 S- N
P00321 PS D 11 - S- N
P00322 DV D 11 - S- N
P00323 PS D 11 - S- J
P00326 DV N  2 0 S- N
P00327 DV N  7 2 SV N
P00328 DV N  7 2 SV N
P00329 PS S  3 - S- N
P00330 PS S  1 - S- N
P00331 DV N  1 0 SV N
P00332 DV S  3 - S- N
P00333 DV S  3 - S- N
P00334 DV S  3 - S- N
P00335 DV N  1 0 S- N
P00336 DV N  1 0 S- N
P00337 DV N  1 0 S- N
P00338 DV N  2 - S- J
P00339 DV S  1 - S- N
P00340 DV N  5 2 SV N
P00341 DV N  1 0 S- N
P00342 DV N  1 0 S- N
P00343 DV N  1 0 S- N
P00344 DV N  1 0 S- N
P00345 DV N  2 0 S- N
P00347 DV N  1 0 S- N
P00353 DV N  2 0 S- N
P00354 DV N  1 - S- N
P00355 DV S  1 - SV N
P00356 DV N 12 - S- N
P00357 DV N  3 0 S- N
P00358 DV N  3 0 S- N
P00359 DV N  1 0 SV N
P00360 DV N  7 0 -V N
P00361 DV N  1 0 SV N
P00362 DV N  1 - S- N
P00363 DV N  7 2 SV N
P00364 DV N  6 0 SV N
P00365 PS S 50 - S- N
P00366 DV N  9 2 SV N
P00367 PS N  5 0 S- N
P00368 PS S 20 - S- N
P00369 DV N  1 0 SV J
P00370 DV N  1 0 S- N
P00371 DV N  1 0 S- N
P00372 DV N  9 2 SV N
P00373 DV N  9 2 SV N
P00374 DV N  9 2 SV N
P00381 DV N  9 2 SV N
P00382 DV N  9 2 SV N
P00383 DV N  9 2 SV N
P00384 DV N  9 2 SV N
P00385 DV N  9 2 SV N
P00390 PS S 50 - S- N
P00391 PS S 10 - S- N
P00392 PS S  6 - S- N
P00401 DV N  6 2 S- N
P00402 DV N  9 2 SV N
P00403 DV N  6 2 -V N
P00404 DV N  6 2 SV N
P00405 DV N  6 2 SV N
P00406 DV N  1 0 SV N
P00407 DV N  9 2 SV N
P00408 DV N  7 2 -V N
P00409 DV N  9 2 S- J
P00410 DV N  6 2 S- N
P00411 DV N  9 2 SV N
P00412 DV N  9 2 SV N
P00413 DV N  9 2 -V N
P00414 DV N  9 2 SV J
P00415 DV N  9 2 SV J
P00416 DV N  6 2 SV J
P00420 DV N  1 0 SV N
P00421 DV S  3 - S- N
P00422 DV N  6 2 SV N
P00423 DV N  9 2 SV N
P00424 DV N  6 2 SV N
P00425 DV N  5 2 -V N
P00426 DV N  9 2 SV N
P00428 DV N  9 2 SV N
P00429 DV N  6 2 S- J
P00430 DV N  6 2 S- J
P00431 DV N  6 2 -V N
P00432 DV N  2 0 -V N
P00433 DV N  2 0 -V J
P00434 DV N  6 2 -V N
P00435 DV N  3 0 -V N
P00436 DV N  3 0 -V N
P00437 DV N  9 2 SV N
P00438 DV N 11 4 SV N
P00439 DV N  2 0 -V N
P00440 DV N  9 2 SV N
P00441 DV N  9 2 SV N
P00442 DV N  9 2 SV N
P00443 DV N  9 2 SV N
P00444 DV N 11 4 SV N
P00445 DV N  1 0 -V N
P00446 DV N 11 4 SV N
P00447 DV N  6 2 SV N
P00448 DV S  7 - S- N
P00449 DV N  3 0 -V N
P00450 DV N  9 2 S- N
P00451 DV N  6 2 S- N
P00452 DV N  9 2 SV N
P00453 DV N  1 0 S- N
P00454 DV N  6 2 SV J
P00455 DV N  7 2 SV N
P00456 DV N  9 2 SV N
P00457 DV N  1 - S- N
P00459 DV N  9 2 -V N
P00460 DV N  9 2 SV N
P00461 DV N  6 2 SV N
P00462 DV N  6 2 SV N
P00463 DV N  9 2 SV N
P00464 DV N  6 2 SV N
P00465 DV N  6 2 SV N
P00467 DV N  6 2 SV N
P00468 DV N  6 2 SV N
P00469 DV N  7 2 SV N
P00470 DV N  2 0 S- N
P00471 DV N  9 2 S- N
P00474 DV N  5 2 -V J
P00475 DV N  6 0 SV N
P00476 DV N  6 2 SV N
P00477 DV N  9 2 S- N
P00478 DV N  6 2 SV N
P00479 DV N  9 2 SV N
P00480 DV N  1 0 S- N
P00481 DV N  3 0 -V N
P00482 DV N  3 0 -V J
P00483 DV N  6 2 -V N
P00484 DV N  8 2 SV J
P00485 DV N  9 2 -V J
P00487 DV N  6 2 SV N
P00488 DV N  6 2 -V N
P00489 DV N  6 2 -V N
P00490 DV N  6 2 -V N
P00491 DV N  6 2 -V N
P00492 DV N  1 0 S- J
P00493 DV N  9 2 -V N
P00494 DV N  9 2 -V N
P00495 DV N  9 2 -V N
P00496 DV N  9 2 -V N
P00497 DV N  9 2 -V N
P00498 DV N  9 2 -V N
P00499 DV N  9 2 -V N
P00503 DV N  9 2 S- N
P00504 DV N  8 4 SV N
P00505 DV N  9 2 S- J
P00506 DV N  9 2 -V N
P00507 DV N  9 2 -V N
P00508 DV N  9 2 -V N
P00509 DV N  9 2 -V N
P00510 DV N  9 2 -V N
P00511 DV N  9 2 -V N
P00512 DV N  9 2 -V N
P00513 DV N  9 2 SV N
P00539 DV N  6 2 SV J
P00540 DV N  6 2 SV N
P00541 DV N  6 2 SV N
P00542 DV N  6 2 SV N
P00543 DV N  6 2 SV N
P00544 DV N  6 2 SV N
P00545 DV N  6 2 SV N
P00546 DV N  6 2 SV N
P00547 DV N  6 2 SV N
P00548 DV N  6 2 SV N
P00550 DV N  9 2 SV N
P00553 DV N  9 2 SV N
P00566 DV N  9 2 SV N
P00570 DV N  9 2 SV N
P00574 DV S  3 - S- N
P00575 DV S  3 - S- N
P00576 DV S  3 - S- N
P00577 DV S  3 - S- N
P00578 DV S  3 - S- N
P00579 DV S  3 - S- N
P00580 DV N  6 2 SV N
P00581 DV N  6 2 SV N
P00582 DV N  6 2 SV N
P00583 DV N  6 2 SV N
P00584 DV N  6 2 SV N
P00591 DV N  9 2 SV J
P00595 DV N  9 2 SV J
P00600 DV S 50 - S- J
P00601 DV S 50 - S- J
P00606 DV N  3 0 -V N
P00611 DV N  9 2 SV N
P00612 DV N  9 2 SV N
P00613 DV N  9 2 -V N
P00614 DV N  9 2 -V J
P00616 DV N  9 2 -V J
P00617 DV N  7 2 -V N
P00620 DV N  9 2 -V N
P00626 DV N  9 2 SV N
P00627 DV N  6 2 SV N
P00628 DV N 10 2 SV N
P00629 DV N 10 2 SV N
P00630 DV N  9 2 SV N
P00631 DV N  6 2 SV -
P00632 DV N  6 2 SV -
P00633 DV N  6 2 SV -
P00635 DV N  9 2 -V N
P00636 DV N  9 2 -V N
P00642 DV N  6 2 SV N
P00643 DV N  6 2 SV N
P00644 DV N  6 2 SV N
P00645 DV N  6 2 SV N
P00646 DV N  6 2 SV N
P00647 DV N  9 2 -V N
P00648 DV N  9 2 -V N
P00649 DV N  9 2 -V N
P00650 DV N  9 2 -V N
P00653 DV N  7 2 -V N
P00654 DV N  7 2 -V N
P00655 DV N  9 2 -V N
P00656 DV N  9 2 -V N
P00657 DV N 10 2 SV J
P00658 DV N 10 2 SV J
P00661 DV N  9 2 -V J
P00662 DV N  9 2 -V J
P00670 DV N  9 2 SV J
P00672 DV N 10 2 SV N
P00673 DV N  9 2 SV N
P00674 DV N  9 2 SV N
P00675 DV N 10 2 SV N
P00676 DV N 10 2 SV N
P00677 DV N 10 2 SV N
P00678 DV N 10 2 SV N
P00679 DV N 10 2 SV N
P00680 DV N 10 2 SV N
P00681 DV N 10 2 SV N
P00682 DV N 10 2 SV N
P00683 DV N 10 2 SV N
P00684 DV N 10 2 SV N
P00685 DV N 10 2 SV N
P00686 DV N 10 2 SV N
P00687 DV N 10 2 SV N
P00688 DV N 10 2 SV N
P00689 DV N 10 2 SV N
P00690 DV N  2 - SV J
P00691 DV N  2 - SV J
P00692 DV N  2 - SV J
P00693 DV N  9 2 SV N
P00694 DV N  9 2 -V N
P00695 DV N 10 2 SV N
P00696 DV N  9 2 -V J
P00697 DV N  9 2 -V J
P00698 DV N  9 2 -V J
P00699 DV N  9 2 -V J
P00701 DV N  9 2 SV N
P00702 DV N  9 2 SV N
P00703 DV N  9 2 SV N
P00704 DV N  9 2 SV N
P00705 DV N  9 2 SV N
P00706 DV N  9 2 SV N
P00707 DV N  9 2 SV N
P00708 DV N  9 2 SV N
P00709 DV N  9 2 SV N
P00710 DV N  9 2 SV N
P00711 DV N  9 2 SV N
P00712 DV N  9 2 SV N
P00713 DV N  9 2 SV N
P00714 DV N  9 2 SV N
P00715 DV N  9 2 SV N
P00716 DV N  9 2 SV N
P00717 DV N  9 2 SV N
P00718 DV N  9 2 SV N
P00719 DV N  9 2 SV N
P00720 DV N  9 2 SV N
P00721 DV N  9 2 SV N
P00722 DV N  9 2 SV N
P00723 DV N  9 2 SV N
P00724 DV N  9 2 SV N
P00725 DV N  9 2 SV N
P00726 DV N  9 2 SV N
P00727 DV N  9 2 SV N
P00728 DV N  9 2 SV N
P00729 DV N  9 2 SV N
P00730 DV N  9 2 SV N
P00745 DV N  9 2 -V N
P00746 DV N  9 2 -V J
P00751 DV N  6 2 SV N
P00752 DV N  6 2 SV N
P00753 DV N  6 2 SV N
P00754 DV N  6 2 SV N
P00755 DV N  6 2 SV N
P00756 DV N  6 2 SV N
P00757 DV N  6 2 SV N
P00758 DV N  6 2 SV N
P00759 DV N  6 2 SV N
P00760 DV N  6 2 SV N
P00761 DV N  1 - SV N
P00762 DV N  1 - SV N
P00763 DV N  6 2 SV N
P00764 DV N  6 2 SV N
P00765 DV N  6 2 SV N
P00766 DV N  6 2 SV N
P00767 DV N  6 2 SV N
P00768 DV N  6 2 SV N
P00769 DV N  1 - SV N
P00771 DV N  6 2 SV -
P00772 DV N  6 2 SV N
P00773 DV N  9 2 SV N
P00774 DV N  9 2 SV N
P00775 DV N  9 2 SV N
P00776 DV N  6 2 SV N
P00777 DV N  6 2 SV N
P00778 DV N  1 - S- J
P00804 DV N  4 0 -V N
P00829 DV N  6 2 SV N
P00830 DV D 11 - S- N
P00831 DV D 11 - S- N
P00832 DV D 11 - S- N
P00833 PS D 11 - S- N
P00834 DV D 11 - S- N
P00835 DV N  2 - S- J
P00836 DV S  6 - S- N
P00837 DV D 11 - S- J
P00839 DV N  1 0 SV N
P00842 DV S  4 - S- N
P00844 DV S 12 - S- N
P00845 DV N  1 0 S- N
P00847 PS S  4 - S- N
P00849 PS D 11 - S- N
P00850 DV S  2 - S- N
P00851 DV N  9 2 SV J
P00852 DV N  9 2 SV J
P00853 DV N  9 2 SV J
P00858 DV N  2 - S- N
P00859 DV S  3 - S- N
P00860 DV N  1 - S- J
P00862 DV S  8 - S- N
P00863 DV N  1 - S- N
P00864 DV D 11 - -V J
P00868 DV N  2 - S- N
P00869 DV N  1 0 S- N
P00870 DV N  9 2 SV N
P00871 DV N  1 - S- N
P00873 DV N  1 0 S- N
P00874 DV N  9 2 SV N
P00875 DV N  1 - S- J
P00876 DV N  6 2 S- N
P00877 DV N  6 2 S- J
P00878 DV N  1 0 S- N
P00879 DV N  1 - S- N
P00880 DV N  1 0 S- N
P00881 DV N  9 2 SV N
P00882 DV N  9 2 SV N
P00883 DV N  9 2 SV N
P00884 DV N  7 3 SV N
P00885 DV N  1 - -V N
P00886 DV N  1 0 S- N
P00887 DV N  2 - S- N
P00890 DV N  1 0 S- N
P00891 DV N  1 0 S- J
P00893 DV N  2 - S- N
P00895 DV N  1 0 SV J
P00896 DV N  1 0 S- N
P00898 DV N  1 0 S- J
P00899 DV S  9 2 SV J
P00900 DV N  2 0 S- N
P00901 DV N  9 2 S- N
P00902 DV N  9 2 S- N
P00903 DV N  9 2 S- N
P00904 DV N  9 2 S- N
P00905 DV N  9 2 S- N
P00906 DV N  9 2 S- N
P00907 DV N  9 2 S- N
P00908 DV N  9 2 S- N
P00909 DV N  9 2 S- N
P00910 DV N  9 2 S- N
P00915 DV N  7 3 SV N
P00916 DV N  7 3 SV N
P00917 DV N  7 3 SV N
P00918 DV N  7 3 SV N
P00919 DV N  7 3 SV N
P00920 DV N  4 0 S- N
P00921 DV N  8 0 S- N
P00922 DV N  4 0 S- N
P00929 DV N  9 2 SV N
P00930 DV N  2 - S- N
P00934 DV D 11 - S- J
P00937 DV S 15 - S- N
P00938 PS H 50 - S- N
P00943 DV N  9 2 S- N
P00944 DV N  9 2 S- N
P00951 DV N  9 2 SV J
P00953 DV N  9 2 SV J
P00955 DV N  9 2 SV J
P00957 DV N  9 2 SV J
P00959 DV N  9 2 SV J
P00960 DV N  5 0 SV J
P00961 DV N  5 0 SV J
P00962 DV N  5 0 SV J
P00963 DV N  5 0 SV J
P00964 DV N  5 0 SV J
P00965 DV N  5 0 SV J
P00966 DV N  5 0 SV J
P00967 DV N  5 0 SV J
P00968 DV N  5 0 SV J
P00969 DV N  5 0 SV J
P00970 DV S 50 - S- N
P00971 DV S 50 - S- N
P00972 DV S 50 - S- N
P00973 DV S 50 - S- N
P00974 DV S 50 - S- N
P00975 DV S 50 - S- N
P00976 DV S 50 - S- N
P00977 DV S 50 - S- N
P00978 DV S 50 - S- N
P00979 DV S 50 - S- N
P00980 DV S 50 - SV N
P00981 DV S 50 - SV N
P00982 DV S 50 - SV N
P00987 DV N  9 2 SV N
P00988 DV N  4 - S- J
P00989 DV N  6 - S- J
P00999 DV N 10 2 SV N
P01000 PS I 30 - S- N
P01001 PS N  6 0 S- J
P01002 PS S 70 - S- N
P01003 PS S 25 - S- N
P01004 PS S  4 - S- N
P01005 PS D 11 - S- N
P01006 PS S 25 - S- N
P01007 PS D 11 - S- N
P01008 PS S 70 - S- J
P01009 PS S 10 - S- J
P01011 PS D 11 - S- N
P01012 PS D 11 - S- N
P01014 PS S 50 - S- N
P01016 PS N  5 0 S- N
P01018 PS S 20 - S- N
P01020 PS S 20 - S- N
P01022 PS S 50 - S- N
P01024 PS S  4 - S- N
P01026 PS S 15 - S- N
P01027 PS S 15 - S- N
P01031 PS S 30 - S- N
P01032 PS S 50 - S- N
P01033 PS S  1 - S- N
P01034 PS S 99 - S- N
P01035 PS S 99 - S- N
P01036 PS S 15 - S- N
P01037 PS S 15 - S- N
P01038 PS S 20 - S- N
P01039 PS S 15 - S- N
P01040 PS S  1 - S- N
P01041 PS S  4 - S- N
P01042 PS S  4 - S- N
P01043 PS S  4 - S- N
P01044 PS S  3 - S- N
P01045 PS S  1 - S- N
P01050 PS S 50 - S- N
P01073 DV N 10 2 SV N
P01074 DV N 10 2 SV N
P01079 DV N 10 2 SV N
P01097 DV N  6 2 -V N
P01100 DV N  9 2 SV N
P01101 DV N  3 0 S- J
P01102 DV S  3 - S- N
P01103 DV N  5 0 S- N
P01104 DV N  4 0 S- N
P01105 DV N  6 0 S- N
P01106 DV N 10 0 S- N
P01107 DV S  7 - S- N
P01108 DV N  1 0 S- N
P01109 DV N  5 2 S- N
P01110 DV S  3 - S- N
P01111 PS N  4 - S- J
P01112 DV S  1 - S- N
P01113 DV S  3 - S- N
P01114 DV N  4 0 S- J
P01118 DV N  9 2 SV N
P01119 DV N  5 2 S- N
P01151 DV S  4 - S- N
P01152 DV N  4 0 S- N
P01153 DV N  9 2 S- N
P01154 DV N  6 2 S- N
P01155 DV S  1 - S- N
P01156 DV N  6 2 SV N
P01157 DV S  4 - S- N
P01158 DV N  2 0 S- N
P01159 DV N  9 2 S- N
P01160 DV N  6 2 S- N
P01161 DV N  9 2 S- N
P01162 DV N  6 2 S- N
P01164 DV N 10 2 SV N
P01165 DV S  1 - S- N
P01166 DV N 10 2 SV N
P01168 DV N 10 2 SV N
P01171 DV S 10 - S- N
P01172 DV S 10 - S- N
P01173 DV N  5 2 S- N
P01174 DV N  5 2 S- N
P01176 DV N 10 2 SV N
P01178 DV N 10 2 SV N
P01179 DV N 10 2 SV N
P01182 DV S  3 - S- N
P01183 DV S  3 - S- N
P01185 DV N 10 2 SV J
P01189 DV N 10 2 S- N
P01190 DV N  9 2 S- N
P01191 DV N  9 2 S- N
P01192 DV N  9 2 S- N
P01193 DV N  9 2 S- N
P01194 DV N  9 2 S- N
P01195 DV N  9 2 S- N
P01213 DV N  9 2 SV N
P01241 DV S  7 - S- N
P01242 DV S  7 - S- N
P01243 DV S  7 - S- N
P01244 DV S  7 - S- N
P01245 DV S  7 - S- N
P01246 DV S  7 - S- N
P01247 DV S  7 - S- N
P01248 DV S  7 - S- N
P01249 DV S  7 - S- N
P01250 DV S  7 - S- N
P01251 DV S  7 - S- N
P01252 DV S  7 - S- N
P01253 DV S  7 - S- N
P01254 DV S  7 - S- N
P01255 DV S  7 - S- N
P01256 DV S  7 - S- N
P01257 DV S  7 - S- N
P01258 DV S  7 - S- N
P01259 DV S  7 - S- N
P01260 DV D 11 - S- N
P01281 DV N  7 3 SV N
P01282 DV N  7 3 SV N
P01283 DV N 10 2 -V N
P01300 DV N  9 2 SV N
P01301 DV N  9 2 SV N
P01302 DV N 10 2 SV J
P01303 DV N  6 2 SV N
P01304 DV N  9 2 SV N
P01307 DV N 10 2 SV N
P01308 DV N  9 2 SV N
P01325 DV N 10 2 SV N
P01326 DV N 10 2 SV J
P01329 DV N 10 2 SV N
P01330 DV N 10 2 SV N
P01331 DV N 10 2 SV N
P01332 DV N 10 2 SV N
P01333 DV N 10 2 SV N
P01334 DV N 10 2 SV N
P01336 DV N 10 2 SV N
P01340 DV N 10 2 SV N
P01341 DV N 10 2 SV N
P01342 DV N 10 2 SV N
P01343 DV N 10 2 SV N
P01344 DV N 10 2 SV N
P01345 DV N 10 2 SV N
P01346 DV N 10 2 SV N
P01347 DV N 10 2 SV N
P01348 DV N 10 2 SV N
P01349 DV N 10 2 SV N
P01350 DV N 10 2 SV N
P01351 DV N 10 2 SV N
P01352 DV N 10 2 SV N
P01353 DV N 10 2 SV N
P01354 DV N 10 2 SV N
P01355 DV N 10 2 SV N
P01356 DV N 10 2 SV N
P01357 DV N 10 2 SV N
P01358 DV N 10 2 SV N
P01359 DV N 10 2 SV N
P01360 DV N 10 2 SV N
P01361 DV N 10 2 SV N
P01362 DV N 10 2 SV N
P01363 DV N 10 2 SV N
P01364 DV N 10 2 SV N
P01365 DV N 10 2 SV N
P01366 DV N 10 2 SV N
P01367 DV N 10 2 SV N
P01368 DV N 10 2 SV N
P01369 DV N 10 2 SV N
P01370 DV N 10 2 SV N
P01371 DV N 10 2 SV N
P01372 DV N 10 2 SV N
P01373 DV N 10 2 SV N
P01374 DV N 10 2 SV N
P01375 DV N 10 2 SV N
P01376 DV N 10 2 SV N
P01377 DV N 10 2 SV N
P01378 DV N 10 2 SV N
P01379 DV N 10 2 SV N
P01380 DV N 10 2 SV N
P01381 DV N 10 2 S- N
P01382 DV N 10 2 S- N
P01383 DV N 10 2 S- N
P01384 DV N 10 2 S- N
P01385 DV N 10 2 S- N
P01386 DV N 10 2 SV N
P01387 DV N 10 2 SV N
P01388 DV N 10 2 SV N
P01389 DV N 10 2 SV N
P01390 DV N 10 2 SV N
P01391 DV N 10 2 SV J
P01392 DV N 10 2 SV J
P01393 DV N  6 2 S- N
P01394 DV N  6 2 SV N
P01395 DV N  9 2 SV N
P01396 DV N  6 2 S- N
P01397 DV N  9 2 SV N
P01398 DV N 10 2 SV N
P01399 DV N  6 2 S- N
P01400 DV S  3 - S- N
P01401 DV D 11 - S- N
P01402 DV S  3 - S- N
P01404 DV N  6 2 S- N
P01417 DV N 10 2 SV N
P01419 DV N 10 2 SV N
P01423 DV N  6 2 SV N
P01424 DV N  6 2 SV N
P01458 DV N 10 2 SV N
P01459 DV N 10 2 SV N
P01473 DV D 11 - S- N
P01474 DV D 11 - S- N
P01475 DV N  2 - S- N
P01476 DV N  1 - S- N
P01477 DV N  1 - S- N
P01478 DV S 34 - S- N
P01479 DV S 11 - S- N
P01480 DV S 50 - S- N
P01481 DV S 50 - S- N
P01482 DV N  6 2 SV N
P01483 DV N  6 2 SV N
P01485 DV D 11 - S- N
P01486 DV N  5 - S- N
P01487 DV D 11 - S- N
P01497 DV N  6 2 SV N
P01498 DV N  6 2 S- N
P01499 DV N  6 2 SV N
P01503 DV N 10 2 SV N
P01504 DV N 10 2 SV N
P01505 DV N 10 2 SV N
P01506 DV N 10 2 SV N
P01507 DV N 10 2 SV N
P01508 DV N 10 2 SV N
P01509 DV N 10 2 SV N
P01510 DV N 10 2 SV N
P01511 DV N  1 - S- N
P01512 DV N  6 2 SV N
P01514 DV N  6 2 SV N
P01515 DV N  6 2 SV N
P01516 DV N  6 2 SV N
P01517 DV N  6 2 SV N
P01518 DV N  6 2 SV N
P01519 DV N  6 2 S- N
P01520 DV N  6 2 S- N
P01521 DV N  6 2 S- N
P01522 DV N  6 2 S- N
P01523 DV N  6 2 S- N
P01524 DV N  6 2 -V N
P01525 DV N  6 2 -V N
P01526 DV N  6 2 -V N
P01527 DV N  6 2 -V N
P01528 DV N  6 2 -V N
P01530 DV N 10 2 SV N
P01531 DV N  6 2 SV N
P01532 DV N  6 2 SV N
P01533 DV N  6 2 SV N
P01534 DV N  6 2 SV N
P01535 DV N  6 2 SV N
P01536 DV N  2 0 SV N
P01537 DV N  2 0 SV N
P01538 DV N  2 0 SV N
P01539 DV N  2 0 SV N
P01540 DV N  2 0 SV N
P01543 DV N 10 2 -V J
P01549 DV N  9 2 SV N
P01550 DV N  6 2 SV N
P01551 DV N  6 2 SV N
P01552 DV N  6 2 SV N
P01553 DV N  6 2 SV N
P01554 DV N  6 2 SV N
P01557 DV N  9 2 -V N
P01558 DV N  9 2 -V N
P01559 DV N  9 2 -V N
P01560 DV N  1 - S- N
P01561 DV N  1 - S- N
P01562 DV N  1 - S- N
P01563 DV N  1 - S- N
P01564 DV N  1 - S- N
P01565 DV N 10 2 SV N
P01566 DV N 10 2 SV N
P01567 DV N 10 2 SV N
P01568 DV N 10 2 SV N
P01569 DV N 10 2 SV N
P01571 DV S  2 - S- N
P01572 DV N  2 - S- N
P01573 DV D 11 - S- N
P01574 DV D 11 - S- N
P01575 DV N  2 - S- N
P01576 DV N  1 - S- N
P01577 DV N  1 - S- N
P01578 DV S 34 - S- N
P01579 DV S 11 - S- N
P01580 DV S 50 - S- N
P01581 DV S 50 - S- N
P01582 DV S  2 - S- N
P01583 DV D 11 - S- N
P01599 DV N 10 2 SV N
P01650 DV D 11 - S- J
P01651 DV S  1 - S- J
P01652 DV D 11 - S- J
P01653 DV N 10 2 -V N
P01658 DV N  9 2 SV N
P01659 DV N  9 2 SV N
P01668 DV N 10 2 SV N
P01680 DV D 11 - S- N
P01681 DV S  1 - S- N
P01682 DV D 11 - S- N
P01683 DV S  6 - S- N
P01684 DV N  3 0 S- N
P01688 DV N 10 2 SV N
P01690 DV N  1 0 S- J
P01691 DV N  1 0 S- J
P01700 DV N  1 - S- N
P01701 DV N  6 2 SV N
P01702 DV N  9 2 S- N
P01703 DV N  9 2 SV N
P01704 DV N  1 - SV N
P01705 DV N  2 - S- N
P01710 DV N  1 - SV N
P01769 DV N 10 2 SV J
P01808 DV N  9 2 -V N
P01831 DV N  6 2 SV N
P01832 DV N 10 2 SV N
P01884 DV N  6 2 SV N
P01885 DV N  1 0 S- N
P01886 DV N  1 0 S- N
P01888 DV N  6 2 SV N
P01889 DV N  6 2 SV N
P01891 DV S  3 - S- N
P01892 DV N 15 0 S- N
P01893 DV D 11 - S- N
P01894 DV D 11 - S- N
P01895 DV S  1 - S- N
P01896 DV S 35 - S- N
P01898 DV N 10 2 -V N
P01899 DV N 10 2 -V N
P01904 DV N 10 2 SV N
P01935 DV N  6 2 S- N
P01936 DV N  6 2 S- N
P01937 DV N  6 2 S- N
P01938 DV N  6 2 S- N
P01939 DV N  6 2 S- N
P01945 DV N  6 2 -V N
P01946 DV N  6 2 -V N
P01947 DV N  6 2 -V N
P01948 DV N  6 2 -V N
P01949 DV N  6 2 -V N
P01955 DV N  1 0 SV N
P01956 DV N  1 0 SV N
P01957 DV N  1 0 SV N
P01958 DV N  1 0 SV N
P01959 DV N  1 0 SV N
P01967 DV N 10 2 SV N
P01968 DV N 10 2 SV N
P01970 DV N  9 2 SV N
P01971 DV N  1 0 S- N
P01972 DV N  1 0 S- N
P01973 DV N  1 0 S- N
P01974 DV N  1 0 S- N
P01975 DV N  1 0 S- N
P01978 DV N 10 2 SV N
P01979 DV N 10 2 SV N
P01981 DV N  7 2 SV N
P01985 DV N  6 2 SV N
P01986 DV N  6 2 SV N
P01987 DV N  6 2 SV N
P01988 DV N  6 2 SV N
P01989 DV N  6 2 SV N
P01990 DV N  6 2 SV N
P01991 DV N  6 2 SV N
P01992 DV N  6 2 SV N
P01993 DV N  6 2 SV N
P01995 DV N  1 0 S- N
P01996 DV N  1 0 S- N
P01997 DV N  1 0 S- N
P01998 DV N  1 0 S- N
P01999 DV N  1 0 S- N
P02036 DV N 10 2 SV N
P02037 DV N 10 2 SV N
P02038 DV N 10 2 SV N
P02039 DV N 10 2 SV N
P02040 DV N 10 2 SV N
P02041 DV N  9 2 S- N
P02071 DV S  1 - S- N
P02073 DV N 10 2 SV N
P02079 DV N 10 2 SV N
P02085 DV N 10 2 SV N
P02086 DV N 10 2 SV N
P02087 DV N 10 2 SV N
P02088 DV N 10 2 SV N
P02089 DV N 10 2 SV N
P02090 DV N 10 2 SV N
P02091 DV N 10 2 SV N
P02092 DV N  9 2 SV N
P02093 DV N  9 2 SV N
P02094 DV N  9 2 SV N
P02095 DV S  4 - S- N
P02096 DV N  2 0 S- N
P02097 DV N 10 2 SV N
P02098 DV N 10 2 SV N
P02099 DV N  1 - SV N
P02100 DV S  3 - S- N
P02101 DV S  3 - S- N
P02102 DV S  3 - S- N
P02103 DV S  3 - S- N
P02104 DV S  3 - S- N
P02105 DV S  3 - S- N
P02106 DV S  3 - S- N
P02107 DV S  3 - S- N
P02108 DV S  1 - S- N
P02109 DV S  1 - S- N
P02110 DV S  1 - S- N
P02111 DV S  3 - S- N
P02112 DV N  3 0 S- N
P02113 DV N  9 2 SV N
P02114 DV N  6 2 S- N
P02115 DV N 10 2 -V N
P02116 DV N 10 2 SV N
P02117 DV N  8 4 S- N
P02118 DV N  9 2 S- N
P02119 DV N  9 2 S- J
P02120 DV N  9 2 S- N
P02121 DV N  6 2 S- N
P02122 DV N  2 0 S- N
P02123 DV N  2 0 S- N
P02124 DV N  2 0 S- N
P02130 DV N  1 0 SV J
P02131 DV D 11 - -V J
P02132 DV N  1 0 SV J
P02133 DV N  1 0 SV J
P02134 DV N  1 0 SV J
P02135 DV N  1 0 SV J
P02136 DV N  1 0 SV J
P02137 DV N  1 0 SV J
P02138 DV N  2 0 SV J
P02139 DV N  2 0 SV J
P02140 DV N  2 0 SV J
P02141 DV N  2 - SV J
P02142 DV S 34 - S- J
P02143 DV S 11 - S- J
P02144 DV S 35 - S- J
P02145 DV S 35 - S- J
P02146 DV S 35 - S- J
P02147 DV S  2 - S- J
P02148 DV S 35 - S- J
P02149 DV S 34 - SV J
P02150 DV S 34 - SV J
P02151 DV S 34 - SV J
P02152 DV S 34 - SV J
P02153 DV S  3 - S- J
P02160 DV N  1 0 SV N
P02161 DV N  1 0 SV N
P02162 DV N  1 0 SV N
P02163 DV N  1 0 SV N
P02164 DV N  1 0 SV N
P02165 DV N  1 0 SV N
P02166 DV N  1 0 SV N
P02167 DV N  1 0 SV N
P02168 DV N  1 0 SV N
P02169 DV N  1 0 SV N
P02170 DV N  1 0 SV J
P02171 DV N  1 0 S- N
P02172 DV N  9 2 SV N
P02173 DV N  7 3 SV N
P02174 DV N  1 0 S- N
P02175 DV N  9 2 SV N
P02176 DV N  7 3 SV N
P02177 DV N  1 0 S- N
P02178 DV N  9 2 SV N
P02179 DV N  7 3 SV N
P02180 DV N  1 0 S- N
P02181 DV N  9 2 SV N
P02182 DV N  7 3 SV N
P02183 DV N  1 0 S- N
P02184 DV N  9 2 SV N
P02185 DV N  7 3 SV N
P02186 DV N  1 0 SV N
P02187 DV N  1 0 SV N
P02188 DV N  1 0 SV N
P02189 DV N  1 0 SV N
P02190 DV N  1 0 SV N
P02192 DV N  9 2 SV N
P02225 DV S  1 - S- N
P02226 DV S  1 - S- N
P02227 DV S  1 - S- N
P02252 DV N  1 - S- N
P02253 DV D 11 - S- N
P02254 DV D 11 - S- N
P02255 DV N  2 - S- N
P02256 DV N  1 - S- N
P02261 DV N  1 - SV N
P02262 DV N  1 - SV J
P02263 DV N  1 - SV N
P02264 DV N  1 0 S- J
P02265 DV N 10 2 S- J
P02266 DV N  2 - S- J
P02267 DV N  2 0 SV N
P02268 DV S  2 - SV N
P02269 DV N  1 0 S- N
P02270 DV N  1 0 S- J
P02271 DV N  1 0 S- J
P02272 DV S  4 - S- N
P02273 DV N  1 0 SV J
P02274 DV N  1 0 SV N
P02275 DV N 10 2 S- J
P02276 DV N 10 2 S- J
P02277 DV N 10 2 S- J
P02278 DV N 10 2 S- J
P02279 DV N 10 2 S- J
P02280 DV N 10 2 S- J
P02281 DV N  7 2 S- J
P02282 DV N  7 2 S- J
P02283 DV N  7 2 S- J
P02284 DV N  7 2 S- J
P02285 DV N  7 2 S- J
P02286 DV N  7 2 S- J
P02287 DV N  7 2 S- J
P02288 DV N  7 2 S- J
P02290 DV N  7 2 S- J
P02291 DV N 10 2 S- J
P02292 DV N  7 2 S- J
P02293 DV N  1 0 S- N
P02294 DV N 10 2 S- N
P02295 DV N 10 2 SV N
P02296 DV N 10 2 SV N
P02297 DV N 10 2 SV N
P02298 DV N  6 2 SV N
P02299 DV N  6 2 SV N
P02300 DV N  6 2 SV N
P02312 DV N  6 2 SV N
P02321 DV N  1 0 S- J
P02324 DV S  5 - S- N
P02338 DV N  1 - S- J
P02345 DV N  7 2 S- N
P02346 DV N  6 0 S- N
P02356 DV N  1 0 S- N
P02358 DV N  4 - S- N
P02375 DV N  1 - S- N
P02376 DV N  6 2 SV N
P02378 DV N  4 - -V J
P02381 DV N 10 2 SV J
P02384 DV N 10 2 SV J
P02387 DV N 10 2 SV J
P02394 DV N 10 2 SV J
P02397 DV N  6 2 SV N
P02404 DV N  6 2 S- J
P02409 DV N  1 0 S- N
P02410 DV N  6 2 SV N
P02414 DV N  6 2 SV N
P02415 DV N  6 2 SV N
P02416 DV N 10 2 -V N
P02417 DV N  6 2 SV N
P02418 DV N 10 2 SV N
P02419 DV N  6 2 SV N
P02427 DV N 10 2 SV N
P02433 DV N  9 2 SV N
P02435 DV N  9 2 -V N
P02454 DV N  1 0 -V J
P02458 DV N 10 2 -V N
P02472 DV N 10 2 -V N
P02473 DV N  1 0 SV N
P02482 DV N  6 2 SV N
P02484 DV N  9 2 SV N
P02485 DV N  9 2 SV N
P02486 DV N  1 - S- N
P02487 DV N  9 2 SV N
P02504 DV N  8 4 S- N
P02505 DV N  6 2 SV N
P02506 DV N 11 4 SV N
P02507 DV N 10 2 SV N
P02508 DV N  6 2 SV N
P02509 DV N  6 2 S- N
P02510 DV N  6 2 S- N
P02511 DV N 10 2 SV J
P02512 DV N  5 2 S- N
P02514 DV N  5 2 S- N
P02515 DV N  1 0 S- N
P02516 DV N  1 0 S- J
P02517 DV S  4 - S- N
P02518 DV N  5 2 SV N
P02519 DV N  1 0 SV N
P02520 DV N  1 0 SV N
P02521 DV N  1 0 SV N
P02522 DV N  1 0 SV N
P02523 DV N  1 0 SV N
P02524 DV N  1 0 SV N
P02525 DV N  1 0 SV N
P02526 DV N  2 0 SV N
P02527 DV N  1 0 SV N
P02528 DV N  5 2 SV N
P02529 DV N  2 0 SV N
P02530 DV N  6 2 S- N
P02531 DV N  1 0 S- N
P02532 DV N  6 2 SV N
P02533 DV N  1 0 SV N
P02534 DV N  6 2 SV N
P02535 DV N  6 2 SV J
P02536 DV N  1 0 SV J
P02537 DV N  6 2 SV J
P02538 DV N  1 0 SV J
P02539 DV N  6 2 S- J
P02543 DV N  6 2 SV N
P02549 DV N  6 2 SV N
P02551 DV N  6 2 SV N
P02552 DV N  6 2 SV N
P02554 DV N  6 2 SV N
P02555 DV N  6 2 SV N
P02556 DV N  1 0 SV N
P02557 DV N  1 0 SV N
P02558 DV N  1 0 SV N
P02559 DV N  1 0 SV N
P02560 DV N  1 0 SV N
P02561 DV N  6 2 SV N
P02562 DV N  6 2 SV N
P02563 DV N  6 2 SV N
P02564 DV N  6 2 SV N
P02565 DV N  6 2 SV N
P02567 DV N  9 2 SV N
P02568 DV N  9 2 SV N
P02569 DV N  9 2 SV N
P02571 DV N  2 0 SV N
P02572 DV N 10 2 SV N
P02573 DV N  2 0 SV N
P02574 DV N 10 2 SV N
P02575 DV N  2 0 SV N
P02576 DV N 10 2 SV N
P02577 DV N  2 0 SV N
P02578 DV N 10 2 SV N
P02579 DV N  2 0 SV N
P02580 DV N 10 2 SV N
P02581 DV N 10 2 SV N
P02582 DV N  1 - SV N
P02589 DV N  6 2 SV N
P02590 DV N 10 2 SV N
P02591 DV N  5 2 SV N
P02592 DV N  8 2 S- N
P02593 DV N 10 2 SV N
P02594 DV N  9 2 -V J
P02595 DV N 10 2 S- N
P02596 DV N  6 2 SV N
P02597 DV N 10 2 SV N
P02598 DV N 10 2 SV N
P02599 DV N 10 2 SV N
P02600 DV N  2 0 S- N
P02601 DV N  2 0 S- N
P02602 DV N  6 2 SV N
P02603 DV N  6 2 SV N
P02604 DV N 10 2 -V N
P02605 DV N 10 2 -V N
P02607 DV N  1 0 S- N
P02608 DV N  9 2 SV N
P02609 DV N  9 2 SV N
P02610 DV N  9 2 SV N
P02614 DV N  1 0 S- N
P02615 DV N  1 0 S- N
P02616 DV N  9 2 -V N
P02618 DV N  6 2 SV N
P02619 DV N  6 2 SV N
P02621 DV N  1 0 SV N
P02622 DV N  1 0 SV N
P02623 DV N  1 0 SV N
P02624 DV N  1 0 SV N
P02625 DV N  1 0 SV N
P02626 DV N 10 2 SV N
P02627 DV N 10 2 SV N
P02628 DV N 10 2 SV N
P02629 DV N 10 2 SV N
P02630 DV N 10 2 SV N
P02641 DV N  1 - S- N
P02651 DV N 10 2 S- N
P02653 DV N 10 2 S- N
P02657 DV N  1 0 S- N
P02658 DV N  1 0 S- N
P02659 DV N 10 2 S- N
P02660 DV N 10 2 S- N
P02663 DV N 10 2 S- N
P02664 DV N 10 2 S- N
P02665 DV N  7 3 SV N
P02666 DV N  7 3 SV N
P02667 DV N  7 3 SV N
P02668 DV N  7 3 SV N
P02669 DV N  7 3 SV N
P02670 DV N  5 2 S- N
P02671 DV N  1 0 S- N
P02672 DV N  1 0 S- N
P02673 DV N  1 0 S- N
P02674 DV N  1 0 S- N
P02675 DV N  1 0 S- N
P02676 DV N  9 2 SV N
P02677 DV N  9 2 SV N
P02678 DV N  9 2 SV N
P02679 DV N  9 2 SV N
P02680 DV N  9 2 SV N
P02681 DV N  7 3 SV N
P02682 DV N  7 3 SV N
P02683 DV N  7 3 SV N
P02684 DV N  7 3 SV N
P02685 DV N  7 3 SV N
P02686 DV N 10 2 S- N
P02687 DV N 10 2 S- N
P02689 DV N  7 3 SV N
P02690 DV N  7 3 SV N
P02691 DV N  7 3 SV N
P02692 DV N  7 3 SV N
P02693 DV N  7 3 SV N
P02694 DV N  6 2 SV N
P02696 DV N  7 3 SV N
P02697 DV N  7 3 SV N
P02698 DV N  7 3 SV N
P02699 DV N  7 3 SV N
P02700 DV N  7 3 SV N
P02731 DV N 10 2 S- N
P02732 DV N 10 2 S- N
P02737 DV N  9 2 -V N
P02738 DV N  9 2 -V N
P02739 DV N  9 2 -V N
P02740 DV N  9 2 -V N
P02741 DV N  9 2 -V N
P02744 DV N 10 2 S- N
P02746 DV S 34 - S- N
P02747 DV N 10 2 S- N
P02750 DV N 10 2 -V N
P02780 DV N 10 2 S- N
P02781 DV N  9 2 -V N
P02782 DV N  9 2 -V N
P02783 DV N  9 2 -V N
P02784 DV N  9 2 -V N
P02785 DV N  9 2 -V N
P02786 DV N 10 2 S- N
P02787 DV N  6 2 SV N
P02788 DV N  6 2 SV N
P02809 DV N  1 0 SV N
P02810 DV N  6 2 SV N
P02811 DV N  2 0 S- N
P02812 DV N  2 0 S- N
P02813 DV N  2 0 S- N
P02814 DV N  2 0 S- N
P02815 DV N  2 0 S- N
P02816 DV N  2 0 S- N
P02817 DV N  2 0 S- N
P02818 DV N  2 0 S- N
P02819 DV N  2 0 S- N
P02820 DV N  2 0 S- N
P02821 DV N  2 0 S- N
P02822 DV N  2 0 S- N
P02823 DV N  2 0 S- N
P02824 DV N  2 0 S- N
P02829 DV N  2 0 S- N
P02830 DV N  6 2 SV N
P02831 DV N 10 2 -V N
P02832 DV N  1 0 S- N
P02833 DV N  9 2 -V N
P02836 DV S  1 - S- N
P02838 DV N  6 2 SV N
P02843 DV N 10 2 SV N
P02844 DV N  1 - SV N
P02845 DV S  1 - S- J
P02846 DV N  1 0 S- N
P02847 DV N  6 2 -V N
P02851 DV N  1 0 SV J
P02852 DV N  1 0 -V J
P02853 DV N  6 2 SV N
P02854 DV N  1 0 S- N
P02855 DV N  1 0 S- N
P02856 DV N  1 0 S- N
P02857 DV S 15 - -V J
P02860 DV N  1 - S- N
P02861 PS N  1 0 S- N
P02864 DV N  1 - S- N
P02865 DV N 10 2 SV N
P02866 DV N 10 2 SV N
P02867 DV N  1 0 S- J
P02872 DV N  1 0 SV N
P02875 DV N 10 2 -V N
P02876 DV N  1 - S- N
P02887 DV N  1 - S- N
P02888 DV N  1 0 S- N
P02889 DV N  1 0 S- J
P02891 DV N 10 2 SV N
P02892 DV N 10 2 SV N
P02894 DV N  7 2 -V N
P02897 DV N  7 2 -V N
P02898 DV N  1 - S- N
P02930 DV D 11 - S- N
P02933 DV D 11 - S- N
P02934 DV N 10 2 SV N
P02935 DV D 11 - S- N
P02936 PS S  4 - S- N
P02939 DV N  1 0 SV N
P02940 DV N  1 0 S- N
P02941 DV D 11 - S- N
P02942 DV D 11 - S- N
P02950 DV N  1 - S- J
P02951 DV D 11 - S- N
P02953 DV D 11 - S- N
P02955 DV D 11 - S- N
P02957 DV D 11 - S- J
P02959 DV N  3 0 SV J
P02960 DV N 10 2 SV N
P02961 DV N 10 2 S- N
P02963 DV D 11 - SV N
P02964 DV D 11 - SV N
P02965 DV D 11 - SV N
P02966 DV D 11 - S- N
P02967 DV N 10 2 SV N
P02968 DV N 10 2 SV N
P02969 DV N 10 2 SV N
P02989 PS N  1 0 SV J
P02990 DV N 10 2 SV N
P02991 DV N  6 2 SV N
P02993 DV N  1 0 SV N
P02994 DV N  1 0 S- N
P02995 DV S 15 - SV N
P02996 DV S  2 - SV N
P02997 DV N  6 0 S- N
P03000 DV S  4 - S- N
P03001 DV N  8 4 S- N
P03003 DV N  8 4 S- N
P03005 DV S  2 - S- N
P03008 DV N  8 4 S- N
P03009 DV N  8 4 S- N
P03010 DV N  8 4 S- N
P03011 DV N  4 - S- N
P03030 DV N  1 - -V N
P03042 DV S 50 - S- N
P03043 DV S 50 - S- N
P03044 DV S 50 - S- N
P03045 DV S 50 - S- N
P03046 DV S 50 - S- N
P03047 DV S 50 - S- N
P03048 DV S 50 - S- N
P03049 DV S  1 - SV N
P03050 DV N  6 2 SV N
P03051 DV N  6 2 SV N
P03052 DV N  6 2 SV N
P03053 DV N  6 2 SV N
P03054 DV N  6 2 SV N
P03055 DV N  6 2 SV N
P03056 DV N  6 2 SV N
P03057 DV N  6 2 SV N
P03058 DV N  6 2 SV N
P03059 DV N  6 2 SV N
P03060 DV N  6 2 SV N
P03061 DV N  6 2 SV N
P03062 DV N  6 2 SV N
P03063 DV N  6 2 SV N
P03064 DV N  6 2 SV N
P03065 DV D 11 - SV N
P03066 DV D 11 - SV N
P03067 DV D 11 - SV N
P03068 DV D 11 - SV N
P03069 DV D 11 - SV N
P03070 DV D 11 - SV N
P03071 DV D 11 - SV N
P03072 DV D 11 - SV N
P03073 DV D 11 - SV N
P03074 DV D 11 - SV N
P03075 DV D 11 - SV N
P03076 DV D 11 - SV N
P03077 DV D 11 - SV N
P03078 DV D 11 - SV N
P03079 DV S  1 - S- N
P03080 DV N  1 0 SV N
P03081 DV S 50 - SV N
P03082 DV S 50 - SV N
P03083 DV S 50 - SV N
P03084 DV S 50 - SV N
P03085 DV S 50 - SV N
P03086 DV S 50 - SV N
P03087 DV S 50 - SV N
P03088 DV S 50 - SV N
P03089 DV S 50 - SV N
P03090 DV S 50 - SV N
P03091 DV S 50 - SV N
P03092 DV S 50 - SV N
P03093 DV S 50 - SV N
P03094 DV S 50 - SV J
P03095 DV S 50 - SV J
P03096 DV S 50 - SV J
P03097 DV S 50 - SV J
P03098 DV S 50 - SV J
P03099 DV S 50 - SV J
P03100 DV D 11 - S- N
P03101 DV D 11 - S- N
P03102 DV N  9 2 S- N
P03103 DV N  9 2 S- N
P03104 DV N  9 2 S- N
P03105 DV N  9 2 S- N
P03106 DV N  9 2 S- N
P03107 DV N  3 0 S- N
P03108 DV S  5 - S- N
P03109 DV N  5 2 S- J
P03110 DV D 11 - S- N
P03111 DV D 11 - S- N
P03112 DV N  9 2 S- N
P03113 DV N  9 2 S- N
P03114 DV N  9 2 S- N
P03115 DV N  9 2 S- N
P03116 DV N  9 2 S- N
P03117 DV N  3 0 S- N
P03118 DV S  5 - S- N
P03120 DV D 11 - S- N
P03121 DV D 11 - S- N
P03122 DV N  9 2 S- N
P03123 DV N  9 2 S- N
P03124 DV N  9 2 S- N
P03125 DV N  9 2 S- N
P03126 DV N  9 2 S- N
P03127 DV N  3 0 S- N
P03128 DV S  5 - S- N
P03130 DV D 11 - S- N
P03131 DV D 11 - S- N
P03132 DV N  9 2 S- N
P03133 DV N  9 2 S- N
P03134 DV N  9 2 S- N
P03135 DV N  9 2 S- N
P03136 DV N  9 2 S- N
P03137 DV N  3 0 S- N
P03138 DV S  5 - S- N
P03140 DV D 11 - S- N
P03141 DV D 11 - S- N
P03142 DV N  9 2 S- N
P03143 DV N  9 2 S- N
P03144 DV N  9 2 S- N
P03145 DV N  9 2 S- N
P03146 DV N  9 2 S- N
P03147 DV N  3 0 S- N
P03148 DV S  5 - S- N
P03150 DV D 11 - S- N
P03151 DV D 11 - S- N
P03152 DV N  9 2 S- N
P03153 DV N  9 2 S- N
P03154 DV N  9 2 S- N
P03155 DV N  9 2 S- N
P03156 DV N  9 2 S- N
P03157 DV N  3 0 S- N
P03158 DV S  5 - S- N
P03160 DV D 11 - S- N
P03161 DV D 11 - S- N
P03162 DV N  9 2 S- N
P03163 DV N  9 2 S- N
P03164 DV N  9 2 S- N
P03165 DV N  9 2 S- N
P03166 DV N  9 2 S- N
P03167 DV N  3 0 S- N
P03168 DV S  5 - S- N
P03170 DV D 11 - S- N
P03171 DV D 11 - S- N
P03172 DV N  9 2 S- N
P03173 DV N  9 2 S- N
P03174 DV N  9 2 S- N
P03175 DV N  9 2 S- N
P03176 DV N  9 2 S- N
P03177 DV N  3 0 S- N
P03178 DV S  5 - S- N
P03180 DV D 11 - S- N
P03181 DV D 11 - S- N
P03182 DV N  9 2 S- N
P03183 DV N  9 2 S- N
P03184 DV N  9 2 S- N
P03185 DV N  9 2 S- N
P03186 DV N  9 2 S- N
P03187 DV N  3 0 S- N
P03188 DV S  5 - S- N
P03189 DV N  1 0 SV N
P03190 DV D 11 - S- N
P03191 DV D 11 - S- N
P03192 DV N  9 2 S- N
P03193 DV N  9 2 S- N
P03194 DV N  9 2 S- N
P03195 DV N  9 2 S- N
P03196 DV N  9 2 S- N
P03197 DV N  3 0 S- N
P03198 DV S  5 - S- N
P03200 DV D 11 - S- N
P03201 DV D 11 - S- N
P03202 DV N  9 2 S- N
P03203 DV N  9 2 S- N
P03204 DV N  9 2 S- N
P03205 DV N  9 2 S- N
P03206 DV N  9 2 S- N
P03207 DV N  3 0 S- N
P03208 DV S  5 - S- N
P03210 DV D 11 - S- N
P03211 DV D 11 - S- N
P03212 DV N  9 2 S- N
P03213 DV N  9 2 S- N
P03214 DV N  9 2 S- N
P03215 DV N  9 2 S- N
P03216 DV N  9 2 S- N
P03217 DV N  3 0 S- N
P03218 DV S  5 - S- N
P03220 DV D 11 - S- N
P03221 DV D 11 - S- N
P03222 DV N  9 2 S- N
P03223 DV N  9 2 S- N
P03224 DV N  9 2 S- N
P03225 DV N  9 2 S- N
P03226 DV N  9 2 S- N
P03227 DV N  3 0 S- N
P03228 DV S  5 - S- N
P03234 DV N  2 - S- N
P03235 DV N  1 - S- N
P03236 DV N  9 2 -V N
P03237 DV N  6 2 SV N
P03238 DV N  9 2 S- N
P03239 DV N  9 2 SV N
P03240 DV N  1 - SV N
P03256 DV N  4 - S- N
P03257 DV N  4 - S- N
P03258 DV N  4 - S- N
P03259 DV N  6 2 SV N
P03260 DV N  7 2 S- J
P03261 DV N  7 2 S- J
P03262 DV N  7 2 S- J
P03263 DV N  7 2 S- J
P03264 DV N  7 2 S- J
P03265 DV N  7 2 S- J
P03266 DV N  1 - S- N
P03267 DV N  6 2 SV N
P03268 DV N  9 2 S- N
P03269 DV N  9 2 SV N
P03270 DV N  1 - SV N
P03271 DV N  2 - S- N
P03272 DV N  1 - S- N
P03273 DV N  6 2 SV N
P03274 DV N  9 2 S- N
P03275 DV N  9 2 SV N
P03276 DV N  1 - SV N
P03277 DV N  2 - S- N
P03278 DV N  6 2 SV N
P03279 DV N  6 2 SV N
P03280 DV N  5 2 SV J
P03281 DV N  5 2 SV J
P03282 DV N  5 2 SV J
P03283 DV N  5 2 SV J
P03284 DV N  5 2 SV J
P03285 DV N  5 2 SV J
P03286 DV N  5 2 SV J
P03287 DV N  5 2 SV J
P03288 DV N  5 2 SV J
P03289 DV N  5 2 SV J
P03290 DV N  5 2 SV J
P03291 DV N  5 2 SV J
P03292 DV N  5 2 SV J
P03293 DV N  5 2 SV J
P03294 DV N  5 2 SV J
P03295 DV N  5 2 SV J
P03296 DV N  5 2 SV J
P03297 DV N  5 2 SV N
P03298 DV N  1 - -V N
P03299 DV N  7 2 -V N
P03300 PS N  5 - S- J
P03301 PS N  5 - S- J
P03302 PS N  5 - S- J
P03303 PS N  5 - S- J
P03304 PS N  5 - S- J
P03305 PS N  5 - S- J
P03306 PS N  5 - S- J
P03307 PS N  5 - S- J
P03308 PS N  5 - S- J
P03309 PS D 11 - S- J
P03310 PS D 11 - S- J
P03311 PS D 11 - S- J
P03312 PS D 11 - S- J
P03313 PS D 11 - S- J
P03314 PS D 11 - S- J
P03315 PS D 11 - S- J
P03316 PS D 11 - S- J
P03317 PS D 11 - S- J
P03318 PS D 11 - S- N
P03319 PS D 11 - S- N
P03320 PS N  4 - S- J
P03321 PS N  4 - S- J
P03322 PS N  4 - S- J
P03323 PS N  4 - S- J
P03324 PS N  4 - S- J
P03325 PS N  4 - S- J
P03326 PS N  4 - S- J
P03327 PS N  4 - S- J
P03328 PS N  4 - S- J
P03329 PS N  4 - S- J
P03330 PS N  4 - S- J
P03331 PS N  4 - S- J
P03332 PS N  4 - S- J
P03333 PS N  4 - S- J
P03334 PS N  4 - S- J
P03335 PS N  4 - S- J
P03336 PS N  4 - S- J
P03337 PS N  4 - S- J
P03338 PS N  4 - S- J
P03339 PS N  4 - S- J
P03340 PS N  4 - S- J
P03341 PS N  4 - S- J
P03342 PS N  4 - S- J
P03343 PS N  4 - S- J
P03344 PS N  4 - S- J
P03345 PS N  4 - S- J
P03346 PS N  4 - S- J
P03348 PS N  4 - S- J
P03349 PS N  4 - S- J
P03350 PS N  4 - S- J
P03351 DV S  1 - S- N
P03352 DV S  1 - S- N
P03353 DV S  1 - S- N
P03354 DV S  1 - S- N
P03355 DV S  1 - S- N
P03356 DV S  1 - S- N
P03357 DV S  1 - S- N
P03358 DV S  1 - S- N
P03359 DV S  1 - S- N
P03360 DV S  1 - S- N
P03405 DV D 11 - S- N
P03406 DV D 11 - S- N
P03407 DV D 11 - S- N
P03410 PS S 50 - S- N
P03411 PS S 50 - S- N
P03412 PS S 50 - S- N
P03413 PS S 50 - S- N
P03414 PS S 50 - S- N
P03454 DV N  6 2 SV N
P03516 DV N  2 - S- N
P03573 DV D 11 - S- N
P03575 DV N  2 - S- N
P03576 DV N  1 - S- N
P03577 DV N  1 - S- N
P03578 DV S 34 - S- N
P03579 DV S 11 - S- N
P03580 DV S 50 - S- N
P03581 DV S 50 - S- N
P03582 DV S  2 - S- N
P03594 DV N 11 4 SV N
P03600 DV N  9 2 SV N
P03601 DV N  9 2 SV N
P03602 DV N  9 2 SV N
P03603 DV N  9 2 SV N
P03604 DV N  9 2 SV N
P03605 DV N  9 2 SV N
P03606 DV N  9 2 SV N
P03607 DV N  9 2 SV N
P03608 DV N  9 2 SV N
P03609 DV N  9 2 SV N
P03635 DV N  9 2 SV N
P03636 DV N  9 2 SV N
P03637 DV N  9 2 SV N
P03638 DV N  9 2 SV N
P03639 DV N  9 2 SV J
P03640 DV N  9 2 SV N
P03641 DV N  9 2 SV N
P03642 DV N  9 2 SV N
P03643 DV N  9 2 SV N
P03644 DV N  9 2 SV N
P03645 DV N  9 2 SV N
P03646 DV N  9 2 SV N
P03647 DV N  9 2 SV N
P03648 DV N  9 2 SV N
P03649 DV N  9 2 SV N
P03660 DV N  9 2 SV N
P03670 DV N  7 2 SV N
P03671 DV N  7 2 SV N
P03696 DV N  7 2 SV N
P03697 DV N  7 2 SV N
P03698 DV N  7 2 SV N
P03699 DV N  9 2 -V N
P03937 PS S 15 - S- N
P04109 DV N  5 2 S- J
P04131 DV N  9 2 SV J
P04154 DV N  2 0 SV N
P04156 DV N  9 2 SV N
P04157 DV D 11 - S- N
P04158 DV D 11 - S- N
P04159 DV N  1 0 S- N
P04160 DV N  1 0 S- N
P04164 DV N  2 0 S- N
P04165 DV N  2 0 S- N
P04166 DV N  6 2 S- N
P04167 DV N  6 2 S- N
P04168 DV N 10 2 -V N
P04169 DV N 10 2 -V N
P04170 DV N  1 0 S- N
P04171 DV N  1 0 S- N
P04172 DV N  1 0 S- N
P04173 DV N  1 0 S- N
P04174 DV N  1 0 S- N
P04175 DV N  1 0 S- N
P04176 DV N 10 2 SV N
P04177 DV N 10 2 SV N
P04181 DV N  1 - S- N
P04182 DV N  1 - S- N
P04183 DV N  1 - S- N
P04204 DV N  9 2 SV N
P04205 DV N  9 2 SV N
P04215 DV N  9 2 SV J
P04216 DV N  9 2 SV J
P04223 DV N  6 2 SV N
P04300 DV N  4 0 SV N
P04400 PS S 50 - S- N
P04410 PS S  3 - S- N
P04411 PS D 11 - S- N
P04412 PS S 25 - S- N
P04413 PS S 25 - S- N
P04414 PS D 11 - S- N
P04573 DV D 11 - S- N
P04575 DV N  2 - S- N
P04576 DV N  1 - S- N
P04577 DV N  1 - S- N
P04578 DV S 34 - S- N
P04579 DV S 11 - S- N
P04580 DV S 50 - S- N
P04581 DV S 50 - S- N
P04582 DV S  2 - S- N
P05000 DV S  1 - S- N
P05001 DV S  1 - S- N
P05002 DV S  1 - S- N
P05003 DV S  1 - S- N
P05004 DV S  1 - S- N
P05005 DV S  1 - S- N
P05006 DV S  1 - S- N
P05007 DV S  1 - S- N
P05008 DV S  1 - S- N
P05009 DV S  1 - S- N
P05010 DV S  1 - S- N
P05012 DV D 11 - S- N
P05013 DV S  1 - S- N
P05014 DV S  1 - S- N
P05015 DV S  1 - S- N
P05016 DV S  1 - S- N
P05017 DV D 11 - S- N
P05018 DV D 11 - S- N
P05019 PS S  1 - S- N
P05020 DV D 11 - S- N
P05021 DV S 40 - S- N
P05022 DV S  1 - S- N
P05023 DV S  1 - S- N
P05024 DV N  6 - S- N
P05025 DV N  6 - S- N
P05026 DV N  6 - S- N
P05027 DV S 10 - S- N
P05028 DV N  6 - S- N
P05029 DV N  6 - S- N
P05030 PS S  1 - S- N
P05040 DV S  1 - S- N
P05041 DV S  1 - S- N
P05042 DV S  1 - S- N
P05043 DV S  1 - S- N
P05044 DV S  6 - S- N
P05045 DV S  1 - S- N
P05046 DV S  1 - S- N
P05047 DV S  1 - S- N
P05048 DV S  6 - S- N
P05049 DV S  6 - S- N
P05050 DV S  6 - S- N
P05051 DV S  6 - S- N
P05052 DV S  1 - S- N
P05090 DV S 50 - S- N
P05091 DV D 11 - S- N
P05092 PS S 10 - S- N
P05101 DV S  1 - S- N
P05102 DV S  1 - S- N
P05103 DV S  1 - S- N
P05104 DV S  1 - S- N
P05105 DV S  2 - S- N
P05106 DV S  1 - S- N
P05107 DV N  3 0 S- N
P05108 PS S  1 - S- N
P05109 PS S  1 - S- N
P05110 DV S 25 - S- N
P05111 PS S  1 - S- N
P05112 PS D 11 - S- N
P05113 DV S  1 - S- N
P05114 DV S  1 - S- N
P05115 DV D 11 - S- N
P05116 DV S  4 - S- N
P05117 DV N  3 0 S- N
P05118 DV N  3 - S- N
P05119 DV S  1 - S- N
P05125 DV N  2 0 S- N
P05171 DV N  9 2 SV N
P05182 DV N  7 3 SV N
P05188 DV N  6 2 SV N
P05189 DV N  1 - SV N
P05200 DV S 50 - S- J
P05201 DV S 50 - S- J
P05202 DV S 50 - S- J
P05220 DV N  5 2 S- N
P05262 DV N  1 - S- N
P05300 DV S  1 - S- N
P05310 DV S  1 - S- N
P05320 DV S  1 - S- N
P05330 DV S  1 - S- N
P05340 DV S  1 - S- N
P05350 DV S  1 - S- N
P05360 DV S  1 - S- N
P05370 DV S  1 - S- N
P05380 DV S  1 - S- N
P05390 DV S  1 - S- N
P05401 DV N  5 2 S- N
P05402 DV S 15 - S- N
P05403 DV S 15 - S- N
P05404 DV S 15 - S- N
P05575 DV N  9 2 SV N
P05576 DV N  9 2 SV N
P05577 DV N  9 2 SV N
P05578 DV N  9 2 SV N
P05579 DV N  9 2 SV N
P05580 DV N  9 2 SV N
P05581 DV N  9 2 SV N
P05582 DV N  9 2 SV N
P05583 DV N  9 2 SV N
P05584 DV N  9 2 SV N
P05585 DV N  9 2 SV N
P05586 DV N  9 2 SV N
P05587 DV N  9 2 SV N
P05588 DV N  9 2 SV N
P05589 DV N  9 2 SV N
P05590 DV N  9 2 SV N
P05591 DV N  9 2 SV N
P05592 DV N  9 2 SV N
P05593 DV N  9 2 SV N
P05594 DV N  9 2 SV N
P05595 DV N  9 2 SV N
P05596 DV N  9 2 SV N
P05597 DV N  9 2 SV N
P05598 DV N  9 2 SV N
P05599 DV N  9 2 SV N
P05695 DV N  6 2 SV N
P06142 DV N 10 2 SV N
P06143 DV N 10 2 SV N
P06144 DV N 10 2 SV N
P06145 DV N 10 2 SV N
P06146 DV N 10 2 SV N
P06147 DV N 10 2 SV N
P06148 DV N 10 2 SV N
P06149 DV N 10 2 SV N
P06150 DV N 10 2 SV N
P06151 DV N 10 2 SV N
P06152 DV N 10 2 SV N
P06153 DV N 10 2 SV N
P06205 DV N  6 2 S- N
P06431 DV N 10 2 SV N
P06432 DV N 10 2 SV N
P06433 DV N 10 2 SV N
P06434 DV N 10 2 SV N
P06650 DV N 10 2 SV N
P06651 DV N 10 2 SV N
P06652 DV N 10 2 SV N
P06653 DV N 10 2 SV N
P06654 DV N 10 2 SV N
P06655 DV N 10 2 SV N
P06656 DV N  6 2 SV N
P06657 DV N  6 2 SV N
P06658 DV N  6 2 SV N
P06659 DV N  6 2 SV N
P06660 DV N  6 2 SV N
P06661 DV N  6 2 SV N
P06662 DV N  6 2 SV N
P06663 DV N  6 2 SV N
P06664 DV N  6 2 SV N
P06665 DV N  6 2 SV N
P06666 DV N  6 2 SV N
P06667 DV N  6 2 SV N
P06668 DV N 10 2 SV N
P06669 DV N 10 2 SV N
P06700 DV N  9 2 SV N
P06701 DV N  9 2 SV N
P06702 DV N  9 2 SV N
P06703 DV N  9 2 SV N
P06704 DV N  9 2 SV N
P06705 DV N  9 2 SV N
P06706 DV N  9 2 SV N
P06707 DV N  9 2 SV N
P06708 DV N  9 2 SV N
P06709 DV N  9 2 SV N
P06710 DV N  9 2 SV N
P06711 DV N  9 2 SV N
P06712 DV N  9 2 SV N
P06713 DV N  9 2 SV N
P06714 DV N  9 2 SV N
P06715 DV N  9 2 SV N
P06716 DV N  9 2 SV N
P06717 DV N  9 2 SV N
P06718 DV N  9 2 SV N
P06719 DV N  9 2 SV N
P06720 DV N  9 2 SV N
P06721 DV N  9 2 SV N
P06722 DV N  9 2 SV N
P06723 DV N  9 2 SV N
P06724 DV N  9 2 SV N
P06725 DV N  9 2 SV N
P06726 DV N  9 2 SV N
P06727 DV N  9 2 SV N
P06728 DV N  9 2 SV N
P06729 DV N  9 2 SV N
P06730 DV N  9 2 SV N
P06731 DV N  9 2 SV N
P06732 DV N  9 2 SV N
P06733 DV N  9 2 SV N
P06734 DV N  9 2 SV N
P06735 DV N  9 2 SV N
P06736 DV N  9 2 SV N
P06737 DV N  9 2 SV N
P06738 DV N  9 2 SV N
P06739 DV N  9 2 SV N
P06740 DV N  9 2 SV N
P06741 DV N  9 2 SV N
P06742 DV N  9 2 SV N
P06743 DV N  9 2 SV N
P06744 DV N  9 2 SV N
P06745 DV N  9 2 SV N
P06746 DV N  9 2 SV N
P06747 DV N  9 2 SV N
P06748 DV N  9 2 SV N
P06749 DV N  9 2 SV N
P06750 DV N  9 2 SV N
P06751 DV N  9 2 SV N
P06752 DV N  9 2 SV N
P06753 DV N  9 2 SV N
P06754 DV N  9 2 SV N
P06755 DV N  9 2 SV N
P06756 DV N  9 2 SV N
P06757 DV N  9 2 SV N
P06758 DV N  9 2 SV N
P06759 DV N  9 2 SV N
P06760 DV N  9 2 SV N
P06761 DV N  9 2 SV N
P06762 DV N  9 2 SV N
P06763 DV N  9 2 SV N
P06764 DV N  9 2 SV N
P06765 DV N  9 2 SV N
P06766 DV N  9 2 SV N
P06767 DV N  9 2 SV N
P06768 DV N  9 2 SV N
P06769 DV N  9 2 SV N
P06770 DV N  9 2 SV N
P06771 DV N  9 2 SV N
P06772 DV N  9 2 SV N
P06773 DV N  9 2 SV N
P06774 DV N  9 2 SV N
P06775 DV N  9 2 SV N
P06776 DV N  9 2 SV N
P06777 DV N  9 2 SV N
P06778 DV N  9 2 SV N
P06779 DV N  9 2 SV N
P06780 DV N  9 2 SV N
P06781 DV N  9 2 SV N
P06782 DV N  9 2 SV N
P06783 DV N  9 2 SV N
P06784 DV N  9 2 SV N
P06785 DV N  9 2 SV N
P06786 DV N  9 2 SV N
P06787 DV N  9 2 SV N
P06788 DV N  9 2 SV N
P06789 DV N  9 2 SV N
P06790 DV N  9 2 SV N
P06791 DV N  9 2 SV N
P06792 DV N  9 2 SV N
P06793 DV N  9 2 SV N
P06794 DV N  9 2 SV N
P06795 DV N  9 2 SV N
P06796 DV N  9 2 SV N
P06797 DV N  9 2 SV N
P06798 DV N  9 2 SV N
P06799 DV N  9 2 SV N
P06800 DV N  9 2 SV N
P06801 DV N  9 2 SV N
P06802 DV N  9 2 SV N
P06803 DV N  9 2 SV N
P06804 DV N  9 2 SV N
P06805 DV N  9 2 SV N
P06806 DV N  9 2 SV N
P06807 DV N  9 2 SV N
P06808 DV N  9 2 SV N
P06809 DV N  9 2 SV N
P06810 DV N  9 2 SV N
P06811 DV N  9 2 SV N
P06812 DV N  9 2 SV N
P06813 DV N  9 2 SV N
P06814 DV N  9 2 SV N
P06815 DV N  9 2 SV N
P06816 DV N  9 2 SV N
P06817 DV N  9 2 SV N
P06818 DV N  9 2 SV N
P06819 DV N  9 2 SV N
P06820 DV N  9 2 SV N
P06821 DV N  9 2 SV N
P06822 DV N  9 2 SV N
P06823 DV N  9 2 SV N
P06824 DV N  9 2 SV N
P06825 DV N  9 2 SV N
P06826 DV N  9 2 SV N
P06827 DV N  9 2 SV N
P06828 DV N  9 2 SV N
P06829 DV N  9 2 SV N
P06830 DV N  9 2 SV N
P06831 DV N  9 2 SV N
P06832 DV N  9 2 SV N
P06833 DV N  9 2 SV N
P06834 DV N  9 2 SV N
P06835 DV N  9 2 SV N
P06836 DV N  9 2 SV N
P06837 DV N  9 2 SV N
P06838 DV N  9 2 SV N
P06839 DV N  9 2 SV N
P06840 DV N  9 2 SV N
P06841 DV N  9 2 SV N
P06842 DV N  9 2 SV N
P06843 DV N  9 2 SV N
P06844 DV N  9 2 SV N
P06845 DV N  9 2 SV N
P06846 DV N  9 2 SV N
P06847 DV N  9 2 SV N
P06848 DV N  9 2 SV N
P06849 DV N  9 2 SV N
P06850 DV N  9 2 SV N
P06851 DV N  9 2 SV N
P06852 DV N  9 2 SV N
P06853 DV N  9 2 SV N
P06854 DV N  9 2 SV N
P06855 DV N  9 2 SV N
P06856 DV N  9 2 SV N
P06857 DV N  9 2 SV N
P06858 DV N  9 2 SV N
P06859 DV N  9 2 SV N
P06860 DV N  9 2 SV N
P06861 DV N  9 2 SV N
P06862 DV N  9 2 SV N
P06863 DV N  9 2 SV N
P06864 DV N  9 2 SV N
P06865 DV N  9 2 SV N
P06866 DV N  9 2 SV N
P06867 DV N  9 2 SV N
P06868 DV N  9 2 SV N
P06869 DV N  9 2 SV N
P06870 DV N  9 2 SV N
P06871 DV N  9 2 SV N
P06872 DV N  9 2 SV N
P06873 DV N  9 2 SV N
P06874 DV N  9 2 SV N
P06875 DV N  9 2 SV N
P06876 DV N  9 2 SV N
P06877 DV N  9 2 SV N
P06878 DV N  9 2 SV N
P06879 DV N  9 2 SV N
P06880 DV N  9 2 SV N
P06881 DV N  9 2 SV N
P06882 DV N  9 2 SV N
P06883 DV N  9 2 SV N
P06884 DV N  9 2 SV N
P06885 DV N  9 2 SV N
P06886 DV N  9 2 SV N
P06887 DV N  9 2 SV N
P06888 DV N  9 2 SV N
P06889 DV N  9 2 SV N
P06890 DV N  9 2 SV N
P06891 DV N  9 2 SV N
P06892 DV N  9 2 SV N
P06893 DV N  9 2 SV N
P06894 DV N  9 2 SV N
P06895 DV N  9 2 SV N
P06896 DV N  9 2 SV N
P06897 DV N  9 2 SV N
P06898 DV N  9 2 SV N
P06899 DV N  9 2 SV N
P06900 DV N  9 2 SV N
P06901 DV N  9 2 SV N
P06902 DV N  9 2 SV N
P06903 DV N  9 2 SV N
P06904 DV N  9 2 SV N
P06905 DV N  9 2 SV N
P06906 DV N  9 2 SV N
P06907 DV N  9 2 SV N
P06908 DV N  9 2 SV N
P06909 DV N  9 2 SV N
P06910 DV N  9 2 SV N
P06911 DV N  9 2 SV N
P06912 DV N  9 2 SV N
P06913 DV N  9 2 SV N
P06914 DV N  9 2 SV N
P06915 DV N  9 2 SV N
P06916 DV N  9 2 SV N
P06917 DV N  9 2 SV N
P06918 DV N  9 2 SV N
P06919 DV N  9 2 SV N
P06920 DV N  9 2 SV N
P06921 DV N  9 2 SV N
P06922 DV N  9 2 SV N
P06923 DV N  9 2 SV N
P06924 DV N  9 2 SV N
P06925 DV N  9 2 SV N
P06926 DV N  9 2 SV N
P06927 DV N  9 2 SV N
P06928 DV N  9 2 SV N
P06929 DV N  9 2 SV N
P06930 DV N  9 2 SV N
P06931 DV N  9 2 SV N
P06932 DV N  9 2 SV N
P06933 DV N  9 2 SV N
P06934 DV N  9 2 SV N
P06935 DV N  9 2 SV N
P06936 DV N  9 2 SV N
P06937 DV N  9 2 SV N
P06938 DV N  9 2 SV N
P06939 DV N  9 2 SV N
P06940 DV N  9 2 SV N
P06941 DV N  9 2 SV N
P06942 DV N  9 2 SV N
P06943 DV N  9 2 SV N
P06944 DV N  9 2 SV N
P06945 DV N  9 2 SV N
P06946 DV N  9 2 SV N
P06947 DV N  9 2 SV N
P06948 DV N  9 2 SV N
P06949 DV N  9 2 SV N
P06950 DV N  9 2 SV N
P06951 DV N  9 2 SV N
P06952 DV N  9 2 SV N
P06953 DV N  9 2 SV N
P06954 DV N  9 2 SV N
P06955 DV N  9 2 SV N
P06956 DV N  9 2 SV N
P06957 DV N  9 2 SV N
P06958 DV N  9 2 SV N
P06961 DV N 10 2 SV N
P06962 DV N 10 2 SV N
P06969 DV N 10 2 SV N
P06970 DV N  9 2 SV N
P06971 DV N  9 2 SV N
P06972 DV N  9 2 SV N
P06973 DV N  9 2 SV N
P06974 DV N  7 3 SV N
P06975 DV N  7 3 SV N
P06976 DV N  9 2 SV N
P06977 DV N  9 2 SV N
P06978 DV N  9 2 SV N
P06979 DV N 10 2 SV N
P06980 DV N 10 2 SV N
P06989 DV N 10 2 SV N
P06990 DV N 10 2 SV N
P06991 DV N 10 2 SV N
P06992 DV N  6 2 SV N
P06993 DV N  6 2 SV N
P06994 DV N  6 2 SV N
P06995 DV N  6 2 SV N
P06996 DV N  6 2 SV N
P06997 DV N  6 2 SV N
P06998 DV N 10 2 SV N
P07107 DV N 10 2 SV N
P07108 DV N 10 2 -V N
P07160 DV N 10 2 -V J
P07172 DV N 10 2 SV N
P07182 DV N 10 2 SV J
P07200 DV N  6 2 SV N
P07201 DV N  6 2 S- N
P07202 DV N  6 2 S- N
P07203 DV N  6 2 S- N
P07204 DV N  6 2 S- N
P07205 DV N  6 2 S- J
P07206 DV N  6 2 S- N
P07207 DV N  6 2 S- N
P07208 DV N  6 2 S- N
P07209 DV N  6 2 S- N
P07210 DV N  6 2 S- N
P07211 DV N  6 2 -V N
P07212 DV N  6 2 -V N
P07213 DV N  6 2 -V N
P07214 DV N  6 2 -V N
P07215 DV N  6 2 -V J
P07216 DV N  6 2 -V N
P07217 DV N  6 2 -V N
P07218 DV N  6 2 -V N
P07219 DV N  6 2 -V N
P07220 DV N  6 2 -V N
P07221 DV N  6 2 S- N
P07222 DV N  6 2 -V N
P07431 DV N 10 2 SV J
P07432 DV N 10 2 SV J
P07433 DV N 10 2 SV J
P07434 DV N 10 2 SV J
P07435 DV N 12 2 SV N
P07441 DV N 10 2 SV N
P07442 DV N 10 2 SV N
P07443 DV N 10 2 SV N
P07444 DV N 10 2 SV N
P07445 DV N 10 2 SV N
P07446 DV N  6 2 SV N
P07447 DV N  6 2 SV N
P07448 DV N  6 2 SV N
P07449 DV N  6 2 SV N
P07450 DV N  6 2 SV N
P07451 DV N  6 2 SV N
P07452 DV N  6 2 SV N
P07453 DV N  6 2 SV N
P07454 DV N  6 2 SV N
P07455 DV N  6 2 SV N
P07456 DV N  6 2 SV N
P07457 DV N  6 2 SV N
P07460 DV N 10 2 SV N
P07461 DV N 10 2 SV N
P07462 DV N 10 2 SV N
P07463 DV N 10 2 SV N
P07464 DV N 10 2 SV N
P07465 DV N 10 2 SV N
P07466 DV N 10 2 SV N
P07467 DV N 10 2 SV N
P07468 DV N 10 2 SV N
P07469 DV N 10 2 SV N
P07470 DV N 10 2 SV N
P07471 DV N 10 2 SV N
P07472 DV N 10 2 SV N
P07473 DV N 10 2 SV N
P07474 DV N 10 2 SV N
P07475 DV N 10 2 SV N
P07476 DV N 10 2 SV N
P07477 DV N 10 2 SV N
P07478 DV N 10 2 SV N
P07479 DV N 10 2 SV N
P07480 DV N 10 2 SV N
P07481 DV N 10 2 SV N
P07482 DV N 10 2 SV N
P07483 DV N 10 2 SV N
P07484 DV N 10 2 SV N
P07485 DV N 10 2 SV N
P07486 DV N 10 2 SV N
P07487 DV N 10 2 SV N
P07488 DV N 10 2 SV N
P07489 DV N 10 2 SV N
P07490 DV N 10 2 SV N
P07491 DV N 10 2 SV N
P07492 DV N 10 2 SV N
P07493 DV N 10 2 SV N
P07494 DV N 10 2 SV N
P07495 DV N 10 2 SV N
P07496 DV N 10 2 SV N
P07609 DV N  4 0 SV N
P07643 DV N 10 2 -V N
P07670 DV N 10 2 SV N
P07692 DV N  1 0 S- N
P07693 DV N  1 0 S- N
P07743 DV N  1 0 S- N
P07744 DV N 10 2 SV N
P07745 DV N  6 2 SV N
P07746 DV N  6 2 SV N
P07777 DV N 10 2 SV J
P07778 DV N 10 2 SV J
P07779 DV N 10 2 SV N
P07784 DV N  5 2 -V N
P07903 DV N 10 2 SV N
P07920 DV N 10 2 SV N
P08003 DV N  1 0 SV N
P08004 DV N  2 0 SV N
P08005 DV N  1 0 SV N
P08016 DV N 10 2 SV N
P08017 DV N 10 2 SV N
P08018 DV N 10 2 SV N
P08019 DV N 10 2 SV N
P08020 DV N 10 2 SV N
P08021 DV N 10 2 SV N
P08022 DV N 10 2 SV N
P08023 DV N 10 2 SV N
P08024 DV N 10 2 SV N
P08025 DV N 10 2 SV N
P08026 DV N 10 2 SV N
P08027 DV N 10 2 SV N
P08028 DV N 10 2 SV N
P08029 DV N 10 2 SV N
P08030 DV N 10 2 SV N
P08031 DV N 10 2 SV N
P08032 DV N 10 2 SV N
P08033 DV N 10 2 SV N
P08034 DV N 10 2 SV N
P08035 DV N 10 2 SV N
P08043 DV N  9 2 SV N
P08044 DV N  9 2 SV N
P08045 DV N  9 2 SV N
P08046 DV N 10 2 SV J
P08047 DV N 10 2 SV J
P08048 DV N 10 2 SV J
P08049 DV N 10 2 SV N
P08050 DV N 10 2 SV N
P08051 DV N 10 2 SV N
P08052 DV N 10 2 SV N
P08053 DV N 10 2 SV N
P08054 DV N 10 2 SV N
P08055 DV N 10 2 SV N
P08056 DV N 10 2 SV J
P08057 DV N 10 2 SV J
P08058 DV N 10 2 SV J
P08059 DV N 10 2 SV J
P08060 DV N 10 2 SV N
P08061 DV N 10 2 SV N
P08062 DV N 10 2 SV N
P08063 DV N 10 2 SV N
P08064 DV N 10 2 SV N
P08065 DV N 10 2 SV N
P08066 DV N 10 2 SV N
P08067 DV N 10 2 SV N
P08068 DV N 10 2 SV N
P08069 DV N 10 2 SV N
P08070 DV N 10 2 SV N
P08071 DV N 10 2 SV N
P08072 DV N 10 2 SV N
P08074 DV N 10 2 SV J
P08075 DV N 10 2 SV J
P08076 DV N 10 2 SV J
P08077 DV N 10 2 SV J
P08078 DV N 10 2 SV J
P08080 DV N 10 2 SV N
P08081 DV N 10 2 SV N
P08082 DV N 10 2 SV N
P08083 DV N 10 2 SV N
P08084 DV N 10 2 SV N
P08101 DV N  1 - S- N
P08102 DV N  1 0 S- J
P08103 DV N  1 0 S- N
P08104 DV N  6 2 S- N
P08105 DV N  7 3 S- N
P08106 DV N  1 0 S- N
P08107 DV N  1 0 S- N
P08108 DV N  2 0 S- N
P08109 DV N  9 2 SV N
P08110 DV N  9 2 SV N
P08111 DV N  6 2 SV N
P08155 DV N 10 2 SV J
P08191 DV N  9 2 SV N
P08208 DV N  1 0 SV N
P08211 DV N 10 2 -V J
P08212 DV N 10 2 -V J
P08213 DV N 10 2 -V J
P08214 DV N 10 2 -V J
P08243 DV N 10 2 -V J
P08244 DV N 10 2 -V J
P08245 DV N 10 2 -V J
P08246 DV N 10 2 -V J
P08247 DV N 10 2 -V J
P08248 DV N 10 2 -V J
P08249 DV N 10 2 -V J
P08250 DV N 10 2 -V J
P08252 PS S  2 - S- N
P08253 DV N 10 2 SV N
P08254 DV N 10 2 SV N
P08255 DV N  1 0 S- N
P08256 DV N 10 2 SV N
P08257 DV N  2 0 SV J
P08258 DV N  1 - SV N
P08259 DV N  1 - SV N
P08260 DV N  2 0 -V N
P08265 DV N 10 2 -V J
P08293 DV N 10 2 SV N
P08307 DV S  3 - S- N
P08309 DV S  4 - S- N
P08310 DV N  2 0 S- N
P08321 DV N  2 0 SV J
P08486 DV N 10 2 -V N
P08500 DV S 50 - S- N
P08501 DV N  9 2 S- N
P08502 DV S 50 - S- N
P08503 DV S 50 - S- N
P08867 DV N  1 - SV N
P09000 DV N  1 - S- J
P09001 PS S 50 - S- N
P09002 PS S 50 - S- N
P09003 PS S 50 - S- N
P09004 DV N  6 2 SV N
P09005 DV N  2 0 SV N
P09006 DV N  2 0 S- N
P09007 DV N  2 0 S- N
P09008 DV N  2 0 SV N
P09009 DV N  2 0 SV N
P09010 DV N 10 2 SV N
P09011 DV N 10 2 SV N
P09012 DV N 10 2 SV N
P09013 DV N 10 2 SV N
P09014 DV N 10 2 SV N
P09015 DV N 10 2 SV N
P09016 DV N 10 2 SV N
P09017 DV N 10 2 SV N
P09018 DV N 10 2 SV N
P09019 DV N 10 2 SV N
P09020 DV N 10 2 SV N
P09021 DV N 10 2 SV N
P09022 DV N 10 2 SV N
P09023 DV N 10 2 SV N
P09024 DV N 10 2 SV N
P09025 DV N 10 2 SV N
P09026 DV N 10 2 SV N
P09027 DV N 10 2 SV N
P09028 DV N 10 2 SV N
P09029 DV N 10 2 SV N
P09036 PS H 50 - S- N
P09037 PS S 50 - S- N
P09038 PS S 50 - S- N
P09039 PS S  1 - S- N
P09040 PS S  1 - S- N
P09041 PS S  1 - S- N
P09100 DV S 50 - SV J
P09101 DV N 10 2 S- N
P09102 DV N 10 2 S- N
P09103 DV N 10 2 S- N
P09104 DV S 34 - S- N
P09105 DV S 34 - S- N
P09106 DV S 34 - S- N
P09107 DV S 34 - S- N
P09108 DV S 34 - S- N
P09109 DV S 34 - S- N
P09110 DV S 34 - S- N
P09111 DV S 34 - S- N
P09112 DV S 34 - S- N
P09113 DV S 34 - S- N
P09114 DV S 34 - S- N
P09115 DV S 34 - S- N
P09116 DV S  2 - S- N
P09117 DV S  2 - S- N
P09118 DV S  2 - S- N
P09119 DV S 11 - S- N
P09120 DV S 11 - S- N
P09121 DV S 11 - S- N
P09122 DV S  1 - S- N
P09123 DV S  1 - S- N
P09124 DV S  1 - S- N
P09125 DV N  9 2 SV N
P09126 DV N  9 2 SV J
P09127 DV N  9 2 SV J
P09128 DV N  9 2 SV J
P09129 DV N  9 2 SV J
P09130 DV N  9 2 SV J
P09131 DV N  9 2 SV N
P09132 DV N  1 0 SV N
P09133 DV N  1 0 SV N
P09134 DV N  6 2 SV N
P09135 DV N  1 0 SV N
P09136 DV N  1 0 SV N
P09137 DV N  1 0 SV N
P09138 DV S  2 - SV N
P09139 DV S  2 - SV N
P09140 DV S  2 - SV N
P09141 DV S  2 - SV N
P09161 DV S 32 - SV N
P09162 DV S 32 - SV N
P09163 DV S 32 - SV N
P09164 DV N  7 2 S- J
P09165 DV S 11 - S- N
P10611 DV N  6 2 SV N
P10612 DV N  6 2 SV N
P10613 DV N  6 2 SV N
P10614 DV N  6 2 SV N
P10615 DV N  6 2 SV N
P10616 DV N  6 2 SV N
P10617 DV N  6 2 SV N
P10618 DV N  6 2 SV N
P10619 DV N  6 2 SV N
P10620 DV N  6 2 SV N
P10621 DV N  6 2 SV N
P10622 DV N  6 2 SV N
P10623 DV N  6 2 SV N
P10624 DV N  6 2 SV N
P10625 DV N  6 2 SV N
P10639 DV N  9 2 SV N
P10982 DV N  6 2 SV N
P10988 DV N  9 2 SV N
P11185 DV N  9 2 SV N
P11302 DV N 10 2 SV N
P11307 DV N  7 3 SV N
P11393 DV N  6 2 S- N
P11396 DV N  6 2 S- N
P11399 DV N  6 2 S- N
P11541 DV N  1 0 SV N
P11542 DV N  1 0 SV N
P11543 DV N  1 0 SV N
P11544 DV N  1 0 SV N
P11545 DV N  1 0 SV N
P11688 DV N  9 2 SV N
P11800 DV N  9 2 S- N
P11905 DV N  2 - SV N
P11906 DV N  9 2 SV N
P11907 DV N  2 - SV N
P11908 DV N  9 2 SV N
P11909 DV N  2 - SV N
P11910 DV N  9 2 SV N
P11911 DV N  2 - SV N
P11912 DV N  9 2 SV N
P11913 DV N  2 - SV N
P11914 DV N  9 2 SV N
P11915 DV N 10 2 SV N
P11916 DV N 10 2 SV N
P11917 DV N 10 2 SV N
P11918 DV N 10 2 SV N
P11919 DV N 10 2 SV N
P11921 DV N  1 - SV -
P11922 DV N  1 - SV -
P11923 DV N  1 - SV -
P11925 DV N  6 2 SV N
P11926 DV N  6 2 SV N
P11927 DV N  6 2 SV N
P11928 DV N  6 2 SV N
P11929 DV N  6 2 SV N
P11965 DV N  1 0 SV N
P11966 DV N  1 0 SV N
P11967 DV N  1 0 SV N
P11968 DV N  1 0 SV N
P11969 DV N  1 0 SV N
P11980 DV N  7 3 S- N
P11981 DV N  7 3 S- N
P12006 DV N  7 2 SV J
P12007 DV N  7 2 SV J
P12008 DV N  7 2 SV J
P12009 DV N  7 2 SV J
P12010 DV N  7 2 SV J
P12011 DV N  7 2 SV J
P12012 DV N  7 2 SV J
P12042 DV N  2 - S- J
P12046 DV N  6 2 SV N
P12047 DV N  6 2 SV N
P12048 DV N  6 2 SV N
P12100 PS N  1 - S- N
P12102 DV N  6 2 S- N
P12146 DV N  6 2 -V N
P12147 DV N  1 - SV N
P12148 DV N  9 2 SV N
P12149 DV N  6 2 SV N
P12170 DV N  1 0 S- N
P12218 DV N 10 2 SV J
P12260 DV N  1 - S- N
P12266 DV S  1 - S- N
P12294 DV N  1 0 S- N
P12323 PS N  1 - S- N
P12338 DV N  6 2 S- N
P12409 DV N  6 2 SV N
P12474 DV N  1 0 S- N
P12531 DV N  9 2 SV N
P12535 DV N  6 2 SV N
P12536 DV N  1 - SV N
P12537 DV N  6 2 SV N
P12538 DV N  7 2 -V N
P12539 DV N  1 0 S- N
P12567 DV N  9 2 SV N
P12568 DV N  9 2 SV N
P12569 DV N  9 2 SV N
P12694 DV N  6 2 SV N
P12695 DV N  6 2 SV N
P12837 DV N  2 0 S- N
P12875 DV N  1 0 S- N
P12889 DV S  1 - S- N
P12895 DV N  2 0 S- N
P12899 DV N  2 0 S- N
P12932 DV N  2 0 S- N
P12940 DV S  1 - S- N
P12957 DV N  2 0 S- N
P13049 DV N  2 0 S- N
P13080 DV N  6 2 SV N
P13081 DV N  6 2 SV N
P13082 DV N  6 2 SV N
P13083 DV N  6 2 SV N
P13084 DV N  6 2 SV N
P13085 DV N  6 2 SV N
P13086 DV N  6 2 SV N
P13087 DV N  6 2 SV N
P13088 DV N  6 2 SV N
P13089 DV N  6 2 SV N
P13090 DV N  6 2 SV N
P13091 DV N  6 2 SV N
P13092 DV N  6 2 SV N
P13093 DV N  6 2 SV N
P13125 DV N  9 2 SV N
P13155 DV N  6 2 SV N
P13178 DV N 10 2 -V N
P13179 DV N  9 2 -V N
P13180 DV N  9 2 SV N
P13183 DV N  7 2 -V N
P13184 DV N  1 - S- N
P13185 DV N  1 - S- N
P13186 DV N  1 - S- N
P13187 DV N  1 - S- N
P13190 DV N  1 0 SV N
P13191 DV N  1 0 SV N
P13192 DV N  1 0 SV N
P13193 DV N  1 0 SV N
P13194 DV N  1 0 SV N
P13195 DV N  1 0 SV J
P13196 DV N  1 0 SV N
P13197 DV N  1 0 SV N
P13198 DV N  1 0 SV N
P13199 DV N  1 0 SV N
P13205 DV N  6 2 S- N
P13215 DV N  6 2 -V N
P13224 DV N  1 - S- N
P13225 DV N  1 - S- N
P13226 DV N  1 - S- N
P13227 DV N  1 - S- N
P13228 DV N  1 - S- -
P13229 DV N  1 - S- N
P13230 DV N  1 - S- N
P13231 DV N  1 - S- N
P13232 DV N  1 - S- N
P13233 DV N  1 - S- N
P13280 DV D 11 - S- N
P13281 DV D 11 - S- N
P13282 DV D 11 - S- N
P13283 DV D 11 - S- N
P13284 DV D 11 - S- N
P13285 DV D 11 - S- N
P13286 DV D 11 - S- N
P13287 DV D 11 - S- N
P13288 DV D 11 - S- N
P13289 DV D 11 - S- N
P13290 DV D 11 - S- N
P13291 DV D 11 - S- N
P13292 DV D 11 - S- N
P13293 DV D 11 - S- N
P13294 DV D 11 - S- N
P13301 DV N 10 2 SV N
P13302 DV N 10 2 SV N
P13303 DV N 10 2 SV N
P13304 DV N 10 2 SV N
P13305 DV N 10 2 SV N
P15013 PS S  1 - S- N
P16865 DV N  6 2 SV N
P16866 DV N  6 2 SV N
P16867 DV N  6 2 SV N
P16868 DV N  6 2 SV N
P16869 DV N  6 2 SV N
P16870 DV N  6 2 SV N
P16941 DV N  6 2 SV N
P17865 DV N  6 2 SV N
P17866 DV N  6 2 SV N
P18060 DV N 10 2 SV J
P18061 DV N 10 2 SV J
P18062 DV N 10 2 SV J
P21000 DV N 10 2 SV N
P23049 DV S  1 - SV -
P23155 DV N  2 0 S- N
P23195 DV N  1 - SV N
P26865 DV N  6 2 SV N
P26866 DV N  6 2 SV N
P26867 DV N  6 2 SV N
P26868 DV N  6 2 SV N
P36865 DV N  6 2 SV N
P36866 DV N  6 2 SV N
P36867 DV N  6 2 SV N
P36868 DV N  6 2 SV N
P36869 DV N  6 2 SV N
P36870 DV N  6 2 SV N
P36871 DV N  6 2 SV N
P40001 PS S 25 - S- N
P40002 DV N 10 0 S- N
P40003 DV N 10 0 S- N
P40004 DV S 10 - S- N
P40005 DV S  8 - S- N
P40006 DV S 10 - S- N
P40007 DV S 10 - S- N
P40008 DV S 10 - S- N
P40009 DV S 10 - S- N
P40010 DV S 10 - S- N
P40011 DV S 50 - S- N
P40012 DV S 50 - S- N
P40013 DV S 50 - S- N
P40014 DV N  5 0 S- N
P40015 DV S  3 - S- N
P40016 DV S  3 - S- N
P40017 DV S 30 - S- N
P40018 DV S 30 - S- N
P40019 DV S 30 - S- N
P41057 DV N  6 2 SV N
P41058 DV N  6 2 SV N
P41059 DV N  6 2 SV N
P46062 DV N  6 2 SV N
P46865 DV N  6 2 SV N
P46866 DV N  6 2 SV N
P56865 DV N  6 2 SV N
P56866 DV N  6 2 SV N
P56867 DV N  6 2 SV N
P56868 DV N  6 2 SV N
P56869 DV N  6 2 SV N
P56870 DV N  6 2 SV N
P66865 DV N  6 2 SV N
P66866 DV N  6 2 SV N
P76865 DV N  6 2 SV N
P86865 DV N  6 2 SV N
P86866 DV N  6 2 SV N
P86867 DV N  6 2 SV N
P86868 DV N  6 2 SV N
PB0001 DV S 16 - S- N
PB0002 DV S 30 - S- N
PB0003 PS N 16 0 S- N
PB0004 DV N 16 0 S- N
PB0005 DV S  3 - S- N
PB0011 DV N  9 2 S- N
PB0012 DV N  9 2 S- N
PB0013 DV N  9 2 S- N
PB0014 DV N  9 2 S- N
PB0015 DV N  9 2 S- N
PB0016 DV N  9 2 S- N
PB0017 DV N  9 2 S- N
PB0018 DV N  9 2 S- N
PB0019 DV N  7 2 S- N
PB0020 DV N  9 2 S- N
PB0021 DV N  7 2 S- N
PB0022 DV N  9 2 S- N
PB0023 DV N  9 2 S- N
PB0024 DV N  9 2 S- N
PB0025 DV N  9 2 S- N
PB0026 DV N  9 2 S- N
PB0027 DV N  9 2 S- N
PB0028 DV N  9 2 S- N
PB0029 DV N  9 2 S- N
PB0030 DV N  9 2 S- N
PB0031 DV N  9 2 S- N
PB1100 DV N  9 2 S- N
PB1101 DV D 11 - S- N
PB1102 DV N  9 2 S- N
PB1103 DV S 10 - S- N
PB1104 DV S 20 - S- N
PB1105 DV S 20 - S- N
PB1106 DV D 11 - S- N
PB1107 DV N  9 2 S- N
PB1108 DV N  9 2 S- N
PB1121 DV N  9 2 S- N
PB1122 DV N  9 2 S- N
PB1123 DV N  9 2 S- N
PB1124 DV N  9 2 S- N
PB1125 DV N  9 2 S- N
PB1126 DV N  9 2 S- N
PB1127 DV N  9 2 S- N
PB1128 DV N  9 2 S- N
PB1129 DV N  9 2 S- N
PB1130 DV N  9 2 S- N
PB1132 DV N  9 2 S- N
PB1133 DV N  9 2 S- N
PB1134 DV N  9 2 S- N
PB1135 DV N  9 2 S- N
PB1136 DV N  9 2 S- N
PB1137 DV N  9 2 S- N
PB1141 DV N  7 2 S- N
PB1142 DV N  7 2 S- N
PB1143 DV N  7 2 S- N
PB1144 DV N  7 2 S- N
PB1145 DV N  7 2 S- N
PB1146 DV N  7 2 S- N
PB1200 DV N  9 2 S- N
PB1201 DV D 11 - S- N
PB1206 DV D 11 - S- N
PB1207 DV N  9 2 S- N
PB1241 DV N  7 2 S- N
PB1242 DV N  7 2 S- N
PB1243 DV N  7 2 S- N
PB1244 DV N  7 2 S- N
PB1245 DV N  7 2 S- N
PB1246 DV N  7 2 S- N
PB1300 DV N  7 2 S- N
PB1301 DV D 11 - S- N
PB1306 DV D 11 - S- N
PB1307 DV N  9 2 S- N
PB1308 DV N  9 2 S- N
PB1321 DV N  9 2 S- N
PB1322 DV N  9 2 S- N
PB1323 DV N  9 2 S- N
PB1324 DV N  9 2 S- N
PB1325 DV N  9 2 S- N
PB1326 DV N  9 2 S- N
PB1327 DV N  9 2 S- N
PB1328 DV N  9 2 S- N
PB1329 DV N  9 2 S- N
PB1330 DV N  9 2 S- N
PB1331 DV N  9 2 S- N
PB1332 DV N  9 2 S- N
PB1333 DV N  9 2 S- N
PB1334 DV N  9 2 S- N
PB1335 DV N  9 2 S- N
PB1336 DV N  9 2 S- N
PB1337 DV N  9 2 S- N
PB1400 DV N  9 2 S- N
PB1401 DV D 11 - S- N
PB1402 DV N  9 2 S- N
PB1403 DV N  9 2 S- N
PB1404 DV N  3 - S- N
PB1405 DV N  3 - S- N
PB1406 DV D 11 - S- N
PB1407 DV N  9 2 S- N
PB1408 DV N  9 2 S- N
PB1409 DV D 11 - S- N
PB1410 DV S 30 - S- N
PB1411 DV N  5 - S- N
PB1412 DV N  9 2 S- N
PB1421 DV N  9 2 S- N
PB1422 DV N  9 2 S- N
PB1423 DV N  9 2 S- N
PB1424 DV N  9 2 S- N
PB1425 DV N  9 2 S- N
PB1426 DV N  9 2 S- N
PB1427 DV N  9 2 S- N
PB1428 DV N  9 2 S- N
PB1429 DV N  9 2 S- N
PB1430 DV N  9 2 S- N
PB1431 DV N  9 2 S- N
PB1432 DV N  9 2 S- N
PB1433 DV N  9 2 S- N
PB1434 DV N  9 2 S- N
PB1435 DV N  9 2 S- N
PB1436 DV N  9 2 S- N
PB1437 DV N  9 2 S- N
PB1441 DV N  7 2 S- N
PB1442 DV N  7 2 S- N
PB1443 DV N  7 2 S- N
PB1444 DV N  7 2 S- N
PB1445 DV N  7 2 S- N
PB1446 DV N  7 2 S- N
PB1500 DV N  9 2 S- N
PB1501 DV D 11 - S- N
PB1502 DV S 20 - S- N
PB1503 DV S 10 - S- N
PB1504 DV S 30 - S- N
PB1505 DV S  5 - S- N
PB1506 DV D 11 - S- N
PB1507 DV N  9 2 S- N
PB1508 DV N  9 2 S- N
PB1521 DV N  9 2 S- N
PB1522 DV N  9 2 S- N
PB1523 DV N  9 2 S- N
PB1524 DV N  9 2 S- N
PB1525 DV N  9 2 S- N
PB1526 DV N  9 2 S- N
PB1527 DV N  9 2 S- N
PB1528 DV N  9 2 S- N
PB1529 DV N  9 2 S- N
PB1530 DV N  9 2 S- N
PB1532 DV N  9 2 S- N
PB1533 DV N  9 2 S- N
PB1534 DV N  9 2 S- N
PB1535 DV N  9 2 S- N
PB1536 DV N  9 2 S- N
PB1537 DV N  9 2 S- N
PB1541 DV N  7 2 S- N
PB1542 DV N  7 2 S- N
PB1543 DV N  7 2 S- N
PB1544 DV N  7 2 S- N
PB1545 DV N  7 2 S- N
PB1546 DV N  7 2 S- N
PB1600 DV N  9 2 S- N
PB1601 DV D 11 - S- N
PB1602 DV N  9 2 S- N
PB1603 DV N  9 2 S- N
PB1604 DV S 30 - S- N
PB1606 DV D 11 - S- N
PB1607 DV N  9 2 S- N
PB1608 DV N  9 2 S- N
PB1621 DV N  9 2 S- N
PB1622 DV N  9 2 S- N
PB1623 DV N  9 2 S- N
PB1624 DV N  9 2 S- N
PB1625 DV N  9 2 S- N
PB1626 DV N  9 2 S- N
PB1627 DV N  9 2 S- N
PB1628 DV N  9 2 S- N
PB1629 DV N  9 2 S- N
PB1630 DV N  9 2 S- N
PB1631 DV N  9 2 S- N
PB1632 DV N  9 2 S- N
PB1633 DV N  9 2 S- N
PB1634 DV N  9 2 S- N
PB1635 DV N  9 2 S- N
PB1636 DV N  9 2 S- N
PB1637 DV N  9 2 S- N
PB1641 DV N  7 2 S- N
PB1642 DV N  7 2 S- N
PB1643 DV N  7 2 S- N
PB1644 DV N  7 2 S- N
PB1645 DV N  7 2 S- N
PB1646 DV N  7 2 S- N
PB1700 DV N  9 2 S- N
PB1701 DV D 11 - S- N
PB1702 DV S 10 - S- N
PB1703 DV S 20 - S- N
PB1704 DV S 20 - S- N
PB1705 DV S  5 - S- N
PB1706 DV D 11 - S- N
PB1707 DV N  9 2 S- N
PB1708 DV N  9 2 S- N
PB1709 DV S 20 - S- N
PB1710 DV N  5 0 S- N
PB1721 DV N  9 2 S- N
PB1722 DV N  9 2 S- N
PB1723 DV N  9 2 S- N
PB1724 DV N  9 2 S- N
PB1725 DV N  9 2 S- N
PB1726 DV N  9 2 S- N
PB1727 DV N  9 2 S- N
PB1728 DV N  9 2 S- N
PB1731 DV N  9 2 S- N
PB1741 DV N  7 2 S- N
PB1742 DV N  7 2 S- N
PB1743 DV N  7 2 S- N
PB1744 DV N  7 2 S- N
PB1745 DV N  7 2 S- N
PB1746 DV N  7 2 S- N
PB1800 DV N  9 2 S- N
PB1801 DV D 11 - S- N
PB1802 DV N  9 2 S- N
PB1806 DV D 11 - S- N
PB1807 DV N  9 2 S- N
PB1808 DV N  9 2 S- N
PB1821 DV N  9 2 S- N
PB1822 DV N  9 2 S- N
PB1823 DV N  9 2 S- N
PB1824 DV N  9 2 S- N
PB1825 DV N  9 2 S- N
PB1826 DV N  9 2 S- N
PB1827 DV N  9 2 S- N
PB1828 DV N  9 2 S- N
PB1829 DV N  9 2 S- N
PB1830 DV N  9 2 S- N
PB1831 DV N  9 2 S- N
PB1832 DV N  9 2 S- N
PB1833 DV N  9 2 S- N
PB1834 DV N  9 2 S- N
PB1835 DV N  9 2 S- N
PB1836 DV N  9 2 S- N
PB1837 DV N  9 2 S- N
PB1841 DV N  7 2 S- N
PB1842 DV N  7 2 S- N
PB1843 DV N  7 2 S- N
PB1844 DV N  7 2 S- N
PB1845 DV N  7 2 S- N
PB1846 DV N  7 2 S- N
PB1900 DV N  9 2 S- N
PB1901 DV D 11 - S- N
PB1902 DV N  9 2 S- N
PB1903 DV N  5 2 S- N
PB1904 DV N  9 2 S- N
PB1905 DV N  9 2 S- N
PB1906 DV D 11 - S- N
PB1907 DV N  9 2 S- N
PB1908 DV N  9 2 S- N
PB1909 DV N  9 2 S- N
PB1910 DV N  9 2 S- N
PB1911 DV N  9 2 S- N
PB1912 DV N  9 2 S- N
PB1913 DV S  7 - S- N
PB1914 DV N  9 2 S- N
PB1921 DV N  9 2 S- N
PB1922 DV N  9 2 S- N
PB1923 DV N  9 2 S- N
PB1924 DV N  9 2 S- N
PB1925 DV N  9 2 S- N
PB1926 DV N  9 2 S- N
PB1927 DV N  9 2 S- N
PB1928 DV N  9 2 S- N
PB1929 DV N  9 2 S- N
PB1930 DV N  9 2 S- N
PB1931 DV N  9 2 S- N
PB1932 DV N  9 2 S- N
PB1933 DV N  9 2 S- N
PB1934 DV N  9 2 S- N
PB1935 DV N  9 2 S- N
PB1936 DV N  9 2 S- N
PB1937 DV N  9 2 S- N
PB1941 DV N  7 2 S- N
PB1942 DV N  7 2 S- N
PB1943 DV N  7 2 S- N
PB1944 DV N  7 2 S- N
PB1945 DV N  7 2 S- N
PB1946 DV N  9 2 S- N
PB2000 DV N  9 2 S- N
PB2001 DV D 11 - S- N
PB2006 DV D 11 - S- N
PB2007 DV N  9 2 S- N
PB2008 DV N  9 2 S- N
PB2021 DV N  9 2 S- N
PB2022 DV N  9 2 S- N
PB2023 DV N  9 2 S- N
PB2024 DV N  9 2 S- N
PB2025 DV N  9 2 S- N
PB2026 DV N  9 2 S- N
PB2027 DV N  9 2 S- N
PB2028 DV N  9 2 S- N
PB2029 DV N  9 2 S- N
PB2030 DV N  9 2 S- N
PB2031 DV N  9 2 S- N
PB2041 DV N  7 2 S- N
PB2042 DV N  7 2 S- N
PB2043 DV N  7 2 S- N
PB2044 DV N  7 2 S- N
PB2045 DV N  7 2 S- N
PB2046 DV N  7 2 S- N
PB2100 DV N  9 2 S- N
PB2101 DV D 11 - S- N
PB2106 DV D 11 - S- N
PB2107 DV N  9 2 S- N
PB2108 DV N  9 2 S- N
PB2121 DV N  9 2 S- N
PB2122 DV N  9 2 S- N
PB2123 DV N  9 2 S- N
PB2124 DV N  9 2 S- N
PB2125 DV N  9 2 S- N
PB2126 DV N  9 2 S- N
PB2127 DV N  9 2 S- N
PB2128 DV N  9 2 S- N
PB2129 DV N  9 2 S- N
PB2130 DV N  9 2 S- N
PB2131 DV N  9 2 S- N
PB2132 DV N  9 2 S- N
PB2133 DV N  9 2 S- N
PB2134 DV N  9 2 S- N
PB2135 DV N  9 2 S- N
PB2136 DV N  9 2 S- N
PB2137 DV N  9 2 S- N
PB2141 DV N  7 2 S- N
PB2142 DV N  7 2 S- N
PB2143 DV N  7 2 S- N
PB2144 DV N  7 2 S- N
PB2145 DV N  7 2 S- N
PB2146 DV N  7 2 S- N
PB2200 DV N  7 2 S- N
PB2201 DV D 11 - S- N
PB2206 DV D 11 - S- N
PB2207 DV N  9 2 S- N
PB2208 DV N  9 2 S- N
PB2221 DV N  9 2 S- N
PB2222 DV N  9 2 S- N
PB2223 DV N  9 2 S- N
PB2224 DV N  9 2 S- N
PB2225 DV N  9 2 S- N
PB2226 DV N  9 2 S- N
PB2227 DV N  9 2 S- N
PB2228 DV N  9 2 S- N
PB2229 DV N  9 2 S- N
PB2230 DV N  9 2 S- N
PB2231 DV N  9 2 S- N
PB2232 DV N  9 2 S- N
PB2233 DV N  9 2 S- N
PB2234 DV N  9 2 S- N
PB2235 DV N  9 2 S- N
PB2236 DV N  9 2 S- N
PB2237 DV N  9 2 S- N
PB2241 DV N  7 2 S- N
PB2242 DV N  7 2 S- N
PB2243 DV N  7 2 S- N
PB2244 DV N  7 2 S- N
PB2245 DV N  7 2 S- N
PB2246 DV N  9 2 S- N
PB2300 DV N  9 2 S- N
PB2301 DV D 11 - S- N
PB2306 DV D 11 - S- N
PB2307 DV N  9 2 S- N
PB2308 DV N  9 2 S- N
PB2321 DV N  9 2 S- N
PB2322 DV N  9 2 S- N
PB2323 DV N  9 2 S- N
PB2324 DV N  9 2 S- N
PB2325 DV N  9 2 S- N
PB2326 DV N  9 2 S- N
PB2327 DV N  9 2 S- N
PB2328 DV N  9 2 S- N
PB2329 DV N  9 2 S- N
PB2330 DV N  9 2 S- N
PB2331 DV N  9 2 S- N
PB2332 DV N  9 2 S- N
PB2333 DV N  9 2 S- N
PB2334 DV N  9 2 S- N
PB2335 DV N  9 2 S- N
PB2336 DV N  9 2 S- N
PB2337 DV N  9 2 S- N
PB2341 DV N  7 2 S- N
PB2342 DV N  7 2 S- N
PB2343 DV N  7 2 S- N
PB2344 DV N  7 2 S- N
PB2345 DV N  7 2 S- N
PB2346 DV N  7 2 S- N
PB2400 DV N  9 2 S- N
PB2401 DV D 11 - S- N
PB2406 DV D 11 - S- N
PB2407 DV N  9 2 S- N
PB2408 DV N  9 2 S- N
PB2421 DV N  9 2 S- N
PB2422 DV N  9 2 S- N
PB2423 DV N  9 2 S- N
PB2424 DV N  9 2 S- N
PB2425 DV N  9 2 S- N
PB2426 DV N  9 2 S- N
PB2427 DV N  9 2 S- N
PB2428 DV N  9 2 S- N
PB2429 DV N  9 2 S- N
PB2430 DV N  9 2 S- N
PB2431 DV N  9 2 S- N
PB2432 DV N  9 2 S- N
PB2433 DV N  9 2 S- N
PB2434 DV N  9 2 S- N
PB2435 DV N  9 2 S- N
PB2436 DV N  9 2 S- N
PB2437 DV N  9 2 S- N
PB2441 DV N  7 2 S- N
PB2442 DV N  7 2 S- N
PB2443 DV N  7 2 S- N
PB2444 DV N  7 2 S- N
PB2445 DV N  7 2 S- N
PB2446 DV N  7 2 S- N
PB2500 DV N  9 2 S- N
PB2501 DV D 11 - S- N
PB2502 DV S 20 - S- N
PB2503 DV S 10 - S- N
PB2504 DV S 40 - S- N
PB2506 DV D 11 - S- N
PB2507 DV N  9 2 S- N
PB2508 DV N  9 2 S- N
PB2509 DV S 20 - S- N
PB2510 DV N  5 0 S- N
PB2521 DV N  9 2 S- N
PB2525 DV N  9 2 S- N
PB2526 DV N  9 2 S- N
PB2527 DV N  9 2 S- N
PB2531 DV N  9 2 S- N
PB2542 DV N  7 2 S- N
PB2543 DV N  7 2 S- N
PB2544 DV N  7 2 S- N
PB2600 DV N  9 2 S- N
PB2601 DV D 11 - S- N
PB2602 DV N  9 2 S- N
PB2603 DV S 10 - S- N
PB2604 DV S 20 - S- N
PB2605 DV S 20 - S- N
PB2606 DV D 11 - S- N
PB2607 DV N  9 2 S- N
PB2608 DV N  9 2 S- N
PB2609 DV S 20 - S- N
PB2610 DV N  5 0 S- N
PB2621 DV N  9 2 S- N
PB2625 DV N  9 2 S- N
PB2626 DV N  9 2 S- N
PB2627 DV N  9 2 S- N
PB2631 DV N  9 2 S- N
PB2642 DV N  7 2 S- N
PB2643 DV N  7 2 S- N
PB2644 DV N  7 2 S- N
PB2700 DV N  9 2 S- N
PB2701 DV D 11 - S- N
PB2702 DV S 10 - S- N
PB2703 DV S 20 - S- N
PB2704 DV S 20 - S- N
PB2706 DV D 11 - S- N
PB2707 DV N  9 2 S- N
PB2708 DV N  9 2 S- N
PB2722 DV N  9 2 S- N
PB2723 DV N  9 2 S- N
PB2724 DV N  9 2 S- N
PB2728 DV N  9 2 S- N
PB2729 DV N  9 2 S- N
PB2730 DV N  9 2 S- N
PB2732 DV N  9 2 S- N
PB2733 DV N  9 2 S- N
PB2734 DV N  9 2 S- N
PB2741 DV N  7 2 S- N
PB2745 DV N  7 2 S- N
PB2746 DV N  7 2 S- N
PB2800 DV N  9 2 S- N
PB2801 DV D 11 - S- N
PB2802 DV S 20 - S- N
PB2803 DV N  9 2 S- N
PB2804 DV S 30 - S- N
PB2806 DV D 11 - S- N
PB2807 DV N  9 2 S- N
PB2808 DV N  9 2 S- N
PB2821 DV N  9 2 S- N
PB2822 DV N  9 2 S- N
PB2823 DV N  9 2 S- N
PB2824 DV N  9 2 S- N
PB2825 DV N  9 2 S- N
PB2826 DV N  9 2 S- N
PB2827 DV N  9 2 S- N
PB2828 DV N  9 2 S- N
PB2829 DV N  9 2 S- N
PB2830 DV N  9 2 S- N
PB2831 DV N  9 2 S- N
PB2832 DV N  9 2 S- N
PB2833 DV N  9 2 S- N
PB2834 DV N  9 2 S- N
PB2835 DV N  9 2 S- N
PB2836 DV N  9 2 S- N
PB2837 DV N  9 2 S- N
PB2841 DV N  7 2 S- N
PB2842 DV N  7 2 S- N
PB2843 DV N  7 2 S- N
PB2844 DV N  7 2 S- N
PB2845 DV N  7 2 S- N
PB2846 DV N  7 2 S- N
PB2900 DV N  9 2 S- N
PB2901 DV D 11 - S- N
PB2906 DV D 11 - S- N
PB2907 DV N  9 2 S- N
PB2908 DV N  9 2 S- N
PB2921 DV N  9 2 S- N
PB2922 DV N  9 2 S- N
PB2923 DV N  9 2 S- N
PB2924 DV N  9 2 S- N
PB2925 DV N  9 2 S- N
PB2926 DV N  9 2 S- N
PB2927 DV N  9 2 S- N
PB2928 DV N  9 2 S- N
PB2929 DV N  9 2 S- N
PB2930 DV N  9 2 S- N
PB2931 DV N  9 2 S- N
PB2932 DV N  9 2 S- N
PB2933 DV N  9 2 S- N
PB2934 DV N  9 2 S- N
PB2935 DV N  9 2 S- N
PB2936 DV N  9 2 S- N
PB2937 DV N  9 2 S- N
PB2941 DV N  7 2 S- N
PB2942 DV N  7 2 S- N
PB2943 DV N  7 2 S- N
PB2944 DV N  7 2 S- N
PB2945 DV N  7 2 S- N
PB2946 DV N  7 2 S- N
PB3000 DV N  7 2 S- N
PB3001 DV D 11 - S- N
PB3002 DV N  7 2 S- N
PB3003 DV N  9 2 S- N
PB3004 DV N  9 2 S- N
PB3005 DV N  9 2 S- N
PB3006 DV D 11 - S- N
PB3007 DV N  9 2 S- N
PB3009 DV S 12 - S- N
PB3100 DV N  7 2 S- N
PB3101 DV D 11 - S- N
PB3102 DV N  7 2 S- N
PB3103 DV N  9 2 S- N
PB3106 DV D 11 - S- N
PB3107 DV N  9 2 S- N
PB3108 DV N  9 2 S- N
PB3121 DV N  9 2 S- N
PB3122 DV N  9 2 S- N
PB3123 DV N  9 2 S- N
PB3128 DV N  9 2 S- N
PB3200 DV S 20 - S- N
PB3201 DV D 11 - S- N
PB3202 DV S 10 - S- N
PB3300 DV N  9 2 S- N
PB3301 DV D 11 - S- N
PB3302 DV N  7 2 S- N
PB3303 DV N  7 2 S- N
PB3306 DV D 11 - S- N
PB3307 DV N  9 2 S- N
PB3308 DV N  9 2 S- N
PB3309 DV N  7 2 S- N
PB3310 DV N  7 2 S- N
PB3311 DV N  7 2 S- N
PB3312 DV N  7 2 S- N
PB3313 DV N  7 2 S- N
PB3314 DV N  7 2 S- N
PB3315 DV N  7 2 S- N
PB3316 DV N  7 2 S- N
PB3317 DV N  7 2 S- N
PB3318 DV N  7 2 S- N
PB3319 DV N  7 2 S- N
PB3320 DV N  7 2 S- N
PB3330 DV N  9 2 S- N
PB3333 DV N  9 2 S- N
PB3334 DV N  9 2 S- N
PB3335 DV N  9 2 S- N
PB3336 DV N  9 2 S- N
PB3337 DV N  9 2 S- N
PB3400 DV S 30 - S- N
PB3401 DV D 11 - S- N
PB3406 DV D 11 - S- N
PE0357 DV S  2 - S- N
PE0404 DV N  6 2 SV N
PE1230 DV N  6 4 S- N
PE1231 DV N  6 4 S- N
PE1232 DV N  6 4 S- N
PE1233 DV N  6 4 S- N
PE1234 DV N  6 4 S- N
PE1235 DV N  6 4 S- N
PE1236 DV N  6 4 S- N
PE1237 DV N  6 4 S- N
PE1238 DV N  6 4 S- N
PE1239 DV N  7 4 S- N
PE1240 DV N  7 4 S- N
PE1241 DV N  7 4 S- N
PE1242 DV N  7 4 S- N
PE1243 DV N  7 4 S- N
PE1244 DV N  7 4 S- N
PE1245 DV N  7 4 S- N
PE1246 DV N  7 4 S- N
PE1247 DV S  1 - S- N
PE1248 DV S  1 - S- N
PE1249 DV S  1 - S- N
PE1250 DV S  1 - S- N
PE1251 DV S  1 - S- N
PE1252 DV S 50 - S- N
PE3400 DV S  2 - S- N
PE3801 PS D 11 - S- J
PE3802 PS D 11 - S- J
PE3803 PS D 11 - S- J
PE3804 DV S  8 - S- N
PE3805 DV S 50 - S- N
PE3806 DV D 11 - S- N
PI0001 DV D 11 - S- N
PI0002 DV D 11 - S- N
PI0003 DV S 25 - S- N
PI0004 DV D 11 - S- N
PI0005 DV S  2 - S- N
PI0006 DV N  5 2 S- N
PI0007 DV N  3 - S- N
PI0008 DV N  7 2 S- N
PI0009 DV S  2 - S- N
PI0010 DV D 11 - S- N
PI0011 DV D 11 - S- N
PI0012 DV S  2 - S- N
PI0013 DV N  5 2 S- N
PI0014 DV N  2 - S- N
PI0015 DV N  7 2 S- N
PI0016 DV D 11 - S- N
PI0017 DV D 11 - S- N
PI0018 DV S  2 - S- N
PI0019 DV N  7 2 S- N
PI0020 DV D 11 - S- N
PI0021 DV D 11 - S- N
PI0022 DV S  2 - S- N
PI0023 DV N  5 2 S- N
PI0024 DV N  2 - S- N
PI0025 DV N  7 2 S- N
PI0026 DV D 11 - S- N
PI0027 DV D 11 - S- N
PI0028 DV S 25 - S- N
PI0029 DV D 11 - S- N
PI0030 DV N  5 2 S- N
PI0031 DV N  2 - S- N
PI0032 DV N  7 2 S- N
PI0033 PS S 50 - S- N
PI0034 PS S 25 - S- N
PI0035 PS S 15 - S- N
PI0036 PS S 15 - S- N
PI0037 PS S 50 - S- N
PI0038 PS S 25 - S- N
PI0039 PS S 15 - S- N
PI0040 PS S 15 - S- N
PI0041 PS S 50 - S- N
PI0042 PS S 50 - S- N
PI0043 PS D 11 - S- N
PI0044 PS D 11 - S- N
PI0045 PS S  1 - S- N
PI0046 PS N  6 2 S- N
PI0047 PS S 50 - S- N
PI0048 PS S 50 - S- N
PI0049 PS D 11 - S- N
PI0050 PS D 11 - S- N
PI0051 PS S  1 - S- N
PI0052 PS N  6 2 S- N
PI0053 PS S 50 - S- N
PI0054 PS S 50 - S- N
PI0055 PS D 11 - S- N
PI0056 PS D 11 - S- N
PI0057 PS S  1 - S- N
PI0058 PS N  6 2 S- N
PI0059 PS D 11 - S- N
PI0060 PS D 11 - S- N
PI0061 PS D 11 - S- N
PI0062 PS D 11 - S- N
PI0063 PS D 11 - S- N
PI0064 PS D 11 - S- N
PI0065 PS D 11 - S- N
PI0066 PS D 11 - S- N
PI0067 PS S 20 - S- N
PI0068 PS S  5 - S- N
PI0069 PS S  5 - S- N
PI0070 PS S  5 - S- N
PI0071 PS S  5 - S- N
PI0072 PS S  5 - S- N
PI0073 PS S  5 - S- N
PI0074 PS S  5 - S- N
PI0075 DV D 11 - S- N
PI0076 DV D 11 - S- N
PI0077 DV D 11 - S- N
PI0078 DV D 11 - S- N
PI0079 DV D 11 - S- N
PI0080 DV D 11 - S- N
PI0081 DV D 11 - S- N
PI0082 DV D 11 - S- N
PI0083 DV D 11 - S- N
PI0084 DV D 11 - S- N
PI0085 DV D 11 - S- N
PI0086 PS S  5 - S- N
PI0087 PS S  5 - S- N
PI0088 DV D 11 - S- N
PI0089 DV D 11 - S- N
PI0090 DV S 25 - S- N
PI0091 DV S  2 - S- N
PI0092 DV S 25 - S- N
PI0093 DV S 25 - S- N
PI0094 DV S 25 - S- N
PI0095 DV S 25 - S- N
PI0096 DV D 11 - S- N
PI0097 DV D 11 - S- N
PI0098 DV S 25 - S- N
PI0099 DV N  6 2 S- N
PI0100 DV D 11 - S- N
PI0101 DV D 11 - S- N
PI0102 DV D 11 - S- N
PI0103 DV D 11 - S- N
PI0104 DV D 11 - S- N
PI0105 DV S 20 - S- N
PI0106 DV D 11 - S- N
PI0107 DV S 11 - S- N
PI0108 DV D 11 - S- N
PI0109 DV S 20 - S- N
PI0110 DV D 11 - S- N
PI0111 DV D 11 - S- N
PI0112 DV S 20 - S- N
PI0113 DV S 20 - S- N
PI0114 DV S 25 - S- N
PI0115 DV D 11 - S- N
PI0116 DV D 11 - S- N
PI0117 DV D 11 - S- N
PI0118 DV S  2 - S- N
PI0119 DV N  5 2 S- N
PI0120 DV S 25 - S- N
PI0121 DV D 11 - S- N
PI0122 DV D 11 - S- N
PI0123 DV D 11 - S- N
PI0124 DV S  2 - S- N
PI0125 DV N  5 2 S- N
PI0126 DV S 25 - S- N
PI0127 DV D 11 - S- N
PI0128 DV D 11 - S- N
PI0129 DV D 11 - S- N
PI0130 DV S  2 - S- N
PI0131 DV N  5 2 S- N
PI0132 PS S 20 - S- N
PI0133 PS S 20 - S- N
PI0134 PS S 20 - S- N
PI0135 PS S 20 - S- N
PI0136 PS S 20 - S- N
PI0137 PS D 11 - S- N
PI0138 PS D 11 - S- N
PI0139 PS D 11 - S- N
PI0140 PS D 11 - S- N
PI0141 PS D 11 - S- N
PI0142 PS D 11 - S- N
PI0143 PS D 11 - S- N
PI0144 PS D 11 - S- N
PI0145 PS D 11 - S- N
PI0146 PS D 11 - S- N
PI0147 PS S  5 - S- N
PI0148 PS S  5 - S- N
PI0149 PS S  5 - S- N
PI0150 PS S  5 - S- N
PI0151 PS S  5 - S- N
PI0152 DV D 11 - S- N
PI0153 DV D 11 - S- N
PI0154 DV D 11 - S- N
PI0155 DV N  2 - S- N
PI0156 DV D 11 - S- N
PI0157 DV S  2 - S- N
PI0158 DV D 11 - S- N
PI0159 DV D 11 - S- N
PI0160 DV D 11 - S- N
PI0161 DV D 11 - S- N
PI0162 DV D 11 - S- N
PI0163 PS N 10 - S- N
PI0164 PS S  2 - S- N
PI0165 PS S 10 - S- N
PI0166 PS S 20 - S- N
PI0167 PS S 20 - S- N
PI0168 PS D 11 - S- N
PI0169 PS D 11 - S- N
PI0170 PS D 11 - S- N
PI0171 PS D 11 - S- N
PI0172 DV S 25 - S- N
PI0173 DV D 11 - S- N
PI0174 DV D 11 - S- N
PI0175 DV D 11 - S- N
PI0176 DV N  9 2 S- N
PI0177 DV N  9 2 S- N
PI0178 DV N  9 2 S- N
PI0179 DV D 11 - S- N
PI0180 DV D 11 - S- N
PI0181 DV N  9 2 S- N
PI0182 DV N  9 2 S- N
PI0183 DV N  9 2 S- N
PI0184 DV N  1 0 S- N
PI0185 DV S 25 - S- N
PI0186 DV D 11 - S- N
PI0187 DV D 11 - S- N
PI0188 DV D 11 - S- N
PI0189 DV N  9 2 S- N
PI0190 DV N  9 2 S- N
PI0191 DV N  9 2 S- N
PI0192 DV D 11 - S- N
PI0193 DV D 11 - S- N
PI0194 DV N  9 2 S- N
PI0195 DV N  9 2 S- N
PI0196 DV N  9 2 S- N
PI0197 DV N  1 0 S- N
PI0198 DV S 25 - S- N
PI0199 DV D 11 - S- N
PI0200 DV D 11 - S- N
PI0201 DV D 11 - S- N
PI0202 DV N  9 2 S- N
PI0203 DV N  9 2 S- N
PI0204 DV N  9 2 S- N
PI0205 DV D 11 - S- N
PI0206 DV D 11 - S- N
PI0207 DV N  9 2 S- N
PI0208 DV N  9 2 S- N
PI0209 DV N  9 2 S- N
PI0210 DV N  1 0 S- N
PI0211 DV S 25 - S- N
PI0212 DV D 11 - S- N
PI0213 DV D 11 - S- N
PI0214 DV D 11 - S- N
PI0215 DV N  9 2 S- N
PI0216 DV N  9 2 S- N
PI0217 DV N  9 2 S- N
PI0218 DV D 11 - S- N
PI0219 DV D 11 - S- N
PI0220 DV N  9 2 S- N
PI0221 DV N  9 2 S- N
PI0222 DV N  9 2 S- N
PI0223 DV N  1 0 S- N
PI0224 DV S  2 - S- N
PI0225 DV D 11 - S- N
PI0226 DV D 11 - S- N
PI0227 DV N  5 2 S- N
PI0228 DV N  1 - S- N
PI0229 DV N  6 2 S- N
PI0230 DV S  2 - S- N
PI0231 DV S 50 - S- N
PI0232 DV S 50 - S- N
PI0233 DV N  1 - S- N
PI0234 DV S  2 - S- N
PI0235 DV S  2 - S- N
PI0236 DV D 11 - S- N
PI0237 DV D 11 - S- N
PI0238 DV N  3 - S- N
PI0239 DV N  9 2 S- N
PI0240 DV D 11 - S- N
PI0241 DV D 11 - S- N
PI0242 DV D 11 - S- N
PI0243 DV N  1 - S- N
PI0244 DV N  8 2 S- N
PI0245 DV N  2 - S- N
PI0246 DV N  9 2 S- N
PI0247 DV N  9 2 S- N
PI0248 DV N  2 - S- N
PI0249 DV N  9 2 S- N
PI0250 DV N  9 2 S- N
PI0251 DV N  2 - S- N
PI0252 DV N  9 2 S- N
PI0253 DV N  9 2 S- N
PI0254 DV N  2 - S- N
PI0255 DV N  9 2 S- N
PI0256 DV N  9 2 S- N
PI0257 DV S 50 - S- N
PI0258 DV S 50 - S- N
PI0259 DV N  9 2 S- N
PI0260 DV D 11 - S- N
PI0261 DV D 11 - S- N
PI0262 DV D 11 - S- N
PI0263 DV N  3 - S- N
PI0264 DV N  3 - S- N
PI0265 DV N  3 - S- N
PI0266 DV N  3 - S- N
PI0267 DV N  6 2 S- N
PI0268 DV S  4 - S- N
PI0269 DV N  6 2 S- N
PI0270 DV N  9 2 S- N
PI0271 DV N  9 2 S- N
PI0272 DV D 11 - S- N
PI0273 DV D 11 - S- N
PI0274 DV N  6 2 S- N
PI0275 DV S  4 - S- N
PI0276 DV N  6 2 S- N
PI0277 DV N  9 2 S- N
PI0278 DV N  9 2 S- N
PI0279 DV D 11 - S- N
PI0280 DV D 11 - S- N
PI0281 DV N  6 2 S- N
PI0282 DV S  4 - S- N
PI0283 DV N  6 2 S- N
PI0284 DV N  9 2 S- N
PI0285 DV N  9 2 S- N
PI0286 DV D 11 - S- N
PI0287 DV D 11 - S- N
PI0288 DV N  6 2 S- N
PI0289 DV S  4 - S- N
PI0290 DV N  6 2 S- N
PI0291 DV N  9 2 S- N
PI0292 DV N  9 2 S- N
PI0293 DV D 11 - S- N
PI0294 DV D 11 - S- N
PI0295 PS D 11 - S- N
PI0296 PS D 11 - S- N
PI0297 DV N  5 2 S- N
PI0298 DV N  5 2 S- N
PI0299 DV N  5 2 S- N
PI0300 DV N  5 2 S- N
PI0301 DV N  5 2 S- N
PI0302 DV N  5 2 S- N
PI0303 DV N  5 2 S- N
PI0304 DV N  5 2 S- N
PI0305 DV N  5 2 S- N
PI0306 DV N  5 2 S- N
PI0307 DV N  5 2 S- N
PI0308 DV N  5 2 S- N
PI0309 DV N  5 2 S- N
PI0310 DV N  5 2 S- N
PI0311 DV S 30 - S- N
PI0312 DV S 30 - S- N
PI0313 PS D 11 - S- N
PI0314 PS D 11 - S- N
PI0315 PS D 11 - S- N
PI0316 PS D 11 - S- N
PI0317 PS D 11 - S- N
PI0318 PS S 30 - S- N
PI0319 PS D 11 - S- N
PI0320 PS D 11 - S- N
PI0321 PS S  5 - S- N
PI0322 PS S  5 - S- N
PI0323 PS S  5 - S- N
PI0324 PS S  5 - S- N
PI0325 PS S  5 - S- N
PI0326 PS S  5 - S- N
PI0327 PS S  5 - S- N
PI0328 DV D 11 - S- N
PI0329 DV D 11 - S- N
PI0330 DV S 50 - S- N
PI0331 DV N  1 - S- N
PI0332 DV N  1 - S- N
PI0333 DV N  5 2 S- N
PI0334 DV N  1 - S- N
PI0335 DV N  1 - S- N
PI0336 DV S 50 - S- N
PI0337 DV D 11 - S- N
PI0338 DV N  9 0 S- N
PI0339 DV N  9 0 S- N
PI0340 DV N  9 0 S- N
PI0341 DV N  9 0 S- N
PI0342 DV N  9 0 S- N
PI0343 DV N  9 0 S- N
PI0344 DV N  9 0 S- N
PI0345 DV N  9 0 S- N
PI0346 DV N  9 0 S- N
PI0347 DV N  9 0 S- N
PI0348 DV N  9 0 S- N
PI0349 DV N  9 0 S- N
PI0350 DV S 25 - S- N
PI0351 DV S 25 - S- N
PI0352 DV S 25 - S- N
PI0353 DV D 11 - S- N
PI0354 DV S 10 - S- N
PI0355 DV S 10 - S- N
PI0356 DV N  9 0 S- N
PI0357 DV N  9 0 S- N
PI0358 DV N  9 0 S- N
PI0359 DV N  9 0 S- N
PI0360 DV N  9 0 S- N
PI0361 DV N  9 0 S- N
PI0362 DV N  9 0 S- N
PI0363 DV N  9 0 S- N
PI0364 DV N  9 0 S- N
PI0365 DV N  9 0 S- N
PI0366 DV N  9 0 S- N
PI0367 DV N  2 0 S- N
PI0368 DV N  9 2 -V N
PI0369 DV N  9 2 -V N
PI0370 DV N  9 2 -V N
PI0371 DV N  9 2 -V N
PI0372 DV N  9 2 -V N
PI0373 DV N  9 2 -V N
PI0374 DV N  9 2 -V N
PI0375 DV N  9 2 -V N
PI0376 DV N  9 2 -V N
PI0377 DV N  9 2 -V N
PI0378 DV D 11 - S- N
PI0379 DV D 11 - S- N
PI0380 DV D 11 - S- N
PI0381 DV S 50 - S- N
PI0382 DV S 12 - S- N
PI0383 DV S 12 - S- N
PI0384 DV S 12 - S- N
PI0385 DV D 11 - S- N
PI0386 DV S 12 - S- N
PI0387 DV S 12 - S- N
PI0388 DV D 11 - S- N
PI0389 DV S 12 - S- N
PI0390 DV D 11 - S- N
PI0391 DV S 12 - S- N
PI0392 DV D 11 - S- N
PI0393 DV N  5 0 S- N
PI0394 DV D 11 0 S- N
PI0395 DV S  3 0 S- N
PI0396 DV S  3 0 S- N
PI0397 DV S  3 0 S- N
PI0398 DV S  3 0 S- N
PI0399 DV S  3 0 S- N
PI0400 DV S  2 0 S- N
PI0401 DV D 11 0 S- N
PI0402 DV S  2 0 S- N
PI0403 DV S  2 0 S- N
PI0404 DV S  2 - S- N
PI0405 DV N  2 - S- N
PI0406 DV S  2 - S- N
PI0407 DV D 11 - S- N
PI0408 DV D 11 - S- N
PI0409 DV S 11 - S- N
PI0410 DV S 50 - S- N
PI0411 DV D 11 0 S- N
PI0412 DV D 11 0 S- N
PI0413 DV N  9 2 S- N
PI0414 DV N  9 2 S- N
PI0415 DV N  9 2 S- N
PI0416 DV N  9 2 S- N
PI0417 DV D 11 0 S- N
PI0418 DV D 11 0 S- N
PI0419 DV D 11 - S- N
PI0420 DV D 11 - S- N
PI0421 DV S 20 - S- N
PI0422 DV S 20 - S- N
PI0423 DV S  3 - S- N
PI0425 DV D 11 0 S- N
PI0426 DV S 50 0 S- N
PI0427 DV D 11 0 S- N
PI0428 DV D 11 0 S- N
PI0429 DV S 50 0 S- N
PI0430 DV S 50 0 S- N
PI0431 PS S  3 0 S- N
PI0432 PS S  3 0 S- N
PI0433 PS S  3 0 S- N
PI0434 PS S  3 0 S- N
PI0435 PS S  3 0 S- N
PI0436 PS S  3 0 S- N
PI0437 PS S  3 - S- N
PI0438 PS S  3 - S- N
PI0439 DV S  3 - S- N
PI0440 PS S  3 - S- N
PI0441 DV S  3 - S- N
PI0442 DV S  3 - S- N
PI0443 DV S  3 - S- N
PI0444 DV S  3 - S- N
PI0445 DV S  3 - S- N
PI0446 PS S  3 - S- N
PI0447 DV S  3 - S- N
PI0448 DV D 11 0 S- N
PI0449 DV D 11 0 S- N
PI0450 DV D 11 0 S- N
PI0451 DV D 11 0 S- N
PI0452 DV D 11 0 S- N
PO0001 PS D 11 - S- N
PO0002 PS D 11 - S- N
PO0003 PS D 11 - S- N
PO0004 PS D 11 - S- N
PO0011 DV S 20 - S- N
PO0012 PS S 25 - S- N
PO0013 PS S 25 - S- N
PO0014 PS S 20 - S- N
PO0015 PS S 20 - S- N
PO0016 PS S 25 - S- N
PO0017 PS S 25 - S- N
PO0018 PS S 20 - S- N
PO0019 PS S 20 - S- N
PO0020 PS S 25 - S- N
PO0021 PS S 25 - S- N
PO0022 PS S 20 - S- N
PO0023 PS S 20 - S- N
PO0025 PS D 11 - S- N
PO0026 PS D 11 - S- N
PO0027 PS D 11 - S- N
PO0028 PS D 11 - S- N
PO0029 PS D 11 - S- N
PO0030 PS D 11 - S- N
PO0031 DV N  2 - S- N
PO0032 DV S 50 - S- N
PO0033 DV S 10 - S- N
PO0034 DV S 50 - S- N
PO0100 PS D 11 - S- N
PO0101 PS D 11 - S- N
PY0006 DV D 11 - -V J
PY0010 DV D 11 - -V J
PY0012 DV D 11 - -V J
PY0013 DV D 11 - -V J
PY0016 DV D 11 - -V J
PY0019 DV D 11 - -V J
PY0020 DV D 11 - -V J
PY0022 DV D 11 - -V J
PY0030 DV D 11 - -V J
PY0055 DV D 11 - -V J
PY0071 DV D 11 - -V J
PY0086 DV D 11 - -V J
PY0121 DV D 11 - -V J
PY0122 DV D 11 - -V J
PY0124 DV D 11 - -V J
PY0126 DV D 11 - -V J
PY0130 DV D 11 - -V J
PY0131 DV D 11 - -V J
PY0135 DV D 11 - -V J
PY0142 DV D 11 - -V J
PY0153 DV D 11 - -V J
PY0163 DV D 11 - -V J
PY0195 DV D 11 - -V J
PY0197 DV D 11 - -V J
PY0198 DV D 11 - -V J
PY0199 DV D 11 - -V J
PY0200 DV D 11 - -V J
PY0218 DV D 11 - -V J
PY0221 DV D 11 - -V J
PY0224 DV D 11 - -V J
PY0226 DV D 11 - -V J
PY0235 DV D 11 - -V J
PY0246 DV D 11 - -V J
PY0261 DV D 11 - -V J
PY0295 DV D 11 - -V J
PY0329 PS D 11 - -V J
PY0342 DV D 11 - -V J
PY0362 DV D 11 - -V J
PY0372 DV D 11 - -V J
PY0403 DV D 11 - -V J
PY0442 DV D 11 - -V J
PY0448 DV D 11 - -V J
PY0468 DV D 11 - -V J
PY0479 DV D 11 - -V J
PY0492 DV D 11 - -V J
PY0613 DV D 11 - -V J
PY0614 DV D 11 - -V J
PY0615 DV D 11 - -V J
PY0648 DV D 11 - -V J
PY0672 DV D 11 - -V J
PY0830 DV D 11 - -V J
PY0850 DV D 11 - -V J
PY1102 DV D 11 - -V J
PY1108 DV D 11 - -V J
PY1113 DV D 11 - -V J
PY1357 DV D 11 - -V J
PY1364 DV D 11 - -V J
PY1371 DV D 11 - -V J
PY1386 DV D 11 - -V J
PY1404 DV D 11 - -V J
PY1423 DV D 11 - -V J
PY1511 DV D 11 - -V J
PY1520 DV D 11 - -V J
PY1525 DV D 11 - -V J
PY1936 DV D 11 - -V J
PY2037 DV D 11 - -V J
PY2038 DV D 11 - -V J
PY2039 DV D 11 - -V J
PY2040 DV D 11 - -V J
PY2095 DV D 11 - -V J
PY2096 DV D 11 - -V J
PY2102 DV D 11 - -V J
PY2263 DV D 11 - -V J
PY2375 DV D 11 - -V J
PY2404 DV D 11 - -V J
PY2482 DV D 11 - -V J
PY2507 DV D 11 - -V J
PY2508 DV D 11 - -V J
PY2518 DV D 11 - -V J
PY2527 DV D 11 - -V J
PY2528 DV D 11 - -V J
PY2532 DV D 11 - -V J
PY2533 DV D 11 - -V J
PY2551 DV D 11 - -V J
PY2559 DV D 11 - -V J
PY2621 DV D 11 - -V J
PY2622 DV D 11 - -V J
PY2623 DV D 11 - -V J
PY2624 DV D 11 - -V J
PY2843 DV D 11 - -V J
PY2993 DV D 11 - -V J
PY3109 DV D 11 - -V J
PY3184 DV D 11 - -V J
PY3185 DV D 11 - -V J
PY3186 DV D 11 - -V J
PY3187 DV D 11 - -V J
PY3205 DV D 11 - -V J
PY3606 DV D 11 - -V J
PY4109 DV D 11 - -V J
PY5044 DV D 11 - -V J
PY6720 DV D 11 - -V J
PY6881 DV D 11 - -V J
PY6927 DV D 11 - -V J
PY6932 DV D 11 - -V J
PY6939 DV D 11 - -V J
PY6940 DV D 11 - -V J
PY6941 DV D 11 - -V J
PY6942 DV D 11 - -V J
PY7203 DV D 11 - -V J
PY7204 DV D 11 - -V J
PY7469 DV D 11 - -V J
PY7471 DV D 11 - -V J
PY8016 DV D 11 - -V J
PY8017 DV D 11 - -V J
PY8018 DV D 11 - -V J
PY8019 DV D 11 - -V J
PY8020 DV D 11 - -V J
PY8021 DV D 11 - -V J
PY8022 DV D 11 - -V J
PY8102 DV D 11 - -V J
PY8259 DV D 11 - -V J
PY9011 DV D 11 - -V J
PZ0006 DV D 11 - -V J
PZ0010 DV D 11 - -V J
PZ0012 DV D 11 - -V J
PZ0013 DV D 11 - -V J
PZ0016 DV D 11 - -V J
PZ0019 DV D 11 - -V J
PZ0020 DV D 11 - -V J
PZ0022 DV D 11 - -V J
PZ0030 DV D 11 - -V J
PZ0055 DV D 11 - -V J
PZ0071 DV D 11 - -V J
PZ0086 DV D 11 - -V J
PZ0121 DV D 11 - -V J
PZ0122 DV D 11 - -V J
PZ0124 DV D 11 - -V J
PZ0126 DV D 11 - -V J
PZ0130 DV D 11 - -V J
PZ0131 DV D 11 - -V J
PZ0135 DV D 11 - -V J
PZ0142 DV D 11 - -V J
PZ0153 DV D 11 - -V J
PZ0163 DV D 11 - -V J
PZ0195 DV D 11 - -V J
PZ0197 DV D 11 - -V J
PZ0198 DV D 11 - -V J
PZ0199 DV D 11 - -V J
PZ0200 DV D 11 - -V J
PZ0218 DV D 11 - -V J
PZ0221 DV D 11 - -V J
PZ0224 DV D 11 - -V J
PZ0226 DV D 11 - -V J
PZ0235 DV D 11 - -V J
PZ0246 DV D 11 - -V J
PZ0261 DV D 11 - -V J
PZ0295 DV D 11 - -V J
PZ0329 PS D 11 - -V J
PZ0342 DV D 11 - -V J
PZ0362 DV D 11 - -V J
PZ0372 DV D 11 - -V J
PZ0403 DV D 11 - -V J
PZ0442 DV D 11 - -V J
PZ0448 DV D 11 - -V J
PZ0468 DV D 11 - -V J
PZ0479 DV D 11 - -V J
PZ0492 DV D 11 - -V J
PZ0613 DV D 11 - -V J
PZ0614 DV D 11 - -V J
PZ0615 DV D 11 - -V J
PZ0648 DV D 11 - -V J
PZ0672 DV D 11 - -V J
PZ0830 DV D 11 - -V J
PZ0850 DV D 11 - -V J
PZ1102 DV D 11 - -V J
PZ1108 DV D 11 - -V J
PZ1113 DV D 11 - -V J
PZ1357 DV D 11 - -V J
PZ1364 DV D 11 - -V J
PZ1371 DV D 11 - -V J
PZ1386 DV D 11 - -V J
PZ1404 DV D 11 - -V J
PZ1423 DV D 11 - -V J
PZ1511 DV D 11 - -V J
PZ1520 DV D 11 - -V J
PZ1525 DV D 11 - -V J
PZ1936 DV D 11 - -V J
PZ2037 DV D 11 - -V J
PZ2038 DV D 11 - -V J
PZ2039 DV D 11 - -V J
PZ2040 DV D 11 - -V J
PZ2095 DV D 11 - -V J
PZ2096 DV D 11 - -V J
PZ2102 DV D 11 - -V J
PZ2263 DV D 11 - -V J
PZ2375 DV D 11 - -V J
PZ2404 DV D 11 - -V J
PZ2482 DV D 11 - -V J
PZ2507 DV D 11 - -V J
PZ2508 DV D 11 - -V J
PZ2518 DV D 11 - -V J
PZ2527 DV D 11 - -V J
PZ2528 DV D 11 - -V J
PZ2532 DV D 11 - -V J
PZ2533 DV D 11 - -V J
PZ2551 DV D 11 - -V J
PZ2559 DV D 11 - -V J
PZ2621 DV D 11 - -V J
PZ2622 DV D 11 - -V J
PZ2623 DV D 11 - -V J
PZ2624 DV D 11 - -V J
PZ2843 DV D 11 - -V J
PZ2993 DV D 11 - -V J
PZ3109 DV D 11 - -V J
PZ3184 DV D 11 - -V J
PZ3185 DV D 11 - -V J
PZ3186 DV D 11 - -V J
PZ3187 DV D 11 - -V J
PZ3205 DV D 11 - -V J
PZ3606 DV D 11 - -V J
PZ4109 DV D 11 - -V J
PZ5044 DV D 11 - -V J
PZ6720 DV D 11 - -V J
PZ6881 DV D 11 - -V J
PZ6927 DV D 11 - -V J
PZ6932 DV D 11 - -V J
PZ6939 DV D 11 - -V J
PZ6940 DV D 11 - -V J
PZ6941 DV D 11 - -V J
PZ6942 DV D 11 - -V J
PZ7203 DV D 11 - -V J
PZ7204 DV D 11 - -V J
PZ7469 DV D 11 - -V J
PZ7471 DV D 11 - -V J
PZ8016 DV D 11 - -V J
PZ8017 DV D 11 - -V J
PZ8018 DV D 11 - -V J
PZ8019 DV D 11 - -V J
PZ8020 DV D 11 - -V J
PZ8021 DV D 11 - -V J
PZ8022 DV D 11 - -V J
PZ8102 DV D 11 - -V J
PZ8259 DV D 11 - -V J
PZ9011 DV D 11 - -V J
RYF195 DV D 11 - S- J
RYF196 PS D 11 - S- J
RYF300 DV D 11 - S- J
RYF900 DV D 11 - S- J
RYF901 DV S 50 - S- J
RYF902 DV D 11 - S- J
RYF903 DV D 11 - S- J
RYF904 DV D 11 - S- J
RYF905 DV S 60 - S- J
RYF906 DV D 11 - S- J
RYF907 DV S 60 - S- J
SETKB  DV S  2 - S- N
'@
# END STANDARD RUBRIEKEN

# Code -> table line. A damaged line means the script itself was changed: stop instead of guessing.
$standardRubrieken = @{}
foreach ($line in ($StandardRubriekenTable -split "`r?`n")) {
    if ($line -cmatch '^([A-Z0-9]{1,6}) +(PS|DV) [A-Z] +[0-9]{1,2} [0-9-] [S-][V-] [JN-]$') {
        $standardRubrieken[$Matches[1]] = $line
    } elseif ($line.Trim()) {
        Stop-WithError "The list of standard rubrieken in this script is damaged (line '$line'). Download the script again."
    }
}

# Returns the attributes of a standard rubriek, or $null for any other code (an eigen rubriek).
function Get-StandardRubriek([string]$code) {
    $line = $standardRubrieken[$code]
    if (-not $line) { return $null }
    $f = $line -split ' +'
    return [PSCustomObject]@{
        Code     = $f[0]
        Gegsoort = $f[1]
        Type     = $f[2]
        Length   = [int]$f[3]
        Decimals = if ($f[4] -eq '-') { 0 } else { [int]$f[4] }
        Stam     = $f[5][0] -eq [char]'S'
        Variabel = $f[5][1] -eq [char]'V'
        Blocked  = $f[6] -eq 'J'
    }
}

# DECISION (owner, #11): The CSV is read and fully validated before the prompts (decision 14), so a
# user with a broken CSV sees the errors at once instead of after answering three questions.
Write-Host "`n=== Beaufort DPIA100 Generator v$ScriptVersion ===" -ForegroundColor Cyan

# -------------------------------------------------------------------------
# File Check & Delimiter Auto-Detection
# -------------------------------------------------------------------------

try {
    $csvFile = Read-CsvFile -filePath $InputCsvPath
} catch {
    Stop-WithError "Could not read '$InputCsvPath': $($_.Exception.Message)"
}
$csvText = $csvFile.Text
Write-Host "Input file        : $InputCsvPath" -ForegroundColor Gray
Write-Host "Output file       : $fullOutputPath" -ForegroundColor Gray
Write-Host "Input encoding    : $($csvFile.EncodingName)" -ForegroundColor Gray
Write-Host "Standard rubrieken: $($standardRubrieken.Count) (Beaufort export $StandardRubriekenDate)" -ForegroundColor Gray

# Ensure file is not empty or header-only
# DECISION: Immediate abort if CSV contains 0 rows or headers only.
$nonEmptyLines = @($csvText -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
if ($nonEmptyLines.Count -eq 0) {
    Stop-WithError "The file '$InputCsvPath' is completely empty. Aborting."
}
if ($nonEmptyLines.Count -eq 1) {
    Stop-WithError "The file '$InputCsvPath' contains only headers and NO data rows. Aborting."
}

$detectedDelimiter = Detect-CsvDelimiter -text $csvText
Write-Host "Detected delimiter: '$detectedDelimiter'" -ForegroundColor Gray

# DECISION: Header row parsed by the script, data rows by ConvertFrom-Csv -Header (decision 9).
$header      = Split-CsvHeader -text $csvText -delimiter $detectedDelimiter

# DECISION (owner): The delimiter hints at the number notation of the Excel that saved the CSV
# (decision 8). Dutch Excel saves with ';' and a decimal comma, English Excel with ',' and a decimal dot.
# A tab-separated file says nothing about the notation.
$numberNotation = switch ($detectedDelimiter) { ';' { 'nl' } ',' { 'en' } default { $null } }
$headerNames = Get-UniqueHeaderNames $header.Fields
# DECISION: ConvertFrom-Csv reads a first line starting with '#' as a '#TYPE' line and drops it without
# a message, so an Excel '#N/B' in the first column of the first data row would vanish. A leading line
# break prevents that; ConvertFrom-Csv skips the blank line, so the row numbers do not change.
$rows = @(("`r`n" + $header.Rest) | ConvertFrom-Csv -Delimiter $detectedDelimiter -Header $headerNames)

if ($rows.Count -eq 0) {
    Stop-WithError "No data records could be parsed from '$InputCsvPath'. Aborting."
}

# -------------------------------------------------------------------------
# Header Mapping (Order Independent)
# -------------------------------------------------------------------------

# Field labels below are used consistently in the mapping overview, missing-header and row error messages.
$colRubriek    = Find-PropertyByPattern $headerNames @("Rubriek")
$colPersNr     = Find-PropertyByPattern $headerNames @("Pers. Nr", "Pers_nr", "PersNr", "Persoon", "Persoonsnummer")
$colDienstv    = Find-PropertyByPattern $headerNames @("Dienstv", "Dienstverband", "Dv")
$colSofV       = Find-PropertyByPattern $headerNames @("S of V", "Stam")
$colWaarde     = Find-PropertyByPattern $headerNames @("Mutatie waarde", "Waarde", "Mutatiewaarde")
$colIngang     = Find-PropertyByPattern $headerNames @("Ingangs datum", "Ingangsdatum")
$colProductie  = Find-PropertyByPattern $headerNames @("Productie", "soort", "Productiesoort", "Productiecode")

$mapping = @(
    [PSCustomObject]@{ Field = "Rubriek";        Column = $colRubriek }
    [PSCustomObject]@{ Field = "Pers. Nr";       Column = $colPersNr }
    [PSCustomObject]@{ Field = "Dienstverband";  Column = $colDienstv }
    [PSCustomObject]@{ Field = "S of V";         Column = $colSofV }
    [PSCustomObject]@{ Field = "Mutatiewaarde";  Column = $colWaarde }
    [PSCustomObject]@{ Field = "Ingangsdatum";   Column = $colIngang }
    [PSCustomObject]@{ Field = "Productiesoort"; Column = $colProductie }
)

# DECISION: Show which CSV column feeds each field. Mutatiewaarde of an eigen rubriek is free text and
# cannot be validated, so a wrong column match there would otherwise go unnoticed.
Write-Host "Column mapping:" -ForegroundColor Gray
foreach ($m in $mapping) {
    $shown = if ($m.Column) { "'$($m.Column)'" } else { "(not found)" }
    Write-Host ("  {0,-15} <- {1}" -f $m.Field, $shown) -ForegroundColor Gray
}

$missingHeaders = @($mapping | Where-Object { -not $_.Column } | ForEach-Object { $_.Field })
if ($missingHeaders.Count -gt 0) {
    Stop-WithError "The following mandatory column headers are missing from the input file: $($missingHeaders -join ', '). Aborting."
}

# DECISION: One CSV column may not feed two fields (a pattern matched a header it was not meant for).
$doubleMapped = @($mapping | Group-Object -Property Column | Where-Object { $_.Count -gt 1 })
if ($doubleMapped.Count -gt 0) {
    $details = ($doubleMapped | ForEach-Object { "'$($_.Name)' -> $(($_.Group | ForEach-Object { $_.Field }) -join ' and ')" }) -join '; '
    Stop-WithError "A CSV column was matched to more than one field: $details. Rename the header(s). Aborting."
}

# -------------------------------------------------------------------------
# Row Validation & Record Building
# -------------------------------------------------------------------------

$validationErrors = [System.Collections.Generic.List[PSCustomObject]]::new()
$warnings         = [System.Collections.Generic.List[string]]::new()
$commaConversions = [System.Collections.Generic.List[string]]::new()
$preparedRecords  = [System.Collections.Generic.List[PSCustomObject]]::new()
# Grouped Waarde warnings (decision 16): empty values, and dates and times not in the Excel-tool format.
$emptyWaardes     = [System.Collections.Generic.List[string]]::new()
$dateWaardes      = [System.Collections.Generic.List[string]]::new()
$timeWaardes      = [System.Collections.Generic.List[string]]::new()
# Rows with a blocked standard rubriek (decision 19) and rewritten dates and times (decision 20).
$blockedRows      = [System.Collections.Generic.List[string]]::new()
$formatConversions = [System.Collections.Generic.List[string]]::new()
$allowedProductie = @("NOR", "COR")

$rowNumber = 1 # 1 is header
$skippedEmptyRows = 0

foreach ($row in $rows) {
    $rowNumber++

    # DECISION: Rows where every cell is blank (e.g. trailing ';;;;;;' lines from Excel) are skipped
    # instead of producing a pile of "field is empty" errors.
    $rowIsEmpty = $true
    foreach ($p in $row.PSObject.Properties) {
        if (-not [string]::IsNullOrWhiteSpace([string]$p.Value)) { $rowIsEmpty = $false; break }
    }
    if ($rowIsEmpty) {
        $skippedEmptyRows++
        continue
    }

    $rawRubriek   = [string]$row.$colRubriek
    $rawPersNr    = [string]$row.$colPersNr
    $rawDienstv   = [string]$row.$colDienstv
    $rawSofV      = [string]$row.$colSofV
    $rawWaarde    = [string]$row.$colWaarde
    $rawIngang    = [string]$row.$colIngang
    $rawProductie = [string]$row.$colProductie

    # 1. Validate Rubriek (Alphanumeric, max 6 chars, mandatory)
    # DECISION (owner): Beaufort rubrieken are 1-6 uppercase characters (e.g. E00101, P00332, SETKB); a
    # shorter code is left-aligned (decision 17, issue #25).
    # Characters are checked first: '#N/B' (an Excel lookup that found nothing) needs that message.
    $cleanRubriek = $rawRubriek.Trim().ToUpper()
    $standard = $null
    if ([string]::IsNullOrWhiteSpace($cleanRubriek)) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Rubriek"; Value = $rawRubriek; Issue = "Field is empty" })
    } elseif ($cleanRubriek -notmatch '^[A-Z0-9]+$') {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Rubriek"; Value = $rawRubriek; Issue = "Contains invalid characters (alphanumeric only)" })
    } elseif ($cleanRubriek.Length -gt 6) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Rubriek"; Value = $rawRubriek; Issue = "Exceeds max length of 6 chars" })
    } else {
        # DECISION (owner): A code that is not a standard rubriek is allowed without a message (decision 17).
        $standard = Get-StandardRubriek $cleanRubriek
    }

    # 2. Validate Pers. Nr (Numeric, max 6 digits, mandatory)
    # DECISION: Documentation mandates no leading zeros ("Er mogen geen voorloopnullen in staan").
    # DECISION: Digits only (see architectural decision 5). No silent cleaning of '1234.0' or '12abc'.
    $cleanPersNr = ""
    $trimPersNr = $rawPersNr.Trim()
    if ([string]::IsNullOrWhiteSpace($trimPersNr)) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Pers. Nr"; Value = $rawPersNr; Issue = "Field is empty" })
    } elseif ($trimPersNr -notmatch '^[0-9]+$') {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Pers. Nr"; Value = $rawPersNr; Issue = "Must contain digits only (no decimals, spaces or letters)" })
    } else {
        $cleanPersNr = $trimPersNr.TrimStart('0')
        if ($cleanPersNr.Length -eq 0) {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Pers. Nr"; Value = $rawPersNr; Issue = "Cannot be zero" })
        } elseif ($cleanPersNr.Length -gt 6) {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Pers. Nr"; Value = $rawPersNr; Issue = "Exceeds max length of 6 digits" })
        }
    }

    # 3. Validate Dienstv (Numeric, max 3 digits, optional)
    # ASSUMPTION: If Dienstv is empty (or all zeros, e.g. '0'), this is a Person-level mutation (PS).
    # If provided, it is a Contract-level mutation (DV) and must not exceed 3 digits.
    # For a standard rubriek the key must match its gegevenssoort (decision 18, below).
    $cleanDienstv = ""
    $dienstvValid = $true
    $trimDienstv = $rawDienstv.Trim()
    if ($trimDienstv.Length -gt 0) {
        if ($trimDienstv -notmatch '^[0-9]+$') {
            $dienstvValid = $false
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Dienstverband"; Value = $rawDienstv; Issue = "Must contain digits only (no decimals, spaces or letters)" })
        } else {
            $cleanDienstv = $trimDienstv.TrimStart('0')
            if ($cleanDienstv.Length -gt 3) {
                $dienstvValid = $false
                $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Dienstverband"; Value = $rawDienstv; Issue = "Exceeds max length of 3 digits" })
            }
        }
    }

    # 4. Validate S of V (Must strictly be 'S' or 'V')
    # DECISION: Beaufort rubric table indicator can only be 'S' (Stam) or 'V' (Variabel).
    $cleanSofV = $rawSofV.Trim().ToUpper()
    $indicator = if ($cleanSofV -in @("S", "STAM")) { "S" } elseif ($cleanSofV -in @("V", "VARIABEL")) { "V" } else { "" }
    if ([string]::IsNullOrWhiteSpace($indicator)) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "S of V"; Value = $rawSofV; Issue = "Must strictly be 'S' (Stam) or 'V' (Variabel)" })
    }

    # 4b. Standard rubriek: key, S of V and Blok ind mut (decisions 18 and 19)
    # DECISION (owner): A mismatch with the standard rubriek is an error; Beaufort rejects it or its effect
    # is unknown. Only checked on a valid Dienstverband, so a broken value gets one error, not two.
    if ($standard) {
        if ($dienstvValid -and $standard.Gegsoort -eq 'DV' -and $cleanDienstv -eq '') {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Dienstverband"; Value = $rawDienstv; Issue = "Rubriek $cleanRubriek is a DV rubriek: fill in the dienstverband (empty or 0 gives a PS key, which Beaufort rejects)" })
        } elseif ($dienstvValid -and $standard.Gegsoort -eq 'PS' -and $cleanDienstv -ne '') {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Dienstverband"; Value = $rawDienstv; Issue = "Rubriek $cleanRubriek is a PS rubriek: leave Dienstverband empty" })
        }
        if ($indicator -eq 'S' -and -not $standard.Stam) {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "S of V"; Value = $rawSofV; Issue = "Rubriek $cleanRubriek is a variabel gegeven: use V" })
        } elseif ($indicator -eq 'V' -and -not $standard.Variabel) {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "S of V"; Value = $rawSofV; Issue = "Rubriek $cleanRubriek is a stamgegeven: use S" })
        }
        if ($standard.Blocked) {
            $blockedRows.Add("row $rowNumber $cleanRubriek")
        }
    }

    # 5. Validate Ingangsdatum (Must parse to a valid date)
    # DECISION: Future dates are permitted by Beaufort, but must be syntactically valid calendar dates.
    $dtIngang = [DateTime]::MinValue
    if (-not (Test-And-ParseDate -rawDate $rawIngang -parsedDate ([ref]$dtIngang))) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Ingangsdatum"; Value = $rawIngang; Issue = "Invalid or unparseable date" })
    }

    # 6. Validate Productiesoort (Must be NOR or COR; defaults to NOR if empty)
    # DECISION: JWP is rejected because it needs a verrekenperiode, which is not implemented (decision 11).
    # Not in current spec; may be added later (.NOT_IN_CURRENT_SPEC B).
    $cleanProductie = if ([string]::IsNullOrWhiteSpace($rawProductie)) { "NOR" } else { $rawProductie.Trim().ToUpper() }
    if ($cleanProductie -eq "JWP") {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Productiesoort"; Value = $rawProductie; Issue = "JWP is not supported by this script (requires a verrekenperiode). Use NOR" })
    } elseif ($cleanProductie -notin $allowedProductie) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Productiesoort"; Value = $rawProductie; Issue = "Must be one of: $($allowedProductie -join ', ')" })
    }

    # 7. Waarde: validate characters, then truncate if > 50 (truncation is a warning only, does not abort)
    # DECISION: Line breaks, tabs and other control characters inside a cell would split or shift the
    # fixed-width record, so they are rejected (fix the cell in the source file).
    # DECISION: Normalized to composed form (NFC) so 'o' + combining diaeresis becomes one 'ö' (1 byte).
    # DECISION: Characters that do not exist in Windows-1252 are rejected instead of becoming '?'.
    # DECISION: Truncation and amount warnings are the ONLY non-fatal exceptions (decision 3).
    # Beaufort field width for Waarde is strictly 50 characters (pos 88-137).
    # DECISION: A Dutch amount (-12,50) gets a decimal dot, as Beaufort requires in a numeric rubriek.
    # DECISION (owner): Amounts Beaufort may read differently ('3.250', '1,234', '1.234,56') get a warning,
    #           never an error: the rubriek may be a text rubriek. The delimiter decides which notation
    #           is expected ($numberNotation, decision 8).
    # DECISION (owner): No forced format on Waarde of an eigen rubriek. Dates, times and other values are
    #       passed through as typed (e.g. '1-11-2026'); Beaufort validates them against the rubriek. Empty is
    #       allowed. Empty values and dates and times in another shape than the Excel tool's get a warning
    #       (decision 16). A standard rubriek is checked by its type (decision 20).
    $cleanWaarde = $rawWaarde.Trim()
    if ($cleanWaarde -match '[\x00-\x1F\x7F]') {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = ($cleanWaarde -replace '[\x00-\x1F\x7F]+', ' <line break/tab> '); Issue = "Contains line break, tab or control character" })
        $cleanWaarde = ($cleanWaarde -replace '[\x00-\x1F\x7F]+', ' ')
    }
    try {
        $cleanWaarde = $cleanWaarde.Normalize([System.Text.NormalizationForm]::FormC)
    } catch {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Contains invalid Unicode characters" })
    }
    $badChars = Get-UnencodableChars $cleanWaarde
    if ($badChars) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Characters not available in Windows-1252: $badChars" })
    }
    # DECISION (owner): The type of a standard rubriek decides the checks (decision 20): N numeric, D date,
    # T and U time; any other type (S, H, I) is text. '' = not a standard rubriek (an eigen rubriek).
    $waardeType = if (-not $standard) { '' } elseif ($standard.Type -in @('N', 'D', 'T', 'U')) { $standard.Type } else { 'S' }
    # Number-like: only digits, dots and commas with an optional minus. A date with dots is not an amount.
    # Amounts (decision 8) only for an eigen rubriek or a standard numeric one: text is written as typed.
    $isNumberLike = $waardeType -in @('', 'N') -and $cleanWaarde -match '^\-?\s*[\d\.,]+$' -and $cleanWaarde -match '\d' -and
                    $cleanWaarde -notmatch '^\d{1,2}\.\d{1,2}\.\d{4}$'
    $rowLabel = "Row $rowNumber (Pers. Nr: $cleanPersNr)"
    if ($isNumberLike -and $cleanWaarde -match '^(\-)?\s*(\d+)(?:,(\d+))?$') {
        $sign = if ($Matches[1] -eq '-') { "-" } else { "" }
        $whole = $Matches[2]
        $fraction = $Matches[3]
        $converted = if ($fraction) { "$sign$whole.$fraction" } else { "$sign$whole" }
        $thousandShape = $fraction -and $whole -match '^[1-9]\d{0,2}$' -and $fraction.Length -eq 3
        if ($thousandShape -and $numberNotation -eq 'en') {
            # English notation: '1,234' is 1234. Not converted, because '1.234' would be read as 1,234.
            $warnings.Add("${rowLabel}: Waarde '$cleanWaarde' has a thousand separator (English notation in a comma-separated file). Written as typed; a numeric rubriek rejects it at processing (PI_2229). Write $whole$fraction.")
            $converted = $cleanWaarde
        } elseif ($thousandShape -and -not $numberNotation) {
            $warnings.Add("${rowLabel}: Waarde '$cleanWaarde' was written as '$converted' (comma read as decimal). If the comma is a thousand separator, write $whole$fraction.")
        } elseif ($fraction) {
            $commaConversions.Add("row $rowNumber '$cleanWaarde' -> '$converted'")
        }
        $cleanWaarde = $converted
    } elseif ($isNumberLike -and $numberNotation -ne 'en' -and $cleanWaarde -match '^\-?\s*[1-9]\d{0,2}\.\d{3}$') {
        # In English notation '3.250' is a plain decimal (3,25), as Beaufort reads it: no warning.
        $warnings.Add("${rowLabel}: Waarde '$cleanWaarde' is read by Beaufort as $($cleanWaarde -replace '\.', ',') in a numeric rubriek. If the dot is a thousand separator, write it without the dot.")
    } elseif ($isNumberLike -and ($cleanWaarde -match '\..*,|,.*\.|\..*\.|,.*,')) {
        $warnings.Add("${rowLabel}: Waarde '$cleanWaarde' looks like an amount with a thousand separator. A numeric rubriek rejects it at processing (PI_2229); write it without, e.g. 1234.56.")
    }

    # DECISION (owner): A standard rubriek checks the value against its type, length and decimals
    # (decision 20). Errors for what Beaufort rejects or would change; dates and times get the shape of the
    # rubriek, which is reported. An empty value is a delete, with the warning below.
    $code = $cleanRubriek
    if ($waardeType -eq 'N' -and $cleanWaarde.Length -gt 0) {
        if ($cleanWaarde -notmatch '^-?[0-9]+(\.[0-9]+)?$') {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Rubriek $code is numeric: write a number with a decimal point, e.g. -1234.56" })
        } else {
            $numberParts = $cleanWaarde.TrimStart('-').Split('.')
            $wholeDigits = $numberParts[0].TrimStart('0').Length
            $significantDecimals = if ($numberParts.Count -gt 1) { $numberParts[1].TrimEnd('0').Length } else { 0 }
            # The length of a numeric rubriek includes the decimal point and the decimals.
            $maxWhole = $standard.Length - $standard.Decimals - $(if ($standard.Decimals -gt 0) { 1 } else { 0 })
            if ($significantDecimals -gt $standard.Decimals) {
                $kept = $cleanWaarde.Split('.')[0] + $(if ($standard.Decimals -gt 0) { '.' + $numberParts[1].Substring(0, $standard.Decimals) } else { '' })
                $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Rubriek $code allows $($standard.Decimals) decimal(s): Beaufort would store $kept" })
            } elseif ($wholeDigits -gt $maxWhole) {
                $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Rubriek $code allows at most $maxWhole digit(s) before the decimal point" })
            }
        }
    } elseif ($waardeType -eq 'D' -and $cleanWaarde.Length -gt 0) {
        $dtWaarde = [DateTime]::MinValue
        if (Test-And-ParseDate -rawDate $cleanWaarde -parsedDate ([ref]$dtWaarde)) {
            $shaped = $dtWaarde.ToString('dd-MM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
            if ($shaped -ne $cleanWaarde) { $formatConversions.Add("row $rowNumber '$cleanWaarde' -> '$shaped'") }
            $cleanWaarde = $shaped
        } else {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Rubriek $code is a date: write dd-mm-yyyy, e.g. 01-11-2026" })
        }
    } elseif ($waardeType -in @('T', 'U') -and $cleanWaarde.Length -gt 0) {
        # T holds hh:mm, U hh:mm:ss. Seconds other than 00 do not fit a T rubriek.
        $isTime = $cleanWaarde -match '^([0-9]{1,2}):([0-9]{2})(?::([0-9]{2}))?$' -and [int]$Matches[1] -le 23 -and
                  [int]$Matches[2] -le 59 -and (-not $Matches[3] -or [int]$Matches[3] -le 59) -and
                  -not ($waardeType -eq 'T' -and $Matches[3] -and $Matches[3] -ne '00')
        if ($isTime) {
            $shaped = '{0:00}:{1}' -f [int]$Matches[1], $Matches[2]
            if ($waardeType -eq 'U') { $shaped += ':' + $(if ($Matches[3]) { $Matches[3] } else { '00' }) }
            if ($shaped -ne $cleanWaarde) { $formatConversions.Add("row $rowNumber '$cleanWaarde' -> '$shaped'") }
            $cleanWaarde = $shaped
        } else {
            $timeShape = if ($waardeType -eq 'T') { 'hh:mm, e.g. 08:30' } else { 'hh:mm:ss, e.g. 08:30:00' }
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Rubriek $code is a time: write $timeShape" })
        }
    } elseif ($waardeType -eq 'S' -and $standard.Length -lt 50 -and $cleanWaarde.Length -gt $standard.Length) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Mutatiewaarde"; Value = $rawWaarde; Issue = "Rubriek $code allows at most $($standard.Length) characters" })
    }

    # DECISION (owner): An empty Waarde is a deliberate delete of the current value in Beaufort; the Excel
    # tool only allowed it through the keyword NULL. For an eigen rubriek dates and times are written as
    # typed, but the Excel tool wrote them as dd-mm-yyyy and hh:mm (and Beaufort type U holds hh:mm:ss), so
    # other shapes get a warning (decision 16). A standard rubriek has its type instead (decision 20).
    if ($cleanWaarde.Length -eq 0) {
        $emptyWaardes.Add("row $rowNumber (Pers. Nr: $cleanPersNr)")
    } elseif ($waardeType -eq '' -and
              $cleanWaarde -match '^[0-9]{1,2}([-/. ])[0-9]{1,2}\1[0-9]{4}$|^[0-9]{1,2}([-/])[0-9]{1,2}\2[0-9]{2}$|^[0-9]{4}([-/.])[0-9]{1,2}\3[0-9]{1,2}$' -and
              $cleanWaarde -notmatch '^[0-9]{2}-[0-9]{2}-[0-9]{4}$') {
        $dateWaardes.Add("row $rowNumber '$cleanWaarde'")
    } elseif ($waardeType -eq '' -and $cleanWaarde -match '^[0-9]{1,2}:[0-9]{2}(:[0-9]{2})?$' -and $cleanWaarde -notmatch '^[0-9]{2}:[0-9]{2}(:[0-9]{2})?$') {
        $timeWaardes.Add("row $rowNumber '$cleanWaarde'")
    }
    if ($cleanWaarde.Length -gt 50) {
        $origWaarde = $cleanWaarde
        $cleanWaarde = $cleanWaarde.Substring(0, 50)
        # DECISION: Display both old and truncated values so user has an audit trail of modified job titles.
        $warnings.Add("Row $rowNumber (Pers. Nr: $cleanPersNr): Waarde was truncated from $($origWaarde.Length) to 50 chars:`n    Original : '$origWaarde'`n    Truncated: '$cleanWaarde'")
    }

    # Queue record data if currently valid
    $preparedRecords.Add([PSCustomObject]@{
        Rubriek      = $cleanRubriek
        PersNr       = $cleanPersNr
        Dienstv      = $cleanDienstv
        Indicator    = $indicator
        Waarde       = $cleanWaarde
        IngangsDatum = $dtIngang.ToString("ddMMyyyy")
        Productie    = $cleanProductie
    })
}

# -------------------------------------------------------------------------
# Abort on Validation Errors
# -------------------------------------------------------------------------

# DECISION: If any field failed validation, display all collected errors and abort.
# Prevents generating a corrupt or partially valid import file.
if ($validationErrors.Count -gt 0) {
    Write-Host "`n[FATAL ERROR] $($validationErrors.Count) validation error(s) found. Process aborted!" -ForegroundColor Red
    Write-Host "No export file was created. Please fix the issues below in your CSV:`n" -ForegroundColor Yellow
    # Explicit width: without a console width (redirected output) Out-String can return an empty table.
    $validationErrors | Format-Table -AutoSize | Out-String -Width 200 | Write-Host -ForegroundColor Red
    Exit-Script 1
}

if ($preparedRecords.Count -eq 0) {
    Stop-WithError "No data rows found (all $skippedEmptyRows row(s) were empty). Aborting."
}

Write-Host "`n[OK] $($preparedRecords.Count) row(s) checked, no errors found." -ForegroundColor Green

# -------------------------------------------------------------------------
# Warnings: listed before the prompts, confirmed after them
# -------------------------------------------------------------------------

# DECISION: Empty values, dates and times are each listed in one warning, so a deliberate bulk delete or a
# column of dates stays readable (decision 16). Only the first few rows are shown.
function Join-Examples([System.Collections.Generic.List[string]]$list) {
    $shownMax = 5
    $more = if ($list.Count -gt $shownMax) { ", and $($list.Count - $shownMax) more" } else { "" }
    return (($list | Select-Object -First $shownMax) -join ', ') + $more
}
if ($blockedRows.Count -gt 0) {
    $warnings.Add("$($blockedRows.Count) row(s) use a standard rubriek that is blocked for indirect mutation (Blok ind mut = J): $(Join-Examples $blockedRows). Beaufort does not import these mutations; enter them in Beaufort itself.")
}
if ($emptyWaardes.Count -gt 0) {
    $warnings.Add("Waarde is empty in $($emptyWaardes.Count) row(s): $(Join-Examples $emptyWaardes). An empty Waarde deletes the current value of the rubriek in Beaufort. Fill in the value if that is not intended.")
}
if ($dateWaardes.Count -gt 0) {
    $warnings.Add("$($dateWaardes.Count) Waarde value(s) look like a date but are not written as dd-mm-yyyy: $(Join-Examples $dateWaardes). The official Excel tool wrote dates as dd-mm-yyyy; in a date rubriek Beaufort may read another shape differently. Write e.g. 01-11-2026.")
}
if ($timeWaardes.Count -gt 0) {
    $warnings.Add("$($timeWaardes.Count) Waarde value(s) look like a time but are not written as hh:mm or hh:mm:ss: $(Join-Examples $timeWaardes). A time rubriek in Beaufort holds hh:mm or hh:mm:ss; Beaufort may read another shape differently. Write e.g. 08:30.")
}

# DECISION: Comma -> dot conversions are listed together, so a long list of amounts stays readable and
# the change is never silent. Only the first few are shown. In a ';' file the comma is the Dutch decimal
# sign, so the conversion is only reported; otherwise it is a warning that needs confirmation.
if ($commaConversions.Count -gt 0) {
    $commaMessage = "Comma changed to dot in $($commaConversions.Count) Waarde value(s), because a numeric rubriek in Beaufort only accepts a dot (a text rubriek that is not a standard rubriek gets the dot too): $(Join-Examples $commaConversions)."
    if ($numberNotation -eq 'nl') {
        Write-Host "`n$commaMessage" -ForegroundColor Gray
    } else {
        $warnings.Insert(0, $commaMessage)
    }
}
# DECISION (owner): A date or time in a standard date or time rubriek gets the rubriek's shape; that is
# certain, so it is only reported, like the Dutch comma conversion (decision 20).
if ($formatConversions.Count -gt 0) {
    Write-Host "`nDate or time written in the shape of the rubriek in $($formatConversions.Count) Waarde value(s): $(Join-Examples $formatConversions)." -ForegroundColor Gray
}

if ($warnings.Count -gt 0) {
    Write-Host "`n[WARNINGS] $($warnings.Count) warning(s). The values are written as shown:" -ForegroundColor Yellow
    foreach ($w in $warnings) {
        Write-Host "  - $w" -ForegroundColor Yellow
    }
}

# -------------------------------------------------------------------------
# Interactive Prompts
# -------------------------------------------------------------------------

Write-Host ""

# 1. Procescode (Default: CMT, max 3 alphanumeric uppercase)
# ASSUMPTION: 'CMT' is the standard default process code defined in Beaufort (Module 200502).
$defaultProces = "CMT"
$procescode = ""
while ([string]::IsNullOrWhiteSpace($procescode)) {
    $inputP = Read-Host "Procescode [Press ENTER for '$defaultProces']"
    $val = if ([string]::IsNullOrWhiteSpace($inputP)) { $defaultProces } else { $inputP.Trim().ToUpper() }
    if ($val -match '^[A-Z0-9]{1,3}$') {
        $procescode = $val
    } else {
        Write-Warning "Procescode must be 1 to 3 alphanumeric characters (e.g. CMT, CON)."
    }
}

# 2. Creatiedatum (Default: today in 'dd mm yyyy')
# ASSUMPTION: Formatted into 8 digits (ddMMyyyy) and right-aligned with 3 spaces (pos 61-71).
$defaultDateStr = (Get-Date).ToString("dd MM yyyy")
$creatieDatumFormatted = ""
while ([string]::IsNullOrWhiteSpace($creatieDatumFormatted)) {
    $inputD = Read-Host "Creatiedatum (dd mm yyyy) [Press ENTER for '$defaultDateStr']"
    $dateStr = if ([string]::IsNullOrWhiteSpace($inputD)) { $defaultDateStr } else { $inputD.Trim() }
    $dtRef = [DateTime]::MinValue
    if (Test-And-ParseDate -rawDate $dateStr -parsedDate ([ref]$dtRef)) {
        $creatieDatumFormatted = $dtRef.ToString("ddMMyyyy")
    } else {
        Write-Warning "Invalid date '$dateStr'. Please use 'dd mm yyyy' or 'dd-mm-yyyy'."
    }
}

# 3. Creatiegebruiker (Mandatory, max 16 chars)
# ASSUMPTION: Username must match an authorized user in Beaufort for batch import (Module 100251).
$creatieGebruiker = ""
while ([string]::IsNullOrWhiteSpace($creatieGebruiker)) {
    $inputU = (Read-Host "Creatiegebruiker [Mandatory]").Trim()
    if ([string]::IsNullOrWhiteSpace($inputU)) {
        Write-Warning "Creatiegebruiker cannot be empty."
    } elseif ($inputU.Length -gt 16) {
        Write-Warning "Creatiegebruiker cannot exceed 16 characters."
    } elseif (Get-UnencodableChars $inputU) {
        Write-Warning "Creatiegebruiker contains characters that do not exist in Windows-1252: $(Get-UnencodableChars $inputU)"
    } else {
        $creatieGebruiker = $inputU
    }
}

# DECISION (owner): With warnings the file is only written after the user types Y (decision 3). Anything
# else cancels, also end of input, so an unattended run never writes a file with unconfirmed warnings.
# Asked after the prompts (decision 14), right before writing.
if ($warnings.Count -gt 0) {
    $confirm = Read-Host "`nCreate the file anyway, with the $($warnings.Count) warning(s) above? Type Y to continue, anything else cancels"
    if ("$confirm".Trim() -notmatch '^(y|yes)$') {
        Stop-WithError "Cancelled. No file was created."
    }
}

# -------------------------------------------------------------------------
# Build Output Lines (Exact 187-Character Layout)
# -------------------------------------------------------------------------

$outputLines = [System.Collections.Generic.List[string]]::new()

foreach ($rec in $preparedRecords) {
    # 1. Procescode (pos 1-3, len 3)
    $f1_Proces = $procescode.PadRight(3).Substring(0, 3)

    # 2. Rubriekcode (pos 4-9, len 6; a shorter code is left-aligned, decision 17)
    $f2_Rubriek = $rec.Rubriek.PadRight(6).Substring(0, 6)

    # 3. Object_id (pos 10-59, len 50)
    # DECISION: DV compound key format is '{pers_nr,6} {dv,3}', space padded to 50 chars.
    # PS compound key is '{pers_nr,6}', space padded to 50 chars.
    # Other object types (BG, IZ, IP, LV, SD, ZG, ZT) are not in current spec (.NOT_IN_CURRENT_SPEC A).
    if ([string]::IsNullOrWhiteSpace($rec.Dienstv)) {
        $objectKey = $rec.PersNr.PadLeft(6, ' ')
    } else {
        $objectKey = "$($rec.PersNr.PadLeft(6, ' ')) $($rec.Dienstv.PadLeft(3, ' '))"
    }
    $f3_ObjectId = $objectKey.PadRight(50, ' ')

    # 4. Indicator (pos 60, len 1)
    $f4_Indicator = $rec.Indicator

    # 5. Creatiedatum (pos 61-71, len 11 -> 3 leading spaces + 8 digits)
    $f5_CreatieDatum = $creatieDatumFormatted.PadLeft(11, ' ')

    # 6. Creatiegebruiker (pos 72-87, len 16 -> right-aligned)
    $f6_Gebruiker = $creatieGebruiker.PadLeft(16, ' ')

    # 7. Waarde (pos 88-137, len 50 -> left-aligned, space padded)
    $f7_Waarde = $rec.Waarde.PadRight(50, ' ')

    # 8. Ingangsdatum (pos 138-148, len 11 -> 3 leading spaces + 8 digits)
    $f8_IngangsDatum = $rec.IngangsDatum.PadLeft(11, ' ')

    # 9. Filler (pos 149-178, len 30)
    # NOTE: Positions 149 to 178 are not described in either Beaufort PDF and are not a field: the Excel
    # tool pads productiesoort to 33 positions (.NOT_IN_CURRENT_SPEC D). Filled with 30 spaces.
    $f9_Filler = "".PadRight(30, ' ')

    # 10. Productiesoort (pos 179-181, len 3)
    $f10_Productie = $rec.Productie.PadRight(3).Substring(0, 3)

    # 11. Verrekenperiode (pos 182-187, len 6)
    # Verrekenperiode is optional in Beaufort (only used for recalculation over a jaarwerkperiode).
    # Not in current spec: always 6 spaces; may be added later together with JWP (.NOT_IN_CURRENT_SPEC B).
    $f11_VerrekenPeriode = "".PadRight(6, ' ')

    # Assemble line: Exactly 187 characters
    $line = "$f1_Proces$f2_Rubriek$f3_ObjectId$f4_Indicator$f5_CreatieDatum$f6_Gebruiker$f7_Waarde$f8_IngangsDatum$f9_Filler$f10_Productie$f11_VerrekenPeriode"

    # INTEGRITY CHECK: Reject any line that does not evaluate to exactly 187 characters.
    if ($line.Length -ne 187) {
        Stop-WithError "Alignment error: Record for Pers. Nr $($rec.PersNr) is $($line.Length) chars instead of 187."
    }

    $outputLines.Add($line)
}

# -------------------------------------------------------------------------
# Export File strictly as Windows-1252 (ANSI)
# -------------------------------------------------------------------------

# DECISION: Enforced single-byte ANSI (Windows-1252) export.
# Guarantees that Dutch accented characters (e.g. 'ö' in Gebiedscoördinator) remain exactly 1 byte,
# preventing internal byte shifts that cut off dates and productiesoort in Beaufort.
# The output path was resolved at the start ($fullOutputPath). The strict encoder guarantees that an
# unencodable character can never silently become '?'.
$replacedStamp = $null
if (Test-Path -LiteralPath $fullOutputPath -PathType Leaf) {
    $replacedStamp = (Get-Item -LiteralPath $fullOutputPath).LastWriteTime.ToString("dd-MM-yyyy HH:mm")
}
try {
    [System.IO.File]::WriteAllLines($fullOutputPath, $outputLines, $strictAnsi)
} catch {
    Stop-WithError "Could not write '$fullOutputPath': $($_.Exception.Message) (is the file open in another program?)"
}

Write-Host "`n[SUCCESS] File successfully created:`n$fullOutputPath" -ForegroundColor Green
Write-Host "Encoding       : Windows-1252 (ANSI - single-byte compatible)" -ForegroundColor Green
Write-Host "Records parsed : $($outputLines.Count)" -ForegroundColor Green
if ($replacedStamp) {
    Write-Host "Replaced an existing file from $replacedStamp" -ForegroundColor Gray
}
if ($skippedEmptyRows -gt 0) {
    Write-Host "Empty rows skipped: $skippedEmptyRows" -ForegroundColor Gray
}
if ($warnings.Count -gt 0) {
    Write-Host "Written with $($warnings.Count) confirmed warning(s), see above." -ForegroundColor Yellow
}

Exit-Script 0
