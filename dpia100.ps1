<#
.SYNOPSIS
    Generates a fixed-width Beaufort DPIA100 import file from an input CSV file.
.DESCRIPTION
    - Automatically detects delimiter (;, ,, \t).
    - Fully independent of column order.
    - Strictly validates all fields and aborts if errors are found.
    - Truncates [Waarde] with a warning if it exceeds 50 characters (showing old & new values).
    - Warns about amounts Beaufort may read differently than intended (comma, thousand separator).
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
    - Only productiesoort NOR is supported (JWP needs a verrekenperiode, which is not implemented).
    - Scope: PS and DV mutations only. See .NOT_IN_CURRENT_SPEC for Beaufort features left out.

.BEAUFORT_LAYOUT_SPECIFICATION
    Total record length: Exactly 187 characters / bytes.
    -----------------------------------------------------------------------------------------
    Pos   1 -   3 (Len  3) : Procescode (Uppercase alphanumeric, e.g. CMT, CON)
    Pos   4 -   9 (Len  6) : Rubriekcode (Uppercase alphanumeric, e.g. E00101)
    Pos  10 -  59 (Len 50) : Object_id (pers_nr [6] + space [1] + dv_vlgnr [3], padded to 50)
    Pos  60 -  60 (Len  1) : Indicator stam/variabel ('S' or 'V')
    Pos  61 -  71 (Len 11) : Creatiedatum (3 leading spaces + ddMMyyyy)
    Pos  72 -  87 (Len 16) : Creatiegebruiker (Right-aligned username, padded with spaces)
    Pos  88 - 137 (Len 50) : Mutatiewaarde (Left-aligned, space padded to 50 chars)
    Pos 138 - 148 (Len 11) : Ingangsdatum (3 leading spaces + ddMMyyyy)
    Pos 149 - 178 (Len 30) : Not described in the Beaufort docs (always 30 spaces, see .NOT_IN_CURRENT_SPEC)
    Pos 179 - 181 (Len  3) : Productiesoort (NOR only in this script; Beaufort also knows JWP)
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
    3. WARNINGS INSTEAD OF ERRORS (Waarde only):
       Two cases are non-fatal: a Waarde longer than 50 characters is truncated (showing the original and
       truncated values; this prevents the #VALUE! error seen in Excel formulas), and amounts Beaufort may
       read differently (decision 8). All warnings are shown before the prompts, and after the prompts the
       file is only written if the user types Y. Anything else, including end of input, cancels like an
       error (decision 14).
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
       Rubrieken are created freely, so '1.234,56' may be a valid text value. The script cannot tell a
       numeric rubriek from a text rubriek, so number-like values are never rejected. Beaufort reads the
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
       Extra decimals are rounded to the rubriek's decimals by Beaufort (owner: seen with manual entry,
       assumed for import).
    9. HEADER ROW PARSED BY THE SCRIPT:
       Excel often adds empty header cells (header ending in ';;'). Windows PowerShell 5.1 aborts on
       duplicate/empty column names, PowerShell 7 renames them. The header row is parsed here and empty
       or duplicate names are made unique, so both versions behave the same.
       Patterns of 1-2 characters (e.g. 'Dv') must match the whole header; longer patterns match anywhere
       in the header. The same CSV column may not be mapped to two fields.
    10. STALE OUTPUT PROTECTION:
       If the run fails, an existing output file (from an earlier run) is renamed to '<name>.old' so an
       outdated import file cannot be imported by mistake.
    11. PRODUCTIESOORT NOR ONLY (current spec; JWP may be added later, see .NOT_IN_CURRENT_SPEC):
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
       The Excel tool has extra columns for loonverdeling, org. eenheid, functiecode and speciale doorbel.
       Field widths for the extra parts are undocumented: verify against Excel tool output before adding.
    B. VERREKENPERIODE (pos 182-187) AND JWP:
       Format ppjjjj (e.g. 132003). Periods 13-15 are only applied if the ingangsdatum is in December,
       otherwise Beaufort recalculates. For recalculation over a jaarwerkperiode the ingangsdatum must equal
       the start date of the last loonperiode. The Excel tool has a 'Verreken periode' column.
    C. RUBRIEK LOOKUP (Excel tool tab 'Rubriek': rubriek_kd, gegsrt_id, type):
       - gegsrt_id determines which object_id parts are required; unknown rubrieken are flagged.
       - type determines the Waarde format: N = numeriek (default), S = tekst, D = datum, T = tijd,
         H = hoofdletters. Currently only amounts are handled (comma -> dot, warnings, decision 8);
         dates, times and uppercase values are written as typed (owner decision, see the Waarde section).
       Left out because rubrieken are created freely in Beaufort and a list is hard to keep up to date.
       Without it, PS vs DV is derived from Dienstv being filled; Beaufort rejects mismatches at import.
    D. POSITIONS 149-178:
       Not described in either PDF: the documented fields end at Ingangsdatum (pos 148) and Productiesoort
       starts at 179. Filled with 30 spaces, which matches the Excel tool output byte for byte.
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

# DECISION: Show which CSV column feeds each field. Mutatiewaarde is free text and cannot be validated,
# so a wrong column match there would otherwise go unnoticed.
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
$allowedProductie = @("NOR")

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
    # DECISION: Beaufort rubrieken are max 6 uppercase characters (e.g. E00101, P00332).
    $cleanRubriek = $rawRubriek.Trim().ToUpper()
    if ([string]::IsNullOrWhiteSpace($cleanRubriek)) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Rubriek"; Value = $rawRubriek; Issue = "Field is empty" })
    } elseif ($cleanRubriek.Length -gt 6) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Rubriek"; Value = $rawRubriek; Issue = "Exceeds max length of 6 chars" })
    } elseif ($cleanRubriek -notmatch '^[A-Z0-9]+$') {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Rubriek"; Value = $rawRubriek; Issue = "Contains invalid characters (alphanumeric only)" })
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
    # Current spec: derived from the data, not from a rubriek lookup (.NOT_IN_CURRENT_SPEC C).
    $cleanDienstv = ""
    $trimDienstv = $rawDienstv.Trim()
    if ($trimDienstv.Length -gt 0) {
        if ($trimDienstv -notmatch '^[0-9]+$') {
            $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Dienstverband"; Value = $rawDienstv; Issue = "Must contain digits only (no decimals, spaces or letters)" })
        } else {
            $cleanDienstv = $trimDienstv.TrimStart('0')
            if ($cleanDienstv.Length -gt 3) {
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

    # 5. Validate Ingangsdatum (Must parse to a valid date)
    # DECISION: Future dates are permitted by Beaufort, but must be syntactically valid calendar dates.
    $dtIngang = [DateTime]::MinValue
    if (-not (Test-And-ParseDate -rawDate $rawIngang -parsedDate ([ref]$dtIngang))) {
        $validationErrors.Add([PSCustomObject]@{ Row = $rowNumber; Column = "Ingangsdatum"; Value = $rawIngang; Issue = "Invalid or unparseable date" })
    }

    # 6. Validate Productiesoort (Must be NOR; defaults to NOR if empty)
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
    # DECISION (owner): No forced format on Waarde. Dates, times and other values are passed through as
    #       typed (e.g. '1-11-2026'); Beaufort validates them against the rubriek. Empty is allowed.
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
    # Number-like: only digits, dots and commas with an optional minus. A date with dots is not an amount.
    $isNumberLike = $cleanWaarde -match '^\-?\s*[\d\.,]+$' -and $cleanWaarde -match '\d' -and
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

# DECISION: Comma -> dot conversions are listed together, so a long list of amounts stays readable and
# the change is never silent. Only the first few are shown. In a ';' file the comma is the Dutch decimal
# sign, so the conversion is only reported; otherwise it is a warning that needs confirmation.
if ($commaConversions.Count -gt 0) {
    $shownMax = 5
    $examples = ($commaConversions | Select-Object -First $shownMax) -join ', '
    $more = if ($commaConversions.Count -gt $shownMax) { ", and $($commaConversions.Count - $shownMax) more" } else { "" }
    $commaMessage = "Comma changed to dot in $($commaConversions.Count) Waarde value(s), because a numeric rubriek in Beaufort only accepts a dot (a text rubriek gets the dot too): $examples$more."
    if ($numberNotation -eq 'nl') {
        Write-Host "`n$commaMessage" -ForegroundColor Gray
    } else {
        $warnings.Insert(0, $commaMessage)
    }
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

    # 2. Rubriekcode (pos 4-9, len 6)
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

    # 9. Externe referentie / Filler (pos 149-178, len 30)
    # NOTE: Positions 149 to 178 are not described in either Beaufort PDF (.NOT_IN_CURRENT_SPEC D).
    # Filled with 30 spaces, which matches the Excel tool output.
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
