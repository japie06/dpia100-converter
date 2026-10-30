<#
.SYNOPSIS
    Embeds the standard PS/DV rubrieken of a Beaufort "Overzicht rubrieken" export (function 210006) in
    dpia100.ps1 and dpia100.html (decision 17).
.DESCRIPTION
    Runs tools\convert-rubrieken.ps1 -OnlyScope -OnlyStandard on the export, takes the export date from the
    page footers, and replaces the generated region between the BEGIN/END STANDARD RUBRIEKEN markers in both
    files with the same table. Then it updates the CSP hashes of the page (tests\web\update-csp.ps1).

    One line per rubriek, fixed columns, sorted by code (ordinal):

        P00001 DV N  6 2 SV N

    code (1-6 characters), gegevenssoort (PS or DV), type (N numeric, S text, D date, T time hh:mm, U time
    hh:mm:ss, ...), length, decimals ('-' = blank), stam/variabel allowed ('S-' stam only, '-V' variabel
    only, 'SV' both) and Blok ind mut ('J' blocked, 'N', '-' = blank, which counts as not blocked).

    These attributes are fixed for a standard rubriek: a customer cannot change them (Beaufort help
    "Rubrieken wijzigen of verwijderen"). Only they go in: no descriptions, values, min/max or referentie
    values (customer settings), and nothing about the customer or the user who made the export.
    Aborts without changing a file on any unexpected value.
.PARAMETER InputPath
    The exported report (.txt, UTF-8). Never commit it.
.PARAMETER ExportDate
    The date of the export as dd-MM-yyyy. Default: the date in the page footers of the report.
.PARAMETER ScriptPath
    Default: dpia100.ps1 in the repository root.
.PARAMETER PagePath
    Default: dpia100.html in the repository root.
.PARAMETER SkipCsp
    Do not run tests\web\update-csp.ps1 afterwards (for tests on copies).
.EXAMPLE
    pwsh tools\update-standard-rubrieken.ps1 -InputPath C:\Temp\OverzichtRubrieken.txt
#>
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [string]$ExportDate = '',
    [string]$ScriptPath = '',
    [string]$PagePath = '',
    [switch]$SkipCsp
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $ScriptPath) { $ScriptPath = Join-Path $root 'dpia100.ps1' }
if (-not $PagePath) { $PagePath = Join-Path $root 'dpia100.html' }

function Stop-WithError([string]$Message, $Details = @()) {
    Write-Host "[ERROR] $Message" -ForegroundColor Red
    foreach ($d in ($Details | Select-Object -First 20)) { Write-Host "        $d" }
    if (@($Details).Count -gt 20) { Write-Host "        ... and $(@($Details).Count - 20) more" }
    exit 1
}

foreach ($p in $InputPath, $ScriptPath, $PagePath) {
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { Stop-WithError "File not found: $p" }
}

# --- Export date ----------------------------------------------------------------------------------------
# The footer of every page reads "HR Core - OPIAnnn - <user>   30 okt 2026 15:13". The user is not kept.
if (-not $ExportDate) {
    $months = @{ jan = 1; feb = 2; mrt = 3; apr = 4; mei = 5; jun = 6; jul = 7; aug = 8; sep = 9; okt = 10; nov = 11; dec = 12 }
    $text = [IO.File]::ReadAllText($InputPath, (New-Object Text.UTF8Encoding($false)))
    $dates = @([regex]::Matches($text, '(?m)^\|\s*HR Core - .*?\s(\d{1,2}) ([a-z]{3}) (\d{4}) \d{2}:\d{2}\s*\|') | ForEach-Object {
        $m = $months[$_.Groups[2].Value]
        if (-not $m) { Stop-WithError "Unknown month '$($_.Groups[2].Value)' in a page footer. Pass -ExportDate dd-MM-yyyy." }
        '{0:00}-{1:00}-{2}' -f [int]$_.Groups[1].Value, $m, $_.Groups[3].Value
    } | Select-Object -Unique)
    if ($dates.Count -ne 1) { Stop-WithError "Could not read one export date from the page footers (found $($dates.Count)). Pass -ExportDate dd-MM-yyyy." }
    $ExportDate = $dates[0]
}
if ($ExportDate -notmatch '^\d{2}-\d{2}-\d{4}$') { Stop-WithError "ExportDate '$ExportDate' is not dd-MM-yyyy." }

# --- Standard PS/DV rubrieken via the cleaner ------------------------------------------------------------
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('dpia100-standard-' + [guid]::NewGuid().ToString('N') + '.csv')
try {
    # 'exit 1' in the cleaner sets the exit code; a successful run leaves it as it was, so reset it first.
    $global:LASTEXITCODE = 0
    & (Join-Path $PSScriptRoot 'convert-rubrieken.ps1') -InputPath $InputPath -OutputPath $tmp -OnlyScope -OnlyStandard
    if ($global:LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tmp)) { Stop-WithError 'convert-rubrieken.ps1 failed; nothing changed.' }
    $rows = @(Import-Csv -LiteralPath $tmp -Delimiter ';' -Encoding UTF8)
} finally {
    Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
}
if ($rows.Count -eq 0) { Stop-WithError 'The export has no standard PS/DV rubrieken.' }

$problems = New-Object Collections.Generic.List[string]
$entries = New-Object Collections.Generic.List[string]
foreach ($r in $rows) {
    $bad = @()
    if ($r.Code -cnotmatch '^[A-Z0-9]{1,6}$') { $bad += "code" }
    if ($r.Gegsoort -cnotmatch '^(PS|DV)$') { $bad += "gegevenssoort '$($r.Gegsoort)'" }
    if ($r.Typ -cnotmatch '^[A-Z]$') { $bad += "type '$($r.Typ)'" }
    if ($r.Len -notmatch '^[1-9][0-9]?$') { $bad += "length '$($r.Len)'" }
    if ($r.Dec -notmatch '^[0-9]?$') { $bad += "decimals '$($r.Dec)'" }
    if ($r.Stam -cnotmatch '^[JN]$' -or $r.Var -cnotmatch '^[JN]$' -or ($r.Stam -eq 'N' -and $r.Var -eq 'N')) { $bad += "stam '$($r.Stam)' / var '$($r.Var)'" }
    if ($r.BlokIndMut -cnotmatch '^[JN]?$') { $bad += "blok ind mut '$($r.BlokIndMut)'" }
    if ($bad.Count -gt 0) { $problems.Add("$($r.Code): $($bad -join ', ')"); continue }
    $sv = $(if ($r.Stam -eq 'J') { 'S' } else { '-' }) + $(if ($r.Var -eq 'J') { 'V' } else { '-' })
    $dec = if ($r.Dec) { $r.Dec } else { '-' }
    $blok = if ($r.BlokIndMut) { $r.BlokIndMut } else { '-' }
    $entries.Add(('{0,-6} {1} {2} {3,2} {4} {5} {6}' -f $r.Code, $r.Gegsoort, $r.Typ, $r.Len, $dec, $sv, $blok))
}
if ($problems.Count -gt 0) { Stop-WithError "Unexpected values ($($problems.Count)); nothing changed." $problems }
# Ordinal sort, so PowerShell 5.1 and 7 give the same order whatever the culture.
$sorted = $entries.ToArray()
[Array]::Sort($sorted, [StringComparer]::Ordinal)

# --- Replace the generated region in both files ------------------------------------------------------------
# Both files are prepared first and only written when both regions were found.
function Update-Region([string]$Path, [string]$Comment, [scriptblock]$Body) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $text = (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes, $(if ($hasBom) { 3 } else { 0 }), $bytes.Length - $(if ($hasBom) { 3 } else { 0 }))
    $nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $pattern = '(?ms)^([ \t]*)' + [regex]::Escape($Comment) + ' BEGIN STANDARD RUBRIEKEN.*?^[ \t]*' + [regex]::Escape($Comment) + ' END STANDARD RUBRIEKEN[^\r\n]*'
    $found = [regex]::Matches($text, $pattern)
    if ($found.Count -ne 1) { Stop-WithError "Expected one BEGIN/END STANDARD RUBRIEKEN region in $Path, found $($found.Count)." }
    $indent = $found[0].Groups[1].Value
    $region = (& $Body $indent) -join $nl
    $new = $text.Substring(0, $found[0].Index) + $region + $text.Substring($found[0].Index + $found[0].Length)
    return [PSCustomObject]@{ Path = $Path; Text = $new; HasBom = $hasBom; Changed = ($new -ne $text) }
}

$note = "BEGIN STANDARD RUBRIEKEN: generated by tools/update-standard-rubrieken.ps1 from the Beaufort export of $ExportDate, do not edit (decision 17)"
$scriptFile = Update-Region $ScriptPath '#' {
    param($indent)
    "$indent# $note"
    "$indent`$StandardRubriekenDate = '$ExportDate'"
    "$indent`$StandardRubriekenTable = @'"
    $sorted
    "'@"
    "$indent# END STANDARD RUBRIEKEN"
}
$pageFile = Update-Region $PagePath '//' {
    param($indent)
    "$indent// $note"
    "${indent}const STANDARD_RUBRIEKEN_DATE = '$ExportDate';"
    "${indent}const STANDARD_RUBRIEKEN_TABLE = String.raw``"
    $sorted
    '`;'
    "$indent// END STANDARD RUBRIEKEN"
}

foreach ($f in $scriptFile, $pageFile) {
    [IO.File]::WriteAllText($f.Path, $f.Text, (New-Object Text.UTF8Encoding($f.HasBom)))
}

$blocked = @($sorted | Where-Object { $_.EndsWith(' J') }).Count
Write-Host "Standard PS/DV rubrieken: $($sorted.Count) (Blok ind mut J: $blocked), export date $ExportDate"
Write-Host "  $ScriptPath $(if ($scriptFile.Changed) { 'updated' } else { 'unchanged' })"
Write-Host "  $PagePath $(if ($pageFile.Changed) { 'updated' } else { 'unchanged' })"
if (-not $SkipCsp) {
    & (Join-Path $root 'tests\web\update-csp.ps1') -Path $PagePath
}
Write-Host 'Add a line to docs/CHANGELOG.md when the list changed, and run the tests.' -ForegroundColor Yellow
