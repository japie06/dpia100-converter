<#
.SYNOPSIS
    Cleans the Beaufort report "Overzicht rubrieken" (function 210006) into one CSV row per rubriek.
.DESCRIPTION
    The report is a printed page layout: a page header, a column header and a footer on every page, and
    per rubriek one fixed-width record line (optionally followed by a wrapped Omschrijving line) plus a
    detail block of six lines. This script drops the page furniture and everything customer-specific
    (municipality, user name, print date), and writes the attributes that matter for dpia100:

      Code Omschrijving Fix Prompt Gegsoort Typ Len Dec BlokIndMut Referentie RefTabel Minimum Maximum Bereik

    Fix = J is a standard rubriek (shipped with Beaufort), Fix = N an "eigen rubriek" made by the customer.
    Gegsoort is the object level (PS, DV, RG, ...). BlokIndMut = J means the rubriek is blocked for indirect
    mutation, so it cannot be imported with function 100254. Typ is N (numeric), S (text), D (date), ...;
    Len includes the comma and decimals; Dec is the number of decimals (numeric only).
    See docs/rubrieken-export.md.

    The column positions are read from the column header line of the report, not hard-coded. The script
    aborts (exit 1, no output file) on anything it does not recognise: no column header, a header that
    differs between pages, a record line it cannot read, a duplicate code, or a number of detail blocks
    that differs from the number of records.

    The export holds the rubriek set of one customer and a user name. Keep it, and the CSV made from it,
    out of the repository (both are git-ignored).
.PARAMETER InputPath
    The exported report (.txt, UTF-8).
.PARAMETER OutputPath
    The CSV to write (UTF-8 with BOM, ';' as delimiter, CRLF, as Dutch Excel opens it). Default:
    rubrieken.csv next to this script.
.PARAMETER OnlyScope
    Keep only gegevenssoort PS and DV, the object types dpia100 writes.
.PARAMETER OnlyStandard
    Keep only standard rubrieken (Fix = J).
.EXAMPLE
    pwsh tools\convert-rubrieken.ps1 -InputPath C:\Temp\OverzichtRubrieken.txt -OnlyScope
#>
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [string]$OutputPath = '',
    [switch]$OnlyScope,
    [switch]$OnlyStandard
)

$ErrorActionPreference = 'Stop'

function Stop-WithError([string]$Message, $Details = @()) {
    Write-Host "[ERROR] $Message" -ForegroundColor Red
    foreach ($d in ($Details | Select-Object -First 20)) { Write-Host "        $d" }
    if (@($Details).Count -gt 20) { Write-Host "        ... and $(@($Details).Count - 20) more" }
    exit 1
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { Stop-WithError "Input file not found: $InputPath" }
if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'rubrieken.csv' }

# DECISION: strict UTF-8 (a BOM is tolerated). The export has no BOM; a legacy code page would turn
# accented names into garbage without any error, so reject instead.
$bytes = [IO.File]::ReadAllBytes($InputPath)
try { $text = (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes) }
catch { Stop-WithError "The input is not valid UTF-8. Export the report again or convert the file." }
$text = $text.TrimStart([char]0xFEFF)
$lines = $text -split "`r?`n"

# --- Column header ------------------------------------------------------------------------------------
# Expected: " Code    Omschrijving   Fix Prompt   soort   soort Typ Len Dec trm  Stam Var Verv Rubr toek Mut mut"
$expectedHeader = 'Code', 'Omschrijving', 'Fix', 'Prompt', 'soort', 'soort', 'Typ', 'Len', 'Dec',
                  'trm', 'Stam', 'Var', 'Verv', 'Rubr', 'toek', 'Mut', 'mut'
$headerLines = @($lines | Where-Object { $_ -match '^ Code\s+Omschrijving\s+Fix\s+Prompt' })
if ($headerLines.Count -eq 0) { Stop-WithError 'No column header line found. Is this the report "Overzicht rubrieken" (210006)?' }
$header = $headerLines[0]
$differing = @($headerLines | Where-Object { $_ -ne $header })
if ($differing.Count -gt 0) { Stop-WithError "The column header differs between pages ($($differing.Count) pages differ from the first)." }
$tokens = @([regex]::Matches($header, '\S+'))
$foundHeader = ($tokens | ForEach-Object { $_.Value }) -join ' '
if ($foundHeader -ne ($expectedHeader -join ' ')) {
    Stop-WithError 'Unexpected column header.' @("found:    $foundHeader", "expected: $($expectedHeader -join ' ')")
}
# A column runs from the start of its header word to the start of the next one.
$col = @()
for ($i = 0; $i -lt $tokens.Count; $i++) {
    $start = $tokens[$i].Index
    $end = if ($i + 1 -lt $tokens.Count) { $tokens[$i + 1].Index } else { [int]::MaxValue }
    $col += , @($start, $end)
}
function Get-Column([string]$Line, [int]$Index) {
    $start = $col[$Index][0]
    if ($Line.Length -le $start) { return '' }
    $len = [Math]::Min($col[$Index][1] - $start, $Line.Length - $start)
    return $Line.Substring($start, $len).Trim()
}
# Column indexes
$cCode = 0; $cOmsch = 1; $cFix = 2; $cPrompt = 3; $cRsoort = 4; $cGsoort = 5; $cTyp = 6; $cLen = 7; $cDec = 8; $cBlokInd = 16

# --- Records and detail blocks ------------------------------------------------------------------------
function Test-RecordLine([string]$Line) {
    if ($Line.Length -lt 2 -or $Line[0] -ne ' ' -or $Line[1] -eq ' ') { return $false }
    (Get-Column $Line $cFix) -match '^[JN]$' -and (Get-Column $Line $cGsoort) -match '^[A-Z0-9]{1,3}$' -and
        (Get-Column $Line $cTyp) -match '^[A-Z]$'
}

$problems = New-Object Collections.Generic.List[string]
$records = New-Object Collections.Generic.List[object]
$blockCount = 0
$current = $null
for ($n = 0; $n -lt $lines.Count; $n++) {
    $line = $lines[$n]
    if (Test-RecordLine $line) {
        $current = [ordered]@{
            Code = Get-Column $line $cCode; Omschrijving = Get-Column $line $cOmsch; Fix = Get-Column $line $cFix
            Prompt = Get-Column $line $cPrompt; Gegsoort = Get-Column $line $cGsoort; Typ = Get-Column $line $cTyp
            Len = Get-Column $line $cLen; Dec = Get-Column $line $cDec
            BlokIndMut = ''; Referentie = ''; RefTabel = ''; Minimum = ''; Maximum = ''; Bereik = ''
            Line = $n + 1; Blocks = 0; Wrapped = $false
        }
        $flag = Get-Column $line $cBlokInd
        # NOTE: older standard rubrieken have blank flags (15 in the 2026-10-30 export, e.g. P00633). Blank is
        # kept blank, not turned into N: the report does not say what it means.
        if ($flag -notmatch '^[JN]?$') { $problems.Add("line $($n + 1): $($current.Code): Blok ind mut is '$flag', expected J, N or blank") }
        $current.BlokIndMut = $flag
        if ($current.Code -notmatch '^\S{1,7}$') { $problems.Add("line $($n + 1): unreadable code '$($current.Code)'") }
        if ($current.Len -notmatch '^\d*$' -or $current.Dec -notmatch '^\d*$') {
            $problems.Add("line $($n + 1): $($current.Code): Len '$($current.Len)' or Dec '$($current.Dec)' is not a number")
        }
        $records.Add([pscustomobject]$current)
        $current = $records[$records.Count - 1]
        continue
    }
    if ($null -eq $current) { continue }
    # Wrapped Omschrijving: the first line after the record, indented to the Omschrijving column.
    if (-not $current.Wrapped -and $current.Blocks -eq 0 -and $line -match '^\s{9}\S' -and $line -notmatch '^\s+(Ingedeeld|Tabel|Veld)\s') {
        $current.Wrapped = $true
        $current.Omschrijving = ($current.Omschrijving + ' ' + $line.Trim()).Trim()
        continue
    }
    if ($line -match '^\s+Ingedeeld\s+(?:[JN]\s+)?Referentie\s+(?:([JN])\s+)?Min/Max\s+(?:[JN]\s+)?Salarisaanlev') {
        $current.Blocks++; $blockCount++
        $current.Referentie = $Matches[1]
        # The next lines hold Tabel/Veld/Logging/Groep code/Exportkenmerken; the labels are fixed text.
        $b = @($lines[($n + 1)..([Math]::Min($n + 5, $lines.Count - 1))])
        if ($b[0] -match '^\s+Tabel\b.*?\bTabel(?:\s+(\S+)(?=\s+Minimum))?') { if ($Matches[1]) { $current.RefTabel = $Matches[1] } }
        else { $problems.Add("line $($n + 2): $($current.Code): no Tabel line in the detail block") }
        if ($b[0] -match 'Minimum\s+(\S*?)\s+TWK vanaf') { $current.Minimum = $Matches[1] }
        if ($b[1] -match 'Maximum\s+(\S*?)\s+DIW') { $current.Maximum = $Matches[1] }
        if ($b[2] -match 'Bereik\s+(\S*?)\s+TWK huidig') { $current.Bereik = $Matches[1] }
        continue
    }
}
foreach ($r in $records) {
    if ($r.Blocks -ne 1) { $problems.Add("line $($r.Line): $($r.Code): $($r.Blocks) detail blocks, expected 1") }
}
if ($records.Count -eq 0) { Stop-WithError 'No rubriek records found.' }
$dups = @($records | Group-Object Code | Where-Object { $_.Count -gt 1 } | ForEach-Object { "duplicate code $($_.Name) ($($_.Count) times)" })
if ($dups.Count -gt 0) { $problems.AddRange([string[]]$dups) }
if ($blockCount -ne $records.Count) { $problems.Insert(0, "$($records.Count) records but $blockCount detail blocks") }
if ($problems.Count -gt 0) { Stop-WithError "The report does not have the expected shape ($($problems.Count) problems); no output written." $problems }

# The report may have been run with selection criteria; the result is then a part of the rubrieken.
$criteria = @($lines | Where-Object { $_ -match '^\s+\|\s+(Rubriek code|Soort rubriek|Gegevenssoort)\s+:\s*[^\s|]' })
if ($criteria.Count -gt 0) { Write-Host "[WARN] The report was run with selection criteria, so it may be incomplete:" -ForegroundColor Yellow; $criteria | ForEach-Object { Write-Host "       $($_.Trim(' |'))" } }

# --- Summary and output -------------------------------------------------------------------------------
Write-Host "Read $($records.Count) rubrieken from $($headerLines.Count) pages."
Write-Host ('  Fix:      ' + (($records | Group-Object Fix | Sort-Object Name -Descending | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '  ') + '   (J standard, N eigen)')
Write-Host ('  Gegsoort: ' + (($records | Group-Object Gegsoort | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '  ') + '  ...')
$scope = @($records | Where-Object { $_.Gegsoort -in 'PS', 'DV' })
Write-Host "  PS+DV:    $($scope.Count), of which Typ $((($scope | Group-Object Typ | Sort-Object Count -Descending | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' ')); Blok ind mut J=$(@($scope | Where-Object BlokIndMut -eq 'J').Count) N=$(@($scope | Where-Object BlokIndMut -eq 'N').Count) blank=$(@($scope | Where-Object BlokIndMut -eq '').Count)"

$out = $records
if ($OnlyScope) { $out = @($out | Where-Object { $_.Gegsoort -in 'PS', 'DV' }) }
if ($OnlyStandard) { $out = @($out | Where-Object { $_.Fix -eq 'J' }) }

function ConvertTo-CsvField([string]$Value) {
    if ($Value -match '[;"\r\n]') { return '"' + $Value.Replace('"', '""') + '"' }
    return $Value
}
$fields = 'Code', 'Omschrijving', 'Fix', 'Prompt', 'Gegsoort', 'Typ', 'Len', 'Dec', 'BlokIndMut', 'Referentie', 'RefTabel', 'Minimum', 'Maximum', 'Bereik'
$sb = New-Object Text.StringBuilder
[void]$sb.Append(($fields -join ';') + "`r`n")
foreach ($r in ($out | Sort-Object Code)) {
    [void]$sb.Append((($fields | ForEach-Object { ConvertTo-CsvField ([string]$r.$_) }) -join ';') + "`r`n")
}
[IO.File]::WriteAllText($OutputPath, $sb.ToString(), (New-Object Text.UTF8Encoding($true)))
Write-Host "[SUCCESS] $($out.Count) rubrieken written to $OutputPath"
