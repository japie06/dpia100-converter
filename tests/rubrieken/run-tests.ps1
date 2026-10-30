<#
.SYNOPSIS
    Tests for tools\convert-rubrieken.ps1 and tools\update-standard-rubrieken.ps1 on a synthetic export
    (tests\rubrieken\sample-export.txt).
.DESCRIPTION
    Runs both tools on Windows PowerShell 5.1 and PowerShell 7 (when installed). The good converter cases
    must produce the expected CSV byte for byte; the bad cases must exit 1 and write no output file.
    The generator writes the standard PS/DV rubrieken into two stub files with the markers of dpia100.ps1
    and dpia100.html; it must keep their encoding and line endings, and change nothing on an error.
    The sample is invented (fake municipality and user) and has the line layout of the real report:
    two pages, a wrapped Omschrijving, Min/Max, a reference table, a blank flag, stam/variabel variants,
    a code shorter than 6, an accent and a ';'. Never put rows of a real export in this folder.
.EXAMPLE
    pwsh tests\rubrieken\run-tests.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent | Split-Path -Parent
$tool = Join-Path $root 'tools\convert-rubrieken.ps1'
$sample = Join-Path $PSScriptRoot 'sample-export.txt'
$work = Join-Path $root 'tests\out\rubrieken'
New-Item -ItemType Directory -Force $work | Out-Null

$engines = @(@{ Name = 'powershell'; Exe = 'powershell.exe' })
if (Get-Command pwsh -ErrorAction SilentlyContinue) { $engines += @{ Name = 'pwsh'; Exe = 'pwsh' } }

$utf8 = New-Object Text.UTF8Encoding($false)
$sampleText = [IO.File]::ReadAllText($sample, $utf8)

# name, input text (null = the sample file itself), extra switches, expected file or $null (must fail)
$colHeader = ' Code    Omschrijving              Fix Prompt               soort   soort Typ Len Dec trm  Stam Var Verv Rubr toek Mut mut'
$secondHeader = $sampleText.IndexOf($colHeader, $sampleText.IndexOf($colHeader) + 1)
$cases = @(
    @{ Name = 'all'; Text = $null; Args = @(); Expected = 'expected-all.csv' }
    @{ Name = 'scope-standard'; Text = $null; Args = @('-OnlyScope', '-OnlyStandard'); Expected = 'expected-scope-standard.csv' }
    @{ Name = 'err-no-column-header'; Text = $sampleText.Replace($colHeader, ''); Args = @(); Expected = $null }
    @{ Name = 'err-header-differs-per-page'; Text = $sampleText.Remove($secondHeader, $colHeader.Length).Insert($secondHeader, $colHeader.Replace(' Dec ', ' Dex ')); Args = @(); Expected = $null }
    @{ Name = 'err-duplicate-code'; Text = $sampleText.Replace(' ZZ0004 ', ' ZZ0003 '); Args = @(); Expected = $null }
    @{ Name = 'err-missing-detail-block'; Text = [regex]::Replace($sampleText, '(?m)^\s+Ingedeeld  N                  Referentie    J[^\r\n]*\r\n', ''); Args = @(); Expected = $null }
    @{ Name = 'err-unknown-flag'; Text = $sampleText.Replace('Voorbeeld tekst                DV   S   50            J   N    J    N   N    J   N', 'Voorbeeld tekst                DV   S   50            J   N    J    N   N    J   X'); Args = @(); Expected = $null }
    @{ Name = 'err-not-utf8'; Text = $null; Bytes = $true; Args = @(); Expected = $null }
)

$failed = 0
foreach ($e in $engines) {
    foreach ($c in $cases) {
        $in = Join-Path $work "$($c.Name).txt"
        $out = Join-Path $work "$($c.Name).$($e.Name).csv"
        Remove-Item $out -ErrorAction SilentlyContinue
        if ($c.Bytes) {
            # Windows-1252 e-acute where UTF-8 is expected
            $b = [IO.File]::ReadAllBytes($sample) + [byte[]](0xE9)
            [IO.File]::WriteAllBytes($in, $b)
        }
        elseif ($null -ne $c.Text) { [IO.File]::WriteAllText($in, $c.Text, $utf8) }
        else { Copy-Item $sample $in -Force }
        $console = & $e.Exe -NoProfile -File $tool -InputPath $in -OutputPath $out @($c.Args) 2>&1 | Out-String
        $code = $LASTEXITCODE
        $problem = $null
        if ($c.Expected) {
            if ($code -ne 0) { $problem = "exit code $code, expected 0" }
            elseif (-not (Test-Path $out)) { $problem = 'no output file' }
            else {
                $a = [IO.File]::ReadAllBytes($out); $x = [IO.File]::ReadAllBytes((Join-Path $PSScriptRoot $c.Expected))
                if (([BitConverter]::ToString($a)) -ne ([BitConverter]::ToString($x))) { $problem = "output differs from $($c.Expected)" }
            }
        }
        else {
            if ($code -ne 1) { $problem = "exit code $code, expected 1" }
            elseif (Test-Path $out) { $problem = 'output file written despite the error' }
            elseif ($console -notmatch '\[ERROR\]') { $problem = 'no [ERROR] line' }
        }
        if ($problem) { $failed++; Write-Host "[FAIL] $($e.Name) $($c.Name): $problem" -ForegroundColor Red; Write-Host $console }
        else { Write-Host "[PASS] $($e.Name) $($c.Name)" -ForegroundColor Green }
    }
}

# --- tools\update-standard-rubrieken.ps1 on stub files ---------------------------------------------------
$generator = Join-Path $root 'tools\update-standard-rubrieken.ps1'
$stubScript = "# before`r`n# BEGIN STANDARD RUBRIEKEN: old`r`n`$StandardRubriekenDate = ''`r`n`$StandardRubriekenTable = @'`r`nOLD    DV N  1 - SV N`r`n'@`r`n# END STANDARD RUBRIEKEN`r`n# after`r`n"
$stubPage = "<script>`n  // BEGIN STANDARD RUBRIEKEN: old`n  const STANDARD_RUBRIEKEN_DATE = '';`n  const STANDARD_RUBRIEKEN_TABLE = String.raw```n``;`n  // END STANDARD RUBRIEKEN`n</script>`n"
# The standard PS/DV rubrieken of the sample: code, gegevenssoort, type, length, decimals, S/V, Blok ind mut.
$table = @('ZZ0003 PS N  1 0 SV J', 'ZZ0004 DV D 11 - S- -', 'ZZ08   DV T  5 - S- N')
$note = 'BEGIN STANDARD RUBRIEKEN: generated by tools/update-standard-rubrieken.ps1 from the Beaufort export of 01-01-2026, do not edit (decision 17)'
$expectScript = "# before`r`n# $note`r`n`$StandardRubriekenDate = '01-01-2026'`r`n`$StandardRubriekenTable = @'`r`n" + (($table | ForEach-Object { "$_`r`n" }) -join '') + "'@`r`n# END STANDARD RUBRIEKEN`r`n# after`r`n"
$expectPage = "<script>`n  // $note`n  const STANDARD_RUBRIEKEN_DATE = '01-01-2026';`n  const STANDARD_RUBRIEKEN_TABLE = String.raw```n" + (($table | ForEach-Object { "$_`n" }) -join '') + "``;`n  // END STANDARD RUBRIEKEN`n</script>`n"
$genCases = @(
    @{ Name = 'generate'; Page = $stubPage; Ok = $true }
    @{ Name = 'err-no-region-in-page'; Page = "<script>`n</script>`n"; Ok = $false }
)
$bom = New-Object Text.UTF8Encoding($true)
foreach ($e in $engines) {
    foreach ($c in $genCases) {
        $ps1 = Join-Path $work "stub.$($e.Name).ps1"
        $html = Join-Path $work "stub.$($e.Name).html"
        [IO.File]::WriteAllText($ps1, $stubScript, $bom)
        [IO.File]::WriteAllText($html, $c.Page, $utf8)
        $console = & $e.Exe -NoProfile -File $generator -InputPath $sample -ScriptPath $ps1 -PagePath $html -SkipCsp 2>&1 | Out-String
        $code = $LASTEXITCODE
        $gotScript = [IO.File]::ReadAllBytes($ps1); $gotPage = [IO.File]::ReadAllBytes($html)
        $problem = $null
        if ($c.Ok) {
            $wantScript = $bom.GetPreamble() + $utf8.GetBytes($expectScript)
            if ($code -ne 0) { $problem = "exit code $code, expected 0" }
            elseif ([BitConverter]::ToString($gotScript) -ne [BitConverter]::ToString($wantScript)) { $problem = "script stub differs:`n$([IO.File]::ReadAllText($ps1))" }
            elseif ([BitConverter]::ToString($gotPage) -ne [BitConverter]::ToString($utf8.GetBytes($expectPage))) { $problem = "page stub differs:`n$([IO.File]::ReadAllText($html))" }
        } else {
            if ($code -ne 1) { $problem = "exit code $code, expected 1" }
            elseif ($console -notmatch '\[ERROR\]') { $problem = 'no [ERROR] line' }
            elseif ([IO.File]::ReadAllText($ps1) -ne $stubScript) { $problem = 'the script stub was changed despite the error' }
        }
        if ($problem) { $failed++; Write-Host "[FAIL] $($e.Name) $($c.Name): $problem" -ForegroundColor Red; Write-Host $console }
        else { Write-Host "[PASS] $($e.Name) $($c.Name)" -ForegroundColor Green }
    }
}
if ($failed) { Write-Host "$failed failed." -ForegroundColor Red; exit 1 }
Write-Host 'All passed.' -ForegroundColor Green
