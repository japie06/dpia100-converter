<#
.SYNOPSIS
    Differential test: random CSV files through dpia100.ps1 and through the core of dpia100.html, which
    must give the same result (decision 15).
.DESCRIPTION
    A seeded generator writes CSV files with a mix of valid and broken content: every input encoding and
    delimiter, renamed, shuffled, duplicate or missing headers, quoted cells (also with the delimiter, a
    quote or a line break inside, spaces before the quote, text after it), blank, whitespace-only and
    ';;;' rows, a '#N/B' first row, NBSP, decomposed accents, characters outside Windows-1252, every amount
    shape of decision 8, dates in all formats and values over 50 characters. A lone CR is left out: there
    PowerShell 5.1 and 7 differ from each other (the page follows 5.1).

    Each file runs through dpia100.ps1 (one process per file) and, all together, through the web harness
    in one headless Edge page. Compared per file: the exit code, the output bytes, the console lines with
    the encoding, delimiter, column mapping, [OK], [WARNINGS], each warning, [FATAL ERROR], [ERROR],
    Records parsed and Empty rows skipped, and the row/column of every validation error. A mismatch prints
    both sides; the CSV stays in tests\out\diff for a rerun (the same seed gives the same files).
.EXAMPLE
    pwsh tests\web\differential.ps1
    pwsh tests\web\differential.ps1 -Count 300 -Seed 7 -Engine powershell
#>
param(
    [int]$Count = 40,
    [int]$Seed = 1,
    [ValidateSet('pwsh', 'powershell')][string]$Engine = 'pwsh'
)

. (Join-Path $PSScriptRoot 'WebTest.ps1')
$script = Join-Path $script:RepoRoot 'dpia100.ps1'
$outDir = Join-Path $script:RepoRoot 'tests\out\diff'
Remove-Item $outDir -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $outDir | Out-Null
$rng = New-Object Random($Seed)
$answers = @('CMT', '07 10 2026', 'TSTUSER', 'Y')

function Pick([object[]]$Items) { return $Items[$rng.Next($Items.Count)] }
function Chance([double]$P) { return $rng.NextDouble() -lt $P }

# --- value pools (valid ones are used for clean files; the full pools for broken files) -------------
$long = 'Senior beleidsmedewerker ruimtelijke ordening en milieu'
$pools = @{
    Rubriek    = @{ ok = @('E00101', 'E00205', 'e00310', ' p00001 ', 'ABC123', 'P00332'); bad = @('E001011', 'E00-01', '', '#N/B', 'E0010ß', 'E 001', ' P1 ', 'E0010') }
    PersNr     = @{ ok = @('123456', '10', '000123', ' 77 ', '2345', "12$([char]0xA0)", '999999'); bad = @('1234.0', '1 234', '', '0', '1234567', '12abc', "$([char]0xFF11)2") }
    Dienstv    = @{ ok = @('', '1', '001', '0', '999', ' 2 '); bad = @('1234', '1a', '1.0', '-1') }
    SofV       = @{ ok = @('S', 'V', 's', ' v ', 'STAM', 'Variabel'); bad = @('X', '', 'Stamm', 'SV') }
    Waarde     = @{ ok = @('Tekst', 'Café de Brug', '12,50', '-12,50', '1,234', '3.250', '1.234,56', '1,234.56', '1.234.567', '12.50',
                           '0.250', '- 7,5', '- 7', '007', '', ' 15 ', '1.11.2026', 'Gebiedscoördinator', '€ 250', "d$([char]0x2019)Ancona", 'a;b',
                           'He said "hi"', $long, "o$([char]0x308)k", "$([char]0xA0)x$([char]0xA0)", '100.000', '12.345', '3250,50', '12,0', 'a,b',
                           '-0.5', "$([char]0x201C)quote$([char]0x201D)", 'Ünïcödé façade', '01-11-2026', '1/11/26', '2026-11-01', '8:30',
                           '08:30', '08:30:00', '8.30')
                     bad = @("$([char]0x133)sbaan", "ok $([char]::ConvertFromUtf32(0x1F600))", "regel1`nregel2", "a`tb", "$([char]0x661)$([char]0x662)", "x$([char]0x2192)y") }
    Ingang     = @{ ok = @('1-11-2026', '1 11 2026', '01112026', '2026-11-01', '1/11/2026', '1.11.2026', '29-02-2028', ' 1 11 2026 ', '1  11 2026', '1-11-0001', '31-12-9999')
                     bad = @('31-02-2026', '', '1-11-26', '2026-1-01', '00-00-0000', '29-02-2027', '1-13-2026', '1-11-2026x') }
    Productie  = @{ ok = @('NOR', 'nor', '', ' NOR ', 'COR', ' cor '); bad = @('JWP', 'ABC', 'NO') }
}
$headerVariants = @{
    Rubriek   = @('Rubriek', 'rubriek', ' Rubriek ', 'Rubriekcode')
    PersNr    = @('Pers. Nr', 'Persoonsnummer', 'PersNr', 'Pers_nr')
    Dienstv   = @('Dienstv', 'Dienstverband', 'Dv', 'DV')
    SofV      = @('S of V', 'Stam', 's of v', 'Stam/Variabel')
    Waarde    = @('Waarde', 'Mutatiewaarde', 'Mutatie waarde', 'Waarde (bedrag, tekst, datum)')
    Ingang    = @('Ingangsdatum', 'Ingangs datum', "Ingangs`r`ndatum")
    Productie = @('Productiesoort', "Productie`r`nsoort", 'Productie', 'Productiecode')
}
$fieldOrder = @('Rubriek', 'PersNr', 'Dienstv', 'SofV', 'Waarde', 'Ingang', 'Productie')
$poolKey = @{ Rubriek = 'Rubriek'; PersNr = 'PersNr'; Dienstv = 'Dienstv'; SofV = 'SofV'; Waarde = 'Waarde'; Ingang = 'Ingang'; Productie = 'Productie' }

# A CSV cell: quoted when needed (or at random), sometimes with the quirks ConvertFrom-Csv tolerates.
function Format-Cell([string]$Value, [string]$Delim) {
    $mustQuote = $Value.Contains($Delim) -or $Value.Contains('"') -or $Value -match "[`r`n]"
    if ($mustQuote -or (Chance 0.15)) {
        $q = '"' + $Value.Replace('"', '""') + '"'
        if (Chance 0.05) { $q = '  ' + $q }
        if (-not $mustQuote -and (Chance 0.05)) { $q = $q + 'z' }
        return $q
    }
    return $Value
}

function New-Csv([int]$Index) {
    $delim = Pick @(';', ';', ';', ',', ',', "`t")
    $nl = if (Chance 0.8) { "`r`n" } else { "`n" }
    $broken = Chance 0.5
    $fields = [System.Collections.Generic.List[string]]::new()
    foreach ($f in $fieldOrder) { $fields.Add($f) }
    if ($broken -and (Chance 0.08)) { $fields.RemoveAt($rng.Next($fields.Count)) }
    $extra = @(Pick @(@(), @(), @('Opmerking'), @(''), @('Opmerking', 'Opmerking'), @('Advies'), @('Persoonswaarde')))
    if (-not $broken) { $extra = @($extra | Where-Object { $_ -ne 'Persoonswaarde' }) }
    $columns = @($fields) + @($extra | ForEach-Object { "#extra:$_" })
    if (Chance 0.5) { $columns = @($columns | Sort-Object { $rng.Next() }) }

    $header = foreach ($c in $columns) {
        $name = if ($c.StartsWith('#extra:')) { $c.Substring(7) } else { Pick $headerVariants[$c] }
        Format-Cell $name $delim
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($header -join $delim) + $(if (Chance 0.1) { $delim + $delim } else { '' }))
    $rows = $rng.Next(1, 8)
    for ($r = 0; $r -lt $rows; $r++) {
        if (Chance 0.06) { $lines.Add('') }
        if (Chance 0.04) { $lines.Add('   ') }
        if (Chance 0.06) { $lines.Add(($columns | ForEach-Object { '' }) -join $delim) }
        $cells = foreach ($c in $columns) {
            if ($c.StartsWith('#extra:')) { Format-Cell (Pick @('', 'x', 'opmerking; met puntkomma', '12,5')) $delim; continue }
            $pool = $pools[$poolKey[$c]]
            $value = if ($broken -and (Chance 0.12)) { Pick $pool.bad } else { Pick $pool.ok }
            if ($r -eq 0 -and $c -eq 'Rubriek' -and (Chance 0.04)) { $value = '#N/B' }
            if ($delim -eq "`t" -and $value.Contains("`t")) { $value = $value.Replace("`t", ' ') }
            Format-Cell $value $delim
        }
        $lines.Add($cells -join $delim)
    }
    $text = ($lines -join $nl) + $(if (Chance 0.85) { $nl } else { '' })

    # Encoding; characters outside Windows-1252 force UTF-8.
    $encName = Pick @('ansi', 'ansi', 'ansi', 'utf8', 'utf8', 'utf8bom', 'utf8bom', 'utf16le', 'utf16be')
    $strict = [Text.Encoding]::GetEncoding(1252, [Text.EncoderExceptionFallback]::new(), [Text.DecoderExceptionFallback]::new())
    if ($encName -eq 'ansi') { try { [void]$strict.GetBytes($text) } catch { $encName = 'utf8' } }
    $bytes = switch ($encName) {
        'ansi'    { $strict.GetBytes($text) }
        'utf8'    { (New-Object Text.UTF8Encoding($false)).GetBytes($text) }
        'utf8bom' { [byte[]](0xEF, 0xBB, 0xBF) + (New-Object Text.UTF8Encoding($false)).GetBytes($text) }
        'utf16le' { [byte[]](0xFF, 0xFE) + [Text.Encoding]::Unicode.GetBytes($text) }
        'utf16be' { [byte[]](0xFE, 0xFF) + [Text.Encoding]::BigEndianUnicode.GetBytes($text) }
    }
    $path = Join-Path $outDir ('{0:D4}.csv' -f $Index)
    [IO.File]::WriteAllBytes($path, [byte[]]$bytes)
    return $path
}

# The console lines both sides must share, and the row/column keys of the validation errors.
$fieldLabel = 'Rubriek|Pers\. Nr|Dienstverband|S of V|Mutatiewaarde|Ingangsdatum|Productiesoort'
function Get-Summary([string]$Console, [string]$CsvPath) {
    $text = $Console.Replace($CsvPath, [IO.Path]::GetFileName($CsvPath))
    $lines = @($text -split '\r?\n')
    $keep = @($lines | Where-Object {
        $_ -match '^(Input encoding|Detected delimiter|\[OK\]|\[WARNINGS\]|\[FATAL ERROR\]|\[ERROR\]|Comma changed|  - |    Original|    Truncated|Records parsed|Empty rows skipped|Written with)' -or
        $_ -match "^  ($fieldLabel)\s+<- "
    } | ForEach-Object { $_.TrimEnd() })
    $errors = @($lines | Where-Object { $_ -match "^\s*(\d+)\s+($fieldLabel)\s" } | ForEach-Object { "$($Matches[1]):$($Matches[2])" } | Sort-Object)
    if ($text -notmatch '\[FATAL ERROR\]') { $errors = @() }
    return [PSCustomObject]@{ Lines = $keep; Errors = $errors }
}

Write-Host "Generating $Count CSV files (seed $Seed) in $outDir"
$files = for ($i = 1; $i -le $Count; $i++) { New-Csv $i }

Write-Host "Running dpia100.ps1 on $Engine ..."
$prevEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$psResults = @{}
try {
    foreach ($csv in $files) {
        $outFile = Join-Path $outDir ('DPIA100_' + [IO.Path]::GetFileNameWithoutExtension($csv) + '.txt')
        $cmd = "[Console]::OutputEncoding = New-Object Text.UTF8Encoding(`$false); & '$script' -InputCsvPath '$csv' -OutputTxtPath '$outFile' -NoPause; exit `$LASTEXITCODE"
        $console = $answers | & $Engine -NoProfile -ExecutionPolicy Bypass -Command $cmd 2>&1 | Out-String
        $psResults[$csv] = [PSCustomObject]@{
            ExitCode = $LASTEXITCODE
            Output   = if (Test-Path $outFile) { [IO.File]::ReadAllBytes($outFile) } else { $null }
            Console  = $console
        }
    }
} finally {
    [Console]::OutputEncoding = $prevEncoding
}

Write-Host 'Running the web core in Edge ...'
$webCases = foreach ($csv in $files) {
    [PSCustomObject]@{ name = $csv; input = [Convert]::ToBase64String([IO.File]::ReadAllBytes($csv)); fileName = [IO.Path]::GetFileName($csv); answers = $answers; outputName = $null }
}
$web = @{}
foreach ($r in (Invoke-WebHarness -Cases @($webCases) -OutDir $outDir).cases) { $web[$r.name] = $r }

$mismatches = 0
$succeeded = 0
foreach ($csv in $files) {
    $ps = $psResults[$csv]
    $w = $web[$csv]
    $problems = [System.Collections.Generic.List[string]]::new()
    if ($ps.ExitCode -ne $w.exitCode) { $problems.Add("exit code: PowerShell $($ps.ExitCode), web $($w.exitCode)") }
    $wOut = if ($w.output) { [Convert]::FromBase64String($w.output) } else { $null }
    if (($null -eq $ps.Output) -ne ($null -eq $wOut)) {
        $problems.Add("output file: PowerShell $(if ($ps.Output) { 'yes' } else { 'no' }), web $(if ($wOut) { 'yes' } else { 'no' })")
    } elseif ($ps.Output -and [Convert]::ToBase64String($ps.Output) -ne $w.output) {
        $problems.Add('output bytes differ')
    }
    $a = Get-Summary $ps.Console $csv
    $b = Get-Summary $w.transcript $csv
    $lineDiff = Compare-Object $a.Lines $b.Lines -SyncWindow 1000
    foreach ($d in $lineDiff) { $problems.Add("line only in $(if ($d.SideIndicator -eq '<=') { 'PowerShell' } else { 'web' }): $($d.InputObject)") }
    if (($a.Errors -join ',') -ne ($b.Errors -join ',')) { $problems.Add("validation errors: PowerShell [$($a.Errors -join ', ')], web [$($b.Errors -join ', ')]") }
    if ($ps.ExitCode -eq 0) { $succeeded++ }
    if ($problems.Count -gt 0) {
        $mismatches++
        Write-Host "MISMATCH  $csv" -ForegroundColor Red
        foreach ($p in $problems) { Write-Host "    - $p" -ForegroundColor Red }
    }
}

Write-Host ''
Write-Host "$Count files, $succeeded made an output file, $($Count - $succeeded) stopped with an error."
if ($mismatches -gt 0) {
    Write-Host "$mismatches of $Count files differ between dpia100.ps1 ($Engine) and dpia100.html." -ForegroundColor Red
    exit 1
}
Write-Host "dpia100.ps1 ($Engine) and dpia100.html agree on all $Count files." -ForegroundColor Green
exit 0
