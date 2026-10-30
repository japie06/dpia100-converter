<#
.SYNOPSIS
    Regression tests for dpia100.ps1 and dpia100.html. Runs every case in tests\cases on Windows
    PowerShell 5.1, PowerShell 7 (when installed) and the web page in headless Edge.
.DESCRIPTION
    Engines: 'powershell' and 'pwsh' run dpia100.ps1 with the answers piped in. 'edge' runs all cases in
    one headless Edge page with the core script of dpia100.html (tests\web\harness.js), which returns
    the exit code, the output bytes and the same console lines as PowerShell. For 'edge' the console
    lines about file paths (Input file, Output file, the path after [SUCCESS], Encoding) are not checked,
    and there is no stale file to rename: the page never touches existing files. With 'edge' the page
    checks also run (page-*: CSP hashes, forbidden APIs, version, 1252 table, smoke test). The UI is tested
    with Playwright (npm test, tests/web/ui.spec.mjs).

    Each folder in tests\cases is one case:
      input.csv            The CSV to convert (exact bytes, any encoding).
      expected.txt         Success case: the exact expected output bytes.
                           No expected.txt = failure case: exit code must be 1, no output file may be
                           written, and an existing output file must be renamed to .old (decision 10).
      expected-output.txt  Optional: lines that must appear in the console output (errors, warnings),
                           in this order.
      unexpected-output.txt Optional: whole console lines that must not appear, e.g. 'Y': the echo of
                           an answer that must not have been read.
      answers.txt          Optional: answers to the prompts, one per line. A failure case without it
                           must stop before the prompts (decision 14): no answer echo, no [OK].
                           Default: CMT / 07 10 2026 / TSTUSER / Y (the Y confirms warnings, if any).
      output-name.txt      Optional: output file name (default DPIA100_<case>.<engine>.txt). For a
                           failure case with a name that does not start with DPIA100, an existing file
                           of that name must be left untouched instead of renamed to .old.
      default-output       Optional marker (empty file): run without -OutputTxtPath, with input.csv
                           copied to out\<case>.<engine>.csv. The script must then write the default
                           out\DPIA100_<case>.<engine>.txt next to it (decision 13).
      extra-rubrieken.txt  Optional: lines in the format of the standard rubrieken table, added to the
                           table for this case only (a copy of the script, a separate Edge page). For
                           types the real list lacks, such as T and U time rubrieken (decision 20).
    Fixtures are stored as raw bytes (.gitattributes: -text). Edit them with care: an editor that saves
    UTF-8 or LF changes the test.
.EXAMPLE
    pwsh tests\run-tests.ps1
    pwsh tests\run-tests.ps1 -Filter err-*
    pwsh tests\run-tests.ps1 -Engine edge
#>
param(
    [string]$Filter = '*',
    # Default: powershell, pwsh (when installed) and edge.
    [ValidateSet('powershell', 'pwsh', 'edge')][string[]]$Engine
)

. (Join-Path $PSScriptRoot 'web\WebTest.ps1')

$script  = Join-Path $PSScriptRoot '..\dpia100.ps1'
$outDir  = Join-Path $PSScriptRoot 'out'
$cases   = Get-ChildItem (Join-Path $PSScriptRoot 'cases') -Directory -Filter $Filter | Sort-Object Name
$engines = if ($Engine) { $Engine } else { @('powershell') + @(Get-Command pwsh -ErrorAction SilentlyContinue | ForEach-Object { 'pwsh' }) + 'edge' }
$ansi    = [System.Text.Encoding]::GetEncoding(1252)
$defaultAnswers = @('CMT', '07 10 2026', 'TSTUSER', 'Y')
# Console lines that only exist on the console (paths, banner), not checked for the web page.
$consoleOnly = '^(=== |Input file|Output file|Encoding\s+:|\\)'

# Record layout (start position, field), used to name the field where a byte mismatch occurs.
$layout = @(
    @(1, 'Procescode'), @(4, 'Rubriek'), @(10, 'Object_id'), @(60, 'S of V'), @(61, 'Creatiedatum'),
    @(72, 'Creatiegebruiker'), @(88, 'Mutatiewaarde'), @(138, 'Ingangsdatum'), @(149, 'Filler'),
    @(179, 'Productiesoort'), @(182, 'Verrekenperiode'), @(188, 'Line ending')
)

function Get-ByteDiff([byte[]]$expected, [byte[]]$actual) {
    $n = [Math]::Min($expected.Length, $actual.Length)
    $i = 0
    while ($i -lt $n -and $expected[$i] -eq $actual[$i]) { $i++ }
    if ($i -eq $n -and $expected.Length -eq $actual.Length) { return $null }
    # Records are 187 bytes + CRLF
    $record = [Math]::Floor($i / 189) + 1
    $pos    = ($i % 189) + 1
    $field  = ($layout | Where-Object { $_[0] -le $pos } | Select-Object -Last 1)[1]
    $line   = { param($b) $s = ($record - 1) * 189; if ($s -ge $b.Length) { '(missing)' } else { $ansi.GetString($b, $s, [Math]::Min(187, $b.Length - $s)).Replace(' ', [char]0xB7) } }
    return "first difference in record $record, pos $pos ($field)`n      expected: $(& $line $expected)`n      actual  : $(& $line $actual)" +
           "`n      lengths : expected $($expected.Length) bytes, actual $($actual.Length) bytes"
}

New-Item -ItemType Directory -Force $outDir | Out-Null
$failed = 0
$total  = 0

# The extra standard rubrieken of a case (extra-rubrieken.txt), or an empty list.
function Get-ExtraRubrieken($CaseDir) {
    $path = Join-Path $CaseDir 'extra-rubrieken.txt'
    if (Test-Path $path) { return @(Get-Content $path | Where-Object { $_.Trim() }) }
    return @()
}

# A copy of dpia100.ps1 with extra lines in its standard rubrieken table.
function New-ScriptWithRubrieken([string]$Name, [string[]]$Extra) {
    $text = [IO.File]::ReadAllText($script)
    $pattern = "(\`$StandardRubriekenTable = @'\r?\n)"
    if ($text -notmatch $pattern) { throw 'No $StandardRubriekenTable in dpia100.ps1.' }
    $text = [regex]::Replace($text, $pattern, { param($m) $m.Groups[1].Value + (($Extra | ForEach-Object { "$_`r`n" }) -join '') })
    $copy = Join-Path $outDir "dpia100-$Name.ps1"
    [IO.File]::WriteAllText($copy, $text, (New-Object Text.UTF8Encoding($true)))
    return $copy
}

function Write-Result([string]$Name, [string]$EngineName, [string[]]$Problems, [string]$Console) {
    if ($Problems.Count -eq 0) {
        Write-Host ("PASS  {0,-32} {1}" -f $Name, $EngineName) -ForegroundColor Green
        return $true
    }
    Write-Host ("FAIL  {0,-32} {1}" -f $Name, $EngineName) -ForegroundColor Red
    foreach ($p in $Problems) { Write-Host "    - $p" -ForegroundColor Red }
    if ($Console) {
        Write-Host "    Console output:" -ForegroundColor DarkGray
        Write-Host (($Console.TrimEnd() -split '\r?\n' | ForEach-Object { "      $_" }) -join "`n") -ForegroundColor DarkGray
    }
    return $false
}

# The web page runs all cases in one Edge launch; the results are checked per case in the loop below.
$webResults = @{}
$webPayload = $null
if ($engines -contains 'edge') {
    if (-not (Find-Edge)) {
        Write-Host "Microsoft Edge was not found; run with -Engine powershell,pwsh to skip the web page." -ForegroundColor Red
        exit 1
    }
    $webCases = foreach ($case in $cases) {
        $answersPath = Join-Path $case.FullName 'answers.txt'
        $namePath    = Join-Path $case.FullName 'output-name.txt'
        [PSCustomObject]@{
            name       = $case.Name
            input      = [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $case.FullName 'input.csv')))
            fileName   = if (Test-Path (Join-Path $case.FullName 'default-output')) { "$($case.Name).edge.csv" } else { 'input.csv' }
            answers    = if (Test-Path $answersPath) { @(Get-Content $answersPath) } else { $defaultAnswers }
            outputName = if (Test-Path $namePath) { (Get-Content $namePath -TotalCount 1).Trim() } else { $null }
            extra      = Get-ExtraRubrieken $case.FullName
        }
    }
    try {
        # One Edge page per distinct table: the cases without extra rubrieken first, all in one page.
        $groups = @($webCases | Group-Object { $_.extra -join "`n" } | Sort-Object { $_.Name -ne '' }, Name)
        for ($i = 0; $i -lt $groups.Count; $i++) {
            $payload = Invoke-WebHarness -Cases @($groups[$i].Group) -OutDir $outDir -ExtraRubrieken @($groups[$i].Group[0].extra) -PageName "web-harness-$i.html"
            if ($i -eq 0) { $webPayload = $payload }
            foreach ($r in $payload.cases) { $webResults[$r.name] = $r }
        }
    } catch {
        Write-Host "FAIL  web harness: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

foreach ($case in $cases) {
    $expectedPath = Join-Path $case.FullName 'expected.txt'
    $mustSucceed  = Test-Path $expectedPath
    $answersPath  = Join-Path $case.FullName 'answers.txt'
    $answers      = if (Test-Path $answersPath) { @(Get-Content $answersPath) } else { $defaultAnswers }
    $outputPath   = Join-Path $case.FullName 'expected-output.txt'
    $mustContain  = if (Test-Path $outputPath) { @(Get-Content $outputPath -Encoding UTF8 | Where-Object { $_ }) } else { @() }
    $absentPath   = Join-Path $case.FullName 'unexpected-output.txt'
    $mustNotShow  = if (Test-Path $absentPath) { @(Get-Content $absentPath -Encoding UTF8 | Where-Object { $_ }) } else { @() }

    $namePath     = Join-Path $case.FullName 'output-name.txt'
    $outName      = if (Test-Path $namePath) { (Get-Content $namePath -TotalCount 1).Trim() } else { $null }
    $foreignName  = $outName -and -not $outName.StartsWith('DPIA100', [StringComparison]::OrdinalIgnoreCase)
    $useDefault   = Test-Path (Join-Path $case.FullName 'default-output')
    $extra        = Get-ExtraRubrieken $case.FullName
    $caseScript   = if ($extra.Count -gt 0 -and ($engines | Where-Object { $_ -ne 'edge' })) { New-ScriptWithRubrieken $case.Name $extra } else { $script }

    foreach ($engine in $engines) {
        $total++
        $problems = [System.Collections.Generic.List[string]]::new()
        if ($engine -eq 'edge') {
            # The page writes nothing until the user saves, so there is no stale file to check (decision 15).
            $web     = $webResults[$case.Name]
            $console = "$($web.transcript)"
            if ($mustSucceed) {
                if ($web.exitCode -ne 0) { $problems.Add("exit code $($web.exitCode), expected 0") }
                if (-not $web.output) {
                    $problems.Add('no output file made')
                } else {
                    $diff = Get-ByteDiff ([IO.File]::ReadAllBytes($expectedPath)) ([Convert]::FromBase64String($web.output))
                    if ($diff) { $problems.Add($diff) }
                }
            } else {
                if ($web.exitCode -ne 1) { $problems.Add("exit code $($web.exitCode), expected 1") }
                if ($web.output) { $problems.Add('an output file was made although the run failed') }
            }
            if ($useDefault -and $web.suggestedName -ne "DPIA100_$($case.Name).edge.txt") {
                $problems.Add("suggested file name '$($web.suggestedName)', expected 'DPIA100_$($case.Name).edge.txt'")
            }
            $expectLines = @($mustContain | Where-Object { $_ -notmatch $consoleOnly })
        } else {
            $outFile = Join-Path $outDir $(if ($outName) { $outName } else { "DPIA100_$($case.Name).$engine.txt" })
            Remove-Item "$outFile*" -Force -ErrorAction SilentlyContinue
            if (-not $mustSucceed) { Set-Content $outFile 'stale output from an earlier run' }

            $inputCsv = Join-Path $case.FullName 'input.csv'
            $outArgs  = @('-OutputTxtPath', $outFile)
            if ($useDefault) {
                # The default output name is derived from the CSV name, so the copy is named to match $outFile.
                $copy = Join-Path $outDir "$($case.Name).$engine.csv"
                Copy-Item -LiteralPath $inputCsv -Destination $copy -Force
                $inputCsv = $copy
                $outArgs  = @()
            }
            $console  = $answers | & $engine -NoProfile -ExecutionPolicy Bypass -File $caseScript -InputCsvPath $inputCsv @outArgs -NoPause 2>&1 | Out-String
            $exitCode = $LASTEXITCODE

            if ($mustSucceed) {
                if ($exitCode -ne 0) { $problems.Add("exit code $exitCode, expected 0") }
                if (-not (Test-Path $outFile)) {
                    $problems.Add('no output file written')
                } else {
                    $diff = Get-ByteDiff ([IO.File]::ReadAllBytes($expectedPath)) ([IO.File]::ReadAllBytes($outFile))
                    if ($diff) { $problems.Add($diff) }
                }
            } elseif ($foreignName) {
                if ($exitCode -ne 1) { $problems.Add("exit code $exitCode, expected 1") }
                if (-not (Test-Path $outFile) -or (Test-Path "$outFile.old")) { $problems.Add('an existing file with a non-DPIA100 name was touched') }
            } else {
                if ($exitCode -ne 1) { $problems.Add("exit code $exitCode, expected 1") }
                if (Test-Path $outFile) { $problems.Add('an output file exists after a failed run') }
                if (-not (Test-Path "$outFile.old")) { $problems.Add('the stale output file was not renamed to .old') }
            }
            $expectLines = $mustContain
        }
        # The lines must appear in this order, so the files also check the order of operations (decision 14).
        $searchFrom = 0
        foreach ($text in $expectLines) {
            $at = $console.IndexOf($text, $searchFrom, [StringComparison]::Ordinal)
            if ($at -lt 0) {
                $why = if ($console.Contains($text)) { 'appears, but too early' } else { 'does not appear' }
                $problems.Add("console output line $why (expected in this order): $text")
            } else {
                $searchFrom = $at + $text.Length
            }
        }
        $consoleLines = @($console -split '\r?\n' | ForEach-Object { $_.Trim() })
        foreach ($text in $mustNotShow) {
            if ($consoleLines -contains $text.Trim()) { $problems.Add("console output has a line it must not have: $text") }
        }
        # Decision 14: a run that fails on the CSV stops before the prompts. Prompt text is not captured
        # when stdin is piped, but each answer read is echoed, so the default answer TSTUSER must not
        # appear. Cases that do reach the prompts (cancelled warnings) have their own answers.txt.
        # The ok-order-* cases expect the echo, so this check cannot pass silently if the echo disappears.
        if (-not $mustSucceed -and -not (Test-Path $answersPath)) {
            if ($console -match '(?m)^TSTUSER\s*$') { $problems.Add('the prompts were asked although the CSV failed validation') }
            if ($console.Contains('[OK]')) { $problems.Add('[OK] was shown although the run failed') }
        }

        if (-not (Write-Result $case.Name $engine $problems $console)) { $failed++ }
    }
}

# Checks of the page itself (only with the edge engine; -Filter page-* runs only these).
if ($engines -contains 'edge') {
    $pageChecks = [ordered]@{
        # CSP equals the canonical policy with the current hashes; no external resources; no network,
        # navigation or code-from-string APIs in the scripts (docs/privacy-en-beveiliging.md).
        'page-security' = { Test-PageSecurity }
        'page-version' = {
            $ps = [regex]::Match([IO.File]::ReadAllText($script), "\`$ScriptVersion = '([^']+)'").Groups[1].Value
            if ($webPayload.version -ne $ps) { "dpia100.html has version '$($webPayload.version)', dpia100.ps1 has '$ps'" }
        }
        # The standard rubrieken table (decision 17) is the same text in the script and the page, with the
        # same export date, and the page parsed every line of it.
        'page-rubrieken' = {
            $ps = [IO.File]::ReadAllText($script)
            $psDate = [regex]::Match($ps, "\`$StandardRubriekenDate = '([^']*)'").Groups[1].Value
            $psTable = @([regex]::Match($ps, "(?s)\`$StandardRubriekenTable = @'\r?\n(.*?)\r?\n'@").Groups[1].Value -split '\r?\n')
            $core = (Get-PageParts).Blocks | Where-Object Id -eq 'dpia100-core'
            $jsDate = [regex]::Match($core.Text, "const STANDARD_RUBRIEKEN_DATE = '([^']*)';").Groups[1].Value
            $jsTable = @([regex]::Match($core.Text, '(?s)const STANDARD_RUBRIEKEN_TABLE = String\.raw`\n(.*?)\n`;').Groups[1].Value -split '\n')
            if ($psDate -notmatch '^\d{2}-\d{2}-\d{4}$') { "dpia100.ps1 has no export date ('$psDate')" }
            if ($jsDate -ne $psDate) { "export date: dpia100.html '$jsDate', dpia100.ps1 '$psDate'" }
            if ($psTable.Count -lt 1000) { "dpia100.ps1 has only $($psTable.Count) standard rubrieken" }
            $diff = @(Compare-Object $psTable $jsTable -SyncWindow 0 | Select-Object -First 3)
            if ($diff.Count -gt 0 -or $psTable.Count -ne $jsTable.Count) {
                "tables differ (ps1 $($psTable.Count) lines, html $($jsTable.Count) lines): " + (($diff | ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" }) -join '; ') + '. Run tools\update-standard-rubrieken.ps1.'
            }
            if ($webPayload.standardRubrieken -ne $jsTable.Count) { "the page parsed $($webPayload.standardRubrieken) of $($jsTable.Count) lines" }
            if ($webPayload.standardDate -ne $jsDate) { "the page reports export date '$($webPayload.standardDate)'" }
        }
        # The strict encoder of the page accepts exactly the characters the strict .NET encoder accepts.
        'page-1252-table' = {
            $strict = [Text.Encoding]::GetEncoding(1252, [Text.EncoderExceptionFallback]::new(), [Text.DecoderExceptionFallback]::new())
            $net = foreach ($ch in [Text.Encoding]::GetEncoding(1252).GetString([byte[]](0..255)).ToCharArray()) {
                try { [void]$strict.GetBytes([string]$ch); [int]$ch } catch { }
            }
            $diff = Compare-Object @($net | Sort-Object) @($webPayload.encodable)
            if ($diff) { "encodable characters differ: " + (($diff | ForEach-Object { 'U+{0:X4} ({1})' -f $_.InputObject, $(if ($_.SideIndicator -eq '<=') { 'only .NET' } else { 'only page' }) }) -join ', ') }
        }
        # The real page starts under its own CSP, with CRLF and with LF line endings.
        'page-smoke-crlf' = {
            $copy = Join-Path $outDir 'dpia100-crlf.html'
            [IO.File]::WriteAllText($copy, ([IO.File]::ReadAllText($script:PagePath) -replace "`r?`n", "`r`n"), (New-Object Text.UTF8Encoding($false)))
            Test-PageStarts $copy
        }
        'page-smoke-lf' = {
            $copy = Join-Path $outDir 'dpia100-lf.html'
            [IO.File]::WriteAllText($copy, ([IO.File]::ReadAllText($script:PagePath) -replace "`r`n", "`n"), (New-Object Text.UTF8Encoding($false)))
            Test-PageStarts $copy
        }
    }
    foreach ($name in $pageChecks.Keys) {
        if ($name -notlike $Filter) { continue }
        $total++
        $problems = @(try { & $pageChecks[$name] } catch { "check failed: $($_.Exception.Message)" }) | Where-Object { $_ }
        if (-not (Write-Result $name 'edge' $problems $null)) { $failed++ }
    }
}

Write-Host ''
if ($total -eq 0) {
    Write-Host "No test cases match '$Filter'." -ForegroundColor Yellow
    exit 1
}
if ($failed -gt 0) {
    Write-Host "$failed of $total test run(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "All $total test run(s) passed." -ForegroundColor Green
exit 0
