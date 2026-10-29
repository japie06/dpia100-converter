<#
.SYNOPSIS
    Regression tests for dpia100.ps1. Runs every case in tests\cases on Windows PowerShell 5.1 and,
    when installed, PowerShell 7.
.DESCRIPTION
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
    Fixtures are stored as raw bytes (.gitattributes: -text). Edit them with care: an editor that saves
    UTF-8 or LF changes the test.
.EXAMPLE
    pwsh tests\run-tests.ps1
    pwsh tests\run-tests.ps1 -Filter err-*
#>
param([string]$Filter = '*')

$script  = Join-Path $PSScriptRoot '..\dpia100.ps1'
$outDir  = Join-Path $PSScriptRoot 'out'
$cases   = Get-ChildItem (Join-Path $PSScriptRoot 'cases') -Directory -Filter $Filter | Sort-Object Name
$engines = @('powershell') + @(Get-Command pwsh -ErrorAction SilentlyContinue | ForEach-Object { 'pwsh' })
$ansi    = [System.Text.Encoding]::GetEncoding(1252)
$defaultAnswers = @('CMT', '07 10 2026', 'TSTUSER', 'Y')

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

    foreach ($engine in $engines) {
        $total++
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
        $console  = $answers | & $engine -NoProfile -ExecutionPolicy Bypass -File $script -InputCsvPath $inputCsv @outArgs -NoPause 2>&1 | Out-String
        $exitCode = $LASTEXITCODE

        $problems = [System.Collections.Generic.List[string]]::new()
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
        # The lines must appear in this order, so the files also check the order of operations (decision 14).
        $searchFrom = 0
        foreach ($text in $mustContain) {
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

        if ($problems.Count -eq 0) {
            Write-Host ("PASS  {0,-32} {1}" -f $case.Name, $engine) -ForegroundColor Green
        } else {
            $failed++
            Write-Host ("FAIL  {0,-32} {1}" -f $case.Name, $engine) -ForegroundColor Red
            foreach ($p in $problems) { Write-Host "    - $p" -ForegroundColor Red }
            Write-Host "    Console output:" -ForegroundColor DarkGray
            Write-Host (($console.TrimEnd() -split '\r?\n' | ForEach-Object { "      $_" }) -join "`n") -ForegroundColor DarkGray
        }
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
