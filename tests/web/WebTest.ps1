<#
.SYNOPSIS
    Helpers to test dpia100.html in headless Microsoft Edge. Dot-sourced by tests/run-tests.ps1 and
    tests/web/differential.ps1; nothing to install (Edge is on every Windows PC and CI runner).
.DESCRIPTION
    msedge.exe hands its work to a detached browser process, so its stdout (--dump-dom) cannot be
    captured. A test page therefore reports with console.log('<marker>:<base64 JSON>'), which Edge
    writes to chrome_debug.log in a throwaway profile (--enable-logging). --virtual-time-budget lets
    timers and promises run before Edge closes the page.
#>

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).ProviderPath
$script:PagePath = Join-Path $script:RepoRoot 'dpia100.html'

function Find-Edge {
    $candidates = @(
        (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe' -ErrorAction SilentlyContinue).'(default)'
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) { return $c }
    }
    return $null
}

# Reads a file that Edge still has open for writing.
function Read-SharedText([string]$Path) {
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    try { return (New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8)).ReadToEnd() } finally { $fs.Dispose() }
}

# Opens an HTML file in headless Edge and waits until the page logs a console line matching $Pattern.
# Returns the first capture group (or $null on timeout) and all console lines of the page.
function Invoke-EdgeLog([string]$HtmlPath, [string]$Pattern, [int]$TimeoutSec = 90) {
    $edge = Find-Edge
    if (-not $edge) { throw 'Microsoft Edge (msedge.exe) was not found.' }
    $profileDir = Join-Path ([IO.Path]::GetTempPath()) ('dpia100-edge-' + [guid]::NewGuid().ToString('N'))
    $url = ([Uri](Resolve-Path -LiteralPath $HtmlPath).ProviderPath).AbsoluteUri
    $edgeArgs = @('--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions',
                  '--disable-sync', '--mute-audio', "--user-data-dir=`"$profileDir`"", '--enable-logging', '--v=0',
                  '--virtual-time-budget=30000', '--dump-dom', $url)
    Start-Process -FilePath $edge -ArgumentList $edgeArgs -WindowStyle Hidden | Out-Null

    $log = Join-Path $profileDir 'chrome_debug.log'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $found = $null
    $text = ''
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec -and $null -eq $found) {
        Start-Sleep -Milliseconds 150
        if (Test-Path -LiteralPath $log) {
            try { $text = Read-SharedText $log } catch { continue }
            $m = [regex]::Match($text, $Pattern)
            if ($m.Success) { $found = $m.Groups[1].Value }
        }
    }

    # Edge closes itself after --dump-dom; wait for that, then remove the profile.
    $deadline = [DateTime]::Now.AddSeconds(15)
    while ((Test-Path -LiteralPath $profileDir) -and [DateTime]::Now -lt $deadline) {
        try { Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 300 }
    }
    if (Test-Path -LiteralPath $profileDir) {
        Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -like "*$profileDir*" } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        Start-Sleep -Milliseconds 500
        Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Only the page's own console lines (not Edge's internal pages).
    $console = @($text -split '\r?\n' | Where-Object { $_ -match ':CONSOLE' -and $_ -match 'source: file:' })
    return [PSCustomObject]@{ Match = $found; Console = $console }
}

# Opens an HTML file in headless Edge and returns the object the page logged as '<Marker>:<base64 JSON>'.
# Throws with the page's console output when no result arrives in time.
function Invoke-EdgePage([string]$HtmlPath, [string]$Marker = 'DPIA100-RESULT', [int]$TimeoutSec = 90) {
    $r = Invoke-EdgeLog $HtmlPath ('"' + [regex]::Escape($Marker) + ':([A-Za-z0-9+/=]*)"') $TimeoutSec
    if ($null -eq $r.Match) {
        throw "No result from $HtmlPath within $TimeoutSec s. Page console:`n$(($r.Console | Select-Object -First 20) -join "`n")"
    }
    $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($r.Match))
    return ($json | ConvertFrom-Json)
}

# page-smoke: the real page (or a copy) starts in Edge under its own CSP, without any CSP violation.
function Test-PageStarts([string]$Path = $script:PagePath) {
    $problems = [System.Collections.Generic.List[string]]::new()
    $r = Invoke-EdgeLog $Path '"DPIA100 converter (\S+) gestart"' 30
    if ($null -eq $r.Match) { $problems.Add("the UI script did not start (blocked by the CSP?). Console:`n      $($r.Console -join "`n      ")") }
    foreach ($line in $r.Console) {
        if ($line -match 'Content Security Policy|Refused to|TrustedHTML|Uncaught') { $problems.Add("console: $line") }
    }
    return $problems
}

# The inline <script> and <style> blocks and the CSP of the page. Text is LF-normalized, as the HTML
# parser sees it (CSP hashes are computed over that).
function Get-PageParts([string]$Path = $script:PagePath) {
    $html = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $norm = $html -replace "`r`n?", "`n"
    $blocks = foreach ($m in [regex]::Matches($norm, '(?s)<(script|style)(\s[^>]*)?>(.*?)</\1>')) {
        $id = [regex]::Match($m.Groups[2].Value, 'id="([^"]+)"').Groups[1].Value
        [PSCustomObject]@{ Tag = $m.Groups[1].Value; Id = $id; Text = $m.Groups[3].Value; Hash = (Get-CspHash $m.Groups[3].Value) }
    }
    $csp = [regex]::Match($norm, '<meta http-equiv="Content-Security-Policy" content="([^"]*)"').Groups[1].Value
    return [PSCustomObject]@{ Html = $html; Blocks = @($blocks); Csp = $csp }
}

function Get-CspHash([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return 'sha256-' + [Convert]::ToBase64String($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))) } finally { $sha.Dispose() }
}

# The one CSP the page must have: no network access of any kind, and only the inline blocks with these
# hashes may run. Trusted Types makes Edge refuse HTML strings in innerHTML and similar sinks.
function Get-PagePolicy($Parts) {
    $scripts = ($Parts.Blocks | Where-Object Tag -eq 'script' | ForEach-Object { "'$($_.Hash)'" }) -join ' '
    $styles  = ($Parts.Blocks | Where-Object Tag -eq 'style' | ForEach-Object { "'$($_.Hash)'" }) -join ' '
    return "default-src 'none'; script-src $scripts; style-src $styles; connect-src 'none'; form-action 'none'; " +
           "base-uri 'none'; require-trusted-types-for 'script'; trusted-types 'none'"
}

# Returns the HTML with the CSP meta tag set to the given policy.
function Set-PagePolicy([string]$Html, [string]$Policy) {
    $pattern = '(<meta http-equiv="Content-Security-Policy" content=")[^"]*(">)'
    if (-not [regex]::IsMatch($Html, $pattern)) { throw 'No Content-Security-Policy meta tag found.' }
    return [regex]::Replace($Html, $pattern, { param($m) $m.Groups[1].Value + $Policy + $m.Groups[2].Value }, 'None')
}

# Static privacy checks of dpia100.html. Returns a list of problems (empty = fine).
function Test-PageSecurity([string]$Path = $script:PagePath) {
    $problems = [System.Collections.Generic.List[string]]::new()
    $parts = Get-PageParts $Path
    $norm = $parts.Html -replace "`r`n?", "`n"

    # page-csp: the policy must be exactly the canonical one, with the current hashes.
    $expected = Get-PagePolicy $parts
    if ($parts.Csp -ne $expected) {
        $problems.Add("CSP does not match the inline blocks. Run tests\web\update-csp.ps1.`n      expected: $expected`n      actual  : $($parts.Csp)")
    }
    foreach ($bad in @("'unsafe-", 'http:', 'https:', ' data:', ' blob:', '*')) {
        if ($parts.Csp.Contains($bad)) { $problems.Add("CSP contains $bad") }
    }
    # The CSP must be in force before anything can load: directly after the charset.
    $head = [regex]::Match($norm, '(?s)<head>(.*?)</head>').Groups[1].Value
    if ($head -notmatch '^\s*<meta charset="utf-8">\s*<meta http-equiv="Content-Security-Policy"') {
        $problems.Add('the CSP meta tag must directly follow <meta charset="utf-8"> at the start of <head>')
    }
    $ids = @($parts.Blocks | ForEach-Object Id)
    foreach ($id in 'dpia100-style', 'dpia100-core', 'dpia100-ui') {
        if ($ids -notcontains $id) { $problems.Add("missing inline block id=`"$id`"") }
    }
    if ($parts.Blocks.Count -ne 3) { $problems.Add("expected 3 inline blocks, found $($parts.Blocks.Count)") }

    # No external resources and no inline event handlers; links only to the project on GitHub.
    $markup = [regex]::Replace($norm, '(?s)<(script|style)(\s[^>]*)?>.*?</\1>', '')
    if ($markup -match '(?i)<(iframe|frame|object|embed|img|video|audio|source|base|form[^>]*action)\b') { $problems.Add("markup contains <$($Matches[1])>") }
    if ($markup -match '(?i)\s(src|srcset|action|formaction|background|poster)\s*=') { $problems.Add("markup has a $($Matches[1])= attribute") }
    if ($markup -match '(?i)\son[a-z]+\s*=') { $problems.Add('markup has an inline event handler') }
    foreach ($m in [regex]::Matches($markup, '(?i)href\s*=\s*"([^"]*)"')) {
        $href = $m.Groups[1].Value
        if ($href -notmatch '^(#[\w-]+|https://github\.com/japie06/dpia100-converter(/[\w./-]*)?)$') { $problems.Add("unexpected link href=`"$href`"") }
    }
    if ($markup -match '(?i)<link\b') { $problems.Add('markup has a <link> element') }

    # page-api-scan: no network, navigation or code-from-string APIs in the scripts. The only href
    # assignment allowed is the Blob URL of the download.
    $forbidden = [ordered]@{
        'fetch'                = '\bfetch\s*\('
        'XMLHttpRequest'       = 'XMLHttpRequest'
        'WebSocket'            = 'WebSocket'
        'EventSource'          = 'EventSource'
        'sendBeacon'           = 'sendBeacon'
        'RTCPeerConnection'    = 'RTCPeerConnection|RTCDataChannel'
        'dynamic import'       = '\bimport\s*\('
        'importScripts'        = 'importScripts'
        'eval'                 = '\beval\s*\('
        'Function constructor' = '\bFunction\s*\('
        'string timer'         = 'set(Timeout|Interval)\s*\(\s*[''"`]'
        'innerHTML'            = 'innerHTML|outerHTML|insertAdjacentHTML|createContextualFragment|DOMParser'
        'document.write'       = 'document\s*\.\s*write'
        'window.open'          = '\bopen\s*\('
        'location'             = '\blocation\b'
        'postMessage'          = 'postMessage'
        'workers'              = '\bWorker\b|SharedWorker|serviceWorker'
        'Image beacon'         = 'new\s+Image\b'
        'src assignment'       = '\.src\s*=|\.srcset\s*='
        'href assignment'      = '\.href\s*=(?!\s*URL\.createObjectURL\()'
        'setAttribute url'     = 'setAttribute\(\s*[''"](src|href|action|formaction|srcset)'
        'navigator.share'      = 'navigator\s*\.\s*share'
    }
    foreach ($block in ($parts.Blocks | Where-Object Tag -eq 'script')) {
        foreach ($name in $forbidden.Keys) {
            foreach ($m in [regex]::Matches($block.Text, $forbidden[$name])) {
                $lineNo = ($block.Text.Substring(0, $m.Index) -split "`n").Count
                $problems.Add("forbidden API '$name' in script $($block.Id), line $lineNo of the block: $($m.Value)")
            }
        }
    }
    return $problems
}

# Builds a page with the core script of dpia100.html (copied byte for byte), harness.js and the cases,
# runs it in Edge and returns the harness payload (version, encodable code points, case results).
# -ExtraRubrieken adds lines to the core's standard rubrieken table, for test rubrieken the real list
# lacks (tests\cases\*\extra-rubrieken.txt); the core is otherwise unchanged.
function Invoke-WebHarness([object[]]$Cases, [string]$OutDir, [string[]]$ExtraRubrieken = @(), [string]$PageName = 'web-harness.html') {
    $core = (Get-PageParts).Blocks | Where-Object { $_.Id -eq 'dpia100-core' }
    if (-not $core) { throw 'No <script id="dpia100-core"> in dpia100.html.' }
    $coreText = $core.Text
    if ($ExtraRubrieken.Count -gt 0) {
        $marker = 'const STANDARD_RUBRIEKEN_TABLE = String.raw`' + "`n"
        if (-not $coreText.Contains($marker)) { throw 'No STANDARD_RUBRIEKEN_TABLE in the core of dpia100.html.' }
        $coreText = $coreText.Replace($marker, $marker + (($ExtraRubrieken | ForEach-Object { "$_`n" }) -join ''))
    }
    $harness = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'harness.js'), [Text.Encoding]::UTF8)
    # '<' is escaped so no case text can close the script element.
    $json = (ConvertTo-Json -InputObject @($Cases) -Depth 5 -Compress).Replace('<', '<')
    $page = "<!doctype html>`n<html><head><meta charset=`"utf-8`"><title>DPIA100 harness</title></head><body>`n" +
            "<script>$coreText</script>`n<script>$harness</script>`n<script>runHarness($json);</script>`n</body></html>`n"
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $path = Join-Path $OutDir $PageName
    [IO.File]::WriteAllText($path, $page, (New-Object Text.UTF8Encoding($false)))
    $result = Invoke-EdgePage $path
    if ($result.error) { throw "Harness error: $($result.error)" }
    return $result
}
