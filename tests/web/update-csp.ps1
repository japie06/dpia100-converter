<#
.SYNOPSIS
    Writes the Content-Security-Policy of dpia100.html with the SHA-256 hashes of its inline script and
    style blocks. Run it after changing any of those blocks; until then the browser refuses to run them
    and tests/run-tests.ps1 (page-csp) fails.
.DESCRIPTION
    The hashes are computed over the block text with CRLF normalized to LF, as the HTML parser sees it,
    so they do not depend on how git checks the file out. To verify a hash by hand (PowerShell):
      $t = '<text between the tags, with LF line endings>'
      [Convert]::ToBase64String([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($t)))
.EXAMPLE
    pwsh tests\web\update-csp.ps1
#>
param([string]$Path)

. (Join-Path $PSScriptRoot 'WebTest.ps1')
if (-not $Path) { $Path = $script:PagePath }

$parts  = Get-PageParts $Path
$policy = Get-PagePolicy $parts
if ($parts.Csp -eq $policy) {
    Write-Host 'The CSP is up to date.' -ForegroundColor Green
    exit 0
}
$html = Set-PagePolicy $parts.Html $policy
[IO.File]::WriteAllText($Path, $html, (New-Object Text.UTF8Encoding($false)))
Write-Host "CSP updated in $Path" -ForegroundColor Green
foreach ($b in $parts.Blocks) { Write-Host ("  {0,-6} {1,-14} {2}" -f $b.Tag, $b.Id, $b.Hash) }
