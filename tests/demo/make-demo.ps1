<#
.SYNOPSIS
    Writes the demo CSV files in tests/demo: every error, warning and report line of the converter.
.DESCRIPTION
    For a visible check in the script or the page, and to show users what the messages look like. Not
    part of the test pipeline. Update this file with every new or changed validation (CLAUDE.md), then run
    it and check the coverage with `node tests\demo\check-demo.mjs`.

    The column Opmerking says what a row shows; the converter does not map it, so it is ignored.
    Each row of 1-fouten.csv has exactly one error. Rubriek codes: E00101 is an eigen rubriek (not in the
    standard list); P-codes are standard rubrieken (look up the line in the table of dpia100.ps1).
    The source is ASCII only: other characters are written as [char] codes.
#>
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot

function Write-Csv([string]$name, [string[]]$header, [object[]]$rows, [string]$delim, [System.Text.Encoding]$enc, [switch]$Bom) {
    $quote = {
        param([string]$v)
        if ($v.Contains($delim) -or $v.Contains('"') -or $v.Contains("`r") -or $v.Contains("`n")) { '"' + $v.Replace('"', '""') + '"' } else { $v }
    }
    $lines = @(($header | ForEach-Object { & $quote $_ }) -join $delim)
    foreach ($r in $rows) { $lines += (($r | ForEach-Object { & $quote ([string]$_) }) -join $delim) }
    $bytes = $enc.GetBytes(($lines -join "`r`n") + "`r`n")
    if ($Bom) { $bytes = [byte[]](0xEF, 0xBB, 0xBF) + $bytes }
    [IO.File]::WriteAllBytes((Join-Path $dir $name), $bytes)
    Write-Host "$name ($($rows.Count) rows)"
}

$header  = @('Rubriek', 'Pers. Nr', 'Dienstv', 'S of V', 'Mutatiewaarde', 'Ingangsdatum', 'Productiesoort', 'Opmerking')
$utf8    = New-Object System.Text.UTF8Encoding($false)
$ansi    = [System.Text.Encoding]::GetEncoding(1252)
$eAcute  = [string][char]0x00E9
$lStroke = [string][char]0x0141

# 1. Every row-level error, one per row (UTF-8 with BOM, so a character outside Windows-1252 fits).
Write-Csv '1-fouten.csv' $header -delim ';' -enc $utf8 -Bom -rows @(
    ,@('E00101', '123456', '1', 'V', '100', '1-11-2026', 'NOR', 'Goede regel ter vergelijking: geeft geen fout')
    # Rubriek
    ,@('', '123456', '1', 'V', '100', '1-11-2026', 'NOR', 'Rubriek is leeg')
    ,@('#N/B', '123456', '1', 'V', '100', '1-11-2026', 'NOR', 'Rubriek met andere tekens dan letters en cijfers (foutwaarde uit Excel)')
    ,@('E001011', '123456', '1', 'V', '100', '1-11-2026', 'NOR', 'Rubriek langer dan 6 tekens')
    # Pers. Nr
    ,@('E00101', '', '1', 'V', '100', '1-11-2026', 'NOR', 'Persoonsnummer is leeg')
    ,@('E00101', '12a456', '1', 'V', '100', '1-11-2026', 'NOR', 'Persoonsnummer met een letter')
    ,@('E00101', '000', '1', 'V', '100', '1-11-2026', 'NOR', 'Persoonsnummer is alleen nullen')
    ,@('E00101', '1234567', '1', 'V', '100', '1-11-2026', 'NOR', 'Persoonsnummer langer dan 6 cijfers')
    # Dienstverband
    ,@('E00101', '123456', '1a', 'V', '100', '1-11-2026', 'NOR', 'Dienstverband met een letter')
    ,@('E00101', '123456', '1234', 'V', '100', '1-11-2026', 'NOR', 'Dienstverband langer dan 3 cijfers')
    ,@('P00001', '123456', '', 'V', '100', '1-11-2026', 'NOR', 'Standaard DV-rubriek zonder dienstverband')
    ,@('P00001', '123456', '0', 'V', '100', '1-11-2026', 'NOR', 'Standaard DV-rubriek met dienstverband 0 (geeft een PS-sleutel)')
    ,@('P00301', '123456', '1', 'S', 'Tekst', '1-11-2026', 'NOR', 'Standaard PS-rubriek met een dienstverband')
    # S of V
    ,@('E00101', '123456', '1', 'X', '100', '1-11-2026', 'NOR', 'S of V is niet S of V')
    ,@('P00063', '123456', '1', 'S', '100', '1-11-2026', 'NOR', 'Standaardrubriek die alleen variabel is, met S')
    ,@('P00027', '123456', '1', 'V', '100', '1-11-2026', 'NOR', 'Standaardrubriek die alleen stam is, met V')
    # Ingangsdatum
    ,@('E00101', '123456', '1', 'V', '100', '32-01-2026', 'NOR', 'Ingangsdatum bestaat niet')
    # Productiesoort
    ,@('E00101', '123456', '1', 'V', '100', '1-11-2026', 'JWP', 'Productiesoort JWP (niet ondersteund)')
    ,@('E00101', '123456', '1', 'V', '100', '1-11-2026', 'XYZ', 'Productiesoort is niet NOR of COR')
    # Mutatiewaarde, every rubriek
    ,@('E00101', '123456', '1', 'V', "regel 1`r`nregel 2", '1-11-2026', 'NOR', 'Waarde met een regelovergang (Alt+Enter in Excel)')
    ,@('E00101', '123456', '1', 'V', "${lStroke}ukasz", '1-11-2026', 'NOR', 'Waarde met een teken dat niet in Windows-1252 bestaat')
    # Mutatiewaarde, standard rubriek by type (decision 20)
    ,@('P00001', '123456', '1', 'V', 'abc', '1-11-2026', 'NOR', 'Standaard numerieke rubriek met tekst')
    ,@('P00001', '123456', '1', 'V', '1.234,56', '1-11-2026', 'NOR', 'Standaard numerieke rubriek met een duizendtalpunt')
    ,@('P00001', '123456', '1', 'V', '12,345', '1-11-2026', 'NOR', 'Standaard numerieke rubriek (2 decimalen) met 3 decimalen')
    ,@('P00001', '123456', '1', 'V', '1234', '1-11-2026', 'NOR', 'Standaard numerieke rubriek (lengte 6 en 2 decimalen) met 4 cijfers voor de komma')
    ,@('P00321', '123456', '', 'S', '1/11/26', '1-11-2026', 'NOR', 'Standaard datumrubriek met een waarde die geen datum is')
    ,@('P00303', '123456', '', 'S', 'ABCDEFG', '1-11-2026', 'NOR', 'Standaard tekstrubriek (lengte 6) met 7 tekens')
)

# 2. Every warning and report line of a ';' file (Windows-1252, as Dutch Excel saves it). Converts after Y.
Write-Csv '2-waarschuwingen.csv' $header -delim ';' -enc $ansi -rows @(
    ,@('E00101', '123456', '', 'S', "Caf$eAcute de Brug", '1-11-2026', 'NOR', 'Gewone tekst: geen melding')
    ,@('E00101', '123456', '1', 'V', '125,50', '1-11-2026', 'NOR', 'Komma wordt punt: alleen gemeld')
    ,@('P00001', '123456', '1', 'V', '12,50', '1-11-2026', 'NOR', 'Standaard numerieke rubriek: komma wordt punt, alleen gemeld')
    ,@('P00332', '123456', '1', 'S', '1,5', '1-11-2026', 'NOR', 'Standaard tekstrubriek: komma blijft staan, geen melding')
    ,@('P00321', '123456', '', 'S', '1-11-2026', '1-11-2026', 'NOR', 'Standaard datumrubriek: datum wordt dd-mm-jjjj, alleen gemeld')
    ,@('E00101', '123456', '1', 'V', '3.250', '1-11-2026', 'NOR', 'Punt als duizendtal? Beaufort leest 3,25')
    ,@('E00101', '123456', '1', 'V', '1.234,56', '1-11-2026', 'NOR', 'Bedrag met punt en komma')
    ,@('E00101', '123456', '1', 'V', 'Deze omschrijving is langer dan vijftig tekens en wordt afgekapt', '1-11-2026', 'NOR', 'Waarde langer dan 50 tekens: wordt afgekapt')
    ,@('P00100', '123456', '1', 'V', '12', '1-11-2026', 'NOR', 'Standaardrubriek geblokkeerd voor indirecte mutatie')
    ,@('E00101', '123456', '1', 'V', '', '1-11-2026', 'NOR', 'Lege waarde: verwijdert de huidige waarde in Beaufort')
    ,@('E00101', '123456', '1', 'V', '1/11/2026', '1-11-2026', 'NOR', 'Datum niet als dd-mm-jjjj')
    ,@('E00101', '123456', '1', 'V', '8:30', '1-11-2026', 'NOR', 'Tijd niet als uu:mm of uu:mm:ss')
)

# 3. Amount warnings that depend on a ',' delimiter (English Excel).
Write-Csv '3-waarschuwingen-komma.csv' $header -delim ',' -enc $ansi -rows @(
    ,@('E00101', '123456', '1', 'V', '12,50', '1-11-2026', 'NOR', 'Komma wordt punt: in een kommabestand een waarschuwing')
    ,@('E00101', '123456', '1', 'V', '1,234', '1-11-2026', 'NOR', 'Komma als duizendtal: blijft staan')
)

# 4. Amount warnings that depend on a tab delimiter (no hint at the notation).
Write-Csv '4-waarschuwingen-tab.csv' $header -delim "`t" -enc $ansi -rows @(
    ,@('E00101', '123456', '1', 'V', '12,50', '1-11-2026', 'NOR', 'Komma wordt punt: in een tabbestand een waarschuwing')
    ,@('E00101', '123456', '1', 'V', '1,234', '1-11-2026', 'NOR', 'Komma als decimaal of als duizendtal? Wordt 1.234')
)

# 5. File-level errors: each stops before the rows are checked, so each needs its own file.
$plain = @('Rubriek', 'Pers. Nr', 'Dienstv', 'S of V', 'Mutatiewaarde', 'Ingangsdatum', 'Productiesoort')
[IO.File]::WriteAllBytes((Join-Path $dir 'bestand-leeg.csv'), [byte[]]@())
Write-Host 'bestand-leeg.csv (0 bytes)'
Write-Csv 'bestand-alleen-kopregel.csv' $plain -delim ';' -enc $ansi -rows @()
Write-Csv 'bestand-alleen-lege-regels.csv' $plain -delim ';' -enc $ansi -rows @(
    ,@('', '', '', '', '', '', '')
    ,@('', '', '', '', '', '', '')
)
Write-Csv 'bestand-kolom-ontbreekt.csv' @('Rubriek', 'Pers. Nr', 'Dienstv', 'Mutatiewaarde', 'Ingangsdatum', 'Productiesoort', 'Opmerking') -delim ';' -enc $ansi -rows @(
    ,@('E00101', '123456', '1', '100', '1-11-2026', 'NOR', 'De kolom S of V ontbreekt')
)
Write-Csv 'bestand-kolom-dubbel.csv' @('Rubriek', 'Persoonswaarde', 'Dienstv', 'S of V', 'Ingangsdatum', 'Productiesoort', 'Opmerking') -delim ';' -enc $ansi -rows @(
    ,@('E00101', '123456', '1', 'V', '1-11-2026', 'NOR', 'Persoonswaarde past bij Pers. Nr en bij Mutatiewaarde')
)
# A UTF-8 BOM followed by an e-acute in Windows-1252 (0xE9): not valid UTF-8.
$broken = $ansi.GetBytes("Rubriek;Pers. Nr;Dienstv;S of V;Mutatiewaarde;Ingangsdatum;Productiesoort`r`nE00101;123456;1;V;Caf$eAcute;1-11-2026;NOR`r`n")
[IO.File]::WriteAllBytes((Join-Path $dir 'bestand-ongeldige-utf8.csv'), [byte[]](0xEF, 0xBB, 0xBF) + $broken)
Write-Host 'bestand-ongeldige-utf8.csv (1 row)'
