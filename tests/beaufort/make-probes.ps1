<#
.SYNOPSIS
    Builds dpia100 probe files to find out what Beaufort accepts at import (100254).
.DESCRIPTION
    Focus: special characters in Waarde, amounts, and the file format (line endings, record length).
    The files are written raw, NOT through dpia100.ps1, because many probes contain things the script
    would reject. Beaufort rejects only the records with errors and reports them, so almost everything
    goes into one file:
      DPIA100_probes.txt      all record-level and line-level probes (P01-P08), CRLF
      DPIA100_P07a-lf.txt     the whole file with LF line endings (cannot be mixed with CRLF probes)
    File names start with DPIA100, otherwise Beaufort does not find them.
    Every text Waarde starts with the probe ID (P02-0x80 = group P02, character 0x80), so each record can
    be recognised in the import report and in 100251. probes\index.txt lists every ID with its content.

    With -Web the probes test the web version instead (round 4): the script writes one probe CSV of
    valid input and lets the page's own core (dpia100.html, in headless Edge, as tests/run-tests.ps1 does)
    turn it into the import file. Everything is in one file:
      web-probes.csv          the input, UTF-8 with BOM and ';' as Dutch Excel saves it
      DPIA100_web-probes.txt  the import file, with the name the page suggests
      index-web.txt           every W-ID with its CSV row
    To also cover the real interface, open web-probes.csv in dpia100.html, make the file with the same
    procescode, creatiedatum and user, and compare: fc /b DPIA100_web-probes.txt <your file>.
.PARAMETER Only
    Probe IDs (prefixes) to include, e.g. P04-5,P04-6,P08 (or W05 with -Web). Use it to repeat a few
    probes without re-importing everything. Without it, all probes are written.
.PARAMETER Web
    Build the web-app probes (W01-W06) through dpia100.html instead of the raw probes.
.EXAMPLE
    pwsh tests\beaufort\make-probes.ps1
.EXAMPLE
    pwsh tests\beaufort\make-probes.ps1 -Only P04-5,P04-6,P08
.EXAMPLE
    pwsh tests\beaufort\make-probes.ps1 -Web
#>
param([string[]]$Only, [switch]$Web)
# With -File, "P04-5,P04-6,P08" arrives as one string.
$Only = @($Only | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# ---- Settings: fill in for the test environment -------------------------------------------
$Proces  = 'CMT'        # procescode of the aanleverend proces
$User    = 'JvdBerg'    # creatiegebruiker that exists in Beaufort (fill in)
$Cdate   = '01012020'   # creatiedatum ddMMyyyy
$Ingang  = '01102020'   # ingangsdatum ddMMyyyy (today or earlier, so it can be processed)
$PersNr  = '10000'        # test person (fill in)
$Dv      = '1'          # dienstverband of the test person
$RubTxt  = 'TST003'     # PS, stam, type S (tekst), length 50
$RubDv   = 'TST002'     # DV, stam, type S (tekst), length 50
$RubNum  = 'TST004'     # DV, stam, type N (numeriek), 2 decimals (variabel rubrieken cannot be created)
# --------------------------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$outDir = Join-Path $PSScriptRoot 'probes'
New-Item -ItemType Directory -Force $outDir | Out-Null
$ansi  = [System.Text.Encoding]::GetEncoding(1252, [System.Text.EncoderExceptionFallback]::new(), [System.Text.DecoderExceptionFallback]::new())
$index = [System.Collections.Generic.List[string]]::new()

# ---- Round 4: probes for the web version (-Web) ---------------------------------------------
if ($Web) {
    . (Join-Path $PSScriptRoot '..\web\WebTest.ps1')
    Get-ChildItem $outDir -File -Filter '*web-probes*' | Remove-Item
    Remove-Item (Join-Path $outDir 'index-web.txt') -ErrorAction SilentlyContinue
    $inv  = [System.Globalization.CultureInfo]::InvariantCulture
    $date = [DateTime]::ParseExact($Ingang, 'ddMMyyyy', $inv)
    $dash = $date.ToString('dd-MM-yyyy', $inv)
    $rows = [System.Collections.Generic.List[string]]::new()

    # One CSV row. Cells are given as a user would type them; the page validates and converts them.
    function Row {
        param($Id, $Waarde, $Rub = $RubTxt, $Pers = $PersNr, $DvCell = '', $SV = 'S', $Datum = $dash, $Prod = 'NOR')
        if ($Only -and -not ($Only | Where-Object { $Id -like "$_*" })) { return }
        $cells = foreach ($c in @($Rub, $Pers, $DvCell, $SV, $Waarde, $Datum, $Prod)) {
            if ($c -match '[;"\r\n]') { '"' + $c.Replace('"', '""') + '"' } else { $c }
        }
        $script:rows.Add($cells -join ';')
        $script:index.Add("$Id  $($cells -join ';')")
    }

    # W01: keys. Leading zeros and spaces are removed; dv 0 makes a PS key (decision 4).
    Row 'W01-1' 'W01-1 PS sleutel'
    Row 'W01-2' 'W01-2 DV sleutel' -Rub $RubDv -DvCell $Dv
    Row 'W01-3' 'W01-3 voorloopnullen persnr en dv' -Rub $RubDv -Pers "00$PersNr" -DvCell "00$Dv"
    Row 'W01-4' 'W01-4 dv 0 wordt PS sleutel' -DvCell '0'
    Row 'W01-5' 'W01-5 spaties rond sleutels' -Rub $RubDv -Pers " $PersNr " -DvCell " $Dv "
    # W02: codes are uppercased; STAM means S; empty productiesoort means NOR. All test rubrieken are stam.
    Row 'W02-1' 'W02-1 rubriek kleine letters' -Rub $RubTxt.ToLower()
    Row 'W02-2' 'W02-2 S of V als STAM' -SV 'STAM'
    Row 'W02-3' 'W02-3 S of V als s' -SV 's'
    Row 'W02-4' 'W02-4 productiesoort leeg' -Prod ''
    Row 'W02-5' 'W02-5 productiesoort nor' -Prod 'nor'
    # W03: every defined Windows-1252 character above 0x7F, typed in UTF-8 and converted by the page.
    foreach ($b in 0x80..0xFF) {
        if ($b -in 0x81, 0x8D, 0x8F, 0x90, 0x9D) { continue }
        $id = 'W03-0x{0:X2}' -f $b
        Row $id ('{0} a{1}b' -f $id, $ansi.GetString([byte[]]@($b)))
    }
    Row 'W03-d' "W03-d Gebiedscoo$([char]0x308)rdinator (losse trema)"
    Row 'W03-q' "W03-q d$([char]0x2019)Ancona $([char]0x201C)citaat$([char]0x201D)"
    Row 'W03-s' 'W03-s puntkomma; en "aanhalingstekens"'
    # W04: the ingangsdatum in every input format the page accepts; all become the same date.
    $formats = 'd M yyyy', 'dd-MM-yyyy', 'd/M/yyyy', 'yyyy-MM-dd', 'ddMMyyyy', 'd.M.yyyy'
    for ($i = 0; $i -lt $formats.Count; $i++) {
        Row ('W04-{0}' -f ($i + 1)) ('W04-{0} datum als {1}' -f ($i + 1), $formats[$i]) -Datum $date.ToString($formats[$i], $inv)
    }
    # W05: amounts in the numeric rubriek. A numeric Waarde cannot hold the ID, so each amount gets its
    # own ingangsdatum (earlier days), which also keeps processing from overwriting them (round 3).
    $amounts = '12,50', '-12,50', '3250,5', '0,25', '12.50', '3250'
    for ($i = 0; $i -lt $amounts.Count; $i++) {
        Row ('W05-{0}' -f ($i + 1)) $amounts[$i] -Rub $RubNum -DvCell $Dv -Datum $date.AddDays(-($i + 1)).ToString('dd-MM-yyyy', $inv)
    }
    # W06: length and empty value. 60 characters are cut to 50 with a warning (decision 3).
    Row 'W06-1' 'W06-1 precies 50 tekens '.PadRight(50, 'x')
    Row 'W06-2' 'W06-2 zestig tekens, wordt afgekapt na vijftig '.PadRight(60, 'y')
    Row 'W06-3' '' -Datum $date.AddDays(-10).ToString('dd-MM-yyyy', $inv)

    if ($rows.Count -eq 0) { throw "No probes match -Only $($Only -join ',')." }
    $csvText  = "Rubriek;Pers. Nr;Dienstv;S of V;Mutatiewaarde;Ingangsdatum;Productiesoort`r`n" + ($rows -join "`r`n") + "`r`n"
    $csvBytes = [byte[]](0xEF, 0xBB, 0xBF) + [System.Text.UTF8Encoding]::new($false).GetBytes($csvText)
    [System.IO.File]::WriteAllBytes((Join-Path $outDir 'web-probes.csv'), $csvBytes)

    # The page's core makes the file, answering the prompts like a user (Y confirms the warnings).
    $case = [PSCustomObject]@{
        name = 'web-probes'; input = [Convert]::ToBase64String($csvBytes); fileName = 'web-probes.csv'
        answers = @($Proces, $Cdate, $User, 'Y'); outputName = $null
    }
    $result = (Invoke-WebHarness -Cases @($case) -OutDir (Join-Path ([IO.Path]::GetTempPath()) 'dpia100-probes')).cases[0]
    if ($result.exitCode -ne 0 -or -not $result.output) {
        Write-Host "The page did not make a file:`n$($result.transcript)" -ForegroundColor Red
        exit 1
    }
    $bytes = [Convert]::FromBase64String($result.output)
    $file  = Join-Path $outDir $result.suggestedName
    [System.IO.File]::WriteAllBytes($file, $bytes)
    if ($bytes.Length -ne $rows.Count * 189) { throw "Expected $($rows.Count) records of 187 bytes + CRLF, got $($bytes.Length) bytes." }
    [System.IO.File]::WriteAllLines((Join-Path $outDir 'index-web.txt'), $index, [System.Text.UTF8Encoding]::new($true))
    $warnings = @($result.transcript -split '\n' | Where-Object { $_ -like '  - *' }).Count
    $sha = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    Write-Host ("{0,-24} {1,3} records {2,6} bytes, {3} warning(s) confirmed" -f $result.suggestedName, $rows.Count, $bytes.Length, $warnings)
    Write-Host "SHA-256 $sha"
    Write-Host "Input: $(Join-Path $outDir 'web-probes.csv')`nIndex: $(Join-Path $outDir 'index-web.txt')"
    exit 0
}

Get-ChildItem $outDir -File | Where-Object Name -notlike '*web-probes*' | Where-Object Name -ne 'index-web.txt' | Remove-Item

# Formatted keys. Not named $ps/$dv: PowerShell variable names ignore case, so $dv would overwrite $Dv.
$psKey = '{0,6}' -f $PersNr
$dvKey = '{0,6} {1,3}' -f $PersNr, $Dv

# One record. Slots are given as they must appear; each is padded/cut to its width so a probe can
# deviate in exactly one field. -Length makes the record too short or too long.
function Rec {
    param($Waarde, $Rub = $RubTxt, $Obj = $psKey, $SV = 'S', $Idate = "   $Ingang", $Proc = $Proces,
          $CdateSlot = "   $Cdate", $UserSlot = $User.PadLeft(16), [int]$Length = 187, $Label = '')
    $l = $Proc.PadRight(3).Substring(0, 3) + $Rub.PadRight(6).Substring(0, 6) + $Obj.PadRight(50).Substring(0, 50) + $SV +
         $CdateSlot.PadRight(11).Substring(0, 11) + $UserSlot.PadRight(16).Substring(0, 16) + $Waarde.PadRight(50).Substring(0, 50) +
         $Idate.PadRight(11).Substring(0, 11) + (' ' * 30) + 'NOR' + (' ' * 6)
    if ($Length -lt 187) { $l = $l.Substring(0, $Length) } elseif ($Length -gt 187) { $l = $l + ('X' * ($Length - 187)) }
    # With -Only, skip probes whose ID (the label, or the first word of Waarde) does not match.
    $id = if ($Label) { $Label } else { ($Waarde.Trim() -split '\s')[0] }
    if ($Only -and -not ($Only | Where-Object { $id -like "$_*" })) { return $null }
    $script:index.Add((("$Label " + $Waarde.Trim())).Trim())
    $l
}

function Section([string]$title) { $script:index.Add("`r`n== $title ==") }

function Write-Probe([string]$name, [object[]]$lines, [string]$newline = "`r`n", [switch]$NoFinalNewline) {
    $lines = @($lines | Where-Object { $null -ne $_ })   # probes skipped by -Only
    if ($lines.Count -eq 0) { return }
    $text  = ($lines -join $newline) + $(if ($NoFinalNewline) { '' } else { $newline })
    $bytes = $ansi.GetBytes($text)
    [System.IO.File]::WriteAllBytes((Join-Path $outDir "$name.txt"), [byte[]]$bytes)
    Write-Host ("{0,-24} {1,3} lines {2,6} bytes" -f "$name.txt", $lines.Count, $bytes.Length)
}

# List[object], not List[string]: PowerShell turns $null (a probe skipped by -Only) into '' for strings.
$all = [System.Collections.Generic.List[object]]::new()

# P01: ASCII punctuation, one character per record
Section 'P01 ASCII punctuation'
foreach ($c in ' !"#$%&''()*+,-./:;<=>?@[\]^_`{|}~'.ToCharArray()) { $all.Add((Rec ('P01-0x{0:X2} a{1}b' -f [int]$c, $c))) }

# P02: every printable Windows-1252 character above 0x7F, one per record.
# Bytes 0x81, 0x8D, 0x8F, 0x90 and 0x9D are undefined in Windows-1252 and left out.
Section 'P02 Windows-1252 characters 0x80-0xFF'
foreach ($b in 0x80..0xFF) {
    if ($b -in 0x81, 0x8D, 0x8F, 0x90, 0x9D) { continue }
    $all.Add((Rec ('P02-0x{0:X2} a{1}b' -f $b, $ansi.GetString([byte[]]@($b)))))
}

# P03: realistic Dutch text and edge positions
# Here-string: PowerShell treats curly quotes as quote characters in normal string literals.
Section 'P03 Dutch text'
$p03 = @'
P03-1 Gebiedscoördinator café één reünie
P03-2 Façade à la carte crème brûlée naïef
P03-3 d’Ancona ’s-Hertogenbosch “quote”
P03-4 € 1250,00 – ½ dag
'@ -split '\r?\n'
foreach ($t in $p03) { $all.Add((Rec $t)) }
$all.Add((Rec ('P03-5 eindigt op e'.PadRight(49, '.') + [char]0xE9)))   # accented character in the last position (137)
$all.Add((Rec ('P03-6 '.PadRight(50, 'x'))))                          # exactly 50 characters
$all.Add((Rec '   P03-7 voorloopspaties'))
$all.Add((Rec 'P03-8 kleine letters en HOOFDLETTERS'))

# P04: object_id alignment and leading zeros (docs: no leading zeros; layout from the Excel tool)
Section 'P04 object_id'
$all.Add((Rec 'P04-1 PS rechts uitgelijnd (Excel tool)'))
$all.Add((Rec 'P04-2 PS links uitgelijnd' -Obj $PersNr))
$all.Add((Rec 'P04-3 PS voorloopnullen' -Obj $PersNr.PadLeft(6, '0')))
$all.Add((Rec 'P04-4 DV Excel tool' -Rub $RubDv -Obj $dvKey))
$all.Add((Rec 'P04-5 DV links, 1 spatie' -Rub $RubDv -Obj "$PersNr $Dv"))
$all.Add((Rec 'P04-6 DV voorloopnullen dv' -Rub $RubDv -Obj ('{0,6} {1}' -f $PersNr, $Dv.PadLeft(3, '0'))))
$all.Add((Rec 'P04-7 PS rubriek met DV sleutel' -Obj $dvKey))
$all.Add((Rec 'P04-8 DV rubriek met PS sleutel' -Rub $RubDv))

# P05: procescode, rubriek and creatiegebruiker
Section 'P05 header fields'
$all.Add((Rec 'P05-1 procescode kleine letters' -Proc $Proces.ToLower()))
$all.Add((Rec 'P05-2 rubriek kleine letters' -Rub $RubTxt.ToLower()))
$all.Add((Rec 'P05-3 gebruiker links uitgelijnd' -UserSlot $User))

# P06: date formats in the creatiedatum / ingangsdatum slots
Section 'P06 dates'
$d10 = "$($Ingang.Substring(0,2))-$($Ingang.Substring(2,2))-$($Ingang.Substring(4))"
$d6  = $Ingang.Substring(0, 4) + $Ingang.Substring(6)
$all.Add((Rec 'P06-1 ingangsdatum dd-mm-eejj' -Idate " $d10"))
$all.Add((Rec 'P06-2 ingangsdatum ddmmjj' -Idate "     $d6"))
$all.Add((Rec 'P06-3 ingangsdatum links' -Idate $Ingang))
$all.Add((Rec 'P06-4 creatiedatum dd-mm-eejj' -CdateSlot " $d10"))

# P08: amounts in a numeric rubriek. Confirms (or not) the script's comma -> dot conversion and the
# thousand-separator rejection (decision 8). A numeric Waarde cannot hold the probe ID, so every value
# is unique; index.txt maps ID -> value.
Section 'P08 amounts (numeric rubriek)'
$amounts = '12.50', '12,50', '-12.50', '-12,50', '3.250', '3250', '1.234,56', '1,234.56', '12.345', '0012.50', '12.5', '12'
for ($i = 0; $i -lt $amounts.Count; $i++) {
    $all.Add((Rec $amounts[$i] -Rub $RubNum -Obj $dvKey -SV 'S' -Label ('P08-{0}' -f ($i + 1))))
}

# P07: line-level probes. Beaufort reads line by line (round 1), so these can share the file.
# They come last; P07c (the very last line) has no line break after it.
Section 'P07 line level'
$all.Add((Rec 'P07f-1 regel van 186 tekens' -Length 186))
$all.Add((Rec 'P07g-1 regel van 188 tekens' -Length 188))
$p07h = Rec 'P07h-1 spaties achteraan weggehaald'
if ($null -ne $p07h) { $p07h = $p07h.TrimEnd() }
$all.Add($p07h)
$all.Add((Rec "P07j-1 a`tb tab in waarde"))
$all.Add((Rec 'P07k dubbel record'))
$all.Add((Rec 'P07k dubbel record'))
if (-not $Only -or ($Only | Where-Object { 'P07i' -like "$_*" })) {
    $all.Add('')                                               # P07i: blank line
    $index.Add('P07i (blank line, expected: PI_1537 with all fields empty)')
}
$all.Add((Rec 'P07c laatste regel zonder regeleinde'))
Write-Probe 'DPIA100_probes' $all -NoFinalNewline

# P07a: LF line endings must be a file of its own
Section 'DPIA100_P07a-lf.txt'
Write-Probe 'DPIA100_P07a-lf' @((Rec 'P07a-1 LF regeleinden'), (Rec 'P07a-2 LF regeleinden')) -newline "`n"

[System.IO.File]::WriteAllLines((Join-Path $outDir 'index.txt'), $index, [System.Text.UTF8Encoding]::new($true))
Write-Host "`nIndex of all probe IDs: $(Join-Path $outDir 'index.txt')"
