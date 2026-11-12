# DPIA100 converter

[![CI](https://github.com/japie06/dpia100-converter/actions/workflows/ci.yml/badge.svg)](https://github.com/japie06/dpia100-converter/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/japie06/dpia100-converter)](https://github.com/japie06/dpia100-converter/releases/latest)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Dit hulpmiddel zet een CSV-bestand met mutaties om naar een bestand `DPIA100_<naam>.txt` dat je in
Beaufort / HR Core kunt importeren (functie *Import externe indirecte mutaties*, 100254).

Deze handleiding heeft drie delen:

- **[Voor gebruikers](#voor-gebruikers)**: de webpagina `dpia100.html`. Dit is wat de meeste mensen
  gebruiken.
- **[Scriptversie](#scriptversie-dpia100ps1)**: hetzelfde met `dpia100.ps1` of `dpia100.bat`.
- **[Voor beheerders](#voor-beheerders)**: de bestanden bijwerken en testen.

# Voor gebruikers

## Zo werkt het

1. Sla je gegevens in Excel op als CSV (*CSV (gescheiden door lijstscheidingsteken)* of *CSV UTF-8*).
   Open `voorbeeld.csv` in Excel om te zien hoe het eruit moet zien.
2. Open `dpia100.html` in **Microsoft Edge** of **Google Chrome**.  
Via internet kan ook: <https://japie06.github.io/dpia100-converter/>.
3. **Stap 1 op de pagina**: sleep je CSV-bestand op de pagina, of klik op het vak om het te kiezen.
4. **Stap 2**: de pagina controleert meteen het hele bestand.
   - Zijn er **fouten**, dan wordt er geen bestand gemaakt. Je ziet per fout de regel, de kolom, de
     waarde en een tip. Pas de regels aan in Excel, sla de CSV op en klik op **Opnieuw controleren**.
   - Kijk onder *Welke kolom is gebruikt?* of elke kolom bij het goede veld hoort.
   - Met *Bekijk de mutaties* zie je precies wat er in het bestand komt.
5. **Stap 3**: vul de **procescode** (bijv. `CMT`), de **creatiedatum** (standaard op vandaag) en je
   **gebruikersnaam in Beaufort** in. Met *Onthoud* hoef je procescode en gebruikersnaam de volgende keer
   niet opnieuw te typen. Waren er [waarschuwingen](#fouten-waarschuwingen-en-meldingen), dan vink je aan
   dat je ze hebt gezien. Klik op **Maak DPIA100_&lt;naam&gt;.txt** en kies de map, bijvoorbeeld de
   *Bestandsuitwisseling*.
6. Importeer het bestand in HR Core met functie 100254, beoordeel de mutaties (100251) en verwerk ze
   (100252).

**De bestandsnaam moet met `DPIA100` beginnen**, anders vindt HR Core het bestand niet. Het bestand
bekijken mag, maar sla het niet op in Kladblok of een andere editor: opslaan kan de tekens of regeleinden
veranderen.

**Privacy.** De pagina werkt volledig op je eigen computer: je CSV-bestand wordt niet verstuurd, ook niet
als je de pagina via internet opent. Gebruik op de pagina geen vertaalfunctie of AI-assistent (zoals
Copilot), want die sturen de inhoud wel door. Uitleg voor IT en privacy:
[docs/privacy-en-beveiliging.md](docs/privacy-en-beveiliging.md).

## Kolommen in de CSV

Zo ziet `voorbeeld.csv` eruit in Excel:

| Rubriek | Pers. Nr | Dienstv | S of V | Mutatiewaarde | Ingangsdatum | Productiesoort |
|---|---|---|---|---|---|---|
| E00101 | 123456 | | S | Café de Brug | 1-11-2026 | NOR |
| E00205 | 123456 | 1 | V | 125,50 | 1-11-2026 | NOR |
| E00310 | 2345 | 2 | S | 36 | 1-12-2026 | |

De eerste regel is een persoonsmutatie (*Dienstv* is leeg), de andere twee zijn
dienstverbandmutaties. De kolommen worden herkend aan de kopregel. De volgorde maakt niet uit en extra
kolommen worden genegeerd.

| Kolom (voorbeeldkop) | Verplicht | Toelichting |
|---|---|---|
| Rubriek | ja | Maximaal 6 letters/cijfers, bijvoorbeeld `E00101`. |
| Pers. Nr | ja | Alleen cijfers, max. 6. Voorloopnullen worden weggehaald. |
| Dienstv | nee | Alleen cijfers, max. 3. **Leeg = persoonsmutatie**, gevuld = dienstverbandmutatie. |
| S of V | ja | `S` (stam) of `V` (variabel). |
| Mutatiewaarde | nee | Max. 50 tekens. Zie [Bedragen](#bedragen) en [Datums, tijden en lege waarden](#datums-tijden-en-lege-waarden). |
| Ingangsdatum | ja | Bijvoorbeeld `1 11 2026`, `01-11-2026`, `01/11/2026`, `2026-11-01` of `01112026`. |
| Productiesoort | nee | `NOR` of `COR`. Leeg wordt `NOR`. |

## Fouten, waarschuwingen en meldingen

- Een **fout** blokkeert: er wordt geen bestand gemaakt tot alle fouten zijn opgelost. Regel 1 is de
  kopregel; de eerste regel met gegevens is regel 2.
- Een **waarschuwing** blokkeert niet, maar je moet haar bevestigen voordat het bestand wordt gemaakt.
  Je krijgt er een bij een waarde die Beaufort anders kan lezen dan jij bedoelt.
- Een **melding** vertelt wat het hulpmiddel zelf heeft aangepast, bijvoorbeeld een komma die een punt
  is geworden. Er wordt nooit iets ongemerkt veranderd.

Hoe streng de controle is, hangt af van de rubriek:

- Een **standaardrubriek** komt met Beaufort mee. Het hulpmiddel kent er ruim 4.200 (persoon en
  dienstverband) en weet wat erin hoort. Het controleert dan ook het dienstverband, S of V en de waarde,
  en geeft een fout als iets niet past. Zie [Standaardrubrieken](#standaardrubrieken).
- Een **eigen rubriek** is elke code die niet in die lijst staat. Die is gewoon toegestaan. Het
  hulpmiddel weet niet of er een bedrag, datum of tekst in hoort. Daarom schrijft het de waarde zoals je
  hem typt en waarschuwt het alleen bij twijfel; Beaufort controleert de rest.

### Bedragen

- Typ bedragen **zonder punt voor duizendtallen**: `3250` of `3250,50`, niet `3.250` of `3.250,50`.
  Beaufort leest `3.250` als 3,25. Zo'n punt komt er ongemerkt in als de cel in Excel is opgemaakt *met
  scheidingsteken voor duizendtallen*; kies dan een opmaak zonder.
- Een decimale komma wordt een punt (`12,50` wordt `12.50`), want Beaufort accepteert alleen een punt.
  Dat wordt gemeld.
- Typ een **0 voor het decimaalteken**: `0,50`, niet `,50`.
- Komt je CSV uit een Excel met Engelse instellingen (scheidingsteken `,`), dan is de punt het
  decimaalteken en geeft een komma in een bedrag een waarschuwing.

### Datums, tijden en lege waarden

- Schrijf een **datum** in de mutatiewaarde als `dd-mm-jjjj` (`01-11-2026`). Tip: geef de kolom in Excel
  de aangepaste opmaak `dd-mm-jjjj`.
- Schrijf een **tijd** als `uu:mm` (`08:30`) of, in een rubriek met seconden, `uu:mm:ss`.
- Een andere vorm, zoals `1-11-2026` of `8:30`, geeft bij een eigen rubriek een waarschuwing. Bij een
  standaardrubriek zet het hulpmiddel de waarde zelf in de goede vorm en meldt dat.
- Een **lege mutatiewaarde verwijdert** de huidige waarde van de rubriek in Beaufort. Dat kan de
  bedoeling zijn, maar ook een vergeten cel. Daarom krijg je altijd een waarschuwing.
- Een waarde van meer dan 50 tekens wordt afgekapt, met een waarschuwing.

### Standaardrubrieken

Voor een standaardrubriek controleert het hulpmiddel:

- **Dienstverband**: een dienstverbandrubriek heeft een dienstverband nodig (leeg of `0` is een fout),
  een persoonsrubriek mag er geen hebben.
- **S of V**: een rubriek die alleen stam of alleen variabel is, accepteert alleen `S` of alleen `V`.
- **Mutatiewaarde**, naar het soort rubriek:
  - getal: moet een getal zijn met niet meer cijfers en decimalen dan de rubriek heeft (`12.345` in een
    rubriek met 2 decimalen is een fout, want Beaufort zou er 12,34 van maken);
  - datum of tijd: moet een datum of tijd zijn;
  - tekst: precies zoals je hem typt (ook `12,50`), en niet langer dan de rubriek toestaat.
- **Geblokkeerd voor indirect muteren** (*Blok ind mut*): zo'n rubriek kan niet met DPIA100 worden
  geïmporteerd. Je krijgt een waarschuwing; voer de mutatie in Beaufort zelf in.

Het aantal standaardrubrieken en de datum van de lijst staan onderaan de pagina. Minimum- en
maximumwaarden en de waarden van referentietabellen controleert alleen Beaufort. Eigen rubrieken kunnen
(nog) niet worden opgegeven ([#24](https://github.com/japie06/dpia100-converter/issues/24)).

### Bijzondere tekens

HR Core accepteert alleen tekens die in de in Windows-1252 (ook wel ANSI of CP1252) tekenset staat. In Excel heet het "West-Europees (Windows)".

Letters met accenten (`é ë ö ç à ü` enz.), `€`, krullende aanhalingstekens (`’ “ ”`),
gedachtestreepjes (`– —`) en `…` zijn geen probleem: Beaufort slaat ze goed op.

Tekens die Beaufort niet kent, geven een fout. Bijvoorbeeld een pijl `→` (Word maakt die soms van `->`),
een vinkje `✓`, Poolse of Turkse letters (`ł`, `ş`) en emoji. 

## Als er iets misgaat

- **De pagina meldt *Deze pagina kan niet starten***: JavaScript is uitgeschakeld of geblokkeerd, of het
  bestand `dpia100.html` is gewijzigd. Gebruik Edge of Chrome en vraag anders je beheerder om een nieuw
  exemplaar.
- **Er komt geen venster *Opslaan als***: in andere browsers, of als het venster op jouw pc is uitgezet,
  komt het bestand in de map *Downloads*. Stond daar al een bestand met dezelfde naam, dan heet het
  nieuwe bijvoorbeeld `DPIA100_mutaties (1).txt`. Gebruik het nieuwste en gooi het oude weg.
- **Opnieuw controleren vraagt om het bestand**: in sommige browsers kan de pagina een gewijzigd bestand
  niet opnieuw lezen. Kies het dan nog een keer.
- **Iets anders, of een fout in het hulpmiddel zelf**: meld het bij je beheerder. Stuur nooit een CSV of
  importbestand met echte gegevens mee naar iemand buiten je organisatie.

## Wat het (nog) niet doet

Alleen persoons- en dienstverbandmutaties (PS en DV) met productiesoort NOR of COR. Andere soorten
mutaties (zoals bankgiro, loonverdeling en ziektegevallen), JWP en de verrekenperiode worden niet
ondersteund. Zie `field-layout.md` voor de technische details.

# Scriptversie (dpia100.ps1)

Het script doet precies hetzelfde als de pagina en maakt hetzelfde bestand. Alles onder
[Kolommen in de CSV](#kolommen-in-de-csv) en
[Fouten, waarschuwingen en meldingen](#fouten-waarschuwingen-en-meldingen) geldt ook hier. De teksten in
het venster zijn Engels.

`dpia100.bat` en `dpia100.ps1` moeten in **dezelfde map** staan.

1. **Sleep het CSV-bestand op `dpia100.bat`.** Dubbelklikken kan ook: je kiest dan zelf een bestand.
   Gebruik in Excel **niet** *CSV (MS-DOS)* of *CSV (Macintosh)*.
2. Controleer onder *Column mapping* welke kolom voor welk veld is gebruikt: links het veld, rechts de
   kolomnaam uit jouw CSV.
3. Staat er **[OK]**, dan beantwoord je de vragen:
   - **Procescode**: Enter voor `CMT`, of een andere code die in Beaufort bestaat als
     aanleverend proces.
   - **Creatiedatum**: Enter voor vandaag, of een datum (`dd mm jjjj`).
   - **Creatiegebruiker**: je Beaufort-gebruikersnaam (verplicht, max. 16 tekens).
   - Bij waarschuwingen: typ `Y` om het bestand toch te maken. Elk ander antwoord stopt zonder bestand.
4. Het bestand komt in **dezelfde map als je CSV-bestand**: `mutaties oktober.csv` wordt
   `DPIA100_mutaties oktober.txt`. Het volledige pad staat bij **[SUCCESS]**.

Zo ziet het venster eruit bij `voorbeeld.csv` (jouw antwoorden staan achter de dubbele punt):

```text
=== Beaufort DPIA100 Generator v1.4.1 ===
Input file        : H:\Mutaties\voorbeeld.csv
Output file       : H:\Mutaties\DPIA100_voorbeeld.txt
Input encoding    : Windows-1252
Standard rubrieken: 4227 (Beaufort export 30-10-2026)
Detected delimiter: ';'
Column mapping:
  Rubriek         <- 'Rubriek'
  Pers. Nr        <- 'Pers. Nr'
  Dienstverband   <- 'Dienstv'
  S of V          <- 'S of V'
  Mutatiewaarde   <- 'Mutatiewaarde'
  Ingangsdatum    <- 'Ingangsdatum'
  Productiesoort  <- 'Productiesoort'

[OK] 3 row(s) checked, no errors found.

Comma changed to dot in 1 Waarde value(s), because a numeric rubriek in Beaufort only accepts a dot (a text rubriek that is not a standard rubriek gets the dot too): row 3 '125,50' -> '125.50'.

Procescode [Press ENTER for 'CMT']:
Creatiedatum (dd mm yyyy) [Press ENTER for '12 11 2026']:
Creatiegebruiker [Mandatory]: TESTUSER

[SUCCESS] File successfully created:
H:\Mutaties\DPIA100_voorbeeld.txt
```

Als er iets misgaat met het script:

- **Oud bestand hernoemd**: na een mislukte poging wordt een eerder gemaakte `DPIA100_<naam>.txt`
  hernoemd naar `DPIA100_<naam>.txt.old`. Zo kun je niet per ongeluk een oud bestand importeren.
- **Script start niet** (*"is not digitally signed"*): je start `dpia100.ps1` vanuit een
  PowerShell-venster terwijl het bestand nog geblokkeerd is, omdat het van internet komt. Start het met
  `dpia100.bat`, of klik met de rechtermuisknop op `dpia100.ps1` → *Eigenschappen* → vink *Blokkering
  opheffen* aan.
- **Melding over *ConstrainedLanguage***: PowerShell is op deze pc dichtgezet. Neem contact op met IT.

# Voor beheerders

## Bijwerken

Een nieuwe versie zet je op dezelfde manier over de oude heen. Het versienummer staat onderaan de
webpagina en bovenaan in het venster van het script. Wat er is veranderd, staat in `docs/CHANGELOG.md`.

De lijst met standaardrubrieken bijwerken (bijvoorbeeld na een nieuwe Beaufort-versie): maak het
overzicht met functie 210006 en draai `pwsh tools\update-standard-rubrieken.ps1 -InputPath <overzicht>`.
Uitleg in [docs/rubrieken-export.md](docs/rubrieken-export.md). Zet het overzicht zelf nooit in de
repository.

## Ontwikkelen en testen

- Notities voor ontwikkelaars staan in `CLAUDE.md`, het bestandsformaat in `field-layout.md`. Openstaand
  werk staat in de [Issues](https://github.com/japie06/dpia100-converter/issues). Zet daar nooit
  persoonsgegevens in.
- Na elke wijziging: `pwsh tests\run-tests.ps1`. Dit test alle gevallen in `tests\cases` op
  PowerShell 5.1, PowerShell 7 en de webpagina (in Edge, zonder venster) en vergelijkt de uitvoer byte
  voor byte. Het controleert ook de beveiliging van de pagina.
- De webpagina heeft daarnaast unit- en klikroutetests: `npm ci` en dan `npm test` (Node.js LTS nodig,
  gebruikt de geïnstalleerde Edge).
- Een regel aanpassen doe je altijd in **beide** bestanden: `dpia100.ps1` en `dpia100.html`.
  `pwsh tests\web\differential.ps1` vergelijkt beide op willekeurige CSV-bestanden.
- Na een wijziging in een script- of stijlblok van `dpia100.html`: `pwsh tests\web\update-csp.ps1`.
  Anders weigert de browser de gewijzigde code (zie
  [docs/privacy-en-beveiliging.md](docs/privacy-en-beveiliging.md)).
- `tests\demo\` bevat CSV-bestanden die elke fout, waarschuwing en melding laten zien, met per regel een
  opmerking. Uitleg in `tests\demo\README.md`.
- `tests\beaufort\` bevat testbestanden om na te gaan wat Beaufort zelf accepteert, met de resultaten in
  `findings.md`.

# Licentie

[MIT](LICENSE). Dit is geen product van Visma of Youforce en wordt niet door hen ondersteund.
