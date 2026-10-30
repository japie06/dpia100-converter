# DPIA100 converter

[![CI](https://github.com/japie06/dpia100-converter/actions/workflows/ci.yml/badge.svg)](https://github.com/japie06/dpia100-converter/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/japie06/dpia100-converter)](https://github.com/japie06/dpia100-converter/releases/latest)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Dit hulpmiddel zet een CSV-bestand met mutaties om naar een bestand `DPIA100_<naam>.txt` dat je in
Beaufort / HR Core kunt importeren (functie *Import externe indirecte mutaties*, 100254). Het doet
hetzelfde als de knop **Maak DPIA100** in het Excel-hulpmiddel, maar controleert de gegevens strenger
voordat er een bestand wordt gemaakt.

Er zijn twee manieren, die precies hetzelfde doen en hetzelfde bestand maken:

- **In de browser** (aanbevolen): een webpagina, `dpia100.html`.
- **Met een script**: sleep de CSV op `dpia100.bat` (PowerShell).

## Wat het doet

- Leest een CSV-bestand dat je vanuit Excel hebt opgeslagen. (scheidingsteken ;, , of tab wordt automatisch herkend).
- Controleert elke regel. Bij één of meer fouten wordt **geen** bestand gemaakt en zie je een lijst met
  alle fouten.
- Maakt het importbestand in het vaste formaat dat Beaufort verwacht. De technische details staan in
  `field-layout.md`.

## Installeren

Voor de browserversie via internet hoef je niets te installeren: open
<https://japie06.github.io/dpia100-converter/>. Wil je de bestanden zelf hebben (bijvoorbeeld op een
netwerkschijf, of om zonder internet te werken):

1. Download het zip-bestand `dpia100-converter-vX.Y.Z.zip` van de nieuwste versie onder
   [**Releases**](https://github.com/japie06/dpia100-converter/releases/latest).
2. Klik met de rechtermuisknop op het zip-bestand → *Eigenschappen* → vink *Blokkering opheffen* aan →
   *OK*. Zie je dat vinkje niet, dan is deze stap niet nodig.
3. Pak het zip-bestand uit in een eigen map. Daarin staan:
   - `dpia100.html`: de browserversie. Dubbelklik erop om hem te openen.
   - `dpia100.bat`: het startbestand voor de scriptversie.
   - `dpia100.ps1`: het script zelf. Laat dit in **dezelfde map** staan als `dpia100.bat`.
   - `voorbeeld.csv`: een voorbeeld van een goed CSV-bestand.
   - `README.md` en `LICENSE`: deze handleiding en de licentie.

Het versienummer staat onderaan de webpagina en bovenaan in het venster van het script.

## Gebruik in de browser

1. Sla je gegevens in Excel op als CSV (*CSV (gescheiden door lijstscheidingsteken)* of *CSV UTF-8*).
   Open `voorbeeld.csv` in Excel om te zien hoe het eruit moet zien.
2. Open de pagina in **Microsoft Edge** of **Google Chrome**: via
   <https://japie06.github.io/dpia100-converter/>, of dubbelklik op `dpia100.html`.
3. **Stap 1, Kies je CSV-bestand**: sleep het bestand op de pagina, of klik op het vak om het te kiezen.
4. **Stap 2, Controle**: de pagina controleert meteen het hele bestand.
   - Zijn er **fouten**, dan zie je een tabel met de regel, de kolom, de waarde en wat er mis is, met een
     tip. Er wordt dan geen bestand gemaakt. Pas de regels aan in Excel, sla de CSV op en klik op
     **Opnieuw controleren**. Je hoeft het bestand niet opnieuw te kiezen.
   - Onder *Welke kolom is gebruikt?* zie je welke kolom uit je CSV bij welk veld hoort. Let vooral op
     *Mutatiewaarde*: die kolom kan niet worden gecontroleerd.
   - **Waarschuwingen** (zie [Bedragen](#bedragen)) blokkeren niet, maar je moet ze bij stap 3 bevestigen.
   - Met *Bekijk de mutaties* zie je precies wat er in het bestand komt.
5. **Stap 3, Bestand maken** (alleen als er geen fouten zijn): vul de **procescode** (meestal `CMT`), de
   **creatiedatum** (standaard vandaag) en je **gebruikersnaam in Beaufort** in. Vink *Onthoud* aan, dan
   hoef je procescode en gebruikersnaam de volgende keer niet opnieuw te typen. Klik op **Maak
   DPIA100_&lt;naam&gt;.txt**.
   - In Edge en Chrome verschijnt een venster *Opslaan als*: kies de map, bijvoorbeeld de
     *Bestandsuitwisseling*. De naam moet met `DPIA100` beginnen, anders slaat de pagina niets op.
   - In andere browsers (of als het opslagvenster op jouw pc is uitgezet) komt het bestand in de map
     *Downloads*. Stond daar al een bestand met dezelfde naam, dan heet het nieuwe bestand bijvoorbeeld
     `DPIA100_mutaties (1).txt`. Gebruik dan dat nieuwste bestand en gooi het oude weg.
6. Importeer het bestand in HR Core met functie 100254, beoordeel de mutaties (100251) en verwerk ze
   (100252).

**Privacy.** De pagina werkt volledig op je eigen computer: je CSV-bestand wordt niet verstuurd, ook niet
als je de pagina via internet opent. Een beveiligingsregel in de pagina blokkeert elk netwerkverkeer;
de pagina werkt ook zonder internetverbinding. Gebruik op de pagina geen vertaalfunctie of AI-assistent
(zoals Copilot), want die sturen de inhoud wel door. Uitleg voor IT en privacy:
[docs/privacy-en-beveiliging.md](docs/privacy-en-beveiliging.md).

## Gebruik met dpia100.bat

1. Sla je gegevens in Excel op als CSV (*CSV (gescheiden door lijstscheidingsteken)* of *CSV UTF-8*).
   Gebruik **niet** *CSV (MS-DOS)* of *CSV (Macintosh)*. Open `voorbeeld.csv` in Excel om te zien hoe
   het eruit moet zien.
2. **Sleep het CSV-bestand op `dpia100.bat`.** Dubbelklikken op `dpia100.bat` kan ook: je kunt dan zelf
   een bestand kiezen.
3. Het script controleert eerst het hele CSV-bestand. Zijn er fouten, dan zie je meteen de lijst met
   fouten (zie [Als er iets misgaat](#als-er-iets-misgaat)) en worden er geen vragen gesteld.
4. Controleer onder *Column mapping* welke kolom voor welk veld is gebruikt. Het script herkent de
   kolommen aan de kopregel (de eerste regel, met de kolomnamen). Links staat het veld, rechts de
   kolomnaam uit jouw CSV. Klopt een koppeling niet, sluit het venster, pas de kolomnaam in Excel aan
   en probeer het opnieuw. Let vooral op *Mutatiewaarde*: die kolom kan het script niet controleren.
5. Staat er **[OK]**, dan beantwoord je de drie vragen:
   - **Procescode**: druk op Enter voor `CMT`, of typ een andere code (max. 3 tekens). Deze moet in
     Beaufort bestaan als aanleverend proces.
   - **Creatiedatum**: druk op Enter voor vandaag, of typ een datum (`dd mm jjjj`).
   - **Creatiegebruiker**: je Beaufort-gebruikersnaam (verplicht, max. 16 tekens).

   Waren er waarschuwingen (zie [Bedragen](#bedragen)), dan vraagt het script daarna nog of het
   bestand toch gemaakt moet worden.
6. Het bestand komt in **dezelfde map als je CSV-bestand**, met de naam van de CSV erin:
   `mutaties oktober.csv` wordt `DPIA100_mutaties oktober.txt`. Het volledige pad staat in het venster
   bij *Output file* en bij **[SUCCESS]**. Zet het bestand in de *Bestandsuitwisseling* naar HR Core en
   importeer het met functie 100254. Beoordeel daarna de mutaties (100251) en verwerk ze (100252).

   **Let op:** HR Core vindt alleen bestanden waarvan de naam begint met `DPIA100`. Wil je het
   bestand hernoemen, laat de naam dan zo beginnen (bijvoorbeeld `DPIA100_oktober.txt`, niet
   `mutaties.txt`). Het bestand bekijken mag, maar sla het niet op in Kladblok of een andere editor:
   opslaan kan de tekens of regeleinden veranderen.

Zo ziet het venster eruit bij `voorbeeld.csv` (jouw antwoorden staan achter de dubbele punt):

```text
=== Beaufort DPIA100 Generator v1.2.0 ===
Input file        : H:\Mutaties\voorbeeld.csv
Output file       : H:\Mutaties\DPIA100_voorbeeld.txt
Input encoding    : Windows-1252
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

Comma changed to dot in 1 Waarde value(s), because a numeric rubriek in Beaufort only accepts a dot (a text rubriek gets the dot too): row 3 '125,50' -> '125.50'.

Procescode [Press ENTER for 'CMT']:
Creatiedatum (dd mm yyyy) [Press ENTER for '28 10 2026']:
Creatiegebruiker [Mandatory]: JANSENP

[SUCCESS] File successfully created:
H:\Mutaties\DPIA100_voorbeeld.txt
```

## Kolommen in de CSV

Zo ziet `voorbeeld.csv` eruit in Excel:

| Rubriek | Pers. Nr | Dienstv | S of V | Mutatiewaarde | Ingangsdatum | Productiesoort |
|---|---|---|---|---|---|---|
| E00101 | 123456 | | S | Café de Brug | 1-11-2026 | NOR |
| E00205 | 123456 | 1 | V | 125,50 | 1-11-2026 | NOR |
| E00310 | 2345 | 2 | S | 36 | 1-12-2026 | |

De eerste regel is een persoonsmutatie (*Dienstv* is leeg), de andere twee zijn
dienstverbandmutaties. De volgorde van de kolommen maakt niet uit. Extra kolommen worden genegeerd.

| Kolom (voorbeeldkop) | Verplicht | Toelichting |
|---|---|---|
| Rubriek | ja | Max. 6 letters/cijfers, bijvoorbeeld `E00101`. |
| Pers. Nr | ja | Alleen cijfers, max. 6. Voorloopnullen worden weggehaald. |
| Dienstv | nee | Alleen cijfers, max. 3. **Leeg = persoonsmutatie**, gevuld = dienstverbandmutatie. |
| S of V | ja | `S` (stam) of `V` (variabel). |
| Mutatiewaarde | nee | Max. 50 tekens; langer wordt afgekapt met een waarschuwing. Bedragen met een komma of punt als decimaalteken (`-12,50`), zonder scheidingsteken voor duizendtallen. Verder vrij in te vullen; Beaufort controleert de waarde. |
| Ingangsdatum | ja | Bijvoorbeeld `1 11 2026`, `01-11-2026`, `01/11/2026`, `2026-11-01` of `01112026`. |
| Productiesoort | nee | Alleen `NOR`. Leeg wordt `NOR`. |

### Bedragen

Typ bedragen **zonder punt voor duizendtallen**: schrijf `3250` of `3250,50`, niet `3.250` of
`3.250,50`. Beaufort leest `3.250` als 3,25.

Zo'n punt komt er ongemerkt in als de cel in Excel is opgemaakt *met scheidingsteken voor
duizendtallen*. Kies in dat geval een opmaak zonder dat scheidingsteken.

Het hulpmiddel kan niet zien of een rubriek een bedrag of een tekst bevat. Daarom weigert het zulke
waarden niet, maar geeft het een **waarschuwing**. Zijn er waarschuwingen, dan wordt het bestand pas
gemaakt als je dat bevestigt: in de browser met het vinkje bij stap 3, in het script door `Y` te typen
(elk ander antwoord stopt zonder bestand).

Een decimale komma wordt een punt (`12,50` wordt `12.50`), want Beaufort accepteert alleen een punt. Dat
wordt altijd gemeld.

> **Voor wie het precies wil weten:** het script kijkt naar het scheidingsteken van de CSV. Een CSV met
> `;` (Excel met Nederlandse instellingen) gebruikt de komma als decimaalteken. Een CSV met `,` (Excel
> met Engelse instellingen) gebruikt de punt: daar is `3.250` gewoon 3,25 en geeft `1,234` een
> waarschuwing.

### Bijzondere tekens

Letters met accenten (`é ë ö ç à ü` enz.) zijn geen probleem.

**Let op:** uit een test blijkt dat Beaufort `€`, krullende aanhalingstekens (`’ “ ”`),
gedachtestreepjes (`– —`), het weglatingsteken `…` en een paar zeldzame tekens (zoals `Š ž ™ ÿ`)
**zonder melding weglaat**: `d’Ancona` wordt `dAncona`. Gebruik daarom gewone tekens: `'`, `"`, `-`,
`...` en `EUR`. Word en Excel maken van `'` en `-` vaak automatisch een krullend aanhalingsteken of een
gedachtestreepje; let daar op. Dit wordt nog nagekeken
([#20](https://github.com/japie06/dpia100-converter/issues/20)).

Tekens die Beaufort helemaal niet kent, worden door het script geweigerd met een foutmelding. Bijvoorbeeld een pijl `→` (Word maakt die soms automatisch van `->`),
een vinkje `✓`, Poolse of Turkse letters (`ł`, `ş`) en emoji.

## Als er iets misgaat

- **Foutenlijst**: pas de genoemde regels aan in de CSV (of in het Excel-bestand) en probeer het
  opnieuw. Regel 1 is de kopregel; de eerste regel met gegevens is regel 2.
- **Oud bestand hernoemd** (script): na een mislukte poging wordt een eerder gemaakte
  `DPIA100_<naam>.txt` hernoemd naar `DPIA100_<naam>.txt.old`. Zo kun je niet per ongeluk een oud
  bestand importeren. De browserversie raakt geen bestaande bestanden aan; zij slaat pas iets op als jij
  op de knop klikt.
- **De pagina meldt *Deze pagina kan niet starten***: JavaScript is uitgeschakeld of geblokkeerd in de
  browser. Gebruik Edge of Chrome, of vraag je IT-afdeling. Het kan ook betekenen dat het bestand
  `dpia100.html` is gewijzigd; download het dan opnieuw.
- **Opnieuw controleren vraagt om het bestand**: in sommige browsers kan de pagina een gewijzigd bestand
  niet opnieuw lezen. Kies het dan nog een keer.
- **Script start niet** (*"is not digitally signed"*): je start `dpia100.ps1` vanuit een
  PowerShell-venster terwijl de bestanden nog geblokkeerd zijn, omdat ze van internet komen. Start het
  script met `dpia100.bat`, of hef de blokkering op zoals bij [Installeren](#installeren), stap 2. Heb
  je de zip al uitgepakt, doe het dan bij `dpia100.ps1`: klik met de rechtermuisknop →
  *Eigenschappen* → vink *Blokkering opheffen* aan.
- **Melding over *ConstrainedLanguage***: PowerShell is op deze pc dichtgezet. Neem contact op met IT.

## Wat het (nog) niet doet

Alleen persoons- en dienstverbandmutaties (PS en DV) met productiesoort NOR. Bankgiro, inzet,
loonverdeling, speciale doorbelasting, ziektegevallen, JWP en de verrekenperiode worden niet
ondersteund. Rubrieken worden niet gecontroleerd aan de hand van een lijst; Beaufort keurt een onjuiste
rubriek af bij het importeren. Zie `field-layout.md` voor de technische details.

## Voor beheerders

- Wijzigingen staan in `docs/CHANGELOG.md`; notities voor ontwikkelaars in `CLAUDE.md`. Openstaand
  werk staat in de [Issues](https://github.com/japie06/dpia100-converter/issues).
- Na elke wijziging: `pwsh tests\run-tests.ps1`. Dit test alle gevallen in `tests\cases` op
  PowerShell 5.1, PowerShell 7 en de webpagina (in Edge, zonder venster) en vergelijkt de uitvoer byte
  voor byte. Het controleert ook de beveiliging van de pagina.
- De webpagina heeft daarnaast unit- en klikroutetests: `npm ci` en dan `npm test` (Node.js LTS nodig,
  gebruikt de geïnstalleerde Edge; niet nodig om de converter te gebruiken).
- Een regel aanpassen doe je altijd in **beide** bestanden: `dpia100.ps1` en `dpia100.html`.
  `pwsh tests\web\differential.ps1` vergelijkt beide op willekeurige CSV-bestanden.
- Na een wijziging in een script- of stijlblok van `dpia100.html`: `pwsh tests\web\update-csp.ps1`.
  Anders weigert de browser de gewijzigde code (zie
  [docs/privacy-en-beveiliging.md](docs/privacy-en-beveiliging.md)).
- `tests\beaufort\` bevat testbestanden om na te gaan wat Beaufort zelf accepteert (bijzondere tekens,
  regeleinden), met de resultaten in `findings.md`.

## Licentie

[MIT](LICENSE). Dit is geen product van Visma of Youforce en wordt niet door hen ondersteund.
