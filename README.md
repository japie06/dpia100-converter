# DPIA100 converter

Dit script zet een CSV-bestand met mutaties om naar een `dpia100.txt` die je in Beaufort kunt importeren
(functie *Import externe indirecte mutaties*, 100254). Het doet hetzelfde als de knop **Maak DPIA100** in
het Excel-hulpmiddel van Raet, maar controleert de gegevens strenger voordat er een bestand wordt gemaakt.

## Wat het doet

- Leest een CSV-bestand (scheidingsteken `;`, `,` of tab wordt automatisch herkend).
- Controleert elke regel. Bij één of meer fouten wordt **geen** bestand gemaakt en zie je een lijst met
  alle fouten tegelijk (regelnummer, kolom, waarde en wat er mis is).
- Maakt een `dpia100.txt` in het vaste Beaufort-formaat (187 tekens per regel, Windows-1252), zodat
  letters als `ö` en `é` goed overkomen.

## Benodigde bestanden

Zet deze twee bestanden in **dezelfde map**:

- `dpia100.ps1`: het script
- `dpia100.bat`: het startbestand met deze inhoud:

```bat
@echo off
REM Drag and drop any .csv file onto this .bat file to process it.
REM Keep this file in the same folder as dpia100.ps1.
REM -NoPause: this launcher has its own pause below, so the script must not ask twice.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dpia100.ps1" -InputCsvPath "%~1" -NoPause
echo.
pause
```

## Gebruik

1. Sla je gegevens in Excel op als CSV (*CSV (gescheiden door lijstscheidingsteken)* of *CSV UTF-8*).
   Gebruik **niet** *CSV (MS-DOS)* of *CSV (Macintosh)*.
2. **Sleep het CSV-bestand op `dpia100.bat`.** Dubbelklikken op de .bat kan ook: je kunt dan zelf een
   bestand kiezen.
3. Beantwoord de drie vragen:
   - **Procescode**: Enter voor `CMT`, of typ een andere code (max. 3 tekens). Deze moet in Beaufort
     bestaan als aanleverend proces.
   - **Creatie datum**: Enter voor vandaag, of typ een datum (`dd mm jjjj`).
   - **Creatie gebruiker**: je Beaufort-gebruikersnaam (verplicht, max. 16 tekens).
4. Controleer in het venster welke CSV-kolom bij welk veld is gebruikt (*Column mapping*).
5. Bij **[SUCCESS]** staat het volledige pad van het gemaakte bestand in het venster. Zet
   `dpia100.txt` in de *files*-map van Beaufort en importeer het met functie 100254. Beoordeel daarna
   de mutaties (100251) en verwerk ze (100252).

   **Let op:** Beaufort vindt alleen bestanden waarvan de naam begint met `DPIA100`. Wil je het
   bestand hernoemen, laat de naam dan zo beginnen (bijvoorbeeld `DPIA100_oktober.txt`, niet
   `mutaties.txt`). Open en bewaar het bestand niet in Kladblok of een andere editor: dat kan de
   tekens of regeleinden veranderen.

## Kolommen in de CSV

De volgorde van de kolommen maakt niet uit. Extra kolommen worden genegeerd.

| Kolom (voorbeeldkop) | Verplicht | Regels |
|---|---|---|
| Rubriek | ja | Max. 6 letters/cijfers, bijvoorbeeld `E00101`. |
| Pers. Nr | ja | Alleen cijfers, max. 6. Voorloopnullen worden weggehaald. |
| Dienstv | nee | Alleen cijfers, max. 3. **Leeg = persoonsmutatie**, gevuld = dienstverbandmutatie. |
| S of V | ja | `S` (stam) of `V` (variabel). |
| Mutatie waarde | nee | Max. 50 tekens; langer wordt afgekapt met een waarschuwing. Bedragen met een komma of punt (`-12,50`), zonder duizendtallen. Verder vrij in te vullen; Beaufort controleert de waarde. |
| Ingangs datum | ja | Bijvoorbeeld `1 11 2026`, `01-11-2026`, `01/11/2026`, `2026-11-01` of `01112026`. |
| Productie soort | nee | Alleen `NOR`. Leeg wordt `NOR`. |

Let op bij **bedragen**: gebruik geen scheidingsteken voor duizendtallen. Beaufort leest `3.250` als
3,25. Schrijf `3250` of `3250,50`. Dit gebeurt ongemerkt als de cel in Excel is opgemaakt *met
scheidingsteken voor duizendtallen*. Het script kan niet zien of een rubriek een bedrag of tekst is, dus
het weigert zulke waarden niet, maar geeft een **waarschuwing**. Een komma wordt een punt (`12,50` wordt
`12.50`). In een CSV met `;` (zo slaat Nederlandse Excel op) is dat de gewone decimale komma en meldt
het script dit alleen. In een CSV met `,` (Engelse Excel) is de punt het decimaalteken: daar is `3.250`
gewoon 3,25 en geeft `1,234` een waarschuwing.

Zijn er waarschuwingen, dan vraagt het script of het bestand toch gemaakt moet worden. Typ `Y` om door
te gaan; elk ander antwoord stopt zonder bestand.

Lege regels (bijvoorbeeld `;;;;;;` onderaan) worden overgeslagen.

**Bijzondere tekens**: letters met accenten (`é ë ö ç à ü` enz.), `€`, krullende aanhalingstekens
(`’ “ ”`) en gedachtestreepjes (`–`) zijn geen probleem. Tekens die Beaufort niet kent worden
geweigerd met een foutmelding, bijvoorbeeld de `ĳ` als één teken (typ `ij`), Poolse of Turkse letters
(`ł`, `ş`) en emoji.

## Als er iets misgaat

- **Foutenlijst**: pas de genoemde regels aan in de CSV (of in het Excel-bestand) en probeer opnieuw.
  Het regelnummer telt de kopregel mee als regel 1.
- **Oud bestand hernoemd**: na een mislukte poging wordt een eerder gemaakte `dpia100.txt` hernoemd
  naar `dpia100.txt.old`. Zo kun je niet per ongeluk een oud bestand importeren.
- **Script start niet** (*"is not digitally signed"*): het script komt van internet of uit een e-mail.
  Klik met rechts op `dpia100.ps1` → *Eigenschappen* → vink *Blokkering opheffen* aan.
- **Melding over *ConstrainedLanguage***: PowerShell is op deze pc dichtgezet. Neem contact op met IT.

## Wat het (nog) niet doet

Alleen persoons- en dienstverbandmutaties (PS en DV) met productiesoort NOR. Bankgiro, inzet,
loonverdeling, speciale doorbelasting, ziektegevallen, JWP en de verrekenperiode worden niet
ondersteund. Rubrieken worden niet gecontroleerd tegen een lijst; Beaufort keurt een onjuiste rubriek
af bij het importeren. Zie `field-layout.md` voor de technische details.

## Voor beheerders

- Wijzigingen staan in `docs/CHANGELOG.md`; notities voor ontwikkelaars in `CLAUDE.md`.
- Na elke wijziging aan het script: `pwsh tests\run-tests.ps1`. Dit test alle gevallen in
  `tests\cases` op PowerShell 5.1 en 7 en vergelijkt de uitvoer byte voor byte.
- `tests\beaufort\` bevat testbestanden om na te gaan wat Beaufort zelf accepteert (bijzondere tekens,
  regeleinden), met de resultaten in `findings.md`.
