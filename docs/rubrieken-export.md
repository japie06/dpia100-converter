# Rubrieken uit Beaufort: het overzicht (functie 210006)

Dit document beschrijft hoe we de rubrieken van Beaufort als gegevens krijgen en waarvoor we ze willen
gebruiken. Het is de basis voor een latere wijziging van `dpia100.ps1` en `dpia100.html`. **Die wijziging is
nog niet gemaakt.** Wat er nu is: het script `tools/convert-rubrieken.ps1` dat het overzicht opschoont, met tests.

## Waarom

Tot nu toe kent de converter geen rubrieken: Beaufort zelf controleert de rubriek bij het verwerken. Het
overzicht laat zien dat dat onderscheid wel te maken is:

- **Standaardrubrieken** (kolom *Fix* = `J`) komen met Beaufort mee. De codes beginnen met `P`, `R` of `S`. Ze zijn
  voor elke gemeente gelijk (voor dezelfde Beaufort-versie), dus we kunnen ze in de converter opnemen, zoals de
  Excel-tool dat ook doet.
- **Eigen rubrieken** (*Fix* = `N`) maakt de gebruiker zelf aan (E, F, K, W, C ...). Die staan niet in de
  standaardlijst en moeten apart worden aangeleverd (zie *Vervolg*).

Wat we ermee willen controleren:

| Gegeven | Gebruik |
|---|---|
| `Blok ind mut` = `J` | De rubriek is geblokkeerd voor indirect muteren. Dat is wat DPIA100 doet, dus **zo'n rubriek kan niet worden geimporteerd**. Eigenaar: geen Beaufort-test nodig. |
| `Geg.soort` PS of DV | Past de rubriek bij het wel of niet invullen van Dienstverband? Nu meldt pas Beaufort het (PI_1650). |
| `Typ`, `Len`, `Dec` | Later: lengte en aantal decimalen van Waarde controleren (nu alleen een waarschuwing boven 50 tekens). |
| `Minimum`, `Maximum`, `Bereik` | Later: waardebereik. |
| `Referentie` | Waarde moet in een referentietabel staan; die waarden zitten **niet** in het overzicht. |

Privacy: in de converter komen alleen rubriekcodes en hun eigenschappen, **nooit waarden** uit het overzicht, en
niets over de gemeente of de gebruiker die het overzicht heeft gemaakt.

## Het overzicht maken

In Beaufort functie **210006 Overzicht rubrieken** starten, zonder selectiecriteria (de converter waarschuwt als er
wel criteria waren ingevuld, want dan is de lijst niet compleet) en als tekstbestand opslaan. Het bestand is groot
(8,9 MB, 106.000 regels voor 7.942 rubrieken) en bevat de rubrieken van deze ene installatie. **Zet het niet in de
repository**; `.gitignore` sluit `Overzicht*Rubrieken*.txt` en `rubrieken*.csv` uit.

## Hoe het bestand eruitziet

UTF-8 zonder BOM, CRLF. Het is een afdruk in pagina's: per pagina een kop (gemeente, titel, paginanummer), bij de
eerste pagina het selectievak, de kolomkop, de rubrieken en een voet (`HR Core - OPIA003 - <gebruiker> <datum>`).
In het voorbeeld 1.608 pagina's.

Per rubriek:

1. Een **recordregel** met vaste kolommen. Is de omschrijving te lang, dan staat de rest op een vervolgregel
   (9 spaties inspringen).
2. Een **detailblok** van zes regels (Ingedeeld, Tabel, Veld, Logging, Groep code, Exportkenmerken, met aan de
   rechterkant Referentie, Min/Max, Salarisaanlev, TWK ...).

De kolomposities (0-gebaseerd) komen uit de kolomkop; het script leest ze uit die regel in plaats van ze vast te
leggen:

| Kolom | Positie | Opmerking |
|---|---|---|
| Code | 1 | 3 tot 6 tekens (zie hieronder) |
| Omschrijving | 9 | 26 tekens, rest op een vervolgregel |
| Fix | 35 | waarde op 36 |
| Prompt | 39 | |
| Rubriek soort | 60 | in dit overzicht overal leeg |
| Geg.soort | 68 | `DV`, `PS`, `RG`, `IN`, `ZG` ... |
| Typ | 74 | `N` numeriek, `S` tekst, `D` datum, `H`, `T`, `I` |
| Len | 78 | inclusief komma en decimalen |
| Dec | 82 | alleen numeriek |
| Opv trm | 86 | opvraagtermijn historie in maanden, mag leeg zijn |
| Stam, Var, Verv, Toek Rubr, Blok toek, Mut | 91, 96, 100, 105, 110, 115 | J/N |
| **Blok ind mut** | 119 | J/N, **mag leeg zijn** |

In het detailblok staan `Referentie J/N` met de tabelnaam, `Min/Max J/N` met `Minimum`, `Maximum`, en `Bereik`
(`+` alleen positief, `-` alleen negatief).

Betekenis van de velden: zie de Beaufort-help *Rubrieken wijzigen of verwijderen* (Youforce 2026-09, lokaal, niet
in de repository).

## Opschonen

```
pwsh tools\convert-rubrieken.ps1 -InputPath C:\Temp\OverzichtRubrieken.txt
pwsh tools\convert-rubrieken.ps1 -InputPath C:\Temp\OverzichtRubrieken.txt -OnlyScope -OnlyStandard -OutputPath C:\Temp\standaard.csv
```

- Werkt op Windows PowerShell 5.1 en 7 met dezelfde uitvoer.
- Schrijft `;`-gescheiden UTF-8 met BOM, zodat Excel het goed opent. Kolommen: `Code;Omschrijving;Fix;Prompt;
  Gegsoort;Typ;Len;Dec;BlokIndMut;Referentie;RefTabel;Minimum;Maximum;Bereik`, gesorteerd op code.
- `-OnlyScope`: alleen `PS` en `DV`. `-OnlyStandard`: alleen *Fix* = `J`.
- **Stopt zonder uitvoer** (exit 1) bij alles wat het niet herkent: geen kolomkop, een kolomkop die per pagina
  verschilt, geen geldige UTF-8, een recordregel die niet te lezen is, een onbekende waarde in `Blok ind mut`,
  dubbele codes, of een ander aantal detailblokken dan rubrieken. Zo verandert er niets stilzwijgend als Beaufort
  het formaat wijzigt.
- Een lege waarde in `Blok ind mut` blijft leeg; het script maakt er geen `N` van.
- Tests: `pwsh tests\rubrieken\run-tests.ps1`. Het voorbeeldbestand `tests/rubrieken/sample-export.txt` is
  verzonnen (gemeente *Voorbeeldstad*, gebruiker *TESTUSER*); zet er nooit regels uit een echt overzicht in.

## Wat erin zit (overzicht van 2026-10-30, Beaufort-productie)

Meetwaarden van het script, ter controle bij een volgende export:

| | Aantal |
|---|---|
| Rubrieken | 7.942 (geen dubbele codes), 1.608 pagina's |
| Standaard (*Fix* = J) / eigen (N) | 7.433 / 509 |
| Gegevenssoort | DV 4.332, RG 2.498, PS 370, IN 123, ZG 117, ZS 36, PL 33, KU 28 ... |
| **PS + DV** | **4.702** (standaard 4.227, eigen 475) |
| PS + DV naar Typ | N 3.198, S 803, D 691, H 7, T 2, I 1 |
| PS + DV met Referentie J / Min-Max | 679 / 801 |
| **PS + DV met Blok ind mut J / N / leeg** | **579 / 4.113 / 10** (alle 579 en alle 10 lege zijn standaard; geen enkele eigen rubriek is geblokkeerd) |

Opvallend:

- Codes zijn niet altijd 6 tekens: 59 rubrieken hebben 3 tot 5 tekens (31 daarvan PS of DV, 1 standaard). De regel
  "rubriek is precies 6 tekens" in `dpia100` klopt dus voor bijna alles, maar niet voor deze; kijk hiernaar voor
  de wijziging.
- 10 standaardrubrieken in PS/DV (en 5 daarbuiten) hebben een lege `Blok ind mut`, bijvoorbeeld P00633. Het
  overzicht zegt niet wat dat betekent.
- Alle standaardcodes beginnen met `P` (4.215), `R` (11) of `S` (1) in PS/DV.

## Vervolg (stap 2, nog niet gemaakt; eerst akkoord van de eigenaar)

1. **Tabel genereren** uit `-OnlyScope -OnlyStandard`: per rubriek code, PS/DV en `Blok ind mut`, plus bron en
   datum van het overzicht. Zowel in `dpia100.ps1` als in `dpia100.html` ingebakken (in de pagina een datablok
   met nieuwe CSP-hash, `update-csp.ps1`), gegenereerd door hetzelfde script zodat ze gelijk blijven.
2. **Regels** (elk een genummerde beslissing, ook in `tests/web/harness.js`, `core.test.mjs`, `tests/cases` en de
   differential-test):
   - Standaardrubriek met `Blok ind mut` = J: fout, kan niet worden geimporteerd.
   - Standaardrubriek: PS/DV volgt uit de tabel; Dienstverband dat daar niet bij past: fout.
   - Onbekende code: toegestaan (kan een eigen rubriek zijn), geen fout.
3. **Eigen rubrieken**: de gebruiker levert zijn eigen overzicht of een kleine CSV aan tijdens het gebruik
   (optioneel `-RubriekenPath` en een bestandskeuze in de pagina). Niets wordt opgeslagen (zie de privacyregels in
   `CLAUDE.md`).
4. Open vragen: wat betekent een lege `Blok ind mut` (voorlopig: niet als geblokkeerd behandelen, en geen
   waarschuwing)? Hoe oud mag de standaardlijst worden (versie en datum in de uitvoer van de converter tonen)?
   `CLAUDE.md` ("No rubriek lookup list") en `.NOT_IN_CURRENT_SPEC C` aanpassen.
