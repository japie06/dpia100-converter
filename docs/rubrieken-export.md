# Rubrieken uit Beaufort: het overzicht (functie 210006)

Dit document beschrijft hoe we de rubrieken van Beaufort als gegevens krijgen en hoe de converter ze
gebruikt. Er zijn twee scripts: `tools/convert-rubrieken.ps1` schoont het overzicht op tot een CSV, en
`tools/update-standard-rubrieken.ps1` zet de standaardrubrieken daaruit in `dpia100.ps1` en `dpia100.html`
(beslissing 17 in de kop van `dpia100.ps1`). Beide hebben tests in `tests/rubrieken`.

## Waarom

Tot versie 1.2 kende de converter geen rubrieken: Beaufort zelf controleerde de rubriek bij het
verwerken. Het overzicht laat zien dat er onderscheid te maken is:

- **Standaardrubrieken** (kolom *Fix* = `J`) komen met Beaufort mee. De codes beginnen met `P`, `R` of `S`. Ze zijn
  voor elke gemeente gelijk (voor dezelfde Beaufort-versie), dus we kunnen ze in de converter opnemen, zoals de
  Excel-tool dat ook doet.
- **Eigen rubrieken** (*Fix* = `N`) maakt de gebruiker zelf aan (E, F, K, W, C ...). Die staan niet in de
  standaardlijst. Een code die niet in de lijst staat, laat de converter zonder melding toe; eigen
  rubrieken opgeven kan nog niet (ideeën in issue #24).

Van een standaardrubriek kan de klant alleen de rubrieksoort, tabbladomschrijving, minimum en maximum,
toekomstmutaties, mutatieverantwoording, opvraagtermijn en de waarden van de referentietabel wijzigen
(Beaufort-help *Rubrieken wijzigen of verwijderen*). Gegevenssoort, datatype, lengte, decimalen, stam/variabel
en *Blok ind mut* liggen vast. Alleen die komen in de converter:

| Gegeven | Gebruik in de converter |
|---|---|
| `Geg.soort` PS of DV | Dienstverband moet erbij passen; anders een fout (beslissing 18). Een DV-rubriek zonder dienstverband keurt Beaufort af (PI_1650). |
| `Stam`, `Var` | S of V moet door de rubriek zijn toegestaan; anders een fout (beslissing 18). |
| `Blok ind mut` = `J` | De rubriek is geblokkeerd voor indirect muteren, dus **niet te importeren met DPIA100** (eigenaar: geen Beaufort-test nodig). Een waarschuwing, geen fout (beslissing 19). Leeg telt als niet geblokkeerd. |
| `Typ`, `Len`, `Dec` | De waarde wordt per type gecontroleerd: getal (`N`), datum (`D`), tijd (`T` `uu:mm`, `U` `uu:mm:ss`) of tekst (beslissing 20). |
| `Minimum`, `Maximum`, `Bereik` | Niet gebruikt: de klant kan ze wijzigen. |
| `Referentie` | Niet gebruikt: de waarden van de referentietabel staan **niet** in het overzicht en zijn per klant. |

Privacy: in de converter komen alleen rubriekcodes en hun vaste eigenschappen, **nooit waarden of
omschrijvingen** uit het overzicht, en niets over de gemeente of de gebruiker die het overzicht heeft
gemaakt.

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
  Gegsoort;Typ;Len;Dec;Stam;Var;BlokIndMut;Referentie;RefTabel;Minimum;Maximum;Bereik`, gesorteerd op code.
- `-OnlyScope`: alleen `PS` en `DV`. `-OnlyStandard`: alleen *Fix* = `J`.
- **Stopt zonder uitvoer** (exit 1) bij alles wat het niet herkent: geen kolomkop, een kolomkop die per pagina
  verschilt, geen geldige UTF-8, een recordregel die niet te lezen is, een onbekende waarde in `Blok ind mut`,
  `Stam` of `Var`, dubbele codes, of een ander aantal detailblokken dan rubrieken. Zo verandert er niets
  stilzwijgend als Beaufort het formaat wijzigt.
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
| Standaard PS + DV naar Stam / Var | alleen stam 2.092, alleen variabel 428, beide 1.707 |
| Standaard PS + DV naar Typ | N 3.035, D 603, S 586, H 2, I 1 (geen `T` of `U`; de eigen rubrieken hebben wel `T`) |

Opvallend:

- Codes zijn niet altijd 6 tekens: 59 rubrieken hebben 3 tot 5 tekens (31 daarvan PS of DV, 1 standaard:
  `SETKB`). Daarom accepteert de converter 1 tot 6 tekens en zet hij een kortere code links in positie 4-9
  (eigenaar). Dat is nog niet in Beaufort getest (issue #25).
- 10 standaardrubrieken in PS/DV (en 5 daarbuiten) hebben een lege `Blok ind mut`, bijvoorbeeld P00633. Het
  overzicht zegt niet wat dat betekent; de converter behandelt leeg als niet geblokkeerd (eigenaar).
- Standaardcodes in PS/DV beginnen met `P` (4.215), `R` (11) of `S` (1), maar ook 3 eigen rubrieken beginnen
  met `R`. De converter leidt dus niets af uit de eerste letter: alleen de lijst telt.
- Alle datumrubrieken hebben lengte 11.

## In de converter

`tools/update-standard-rubrieken.ps1` maakt de lijst en zet hem in beide bestanden:

```
pwsh tools\update-standard-rubrieken.ps1 -InputPath C:\Temp\OverzichtRubrieken.txt
```

- Gebruikt `convert-rubrieken.ps1 -OnlyScope -OnlyStandard` en haalt de exportdatum uit de paginavoet
  (anders `-ExportDate dd-MM-yyyy`).
- Schrijft per rubriek één regel met vaste kolommen, bijvoorbeeld `P00001 DV N  6 2 SV N`: code,
  gegevenssoort, type, lengte, decimalen (`-` = leeg), toegestaan `S-` (alleen stam), `-V` (alleen
  variabel) of `SV`, en *Blok ind mut* (`J`, `N`, `-` = leeg).
- Vervangt het gebied tussen `BEGIN STANDARD RUBRIEKEN` en `END STANDARD RUBRIEKEN` in `dpia100.ps1` (een
  here-string) en in het core-script van `dpia100.html` (een `String.raw`-tekst). In de pagina staat de lijst
  in het core-script zelf, niet in een apart blok: de pagina mag maar drie inline blokken hebben (CSP).
  Daarna werkt het de CSP-hashes bij.
- Stopt zonder iets te wijzigen bij een onverwachte waarde of als een van de twee gebieden ontbreekt.
- De test `page-rubrieken` (`tests/run-tests.ps1`) controleert dat de lijst en de datum in beide bestanden
  gelijk zijn.

Het script toont `Standard rubrieken: 4227 (Beaufort export 30-10-2026)`; de pagina toont hetzelfde onderaan
en bij de controle. Er is geen waarschuwing voor een oude lijst: een nieuwere standaardrubriek wordt dan
behandeld als eigen rubriek, dus toegestaan. Bij een nieuwe Beaufort-versie: overzicht opnieuw maken,
script draaien, de aantallen vergelijken met de tabel hierboven, tests draaien en een regel in
`docs/CHANGELOG.md` zetten.

## Eigen rubrieken

Nog niet ingebouwd. Ideeën om ze te laten aanleveren of bewaren, met de eisen, staan in issue #24.
