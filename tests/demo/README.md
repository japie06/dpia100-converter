# Demobestanden: alle meldingen van de converter

Deze CSV-bestanden laten elke fout, waarschuwing en melding van de converter zien. Gebruik ze om de
converter zelf te bekijken (script of webpagina), of om een collega te laten zien hoe de meldingen eruitzien.
De kolom **Opmerking** zegt wat een regel laat zien; de converter gebruikt die kolom niet.

| Bestand | Wat je ziet |
|---|---|
| `1-fouten.csv` | Elke fout in een regel, één fout per regel. Regel 2 is goed, ter vergelijking. Er wordt geen bestand gemaakt. |
| `2-waarschuwingen.csv` | Alle waarschuwingen en meldingen van een bestand met `;` (Nederlandse Excel): komma wordt punt, datum in de vorm van de rubriek, duizendtallen, afkappen boven 50 tekens, geblokkeerde rubriek, lege waarde, datum en tijd in een andere vorm. Na bevestigen (script: `Y` typen, pagina: het vinkje) wordt het bestand gemaakt. |
| `3-waarschuwingen-komma.csv` | Bedragen in een bestand met `,` (Engelse Excel). |
| `4-waarschuwingen-tab.csv` | Bedragen in een bestand met tabs. |
| `bestand-*.csv` | Fouten in het bestand zelf. Die stoppen vóór de regels worden gecontroleerd, dus elk heeft een eigen bestand: leeg, alleen een kopregel, alleen lege regels, een ontbrekende kolom, één kolom voor twee velden, en een bestand dat zegt UTF-8 te zijn maar het niet is. |

Fouten bij de vragen na de controle probeer je met de hand, bij `2-waarschuwingen.csv`: procescode `ABCD`,
creatiedatum `31-02-2026`, een lege gebruikersnaam of een van meer dan 16 tekens.

Niet te tonen met deze bestanden:
- Een tijd in een standaard tijdrubriek (type T of U): de huidige lijst met standaardrubrieken heeft er
  geen. De regressietests dekken dit met testrubrieken (`*-standard-time`).
- Een ongeldig Unicode-teken (een losse surrogaat): bij het lezen van een bestand wordt dat al vervangen.

Rubrieken: `E00101` staat niet in de lijst met standaardrubrieken en geldt dus als eigen rubriek; de
`P`-codes zijn standaardrubrieken.

## Bijwerken

Deze bestanden horen niet bij de geautomatiseerde tests, maar worden bij elke nieuwe of gewijzigde
controle bijgewerkt:
1. Pas `make-demo.ps1` aan (een regel met opmerking) en draai `pwsh tests\demo\make-demo.ps1`.
2. Draai `node tests\demo\check-demo.mjs`. Dat haalt alle fout- en waarschuwingscodes uit de kern van
   `dpia100.html`, stuurt de demobestanden erdoorheen en noemt elke code die geen bestand laat zien. Het
   controleert ook dat elke regel van `1-fouten.csv` precies één fout geeft.
