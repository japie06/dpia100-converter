# Privacy en beveiliging: webversie DPIA100 converter

**Voor:** CISO, IT-beheer en functionaris gegevensbescherming. **Versie:** 1.3.0.
**Broncode:** <https://github.com/japie06/dpia100-converter> (MIT).

## Doel en werking
`dpia100.html` zet een CSV-bestand met HR-mutaties om naar een importbestand voor Beaufort (functie
100254). Het is één statisch HTML-bestand dat volledig in de browser van de medewerker draait, lokaal
of via GitHub Pages.

## Gegevensverwerking

| | |
|---|---|
| Gegevens | Mutaties uit de CSV: persoonsnummer, dienstverband, rubriek, waarde (zoals salarisbedragen), datums |
| Verwerking | Alleen in het werkgeheugen van de browser, op de werkplek |
| Opslag | Geen. Alleen het importbestand dat de gebruiker zelf opslaat, en optioneel (opt-in) procescode en gebruikersnaam in de browser (`localStorage`) |
| Doorgifte | Geen. Er zijn geen servers, geen derde partijen en dus geen verwerker |

## Maatregelen
- **Privacy by design en by default (AVG art. 25, art. 5 lid 1c).** Er wordt niets verzameld, gelogd
  of bewaard. Onthouden gebeurt alleen na een expliciet vinkje.
- **Netwerkisolatie via Content Security Policy (W3C CSP Level 3).** `default-src 'none'` en
  `connect-src 'none'`: de pagina kan technisch geen verbinding maken en geen externe bronnen laden.
  Formulierverzending is uitgeschakeld (`form-action 'none'`).
- **Integriteit van de code.** Alleen scripts met de exacte SHA-256-hash uit de CSP worden
  uitgevoerd. Gewijzigde of geïnjecteerde code weigert de browser. Trusted Types en de uitsluitend
  tekstuele DOM-opbouw voorkomen XSS (OWASP).
- **Geen supply-chainrisico.** Geen externe bibliotheken, CDN's, lettertypen of build-stap. Het
  gepubliceerde bestand is de broncode.
- **Geautomatiseerde controle (CI).** Bij elke wijziging wordt gecontroleerd dat de CSP klopt, dat de
  code geen netwerk- of navigatiefuncties bevat, en dat de pagina in Microsoft Edge start zonder
  CSP-overtredingen. Alle testgevallen moeten hetzelfde resultaat geven als de PowerShell-versie.
- **Transparantie.** Open source en voor iedereen controleerbaar. Wijzigingen worden bijgehouden in
  `docs/CHANGELOG.md`.

Deze maatregelen sluiten aan op AVG art. 32 en de BIO (beveiliging van verwerking, veilige
ontwikkeling).

## Zelf verifiëren
1. **Offline:** verbreek de netwerkverbinding; de pagina werkt volledig.
2. **Netwerkverkeer:** ontwikkelhulpmiddelen (F12), tabblad *Netwerk*: tijdens gebruik geen verzoeken.
3. **Integriteit:** `tests/web/update-csp.ps1` herberekent de hashes; de uitkomst moet gelijk zijn aan
   de CSP in het bestand.

## Risico's en aanbevelingen
- **Browserfuncties buiten de pagina**, zoals vertalen en AI-assistenten (Copilot), kunnen pagina-inhoud
  naar externe diensten sturen. De pagina vraagt om niet te vertalen en waarschuwt de gebruiker.
  Aanbevolen: beperk dit via Edge-beleid (onder meer `TranslateEnabled`, Copilot-instellingen).
- **Browserextensies** met toegang tot alle websites kunnen pagina-inhoud lezen. Beheer extensies centraal.
- **Het opgeslagen importbestand** bevat persoonsgegevens en valt onder het gewone beheer van de
  werkplek (opslaglocatie, synchronisatie, bewaartermijn).
- **Lokaal of GitHub Pages.** Een lokaal of op een netwerkschijf geplaatst exemplaar geeft IT volledige
  controle over de versie. Via GitHub Pages ziet GitHub alleen het ophalen van de pagina, niet de
  gegevens.
