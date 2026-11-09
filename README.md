# Shopfloor Hausen ↔ Berlin

Web-App für das kurze Shopfloor-Meeting (täglich / jeden 2. Tag) zwischen Hausen und Berlin.
Läuft im Browser (Windows-PC, Chrome/Safari auf dem Handy). Anmeldung mit dem Firmen-Microsoft-Konto.
Daten liegen in einer Excel-Datei auf SharePoint (Sprintex); der Lagerbestand Hausen wird **nur gelesen**
aus derselben Datei/Blatt wie die Teileliste am Raspberry Pi (`Lagerübersicht Pi`).

## Was die App abbildet

| Bereich | Wer | Inhalt |
|---|---|---|
| Bestellungen | Hausen | Bestellpositionen aus OrgaMax (BestellNr + Pos, S-Nr aus Lagerliste, Menge, Wunschtermin, LB-Nr). Berlin trägt den bestätigten Termin ein. Farben: gelb = teilgeliefert, hellrot = in `due_soon_days` (3) Tagen fällig, kräftig rot = überfällig. |
| Versand | Berlin (auch Hausen → Berlin) | Jedes Paket inkl. Express wird gemeldet. Die App vergibt eine **LieferID** (`L-JJJJMMTT-NN`), die auf Lieferschein und Paket geschrieben wird. Ohne Bestellung nur mit Grund. |
| Wareneingang | Hausen | Gezählte Menge, QS-Ergebnis (i.O./teilweise/n.i.O.), Menge i.O., **QS geprüft von**. Pakete ohne Versandmeldung können direkt erfasst werden. |
| Probleme | beide | Problemkarten: Problem, Kategorie, Bezug, Verantwortlich, Maßnahme, Termin, Status. |
| Meeting | beide | 6 Schritte, ca. 2 min je Schritt: Kennzahlen → offene Bestellungen (mit Lager Hausen) → unterwegs → Eingänge & QS → Probleme → Protokoll speichern. |

Kennzahlen (Ampel, Ziele in `config.json` → `kpi_targets`): Liefertreue, Mengentreue (erhalten = versendet),
QS-Quote, Anteil Eingänge mit Bestellbezug, Eingänge ohne QS, überfällige Positionen, offene Probleme.
Zeitraum = seit dem letzten gespeicherten Meeting (sonst 7 Tage).

## Dateien

| Datei | Zweck |
|---|---|
| `index.html`, `app.js`, `logic.js` | Die App (reines HTML/JavaScript, kein Build). `logic.js` enthält Status- und Kennzahlregeln. |
| `config.json` | Alle Einstellungen: App-Registrierung, Freigabelinks, Lagerspalten, Ziele. Keine Adressen im Code. |
| `schema.json` | Tabellen und Spalten der Excel-Datei (gemeinsam für App und Generator). |
| `vendor/msal-browser.min.js` | Microsoft-Anmeldebibliothek, fest Version 2.38.3. |
| `staticwebapp.config.json` | Einstellungen für Azure Static Web Apps (sperrt `tests/`, `tools/`, `.xlsx`). |
| `tools/make_workbook.py` | Erzeugt die leere Excel-Datei `Shopfloor_Hausen_Berlin.xlsx` aus `schema.json`. |
| `tests/logic.test.html` | Tests der Regeln. |
| `tests/app.test.html` | Ablauftest der ganzen App gegen eine simulierte Excel-Datei (kein Microsoft-Zugriff). |
| `tests/phone-preview.html` | Ansicht in 375 px Breite (Handy) für Layout-Kontrolle. |

## Einrichtung (einmalig)

### 1. Excel-Datei anlegen
```
python tools/make_workbook.py Shopfloor_Hausen_Berlin.xlsx
```
Datei auf die Sprintex-SharePoint-Seite hochladen (gleiche Seite wie die Lagerliste).
**Teilen → Link kopieren → „Personen in Ihrer Organisation“ mit Bearbeitungsrecht** und den Link in
`config.json` bei `shopfloor_share_link` eintragen. Zum Testen zuerst eine Kopie (`Shopfloor_TEST.xlsx`) verwenden.

Berliner Kollegen brauchen Bearbeitungsrecht auf diese Datei und Leserecht auf die Lagerliste.

### 2. App-Registrierung ergänzen (Entra ID / Azure-Portal)
Es wird die vorhandene Registrierung der Pi-Teileliste genutzt (`client_id`/`tenant_id` in `config.json`).
1. Azure-Portal → Microsoft Entra ID → App-Registrierungen → die Registrierung öffnen.
2. **Authentifizierung → Plattform hinzufügen → Single-Page-Anwendung**.
3. Umleitungs-URIs:
   - `http://localhost:8080/` (lokaler Test)
   - `https://henrikvhebel.github.io/shopfloor-hausen-berlin/` (GitHub Pages, Schritt 3)
4. Speichern. Die bestehende Einstellung „Öffentliche Clientflows zulassen“ (Pi, Gerätecode) bleibt unverändert.
5. API-Berechtigungen: `Files.ReadWrite.All` (delegiert) ist bereits vorhanden. Falls Benutzer beim ersten
   Anmelden nach Zustimmung gefragt werden und das nicht dürfen: „Administratorzustimmung erteilen“.

### 3. Hosting über GitHub Pages
Repository: `https://github.com/HenrikvHebel/shopfloor-hausen-berlin` (öffentlich – GitHub Pages ist im
kostenlosen Konto nur für öffentliche Repositories verfügbar).
App-Adresse: **`https://henrikvhebel.github.io/shopfloor-hausen-berlin/`**

1. Pages ist eingerichtet: Branch `main`, Ordner `/` (Settings → Pages).
2. Änderungen veröffentlichen: in GitHub Desktop committen und „Push origin“. Nach 1–2 Minuten ist die neue
   Version online (der Browser kann bis zu 10 min die alte Version zwischenspeichern; `config.json` wird immer frisch geladen).
3. Link an Hausen und Berlin verteilen. Auf dem Handy: in Safari/Chrome öffnen → „Zum Home-Bildschirm“.

Hinweise:
- Der Code und `config.json` (Client-ID, Tenant-ID, Freigabelinks) sind öffentlich lesbar. Das sind keine
  Geheimnisse; Daten kommen erst nach Microsoft-Anmeldung und nur mit Berechtigung auf die Excel-Dateien.
  Deshalb müssen beide Freigabelinks Links vom Typ **„Personen in Ihrer Organisation“** sein, nie „Jeder“.
- Keine echten Daten, Passwörter oder Patientendaten ins Repository einchecken. Die erzeugte `.xlsx` ist per
  `.gitignore` ausgeschlossen.
- `staticwebapp.config.json` wird nur bei einem späteren Umzug auf Azure Static Web Apps gebraucht.

## Lokal starten
```
python -m http.server 8080 --bind 127.0.0.1
```
Dann `http://localhost:8080/` öffnen.

## Tests
```
python -m http.server 8099 --bind 127.0.0.1
```
- `http://127.0.0.1:8099/tests/logic.test.html` → „RESULT: ALL PASSED“
- `http://127.0.0.1:8099/tests/app.test.html` → „RESULT: ALL PASSED“ (kompletter Ablauf Bestellung → Versand →
  Eingang/QS → Problem → Meeting, gegen simulierte Excel-Tabellen)

## Regeln für die Excel-Datei
- Spalte **LB-Nr** (Laufbandnummern, Format `1234/26`, mehrere mit Komma) ist optional: fehlt sie in einer älteren Datei, läuft die App ohne das Feld. Nachrüsten: in `tblBestellungen` rechts neben „Bemerkung“ eine Spalte mit Überschrift `LB-Nr` anfügen, Zellformat Text.
- Tabellen- und Spaltennamen nicht ändern; zusätzliche Spalten sind erlaubt.
- Die App findet Datensätze über BestellNr+Pos, LieferID bzw. Nr und prüft vor jedem Ändern, dass der
  Datensatz eindeutig ist.
- Keine personenbezogenen Patientendaten eintragen.
