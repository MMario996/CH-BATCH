# CH-BATCH

Eigenständiges Apps-Script-Projekt (Web-App) für den Workaround zu Phrase-Support-Ticket **#243154** (Phrase-intern: **DISC-790**):

> Phrase TMS rendert in Excel eingebettete Raster-Bilder (PNG/JPG) nicht im In-Context-Preview des CAT-Editors – unabhängig vom Import-Filter. Laut Phrase-Support eine aktuelle Produkt-Grenze, kein Einstellungsproblem.

**Workaround (Phrase-Support-Option 2):** Screenshots extern hosten, die URL statt des Bildes in eine Spalte schreiben, diese Spalte beim Multilingual-Excel-Import in Phrase als **Context Note**-Spalte zuordnen.

Dieses Tool ist ein manuelles Pre-Batch-Werkzeug: Google-Sheet-Link eintragen, verarbeiten lassen, danach die vorbereitete Datei ganz normal ins Portal/Phrase einreichen. Es greift **nicht** automatisch in irgendeinen Einreichungs-Prozess ein.

## Was es macht

1. Liest alle **eingebetteten (über Zellen liegenden) Bilder** aus einem Google Sheet – über **alle Tabellenblätter** hinweg.
2. Veröffentlicht jedes Bild unter einem eigenen, nicht erratbaren Pfad auf **Firebase Hosting**.
3. Schreibt die öffentliche URL in dieselbe Zeile zurück, in eine **gemeinsame Spalte für alle Tabs** (eine Spalte rechts von der breitesten Tabelle) – wichtig, weil Phrases "Identify context note column" nur eine einzige Einstellung pro Import-Profil ist und für alle Blätter gleichermaßen gilt.

## Einmaliges Setup

1. Dieses Repo in ein Apps-Script-Projekt übernehmen (`clasp push` oder manuell kopieren).
2. **Deploy → New deployment → Web app** erstellen (Execute as: Me, Access: nach Bedarf, z. B. "Anyone within [Domain]").
3. Im Apps-Script-Editor: **Project Settings → Script Properties**, zwei Einträge anlegen:
   - `SERVICE_ACCOUNT_JSON` – kompletter Inhalt eines Firebase-Service-Account-Key-JSON (Rolle **Firebase Hosting Admin** auf dem Firebase-Projekt, das die Ziel-Hosting-Site besitzt).
   - `FIREBASE_SITE_ID` – die Firebase-Hosting-Site-ID, unter der die Bilder veröffentlicht werden (z. B. `kaercher-course-preview`).
4. Web-App-URL öffnen.

## Voraussetzung in Phrase (einmalig, pro Account)

**Settings → Access and Security → "Allow loading of external content in editors"** aktivieren.

## Nutzung

1. Datei muss ein **Google Sheet** sein (keine rohe `.xlsx`) – eine hochgeladene `.xlsx` lässt sich in Google Drive per Rechtsklick → "Öffnen mit" → "Google Tabellen" einmalig konvertieren.
2. Sheet-Link (oder rohe Spreadsheet-ID) in die Web-App eintragen, **Verarbeiten** klicken.
3. Ergebnis prüfen: Spaltenbuchstabe (z. B. `M`) notieren.
4. In Phrase, im jeweiligen Projekt-Template unter **File import settings → MS Excel – Multilingual**, das Feld **"Identify context note column"** auf diesen Buchstaben setzen (einmalig pro Template – gilt danach für alle künftigen Einreichungen mit diesem Template).
5. Das (jetzt um die URL-Spalte ergänzte) Google Sheet ganz normal als Hauptdatei einreichen.

## Dateien

- `Code.gs` – Web-App-Einstiegspunkt (`doGet`), API-Funktionen (`apiProcessSheet`, `apiGetConfigStatus`).
- `Batch.gs` – Kernlogik: alle Tabellenblätter verarbeiten, gemeinsame Zielspalte bestimmen.
- `ImageExtraction.gs` – liest eingebettete Bilder direkt aus der OOXML-Struktur eines (kurz als XLSX exportierten) Google Sheets, da `SpreadsheetApp.OverGridImage` kein `getBlob()` hat.
- `FirebaseHosting.gs` – Service-Account-OAuth (JWT-Bearer-Flow) + schlanker Firebase-Hosting-Deploy-Ablauf.
- `Index.html` – einfaches Formular-UI.
