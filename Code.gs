/**
 * Code.gs — CH-BATCH
 *
 * Standalone-Web-App fuer den Workaround zu Phrase-Support-Ticket #243154
 * (Phrase-intern: DISC-790): Phrase TMS rendert in Excel eingebettete
 * Raster-Bilder (PNG/JPG) nicht im In-Context-Preview des CAT-Editors,
 * unabhaengig vom Import-Filter - laut Phrase-Support eine aktuelle
 * Produkt-Grenze, kein Einstellungsproblem.
 *
 * Phrase-Support-Option 2 (die dieses Tool umsetzt): Screenshots extern
 * hosten, die URL statt des Bildes in eine Spalte schreiben, diese Spalte
 * beim Multilingual-Excel-Import in Phrase als Context-Note-Spalte
 * zuordnen. Voraussetzung in Phrase (einmalig, pro Account): "Allow
 * loading of external content in editors" unter Access and Security.
 *
 * Nutzung: Web-App-URL oeffnen, Link (oder rohe ID) eines Google Sheets
 * eintragen, "Verarbeiten" klicken. Das Sheet MUSS ein Google Sheet sein
 * (keine rohe .xlsx) - eine hochgeladene .xlsx laesst sich in Drive per
 * Rechtsklick -> "Oeffnen mit" -> "Google Tabellen" einmalig konvertieren.
 *
 * Verarbeitet ALLE Tabellenblaetter des Sheets und schreibt die Bild-URLs
 * blattuebergreifend in DIESELBE Spalte (eine Spalte rechts von der
 * breitesten Tabelle) - Phrase hat nur EINE globale "Identify context note
 * column"-Einstellung pro Multilingual-Excel-Import-Profil, die fuer alle
 * Blaetter gleichermassen gilt.
 *
 * Einmaliges Setup (Apps Script Editor -> Project Settings -> Script
 * Properties):
 *   SERVICE_ACCOUNT_JSON  Kompletter Inhalt eines Firebase-Service-Account-
 *                         Key-JSON (Rolle "Firebase Hosting Admin" auf dem
 *                         Firebase-Projekt, das die Ziel-Hosting-Site
 *                         besitzt).
 *   FIREBASE_SITE_ID      Firebase-Hosting-Site-ID, unter der die Bilder
 *                         veroeffentlicht werden (z.B. "kaercher-course-preview").
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("CH-BATCH: Screenshot Context Notes")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function apiGetConfigStatus() {
  var props = PropertiesService.getScriptProperties();
  return {
    hasServiceAccount: !!props.getProperty("SERVICE_ACCOUNT_JSON"),
    hasSiteId:          !!props.getProperty("FIREBASE_SITE_ID"),
    siteId:              props.getProperty("FIREBASE_SITE_ID") || ""
  };
}

/**
 * @param {string} sheetUrlOrId Google-Sheet-Link oder rohe Spreadsheet-ID
 * @return {Object} {success, count, column, columnLetter, sheetsProcessed, results:[{sheet,row,imageUrl}], error}
 */
function apiProcessSheet(sheetUrlOrId) {
  var props = PropertiesService.getScriptProperties();
  var siteId = props.getProperty("FIREBASE_SITE_ID");
  if (!siteId) {
    return { success: false, error: "Script Property FIREBASE_SITE_ID ist nicht gesetzt (Project Settings → Script Properties)." };
  }
  if (!props.getProperty("SERVICE_ACCOUNT_JSON")) {
    return { success: false, error: "Script Property SERVICE_ACCOUNT_JSON ist nicht gesetzt (Project Settings → Script Properties)." };
  }

  var spreadsheetId = extractSpreadsheetId_(sheetUrlOrId);
  if (!spreadsheetId) {
    return { success: false, error: "Konnte keine Spreadsheet-ID aus der Eingabe lesen. Bitte den vollen Sheet-Link oder die reine ID einfuegen." };
  }

  try {
    return hostAllSheetsImagesForContextNotes_(spreadsheetId, siteId);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function extractSpreadsheetId_(input) {
  input = String(input || "").trim();
  if (!input) return "";
  var m = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(input)) return input; // sieht schon nach roher ID aus
  return "";
}
