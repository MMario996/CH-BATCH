/**
 * Batch.gs
 *
 * Kernlogik: alle Tabellenblaetter eines Spreadsheets verarbeiten, mit EINER
 * gemeinsamen Zielspalte fuer alle Blaetter. Phrases "Identify context note
 * column" ist eine einzige Einstellung pro Multilingual-Excel-Import und
 * gilt damit fuer jedes Blatt gleichermassen - jedes Blatt seine eigene
 * "naechste freie Spalte" waehlen zu lassen wuerde diese Zuordnung kaputt
 * machen, sobald die Blaetter unterschiedlich breit sind.
 *
 * ZEILEN-ZUORDNUNG: Der XML-Anker eines Bildes (<xdr:from><xdr:row>) ist NUR
 * die Start-Zeile, nicht zwingend die Zeile, in der das Bild optisch sitzt.
 * Live getestet an einer echten Datei: Screenshots waren an einer kurzen
 * "Section Name"-Kopfzeile verankert, ragten aber fast komplett in die viel
 * hoehere "Intro Text"-Zeile darunter (Zeilenumbruch, Auto-Hoehe) hinein -
 * fuer Menschen "gehoert" das Bild sichtbar zur unteren Zeile, nicht zur
 * Ankerzeile. Da Auto-Hoehe-Zeilen ihre tatsaechliche Renderhoehe NICHT in
 * der xlsx-XML speichern, laesst sich das nicht rein aus der Datei berechnen
 * - resolveVisualRow_() fragt stattdessen die live gerenderte Zeilenhoehe
 * (SpreadsheetApp) ab und bestimmt darueber, in welcher Zeile die Bild-
 * UNTERKANTE tatsaechlich liegt - die URL landet also in der Zeile, in der
 * das Bild sichtbar endet (direkt darunter), nie mittendrin.
 *
 * SPALTEN-WAHL: Viele Formular-Vorlagen (z.B. Kärcher-Trainingsformulare)
 * haben bereits eine eigene, vorgefertigte "Screenshot"-Spalte, in der die
 * Bilder liegen - deren Kopfzeile kann irgendwo sitzen (nicht zwingend
 * Zeile 1). Wird eine Spalte gefunden, deren Zelle (in den ersten
 * SCREENSHOT_HEADER_SEARCH_MAX_ROWS_ Zeilen) den Text "Screenshot" traegt -
 * und ist diese Spalte ueber alle Tabellenblaetter hinweg eindeutig (bzw.
 * kommt nur einmal vor) -, wird DIESE Spalte wiederverwendet statt eine neue
 * anzulegen. Sonst faellt das Skript auf die alte Logik zurueck (neue Spalte
 * rechts von der breitesten Tabelle, Header in Zeile 1). Grund: liegt die
 * URL in genau der Spalte, die auch die (von Phrase nicht gerenderten)
 * Bilder enthaelt, bleibt in Phrases CAT-Editor keine leere Spalte mehr an
 * der Stelle zurueck, wo vorher nur das (unsichtbare) Bild war.
 *
 * KOLLISIONEN: Landen zwei Bilder auf derselben Zielzeile, werden ihre URLs
 * in dieselbe Zelle zusammengefuehrt (zeilenumbruch-getrennt) statt dass die
 * zweite die erste stillschweigend ueberschreibt.
 */

var IMAGE_HOSTING_FOLDER_PREFIX_ = "context-images";
var EMU_PER_PIXEL_ = 9525; // Standard-OOXML-Konstante: 914400 EMU/Zoll / 96px/Zoll
var SCREENSHOT_COLUMN_HEADER_TEXT_ = "screenshot";
var SCREENSHOT_HEADER_SEARCH_MAX_ROWS_ = 30;

/**
 * @param {string} spreadsheetId
 * @param {string} siteId Firebase-Hosting-Site-ID
 * @return {Object} {success, count, column, columnLetter, sheetsProcessed, results:[{sheet,row,imageUrl}], deploy, error}
 */
function hostAllSheetsImagesForContextNotes_(spreadsheetId, siteId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheets = ss.getSheets();
  var xlsxBlob = exportSpreadsheetAsXlsx_(spreadsheetId);

  var perSheetImages = {}; // sheetName -> [{row (visuell, 1-based), blob}]
  var maxLastColumn = 0;
  sheets.forEach(function (sh) {
    maxLastColumn = Math.max(maxLastColumn, sh.getLastColumn());
    var anchors = xlsxExtractImagesForSheet_(xlsxBlob, sh.getName());
    if (!anchors.length) return;

    var resolved = anchors.map(function (a) {
      var row = resolveVisualRow_(sh, a.row0, a.rowOffEmu, a.extCyEmu, a.toRow0);
      return { row: row, blob: a.blob };
    });
    perSheetImages[sh.getName()] = resolved;
  });

  var sheetNamesWithImages = Object.keys(perSheetImages);
  var totalImages = sheetNamesWithImages.reduce(function (sum, k) { return sum + perSheetImages[k].length; }, 0);
  if (!totalImages) {
    return { success: false, error: "Keine eingebetteten Bilder in diesem Spreadsheet gefunden (geprueft: " + sheets.length + " Blatt/Blätter)." };
  }

  var existingScreenshotCol = findScreenshotColumnAcrossSheets_(sheets);
  var reusedExistingColumn = !!existingScreenshotCol;
  var col = existingScreenshotCol || (maxLastColumn + 1);
  var colLetter = columnToLetter_(col);

  var folderToken = Utilities.getUuid().replace(/-/g, "").substring(0, 12);
  var pathPrefix = IMAGE_HOSTING_FOLDER_PREFIX_ + "/" + folderToken;

  var fileEntries = [];
  var idx = 0;
  sheetNamesWithImages.forEach(function (sheetName) {
    perSheetImages[sheetName].forEach(function (img) {
      idx++;
      var ext = imgExtFromContentType_(img.blob.getContentType());
      img.fileName = "img" + idx + "_row" + img.row + ext; // Sheet-Name kann Sonderzeichen enthalten -> Index statt Name im Pfad
      fileEntries.push({ path: img.fileName, blob: img.blob });
    });
  });

  var accessToken = getFirebaseAccessToken_();
  var deployResult = deployBlobsToFirebaseHosting_(accessToken, siteId, fileEntries, pathPrefix);

  var results = [];
  sheetNamesWithImages.forEach(function (sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!reusedExistingColumn) {
      var headerCell = sh.getRange(1, col);
      if (!String(headerCell.getValue() || "").trim()) headerCell.setValue("Screenshot URL");
    }

    var urlsByRow = {}; // row -> [urls] - sammelt Kollisionen statt sie zu ueberschreiben
    perSheetImages[sheetName].forEach(function (img) {
      var url = "https://" + siteId + ".web.app/" + pathPrefix + "/" + img.fileName;
      if (!urlsByRow[img.row]) urlsByRow[img.row] = [];
      urlsByRow[img.row].push(url);
      results.push({ sheet: sheetName, row: img.row, imageUrl: url });
    });

    Object.keys(urlsByRow).forEach(function (rowStr) {
      sh.getRange(Number(rowStr), col).setValue(urlsByRow[rowStr].join("\n"));
    });
  });

  try {
    var who = Session.getActiveUser().getEmail() || "unknown";
    console.log("CH-BATCH: " + results.length + " image(s) hosted by " + who +
      " for spreadsheet " + spreadsheetId + " -> column " + colLetter +
      (reusedExistingColumn ? " (vorhandene 'Screenshot'-Spalte wiederverwendet)" : " (neue Spalte angelegt)"));
  } catch (e) {}

  return {
    success:              true,
    count:                results.length,
    column:               col,
    columnLetter:         colLetter,
    reusedExistingColumn: reusedExistingColumn,
    sheetsProcessed:      sheetNamesWithImages,
    results:              results,
    deploy:               deployResult
  };
}

/**
 * Sucht ueber alle Tabellenblaetter hinweg eine gemeinsame Spalte, deren
 * Header-Zelle "Screenshot" heisst (z.B. eine in der Vorlage bereits
 * vorhandene Bild-Spalte). Liefert die Spalte nur zurueck, wenn sie eindeutig
 * ist (in jedem Blatt, das ueberhaupt eine "Screenshot"-Spalte hat, dieselbe
 * Spaltennummer) - bei Widerspruch oder wenn keine gefunden wird, gibt es
 * null zurueck und der Aufrufer faellt auf die alte "neue Spalte"-Logik
 * zurueck, statt zu raten.
 * @param {Array<Sheet>} sheets
 * @return {?number} 1-basierte Spaltennummer, oder null
 */
function findScreenshotColumnAcrossSheets_(sheets) {
  var foundCols = {};
  sheets.forEach(function (sh) {
    var col = findScreenshotColumnInSheet_(sh);
    if (col) foundCols[col] = true;
  });

  var distinctCols = Object.keys(foundCols).map(Number);
  if (distinctCols.length === 1) return distinctCols[0];
  if (distinctCols.length > 1) {
    console.warn("CH-BATCH: Mehrere unterschiedliche 'Screenshot'-Spalten ueber die Tabellenblaetter " +
      "gefunden (Spalten " + distinctCols.join(", ") + ") - lege stattdessen sicherheitshalber eine neue Spalte an.");
  }
  return null;
}

/** Sucht in den ersten SCREENSHOT_HEADER_SEARCH_MAX_ROWS_ Zeilen eines Blatts eine Zelle mit dem Text "Screenshot". */
function findScreenshotColumnInSheet_(sheet) {
  var lastRow = Math.min(sheet.getLastRow(), SCREENSHOT_HEADER_SEARCH_MAX_ROWS_);
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return null;

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c] || "").trim().toLowerCase() === SCREENSHOT_COLUMN_HEADER_TEXT_) return c + 1;
    }
  }
  return null;
}

/**
 * Bestimmt die 1-basierte Zeile, in der ein Bild optisch sichtbar ENDET
 * (seine Unterkante), statt nur seiner XML-Ankerzeile - siehe Erklaerung
 * oben im Datei-Header. Die URL soll direkt darunter stehen, nie mittendrin
 * im Bild. Summiert dafuer die LIVE gerenderten Zeilenhoehen (in Pixel,
 * ueber SpreadsheetApp.getRowHeight - funktioniert auch fuer Auto-Hoehe-
 * Zeilen, im Gegensatz zur statischen xlsx-XML) ab row0 auf, bis die
 * Unterkante erreicht ist.
 *
 * @param {Sheet} sheet
 * @param {number} row0        0-basierte XML-Ankerzeile
 * @param {number} rowOffEmu   Offset in EMU ab Zeilenanfang
 * @param {number} [extCyEmu]  Bildhoehe in EMU (oneCellAnchor) - falls gesetzt, wird ueber Zeilenhoehen die Zeile der Unterkante gesucht
 * @param {number} [toRow0]    0-basierte End-Zeile (twoCellAnchor, falls extCyEmu fehlt) - im Anker bereits explizit angegeben, keine Naeherung noetig
 * @return {number} 1-basierte sichtbare Zeile
 */
function resolveVisualRow_(sheet, row0, rowOffEmu, extCyEmu, toRow0) {
  var maxRow = sheet.getMaxRows();

  if (extCyEmu == null) {
    // twoCellAnchor: Endzeile steht explizit im Anker (<xdr:to><xdr:row>).
    var bottomRow0 = toRow0 != null ? toRow0 : row0;
    return Math.min(bottomRow0 + 1, maxRow);
  }

  var remainingEmu = (rowOffEmu || 0) + extCyEmu; // volle Bildhoehe -> Ziel ist die Unterkante, nicht die Mitte
  var currentRow1Based = row0 + 1;

  while (currentRow1Based <= maxRow) {
    var rowHeightEmu = sheet.getRowHeight(currentRow1Based) * EMU_PER_PIXEL_;
    if (remainingEmu < rowHeightEmu) return currentRow1Based;
    remainingEmu -= rowHeightEmu;
    currentRow1Based++;
  }
  return maxRow;
}

/** 1-basierte Spaltennummer -> Buchstabe(n), z.B. 9 -> "I", 27 -> "AA" (wie Phrase's "Identify ... column"-Felder es erwarten). */
function columnToLetter_(col) {
  var letter = "";
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
