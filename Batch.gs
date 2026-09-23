/**
 * Batch.gs
 *
 * Kernlogik: alle Tabellenblaetter eines Spreadsheets verarbeiten, mit EINER
 * gemeinsamen Zielspalte fuer alle Blaetter. Phrases "Identify context note
 * column" ist eine einzige Einstellung pro Multilingual-Excel-Import und
 * gilt damit fuer jedes Blatt gleichermassen - jedes Blatt seine eigene
 * "naechste freie Spalte" waehlen zu lassen wuerde diese Zuordnung kaputt
 * machen, sobald die Blaetter unterschiedlich breit sind.
 */

var IMAGE_HOSTING_FOLDER_PREFIX_ = "context-images";

/**
 * @param {string} spreadsheetId
 * @param {string} siteId Firebase-Hosting-Site-ID
 * @return {Object} {success, count, column, columnLetter, sheetsProcessed, results:[{sheet,row,imageUrl}], deploy, error}
 */
function hostAllSheetsImagesForContextNotes_(spreadsheetId, siteId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheets = ss.getSheets();
  var xlsxBlob = exportSpreadsheetAsXlsx_(spreadsheetId);

  var perSheetImages = {}; // sheetName -> [{row, blob}]
  var maxLastColumn = 0;
  sheets.forEach(function (sh) {
    maxLastColumn = Math.max(maxLastColumn, sh.getLastColumn());
    var imgs = xlsxExtractImagesForSheet_(xlsxBlob, sh.getName());
    if (imgs.length) perSheetImages[sh.getName()] = imgs;
  });

  var sheetNamesWithImages = Object.keys(perSheetImages);
  var totalImages = sheetNamesWithImages.reduce(function (sum, k) { return sum + perSheetImages[k].length; }, 0);
  if (!totalImages) {
    return { success: false, error: "Keine eingebetteten Bilder in diesem Spreadsheet gefunden (geprueft: " + sheets.length + " Blatt/Blätter)." };
  }

  var col = maxLastColumn + 1;
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
    var headerCell = sh.getRange(1, col);
    if (!String(headerCell.getValue() || "").trim()) headerCell.setValue("Screenshot URL");
    perSheetImages[sheetName].forEach(function (img) {
      var url = "https://" + siteId + ".web.app/" + pathPrefix + "/" + img.fileName;
      sh.getRange(img.row, col).setValue(url);
      results.push({ sheet: sheetName, row: img.row, imageUrl: url });
    });
  });

  try {
    var who = Session.getActiveUser().getEmail() || "unknown";
    console.log("CH-BATCH: " + results.length + " image(s) hosted by " + who +
      " for spreadsheet " + spreadsheetId + " -> column " + colLetter);
  } catch (e) {}

  return {
    success:         true,
    count:           results.length,
    column:          col,
    columnLetter:    colLetter,
    sheetsProcessed: sheetNamesWithImages,
    results:         results,
    deploy:          deployResult
  };
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
