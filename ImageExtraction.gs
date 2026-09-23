/**
 * ImageExtraction.gs
 *
 * Liest eingebettete (ueber Zellen liegende) Bilder aus einem Google Sheet.
 *
 * SpreadsheetApp.OverGridImage hat KEIN getBlob() - live getestet, wirft
 * "img.getBlob is not a function". Apps Script bietet darueber also keinen
 * Weg an die Roh-Bytes eines ueber Zellen liegenden Bildes. Deshalb wird
 * das Sheet stattdessen kurz als XLSX exportiert (Drive-Export-Endpoint,
 * mit dem "spreadsheets"/"drive.readonly"-OAuth-Scope), das XLSX als ZIP
 * entpackt und Bilder + Zeilen-Anker direkt aus den OOXML-Drawing-XMLs
 * gelesen (xl/drawings/drawing*.xml + die zugehoerigen .rels-Dateien) -
 * das ist dieselbe Struktur, die auch eine direkt hochgeladene .xlsx-Datei
 * hat, das Ergebnis ist also identisch zu einer direkten .xlsx-Verarbeitung.
 */

var XDR_NS_  = XmlService.getNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
var XDRA_NS_ = XmlService.getNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
var XDRR_NS_ = XmlService.getNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

function exportSpreadsheetAsXlsx_(spreadsheetId) {
  var url = "https://docs.google.com/spreadsheets/d/" + encodeURIComponent(spreadsheetId) + "/export?format=xlsx";
  var res = UrlFetchApp.fetch(url, {
    headers:            { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error("Konnte Spreadsheet nicht als XLSX exportieren (HTTP " + res.getResponseCode() + ").");
  }
  return res.getBlob().setContentType("application/zip");
}

/**
 * @param {Blob} xlsxBlob
 * @param {string} sheetName Name des Tabs, wie er im Google Sheet steht (leer -> erstes Blatt).
 * @return {Array<{row0:number, rowOffEmu:number, extCyEmu:number, blob:Blob}|{row0:number, rowOffEmu:number, toRow0:number, blob:Blob}>}
 *   Rohe Anker-Daten (0-basierte Start-Zeile + EMU-Offsets) - noch NICHT die
 *   sichtbare Zeile. Die wird erst ueber resolveVisualRow_() (Batch.gs) mit
 *   den tatsaechlichen, live gerenderten Zeilenhoehen aufgeloest, weil
 *   Auto-Hoehe-Zeilen (Zeilenumbruch) ihre echte Hoehe nicht in der XML
 *   speichern.
 */
function xlsxExtractImagesForSheet_(xlsxBlob, sheetName) {
  var files = Utilities.unzip(xlsxBlob);
  var byName = {};
  files.forEach(function (f) { byName[f.getName()] = f; });

  var worksheetFile = xlsxFindWorksheetFileForName_(byName, sheetName);
  if (!worksheetFile) return [];

  var sheetFileName = worksheetFile.getName(); // "xl/worksheets/sheetN.xml"
  var sheetIndexMatch = sheetFileName.match(/sheet(\d+)\.xml$/);
  if (!sheetIndexMatch) return [];
  var sheetIndex = sheetIndexMatch[1];

  var sheetRelsFile = byName["xl/worksheets/_rels/sheet" + sheetIndex + ".xml.rels"];
  if (!sheetRelsFile) return []; // kein Drawing fuer dieses Blatt verknuepft

  var sheetRelsRoot = XmlService.parse(sheetRelsFile.getDataAsString()).getRootElement();
  var relNs = sheetRelsRoot.getNamespace();
  var drawingRel = sheetRelsRoot.getChildren("Relationship", relNs).filter(function (r) {
    return /\/drawing$/.test(r.getAttribute("Type").getValue());
  })[0];
  if (!drawingRel) return [];

  var drawingTarget = drawingRel.getAttribute("Target").getValue(); // z.B. "../drawings/drawing1.xml"
  var drawingPath = xlsxResolveRelativePath_("xl/worksheets/", drawingTarget);
  var drawingFile = byName[drawingPath];
  if (!drawingFile) return [];

  var drawingIndexMatch = drawingPath.match(/drawing(\d+)\.xml$/);
  var drawingIndex = drawingIndexMatch ? drawingIndexMatch[1] : "1";
  var drawingRelsFile = byName["xl/drawings/_rels/drawing" + drawingIndex + ".xml.rels"];

  var ridToMedia = {};
  if (drawingRelsFile) {
    var drRelsRoot = XmlService.parse(drawingRelsFile.getDataAsString()).getRootElement();
    var drRelNs = drRelsRoot.getNamespace();
    drRelsRoot.getChildren("Relationship", drRelNs).forEach(function (r) {
      var target = r.getAttribute("Target").getValue();
      ridToMedia[r.getAttribute("Id").getValue()] = xlsxResolveRelativePath_("xl/drawings/", target);
    });
  }

  var drawRoot = XmlService.parse(drawingFile.getDataAsString()).getRootElement();
  var twoCellAnchors = drawRoot.getChildren("twoCellAnchor", XDR_NS_);
  var oneCellAnchors = drawRoot.getChildren("oneCellAnchor", XDR_NS_);

  var results = [];

  // oneCellAnchor: fixe Groesse (<xdr:ext cx cy>) ab einem Anker-Punkt - der
  // haeufigere Fall bei aus Google Sheets exportierten Screenshots.
  oneCellAnchors.forEach(function (anchor) {
    var from = anchor.getChild("from", XDR_NS_);
    if (!from) return;
    var rowEl = from.getChild("row", XDR_NS_);
    var rowOffEl = from.getChild("rowOff", XDR_NS_);
    var row0 = rowEl ? parseInt(rowEl.getText(), 10) : 0;
    var rowOffEmu = rowOffEl ? parseInt(rowOffEl.getText(), 10) : 0;

    var extEl = anchor.getChild("ext", XDR_NS_);
    var extCyEmu = extEl && extEl.getAttribute("cy") ? parseInt(extEl.getAttribute("cy").getValue(), 10) : 0;

    var mediaBlob = extractMediaBlobFromAnchor_(anchor, ridToMedia, byName);
    if (!mediaBlob) return;

    results.push({ row0: row0, rowOffEmu: rowOffEmu, extCyEmu: extCyEmu, blob: mediaBlob });
  });

  // twoCellAnchor: Start- UND Endzelle explizit angegeben - Hoehe ergibt sich
  // aus der Differenz, dafuer reicht als Naeherung die Mitte zwischen beiden
  // Zeilen (die tatsaechliche Pixelhoehe pro Zwischenzeile wird ohnehin erst
  // beim Aufloesen der sichtbaren Zeile ueber die Live-Zeilenhoehen ermittelt).
  twoCellAnchors.forEach(function (anchor) {
    var from = anchor.getChild("from", XDR_NS_);
    var to = anchor.getChild("to", XDR_NS_);
    if (!from || !to) return;
    var rowEl = from.getChild("row", XDR_NS_);
    var rowOffEl = from.getChild("rowOff", XDR_NS_);
    var toRowEl = to.getChild("row", XDR_NS_);
    var row0 = rowEl ? parseInt(rowEl.getText(), 10) : 0;
    var rowOffEmu = rowOffEl ? parseInt(rowOffEl.getText(), 10) : 0;
    var toRow0 = toRowEl ? parseInt(toRowEl.getText(), 10) : row0;

    var mediaBlob = extractMediaBlobFromAnchor_(anchor, ridToMedia, byName);
    if (!mediaBlob) return;

    results.push({ row0: row0, rowOffEmu: rowOffEmu, toRow0: toRow0, blob: mediaBlob });
  });

  return results;
}

/** Extrahiert aus einem Anchor-Element (<xdr:pic><xdr:blipFill><a:blip r:embed="...">) das zugehoerige Bild-Blob, oder null. */
function extractMediaBlobFromAnchor_(anchor, ridToMedia, byName) {
  var pic = anchor.getChild("pic", XDR_NS_);
  if (!pic) return null;
  var blipFill = pic.getChild("blipFill", XDR_NS_);
  if (!blipFill) return null;
  var blip = blipFill.getChild("blip", XDRA_NS_);
  if (!blip) return null;
  var embedAttr = blip.getAttribute("embed", XDRR_NS_);
  if (!embedAttr) return null;

  var mediaPath = ridToMedia[embedAttr.getValue()];
  if (!mediaPath || !byName[mediaPath]) return null;
  return byName[mediaPath].copyBlob();
}

/** Findet die xl/worksheets/sheetN.xml-Datei, die zum gegebenen Tab-Namen gehoert (leer -> erstes Blatt in Tab-Reihenfolge). */
function xlsxFindWorksheetFileForName_(byName, sheetName) {
  var workbookFile = byName["xl/workbook.xml"];
  var workbookRelsFile = byName["xl/_rels/workbook.xml.rels"];
  if (!workbookFile || !workbookRelsFile) return null;

  var wbRoot = XmlService.parse(workbookFile.getDataAsString()).getRootElement();
  var wbNs = wbRoot.getNamespace();
  var rNs = XmlService.getNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
  var sheetsEl = wbRoot.getChild("sheets", wbNs);
  if (!sheetsEl) return null;

  var sheetEl = sheetsEl.getChildren("sheet", wbNs).filter(function (s) {
    return s.getAttribute("name").getValue() === sheetName;
  })[0];
  if (!sheetEl) sheetEl = sheetsEl.getChildren("sheet", wbNs)[0]; // Fallback: erstes Blatt
  if (!sheetEl) return null;

  var rIdAttr = sheetEl.getAttribute("id", rNs);
  if (!rIdAttr) return null;
  var rId = rIdAttr.getValue();

  var relsRoot = XmlService.parse(workbookRelsFile.getDataAsString()).getRootElement();
  var relNs = relsRoot.getNamespace();
  var rel = relsRoot.getChildren("Relationship", relNs).filter(function (r) {
    return r.getAttribute("Id").getValue() === rId;
  })[0];
  if (!rel) return null;

  var target = rel.getAttribute("Target").getValue(); // z.B. "worksheets/sheet1.xml"
  var path = xlsxResolveRelativePath_("xl/", target);
  return byName[path] || null;
}

/** Loest ein relatives OOXML-Target (z.B. "../media/image1.png") gegen ein Basisverzeichnis auf. */
function xlsxResolveRelativePath_(baseDir, target) {
  var parts = baseDir.replace(/\/$/, "").split("/").concat(target.split("/"));
  var resolved = [];
  parts.forEach(function (p) {
    if (p === "" || p === ".") return;
    if (p === "..") resolved.pop();
    else resolved.push(p);
  });
  return resolved.join("/");
}

function imgExtFromContentType_(ct) {
  ct = String(ct || "").toLowerCase();
  if (ct.indexOf("png")  !== -1) return ".png";
  if (ct.indexOf("jpeg") !== -1 || ct.indexOf("jpg") !== -1) return ".jpg";
  if (ct.indexOf("gif")  !== -1) return ".gif";
  if (ct.indexOf("webp") !== -1) return ".webp";
  return ".png";
}
