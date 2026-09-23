/**
 * FirebaseHosting.gs
 *
 * OAuth-Access-Token fuer den Firebase-Hosting-Service-Account (JWT-Bearer-
 * Flow, gecached fuer ~50 Minuten) + ein schlanker Deploy-Ablauf ueber die
 * offizielle Firebase Hosting REST API:
 *   1. Neue Version anlegen
 *   2. Fuer jede Datei SHA256 des GZIP-komprimierten Inhalts berechnen und
 *      per populateFiles melden
 *   3. Nur die Dateien hochladen, die Firebase als "noch nicht bekannt"
 *      zurueckmeldet (Content-Hash-Dedupe)
 *   4. Version finalisieren
 *   5. Version als Release veroeffentlichen -> live unter der Hosting-URL
 *
 * Script Property "SERVICE_ACCOUNT_JSON" muss den kompletten Inhalt des
 * Service-Account-Key-JSON enthalten (Rolle "Firebase Hosting Admin").
 */

var FIREBASE_SCOPE_ =
  "https://www.googleapis.com/auth/firebase.hosting " +
  "https://www.googleapis.com/auth/cloud-platform";
var FIREBASE_HOSTING_API_ = "https://firebasehosting.googleapis.com/v1beta1";

function getFirebaseAccessToken_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = "firebase_access_token";
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var serviceAccountJsonRaw = PropertiesService.getScriptProperties().getProperty("SERVICE_ACCOUNT_JSON");
  if (!serviceAccountJsonRaw) throw new Error("Script Property 'SERVICE_ACCOUNT_JSON' ist nicht gesetzt.");
  var sa = JSON.parse(serviceAccountJsonRaw);

  var nowSeconds = Math.floor(Date.now() / 1000);
  var header = { alg: "RS256", typ: "JWT" };
  var claimSet = {
    iss:   sa.client_email,
    scope: FIREBASE_SCOPE_,
    aud:   "https://oauth2.googleapis.com/token",
    iat:   nowSeconds,
    exp:   nowSeconds + 3600
  };

  var signedJwt = signJwtRs256_(header, claimSet, sa.private_key);

  var response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method:      "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion:  signedJwt
    },
    muteHttpExceptions: true
  });

  var body = JSON.parse(response.getContentText());
  if (!body.access_token) {
    throw new Error("Konnte kein Access-Token holen. Antwort von Google: " + response.getContentText());
  }

  cache.put(cacheKey, body.access_token, 50 * 60);
  return body.access_token;
}

function signJwtRs256_(header, claimSet, privateKeyPem) {
  var encodedHeader = base64UrlEncodeString_(JSON.stringify(header));
  var encodedClaimSet = base64UrlEncodeString_(JSON.stringify(claimSet));
  var signingInput = encodedHeader + "." + encodedClaimSet;
  var signatureBytes = Utilities.computeRsaSha256Signature(signingInput, privateKeyPem);
  var encodedSignature = base64UrlEncodeBytes_(signatureBytes);
  return signingInput + "." + encodedSignature;
}

function base64UrlEncodeString_(str) {
  return base64UrlEncodeBytes_(Utilities.newBlob(str).getBytes());
}

function base64UrlEncodeBytes_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, "");
}

function fbFetch_(accessToken, method, path, payload) {
  var options = {
    method:             method,
    headers:            { Authorization: "Bearer " + accessToken },
    muteHttpExceptions: true
  };
  if (payload !== null) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }
  var response = UrlFetchApp.fetch(FIREBASE_HOSTING_API_ + path, options);
  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 300) {
    throw new Error("Firebase Hosting API Fehler (" + code + ") bei " + path + ": " + text);
  }
  return text ? JSON.parse(text) : {};
}

function bytesToHex_(bytes) {
  return bytes.map(function (b) {
    var unsigned = b < 0 ? b + 256 : b;
    var hex = unsigned.toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  }).join("");
}

/**
 * Deployt beliebige Blobs unter einem gemeinsamen Pfad-Prefix auf Firebase
 * Hosting und veroeffentlicht sie sofort als neues Release.
 * @param {string} accessToken
 * @param {string} siteId      Firebase-Hosting-Site-ID
 * @param {Array}  fileEntries [{path, blob}, ...] - "path" relativ zu pathPrefix
 * @param {string} pathPrefix
 * @return {Object} {versionName, uploadedFileCount, totalFileCount}
 */
function deployBlobsToFirebaseHosting_(accessToken, siteId, fileEntries, pathPrefix) {
  var hashToBlob = {};
  var pathToHash = {};
  var prefix = String(pathPrefix || "").replace(/^\/|\/$/g, "");

  fileEntries.forEach(function (entry) {
    var relativePath = "/" + prefix + "/" + String(entry.path || "").replace(/^\/+/, "");
    var gzipped = Utilities.gzip(entry.blob);
    var digestBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, gzipped.getBytes());
    var hash = bytesToHex_(digestBytes);
    pathToHash[relativePath] = hash;
    hashToBlob[hash] = gzipped;
  });

  var versionResp = fbFetch_(accessToken, "post", "/sites/" + siteId + "/versions", {});
  var versionName = versionResp.name;

  var populateResp = fbFetch_(accessToken, "post", "/" + versionName + ":populateFiles", { files: pathToHash });
  var requiredHashes = populateResp.uploadRequiredHashes || [];
  var uploadUrl = populateResp.uploadUrl;

  requiredHashes.forEach(function (hash) {
    var uploadResp = UrlFetchApp.fetch(uploadUrl + "/" + hash, {
      method:             "post",
      headers:            { Authorization: "Bearer " + accessToken, "Content-Type": "application/octet-stream" },
      payload:            hashToBlob[hash].getBytes(),
      muteHttpExceptions: true
    });
    if (uploadResp.getResponseCode() >= 300) {
      throw new Error("Upload fehlgeschlagen fuer Hash " + hash + ": " + uploadResp.getContentText());
    }
  });

  fbFetch_(accessToken, "patch", "/" + versionName + "?updateMask=status", { status: "FINALIZED" });
  fbFetch_(accessToken, "post", "/sites/" + siteId + "/releases?versionName=" + encodeURIComponent(versionName), null);

  return { versionName: versionName, uploadedFileCount: requiredHashes.length, totalFileCount: fileEntries.length };
}
