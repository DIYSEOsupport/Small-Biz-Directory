/**
 * Directory form receiver
 *
 * Receives what people type into form.html and adds it as a new row in
 * the same Google Sheet the directory app (index.html) reads.
 *
 *  - People sign in with Google on form.html. This script checks the
 *    sign-in with Google, so nobody can pretend to be someone else.
 *  - The first Google account to add a business name "owns" it. Only that
 *    account can replace the listing later. We store a scrambled key, never
 *    the email, in a column called "Owner Key (private)". The directory app
 *    never shows that column.
 *  - Logos are saved in your Google Drive and made viewable by link.
 *  - The Google Sheet stays PRIVATE. The directory reads a public-only copy
 *    from this script's web address (see doGet). Names, titles, personal
 *    contact info and owner keys are never sent to the public.
 *
 * HOW TO USE
 *   1. Go to script.google.com and click "New project".
 *   2. Delete the code that is there and paste in this whole file.
 *   3. Paste your Google Client ID into CLIENT_ID below.
 *   4. In the dropdown at the top (it may say "myFunction"), choose
 *      authorizeOnce, click Run, and click Allow.
 *   5. Click Deploy > New deployment > the gear > Web app.
 *        Execute as: Me
 *        Who has access: Anyone
 *   6. Click Deploy and copy the Web app URL.
 *   7. In form.html, replace PASTE_WEB_APP_URL_HERE with that URL.
 *   If you change this code later: Deploy > Manage deployments > pencil >
 *   Version: New version > Deploy.
 */

// ---------------------------------------------------------------
// Settings
// ---------------------------------------------------------------
var CLIENT_ID = '429873674496-c3ld7ru88i8b0acssaeh26adtvfok12e.apps.googleusercontent.com';
var SHEET_ID = '17vDXZS_tXVWAerdp0azT896Sg8cQYxCBN8zNl2d6Rec';
var SHEET_GID = 1342044865;
var LOGO_FOLDER_NAME = 'Directory Logos';
var OWNER_HEADER = 'Owner Key (private)';

// Google accounts that may change any listing (put your own email here)
var ADMIN_EMAILS = ['diyseo.support@gmail.com'];

// Rows added before sign-in existed have no owner.
// true  = the first person to submit that business name claims it.
// false = those listings can only be changed by you in the sheet.
var ALLOW_CLAIM_OLD_ROWS = true;

var EMAIL_RE = /^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}$/;
var URL_RE = /^(https?:\/\/)?([A-Za-z0-9\-]+\.)+[A-Za-z]{2,}(:\d+)?([\/?#]\S*)?$/;

// ---------------------------------------------------------------
// Run this once so Google asks for permission
// ---------------------------------------------------------------
function authorizeOnce() {
  SpreadsheetApp.openById(SHEET_ID).getName();
  getLogoFolder_();
  PropertiesService.getScriptProperties().getKeys();
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test',
      { muteHttpExceptions: true });
  Logger.log('All set. Now click Deploy > New deployment.');
}

// ---------------------------------------------------------------
// Web app entry points
// ---------------------------------------------------------------
function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService.createTextOutput('Directory form receiver is running.');
  }
  return ContentService
      .createTextOutput(publicCsv_())
      .setMimeType(ContentService.MimeType.CSV);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);

    var email = verifyToken_(data.idToken);
    if (!email) {
      return json_({ ok: false, error: 'LOGIN_EXPIRED' });
    }

    var problem = validate_(data);
    if (problem) {
      return json_({ ok: false, error: problem });
    }

    var sheet = findSheet_();
    var ownerCol = ensureOwnerColumn_(sheet);
    var ownerKey = ownerKey_(email);
    var isAdmin = ADMIN_EMAILS.map(function (a) {
      return String(a).toLowerCase();
    }).indexOf(email) !== -1;

    if (!isAdmin && !ownerAllowed_(sheet, ownerCol, data.businessName, ownerKey)) {
      return json_({ ok: false, error: 'WRONG_OWNER' });
    }

    var logoUrl = saveLogo_(data);
    addRow_(sheet, data, logoUrl, ownerKey);
    CacheService.getScriptCache().remove('directory_csv');
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {
      // nothing to release
    }
  }
}

function json_(obj) {
  return ContentService
      .createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------
// Google sign-in check
// ---------------------------------------------------------------
function verifyToken_(token) {
  if (!token) return '';
  var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
      { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return '';
  var info = JSON.parse(resp.getContentText());
  if (info.aud !== CLIENT_ID) return '';
  if (String(info.email_verified) !== 'true') return '';
  if (Number(info.exp) * 1000 < Date.now()) return '';
  return String(info.email || '').toLowerCase();
}

// ---------------------------------------------------------------
// Checks
// ---------------------------------------------------------------
function validate_(d) {
  if (!d.businessName || !String(d.businessName).trim()) return 'BAD_NAME';

  d.phone = formatPhone_(d.phone);
  if (d.phone && d.phone.replace(/\D/g, '').length !== 10) return 'BAD_PHONE';

  d.email = String(d.email || '').trim().toLowerCase();
  if (d.email && !EMAIL_RE.test(d.email)) return 'BAD_EMAIL';
  if (d.how && String(d.how).indexOf('@') !== -1 && !EMAIL_RE.test(String(d.how).trim())) {
    return 'BAD_EMAIL';
  }

  // Web addresses are saved without https:// (just mybusiness.com)
  var urlFields = ['website', 'facebook', 'ndBiz', 'ndPersonal'];
  for (var i = 0; i < urlFields.length; i++) {
    var key = urlFields[i];
    var v = String(d[key] || '').trim();
    if (v && !URL_RE.test(v)) return 'BAD_URL';
    d[key] = v.replace(/^https?:\/\//i, '').replace(/\/+$/, '');
  }
  return '';
}

function formatPhone_(value) {
  var digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 11 && digits.charAt(0) === '1') {
    digits = digits.slice(1);
  }
  if (digits.length !== 10) {
    return String(value || '').trim();
  }
  return digits.slice(0, 3) + '-' + digits.slice(3, 6) + '-' + digits.slice(6);
}

// ---------------------------------------------------------------
// Owner key
// ---------------------------------------------------------------
function nameKey_(name) {
  return String(name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('OWNER_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('OWNER_SECRET', secret);
  }
  return secret;
}

function ownerKey_(email) {
  var bytes = Utilities.computeHmacSha256Signature(
      'owner|' + String(email).toLowerCase(), getSecret_());
  return bytes.map(function (b) {
    return ('0' + (b & 255).toString(16)).slice(-2);
  }).join('');
}

// Adds the "Owner Key (private)" column if it is not there yet
function ensureOwnerColumn_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === OWNER_HEADER) {
      return i + 1;
    }
  }
  sheet.getRange(1, lastCol + 1).setValue(OWNER_HEADER);
  return lastCol + 1;
}

// True when this business name is new, or this account owns it
function ownerAllowed_(sheet, ownerCol, businessName, ownerKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return true;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().trim().indexOf('business name') === 0) {
      nameCol = i + 1;
      break;
    }
  }
  if (nameCol < 0) return true;

  var names = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues();
  var keys = sheet.getRange(2, ownerCol, lastRow - 1, 1).getValues();
  var wanted = nameKey_(businessName);
  var found = false;
  var anyOwner = false;

  for (var r = 0; r < names.length; r++) {
    if (nameKey_(names[r][0]) !== wanted) continue;
    found = true;
    var stored = String(keys[r][0] || '');
    if (stored) {
      anyOwner = true;
      if (stored === ownerKey) return true;
    }
  }

  if (!found) return true;
  if (!anyOwner) return ALLOW_CLAIM_OLD_ROWS;
  return false;
}

// ---------------------------------------------------------------
// Public copy of the sheet (only the columns that may be shown)
// ---------------------------------------------------------------
function isPublicHeader_(header) {
  var k = String(header).toLowerCase().replace(/\s+/g, ' ').trim();
  if (!k) return false;
  if (k === 'timestamp' || k === 'name' || k === 'title') return false;
  if (/how to contact|owner key|private|email address/.test(k)) return false;
  if (/nextdoor/.test(k) && /personal|contact/.test(k)) return false;
  return /business name|public|category|about your business|other details|describe|address|city|state|zip|postal|served|area|phone|website|customers|facebook|nextdoor|slogan|additional info|logo/.test(k);
}

function publicCsv_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('directory_csv');
  if (cached) return cached;

  var sheet = findSheet_();
  var values = sheet.getDataRange().getValues();
  var keep = [];
  for (var c = 0; c < values[0].length; c++) {
    if (isPublicHeader_(values[0][c])) keep.push(c);
  }
  var lines = values.map(function (row) {
    return keep.map(function (c) {
      var v = row[c];
      return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    }).join(',');
  });
  var csv = lines.join('\n');

  try {
    cache.put('directory_csv', csv, 60);
  } catch (tooBig) {
    // the file is too big to cache; that is fine
  }
  return csv;
}

// ---------------------------------------------------------------
// Logo: save the picture in Drive and return a link
// ---------------------------------------------------------------
function saveLogo_(data) {
  if (!data.logoData) {
    return '';
  }
  var match = /^data:(image\/[\w.+-]+);base64,(.+)$/.exec(data.logoData);
  if (!match) {
    return '';
  }
  var mime = match[1];
  var ext = mime.split('/')[1].replace('jpeg', 'jpg').replace(/\+.*$/, '');
  var safeName = String(data.businessName || 'logo')
      .replace(/[^\w\- ]+/g, '').trim() || 'logo';
  var blob = Utilities.newBlob(
      Utilities.base64Decode(match[2]), mime, safeName + '.' + ext);

  var file = getLogoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/open?id=' + file.getId();
}

function getLogoFolder_() {
  var folders = DriveApp.getFoldersByName(LOGO_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(LOGO_FOLDER_NAME);
}

// ---------------------------------------------------------------
// Sheet: add one row that matches the column headings already there
// ---------------------------------------------------------------
function addRow_(sheet, data, logoUrl, ownerKey) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = headers.map(function (header) {
    if (String(header).trim() === OWNER_HEADER) return ownerKey;
    return valueFor_(header, data, logoUrl);
  });
  sheet.appendRow(row);
}

function findSheet_() {
  var sheets = SpreadsheetApp.openById(SHEET_ID).getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) {
      return sheets[i];
    }
  }
  throw new Error('Could not find the sheet tab. Check SHEET_GID.');
}

// Picks the right answer for a column heading
function valueFor_(header, d, logoUrl) {
  var k = String(header).toLowerCase().replace(/\s+/g, ' ').trim();

  if (k === 'timestamp') return new Date();
  if (k === 'name') return clean_(d.name);
  if (k === 'title') return clean_(d.title);
  if (/how to contact/.test(k)) return clean_(d.how);

  if (/nextdoor/.test(k)) {
    return clean_(/personal|contact/.test(k) ? d.ndPersonal : d.ndBiz);
  }
  if (/facebook/.test(k)) return clean_(d.facebook);

  if (/public/.test(k) && /name/.test(k)) return clean_(d.pubName);
  if (/public/.test(k) && /title/.test(k)) return clean_(d.pubTitle);

  if (k.indexOf('business name') === 0) return clean_(d.businessName);
  if (/describe/.test(k)) return clean_(d.describe);
  if (/category/.test(k)) return clean_(d.category);
  if (k.indexOf('about your business') === 0) return clean_(d.about);
  if (k.indexOf('other details') === 0) return clean_(d.other);

  if (/area|city served|served/.test(k)) return clean_(d.area);
  if (/street|address/.test(k) && !/e-?mail|web/.test(k)) return clean_(d.street);
  if (/\bcity\b/.test(k)) return clean_(d.city);
  if (/\bstate\b/.test(k)) return clean_(d.state);
  if (/\bzip\b|postal/.test(k)) return clean_(d.zip);

  if (/phone/.test(k)) return clean_(d.phone);
  if (/website|web site/.test(k)) return clean_(d.website);
  if (/customers/.test(k) && /e-?mail/.test(k)) return clean_(d.email);

  if (/slogan|additional info/.test(k)) return clean_(d.slogan);
  if (/logo/.test(k)) return logoUrl || '';

  return '';
}

// Stops a typed answer from being read as a spreadsheet formula
function clean_(value) {
  var v = String(value == null ? '' : value).trim();
  return /^[=+@-]/.test(v) ? "'" + v : v;
}
