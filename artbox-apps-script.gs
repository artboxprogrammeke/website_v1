/**
 * ArtBox Sponsor Form — Google Apps Script backend
 *
 * On every form submission:
 *   1. Adds a row to the "Sponsorships" tab of this Google Sheet
 *   2. Emails the ArtBox account (cc Nimo) with the details
 *   3. Emails the sponsor a copy of their sponsorship
 */

var NOTIFY_TO = "artboxprogramme@gmail.com";
var NOTIFY_CC = "nimo.kanina@gmail.com";
var SHEET_NAME = "Sponsorships";

var GALLERY_FOLDER_ID  = "1CYqGeseRF06Ooe3vPAhV7c1dmRPKqjxQ";
var GALLERY_SHEET_NAME = "Gallery Submissions";

function doGet(e) {
  try {
    if (e.parameter && e.parameter.formType === "gallery") {
      return handleGalleryRead();
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "Unknown request" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  // Gallery uploads arrive as application/x-www-form-urlencoded
  // with a `meta` field — route them before the sponsorship JSON parse
  if (e.parameters && e.parameters.meta) {
    try {
      var meta = JSON.parse(e.parameters.meta[0]);
      return handleGalleryUpload(e, meta);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  try {
    var d = JSON.parse(e.postData.contents);

    // --- 1. Record in the sheet ---
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Name", "Organisation", "Phone", "Email",
        "ArtBoxes", "Total (KSh)", "M-PESA Code",
        "School Name(s)", "School Location", "Contact Name", "Contact Role",
        "Contact Phone", "Contact Email", "Duplicate code?"
      ]);
      sheet.setFrozenRows(1);
    }

    // Flag if this M-PESA code was already submitted
    var existingCodes = sheet.getLastRow() > 1
      ? sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues().flat()
      : [];
    var duplicate = existingCodes.indexOf(d.code) !== -1 ? "YES — check!" : "";

    var hasSchool = d.schoolName || d.schoolLocation || d.contactName;

    sheet.appendRow([
      new Date(), d.name, d.org || "", d.phone, d.email,
      d.boxes, d.total, d.code,
      d.schoolName || "", d.schoolLocation || "", d.contactName || "",
      d.contactRole || "", d.contactPhone || "", d.contactEmail || "",
      duplicate
    ]);

    var totalFmt = "KSh " + Number(d.total).toLocaleString("en-KE");
    var boxWord = d.boxes == 1 ? "ArtBox" : "ArtBoxes";

    // Shared school block, built once and reused in both emails
    var schoolLines = hasSchool
      ? "\n--- School preference ---\n" +
        (d.schoolName     ? "School: " + d.schoolName + "\n" : "") +
        (d.schoolLocation ? "Location: " + d.schoolLocation + "\n" : "") +
        (d.contactName    ? "Contact: " + d.contactName + (d.contactRole ? " (" + d.contactRole + ")" : "") + "\n" : "") +
        (d.contactPhone   ? "Contact phone: " + d.contactPhone + "\n" : "") +
        (d.contactEmail   ? "Contact email: " + d.contactEmail + "\n" : "")
      : "";

    // --- 2. Notify the ArtBox account (cc Nimo) ---
    var teamBody =
      "New ArtBox sponsorship received!\n\n" +
      "Name: " + d.name + "\n" +
      (d.org ? "Organisation: " + d.org + "\n" : "") +
      "Phone: " + d.phone + "\n" +
      "Email: " + d.email + "\n\n" +
      "ArtBoxes: " + d.boxes + "\n" +
      "Total: " + totalFmt + "\n" +
      "M-PESA code: " + d.code + "\n" +
      (duplicate ? "\n⚠️ NOTE: this M-PESA code was submitted before — please verify.\n" : "") +
      schoolLines +
      "\nRecorded in the Sponsorships sheet.";

    MailApp.sendEmail({
      to: NOTIFY_TO,
      cc: NOTIFY_CC,
      subject: "🎨 New ArtBox sponsorship — " + d.name + " (" + d.boxes + " " + boxWord + ")",
      body: teamBody
    });

    // --- 3. Confirmation copy to the sponsor ---
    var sponsorBody =
      "Hi " + d.name + ",\n\n" +
      "Asante sana for sponsoring " + d.boxes + " " + boxWord + "! " +
      "Here is a copy of your sponsorship for your records:\n\n" +
      "ArtBoxes sponsored: " + d.boxes + "\n" +
      "Total: " + totalFmt + "\n" +
      "M-PESA code: " + d.code + "\n" +
      (d.org ? "Organisation: " + d.org + "\n" : "") +
      schoolLines +
      "\nEvery ArtBox gives 50 students a full term of weekly art sessions — " +
      "supplies, activities and all. We'll confirm your M-PESA payment for your " +
      d.boxes + " " + boxWord + " and follow up with next steps. " +
      (hasSchool
        ? "We'll be in touch about the school you named to coordinate delivery.\n\n"
        : "\n\n") +
      "With thanks,\n" +
      "The ArtBox Team\n" +
      "ArtBox — A School Art Programme\n" +
      "artboxprogramme@gmail.com · +254 714 435 756";

    MailApp.sendEmail({
      to: d.email,
      subject: "Your ArtBox sponsorship — thank you! 🎨",
      body: sponsorBody
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── GALLERY UPLOAD ───────────────────────────────────────────────────────────

function handleGalleryUpload(e, meta) {
  var school        = (meta.school        || "").trim();
  var submitterName = (meta.submitterName || "").trim();
  var submitterPos  = (meta.submitterPos  || "").trim();
  var childName     = (meta.childName     || "").trim();
  var childAge      = (meta.childAge      || "").toString().trim();
  var childDesc     = (meta.childDesc     || "").trim();

  if (!school || !childName) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "Missing required fields" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!e.parameters || !e.parameters.photo) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "No photo received" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Save photo to Drive
  var rootFolder   = DriveApp.getFolderById(GALLERY_FOLDER_ID);
  var schoolSlug   = slugify(school);
  var schoolFolder = getOrCreateFolder(rootFolder, schoolSlug);

  var b64   = e.parameters.photo[0];
  var bytes = Utilities.base64Decode(b64);
  var blob  = Utilities.newBlob(bytes, "image/jpeg", buildFilename(schoolSlug, childName, childAge));
  var file   = schoolFolder.createFile(blob);
  var fileId = file.getId();
  UrlFetchApp.fetch(
    "https://www.googleapis.com/drive/v3/files/" + fileId + "/permissions",
    {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      payload: JSON.stringify({ role: "reader", type: "anyone" }),
      muteHttpExceptions: true
    }
  );

  // Log one row per child to Gallery Submissions sheet
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GALLERY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(GALLERY_SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "School", "Submitted By Name", "Submitted By Position",
      "Child First Name", "Age", "Description", "Drive File ID"
    ]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(), school, submitterName, submitterPos,
    childName, childAge, childDesc, fileId
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, fileId: fileId }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GALLERY READ ─────────────────────────────────────────────────────────────

function handleGalleryRead() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GALLERY_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, rows: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var rows = [];
  data.forEach(function (row) {
    var fileId = (row[7] || "").toString().trim();
    if (!fileId) return;
    rows.push({
      timestamp:   row[0] ? new Date(row[0]).toISOString() : "",
      school:      row[1] || "",
      childName:   row[4] || "",
      age:         row[5] ? row[5].toString() : "",
      description: row[6] || "",
      fileId:      fileId
    });
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function getOrCreateFolder(parent, name) {
  var iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}

function slugify(str) {
  return str.trim()
            .replace(/[^a-zA-Z0-9\s-]/g, "")
            .trim()
            .replace(/\s+/g, "-");
}

function buildFilename(schoolSlug, childName, age) {
  var now = new Date();
  var pad = function (n) { return n < 10 ? "0" + n : String(n); };
  var stamp = now.getFullYear() +
              pad(now.getMonth() + 1) +
              pad(now.getDate()) + "-" +
              pad(now.getHours()) +
              pad(now.getMinutes()) +
              pad(now.getSeconds());
  return schoolSlug + "_" + childName.replace(/[^a-zA-Z]/g, "") + "_" + age + "_" + stamp + ".jpg";
}
