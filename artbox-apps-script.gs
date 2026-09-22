// ArtBox — Google Apps Script backend
// Handles: sponsorship form + gallery photo uploads
// Deploy as Web App: Execute as Me, Anyone can access

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var GALLERY_FOLDER_ID = "1CYqGeseRF06Ooe3vPAhV7c1dmRPKqjxQ";

// ─── ROUTING ────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var formType = "";
    var meta = {};

    // Gallery uploads come as multipart FormData (photo + meta JSON field)
    if (e.parameters && e.parameters.meta) {
      meta = JSON.parse(e.parameters.meta[0]);
      formType = meta.formType || "";
    } else if (e.postData && e.postData.contents) {
      meta = JSON.parse(e.postData.contents);
      formType = meta.formType || "";
    }

    if (formType === "gallery") {
      return handleGalleryUpload(e, meta);
    } else if (formType === "sponsorship") {
      return handleSponsorship(meta);
    } else {
      // Legacy: original sponsorship submissions without formType field
      return handleSponsorship(meta);
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet(e) {
  try {
    var formType = e.parameter && e.parameter.formType;
    if (formType === "gallery") {
      return handleGalleryRead();
    }
    return jsonResponse({ ok: false, error: "Unknown request" });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ─── SPONSORSHIP ─────────────────────────────────────────────────────────────

var NOTIFY_TO = "artboxprogramme@gmail.com";
var NOTIFY_CC = "nimo.kanina@gmail.com";

function handleSponsorship(d) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, "Sponsorships", [
    "Timestamp", "Name", "Organisation", "Phone", "Email",
    "ArtBoxes", "Total (KSh)", "M-PESA Code",
    "School Name", "School Location",
    "Contact Name", "Contact Role", "Contact Phone", "Contact Email"
  ]);

  sheet.appendRow([
    new Date(),
    d.name || "",
    d.org || "",
    d.phone || "",
    d.email || "",
    d.boxes || "",
    d.total || "",
    d.code || "",
    d.schoolName || "",
    d.schoolLocation || "",
    d.contactName || "",
    d.contactRole || "",
    d.contactPhone || "",
    d.contactEmail || ""
  ]);

  // Email notification
  try {
    var hasSchool = d.schoolName || d.schoolLocation || d.contactName;
    var lines = [
      "New ArtBox sponsorship received.",
      "",
      "Name:        " + (d.name || ""),
      d.org ? "Organisation: " + d.org : null,
      "Phone:       " + (d.phone || ""),
      "Email:       " + (d.email || ""),
      "",
      "ArtBoxes:    " + (d.boxes || ""),
      "Total paid:  KSh " + (d.total || ""),
      "M-PESA code: " + (d.code || ""),
      hasSchool ? "" : null,
      hasSchool ? "── School preference ──────────────" : null,
      d.schoolName     ? "School:          " + d.schoolName     : null,
      d.schoolLocation ? "Location:        " + d.schoolLocation : null,
      d.contactName    ? "Contact:         " + d.contactName + (d.contactRole ? " (" + d.contactRole + ")" : "") : null,
      d.contactPhone   ? "Contact phone:   " + d.contactPhone   : null,
      d.contactEmail   ? "Contact email:   " + d.contactEmail   : null
    ].filter(function(l) { return l !== null; }).join("\n");

    // Notify ArtBox team
    MailApp.sendEmail({
      to:      NOTIFY_TO,
      cc:      NOTIFY_CC,
      subject: "ArtBox Sponsorship — " + (d.name || "unknown") + " (" + (d.boxes || "?") + " box)",
      body:    lines
    });

    // Confirmation to sponsor
    if (d.email) {
      var boxWord   = d.boxes == 1 ? "ArtBox" : "ArtBoxes";
      var confirmed = [
        "Hi " + (d.name || "there") + ",",
        "",
        "Thank you for sponsoring " + (d.boxes || "an") + " " + boxWord + "! Your support gives " + ((d.boxes || 1) * 50) + " children a full term of art.",
        "",
        "Here's a summary of your sponsorship:",
        "",
        "  ArtBoxes:    " + (d.boxes || ""),
        "  Total paid:  KSh " + (d.total || ""),
        "  M-PESA code: " + (d.code || ""),
        d.schoolName ? "  School:       " + d.schoolName : null,
        "",
        "We'll confirm your M-PESA payment shortly. If you named a school, we'll be in touch to coordinate delivery.",
        "",
        "With thanks,",
        "Jeremy & Nimo",
        "ArtBox — artboxprogramme@gmail.com"
      ].filter(function(l) { return l !== null; }).join("\n");

      MailApp.sendEmail({
        to:      d.email,
        subject: "Your ArtBox sponsorship — thank you!",
        body:    confirmed
      });
    }
  } catch (mailErr) {
    // Don't fail the whole submission if email errors
  }

  return jsonResponse({ ok: true });
}

// ─── GALLERY UPLOAD ───────────────────────────────────────────────────────────

function handleGalleryUpload(e, meta) {
  var school        = (meta.school || "").trim();
  var submitterName = (meta.submitterName || "").trim();
  var submitterPos  = (meta.submitterPos || "").trim();
  var childName     = (meta.childName || "").trim();
  var childAge      = (meta.childAge || "").toString().trim();
  var childDesc     = (meta.childDesc || "").trim();

  if (!school || !childName) {
    return jsonResponse({ ok: false, error: "Missing required fields" });
  }

  // Get or create school subfolder
  var rootFolder   = DriveApp.getFolderById(GALLERY_FOLDER_ID);
  var schoolSlug   = slugify(school);
  var schoolFolder = getOrCreateFolder(rootFolder, schoolSlug);

  // Save photo
  var fileId = "";
  if (e.parameters && e.parameters.photo) {
    // File arrives as base64 via FormData workaround
    var b64   = e.parameters.photo[0];
    var bytes = Utilities.base64Decode(b64);
    var blob  = Utilities.newBlob(bytes, "image/jpeg", buildFilename(schoolSlug, childName, childAge));
    var file  = schoolFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = file.getId();
  } else {
    return jsonResponse({ ok: false, error: "No photo received" });
  }

  // Log to sheet
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, "Gallery Submissions", [
    "Timestamp", "School", "Submitted By Name", "Submitted By Position",
    "Child First Name", "Age", "Description", "Drive File ID"
  ]);

  sheet.appendRow([
    new Date(),
    school,
    submitterName,
    submitterPos,
    childName,
    childAge,
    childDesc,
    fileId
  ]);

  return jsonResponse({ ok: true, fileId: fileId });
}

// ─── GALLERY READ ─────────────────────────────────────────────────────────────

function handleGalleryRead() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Gallery Submissions");

  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ ok: true, rows: [] });
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var rows = [];

  data.forEach(function(row) {
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

  return jsonResponse({ ok: true, rows: rows });
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground("#f0f0f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

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
  var pad = function(n) { return n < 10 ? "0" + n : n.toString(); };
  var datePart = now.getFullYear().toString() +
                 pad(now.getMonth() + 1) +
                 pad(now.getDate()) +
                 "-" +
                 pad(now.getHours()) +
                 pad(now.getMinutes()) +
                 pad(now.getSeconds());
  var safeName = childName.replace(/[^a-zA-Z]/g, "");
  return schoolSlug + "_" + safeName + "_" + age + "_" + datePart + ".jpg";
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
