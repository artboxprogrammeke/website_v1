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

function doPost(e) {
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
