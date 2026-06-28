/* ====================================================================
   SCRIM REGISTRATION BACKEND — Google Apps Script
   ----------------------------------------------------------------
   SETUP (do this once):
   1. Go to https://script.google.com → New Project
   2. Delete the default code, paste this whole file in
   3. Edit RECIPIENTS / SHEET_ID below if needed
   4. Click Deploy → New deployment → type: Web app
        - Execute as: Me
        - Who has access: Anyone
   5. Copy the Web App URL it gives you
   6. Paste that URL into APPS_SCRIPT_URL in index.html's CONFIG block
   ==================================================================== */

const RECIPIENTS = "imamericanyess@gmail.com,kumarabhinav7349@gmail.com";

// Optional: paste a Google Sheet ID here to also log every entry as a row.
// Leave as-is to skip sheet logging (email will still work).
const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";

const DRIVE_FOLDER_NAME = "BGMI Scrim Payment Screenshots";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    let screenshotUrl = "Not provided";
    if (data.screenshotBase64) {
      screenshotUrl = saveScreenshotToDrive(data);
    }

    MailApp.sendEmail({
      to: RECIPIENTS,
      subject: "New Scrim Registration — " + (data.teamName || "Unnamed Team"),
      body: buildEmailBody(data, screenshotUrl)
    });

    if (SHEET_ID && SHEET_ID !== "PASTE_YOUR_GOOGLE_SHEET_ID_HERE") {
      logToSheet(data, screenshotUrl);
    }

    return jsonResponse({ status: "success" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function saveScreenshotToDrive(data) {
  const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const base64 = data.screenshotBase64.split(",").pop(); // strip data:...;base64, prefix
  const bytes = Utilities.base64Decode(base64);
  const filename = (data.teamName || "team").replace(/[^a-z0-9]/gi, "_") + "_" + Date.now() + ".png";
  const blob = Utilities.newBlob(bytes, data.screenshotMime || "image/png", filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function buildEmailBody(data, screenshotUrl) {
  const lines = [];
  lines.push("TEAM: " + (data.teamName || "—"));
  lines.push("CAPTAIN (Google account): " + data.googleName + " <" + data.googleEmail + ">");
  lines.push("WHATSAPP: " + data.whatsapp);
  lines.push("TXN REF SHOWN ON QR: " + (data.txnRef || "—"));
  lines.push("");
  for (let i = 1; i <= 4; i++) {
    lines.push("Player " + i + ": " + data["p" + i + "name"] + "  |  UID: " + data["p" + i + "uid"]);
  }
  lines.push("");
  lines.push("Payment screenshot: " + screenshotUrl);
  lines.push("Submitted: " + new Date().toLocaleString());
  lines.push("");
  lines.push("⚠ Manual step: verify the ₹20 payment actually landed before confirming this team's slot.");
  return lines.join("\n");
}

function logToSheet(data, screenshotUrl) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  sheet.appendRow([
    new Date(),
    data.teamName, data.googleName, data.googleEmail, data.whatsapp,
    data.p1name, data.p1uid,
    data.p2name, data.p2uid,
    data.p3name, data.p3uid,
    data.p4name, data.p4uid,
    data.txnRef, screenshotUrl,
    "Pending" // payment verification status — update manually
  ]);
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
