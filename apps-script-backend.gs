/* ====================================================================
   SCRIM REGISTRATION BACKEND — Google Apps Script
   ----------------------------------------------------------------
   Does two jobs:
   1. doPost  -> receives form submissions, emails both addresses,
                 saves payment screenshot to Drive, logs a row.
   2. doGet   -> serves the LIVE SETTINGS (map, schedule, prizes,
                 links, entry fee) that index.html reads on load.
                 Edit these by editing the Google Sheet — no code
                 changes, no re-uploading to GitHub.

   SETUP (do this once):
   1. Make a blank Google Sheet. Copy its ID from the URL
      (the long string between /d/ and /edit) into SHEET_ID below.
   2. Go to https://script.google.com -> New Project
   3. Delete the default code, paste this whole file in
   4. Edit RECIPIENTS / SHEET_ID below
   5. Deploy -> New deployment -> type: Web app
        - Execute as: Me
        - Who has access: Anyone
   6. Copy the Web App URL it gives you
   7. Paste that URL into APPS_SCRIPT_URL in index.html's CONFIG block
   8. Open the Sheet once -> a "Settings" tab will appear automatically
      the first time the site loads (or run setupSettingsSheet() manually
      from the Apps Script editor's "Run" button). Edit the Value column
      any time you want to change what's shown on the site.
   ==================================================================== */

const RECIPIENTS = "imamericanyess@gmail.com,kumarabhinav7349@gmail.com";

// Paste your Google Sheet ID here (required for both settings + logging).
const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";

const DRIVE_FOLDER_NAME = "BGMI Scrim Payment Screenshots";

// Default values used to auto-create the Settings tab the first time.
// After that, EDIT THE SHEET ITSELF — these defaults are only a seed.
const DEFAULT_SETTINGS = [
  ["mode", "Erangel · TPP Squad"],
  ["schedule", "Sat & Sun · 8 PM IST"],
  ["winPrize", "₹20 / kill"],
  ["mvpPrize", "₹5 / kill"],
  ["liveStream", "https://www.youtube.com/@VirgoYT707"],
  ["whatsappGroup", "https://chat.whatsapp.com/E8lXkGkA4ShE9yRrnOWhcC"],
  ["entryFee", "20"],
  ["upiId", "princeraj21@fam"],
  ["upiPayeeName", "PRINCE"],
  ["roomId", ""],
  ["roomTime", ""]
];

/* ================== doGet: serves live settings ================== */
// Only these keys are ever sent to the public website. roomId/roomTime are
// deliberately left out here — they go out by private email only, via
// broadcastRoomId() below, never through this public endpoint.
const PUBLIC_SETTINGS_KEYS = [
  "mode", "schedule", "winPrize", "mvpPrize",
  "liveStream", "whatsappGroup", "entryFee", "upiId", "upiPayeeName"
];

function doGet(e) {
  try {
    const sheet = getOrCreateSettingsSheet();
    const rows = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < rows.length; i++) {
      const key = rows[i][0];
      const value = rows[i][1];
      if (key && PUBLIC_SETTINGS_KEYS.indexOf(key) !== -1) settings[key] = value;
    }
    return jsonResponse({ status: "success", settings: settings });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

/* ================== doPost: form submissions ================== */
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

    logToSheet(data, screenshotUrl);

    return jsonResponse({ status: "success" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function saveScreenshotToDrive(data) {
  const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const base64 = data.screenshotBase64.split(",").pop();
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
  lines.push("⚠ Manual step: verify the payment actually landed before confirming this team's slot.");
  return lines.join("\n");
}

function logToSheet(data, screenshotUrl) {
  const sheet = getOrCreateRegistrationsSheet();
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

/* ================== sheet + folder helpers ================== */
function getOrCreateRegistrationsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Registrations");
  if (!sheet) {
    sheet = ss.insertSheet("Registrations");
    sheet.appendRow([
      "Timestamp", "Team", "Captain Name", "Captain Email", "WhatsApp",
      "P1 Name", "P1 UID", "P2 Name", "P2 UID", "P3 Name", "P3 UID",
      "P4 Name", "P4 UID", "Txn Ref", "Screenshot", "Status"
    ]);
  }
  return sheet;
}

function getOrCreateSettingsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["Key", "Value"]);
    DEFAULT_SETTINGS.forEach(row => sheet.appendRow(row));
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
  }
  return sheet;
}

// Run this once manually from the Apps Script editor (Run button) if you
// want the Settings tab created before the site has loaded for the first time.
function setupSettingsSheet() {
  getOrCreateSettingsSheet();
  getOrCreateRegistrationsSheet();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ====================================================================
   APPROVAL NOTIFICATION
   ----------------------------------------------------------------
   This is NOT a simple trigger — it needs to be wired up manually
   as an "installable trigger" (see instructions below), because
   sending email requires authorization that simple onEdit() doesn't
   have.

   SETUP (once):
   1. Save this file
   2. Left sidebar -> click the clock icon ("Triggers")
   3. + Add Trigger
   4. Choose function: onStatusChange
   5. Event source: From spreadsheet
   6. Event type: On edit
   7. Save -> approve the permissions popup if asked

   USAGE:
   Open the Registrations tab in the Sheet, find the row for the team
   you've checked, change its "Status" cell from "Pending" to
   "Approved" -> the captain instantly gets a confirmation email.
   ==================================================================== */
function onStatusChange(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== "Registrations") return;
    if (e.range.getColumn() !== 16) return; // "Status" is column 16 (P)

    const newStatus = String(e.range.getValue()).trim().toLowerCase();
    if (newStatus !== "approved" && newStatus !== "confirmed") return;

    const row = e.range.getRow();
    const rowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
    const teamName = rowData[1];
    const captainEmail = rowData[3];
    if (!captainEmail) return;

    MailApp.sendEmail({
      to: captainEmail,
      subject: "✅ Squad Confirmed — " + teamName,
      body: "Your squad \"" + teamName + "\" is confirmed for the scrim!\n\n" +
            "Room ID and lobby time will be posted in the WhatsApp group shortly before the match — keep notifications on.\n\n" +
            "See you in the lobby.\n— VIRGO YT 707"
    });
  } catch (err) {
    // Don't let a failed email break the sheet edit.
  }
}

/* ====================================================================
   ROOM ID BROADCAST
   ----------------------------------------------------------------
   Run this manually right before the match starts. It emails the
   room ID + lobby time to every team currently marked "Approved" —
   nobody else sees it, and it's never exposed on the public website.

   USAGE (every match):
   1. Open the Settings tab -> fill in "roomId" and "roomTime"
   2. In this script editor, tap the function dropdown (next to the
      Run button, top toolbar) -> select "broadcastRoomId"
   3. Tap Run
   4. Check Execution log (left sidebar, clock-with-list icon) to see
      how many teams were notified
   ==================================================================== */
function broadcastRoomId() {
  const settingsSheet = getOrCreateSettingsSheet();
  const settingsRows = settingsSheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < settingsRows.length; i++) {
    if (settingsRows[i][0]) settings[settingsRows[i][0]] = settingsRows[i][1];
  }

  const roomId = settings.roomId;
  const roomTime = settings.roomTime;
  if (!roomId) {
    Logger.log("Stop: 'roomId' is empty in the Settings tab. Fill it in first.");
    return;
  }

  const regSheet = getOrCreateRegistrationsSheet();
  const rows = regSheet.getDataRange().getValues();
  let sentCount = 0;

  for (let i = 1; i < rows.length; i++) {
    const status = String(rows[i][15]).trim().toLowerCase(); // column 16 = Status
    const teamName = rows[i][1];
    const captainEmail = rows[i][3];
    if ((status === "approved" || status === "confirmed") && captainEmail) {
      MailApp.sendEmail({
        to: captainEmail,
        subject: "🎮 Room ID & Time — " + teamName,
        body: "Squad: " + teamName + "\n\n" +
              "ROOM ID: " + roomId + "\n" +
              (roomTime ? "LOBBY TIME: " + roomTime + "\n" : "") +
              "\nJoin a few minutes early. Don't share this ID outside your squad.\n\nSee you in the lobby!\n— VIRGO YT 707"
      });
      sentCount++;
    }
  }
  Logger.log(sentCount + " approved team(s) notified with the room ID.");
}
