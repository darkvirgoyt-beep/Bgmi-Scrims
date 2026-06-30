/* ====================================================================
   SCRIM REGISTRATION BACKEND — Google Apps Script
   ----------------------------------------------------------------
   Three jobs:
   1. doGet   -> serves PUBLIC live settings (prizes, links, fee, the
                 currently active map + live capacity) that index.html
                 reads on load. Also serves ADMIN data when called with
                 ?admin=1&password=... (used by admin.html).
   2. doPost  -> receives form submissions, FamPay payment notifications
                 (from MacroDroid), and admin actions (map switch, room
                 ID broadcast, approve team) from admin.html.
   3. Round sheets -> every registration is written into a sheet named
                 "<Map>_<YYYY-MM-DD>" (e.g. "Livik_2026-06-30"). Whenever
                 you change "activeMap" (via the Sheet or the admin
                 panel), a brand new round sheet is used automatically —
                 old rounds are untouched, capacity always starts at 0
                 for a new round.

   SETUP (do this once):
   1. Make a blank Google Sheet. Copy its ID from the URL
      (the long string between /d/ and /edit) into SHEET_ID below.
   2. Go to https://script.google.com -> New Project
   3. Delete the default code, paste this whole file in
   4. Edit RECIPIENTS / SHEET_ID below
   5. Save. Deploy -> New deployment -> type: Web app
        - Execute as: Me
        - Who has access: Anyone
   6. Copy the Web App URL it gives you (ends in /exec)
   7. Paste that URL into APPS_SCRIPT_URL in BOTH index.html's CONFIG
      block and admin.html's CONFIG block
   8. Open the Sheet -> a "Settings" tab appears automatically the first
      time the site loads (or run setupSettingsSheet() once from the
      Apps Script editor's Run button to create it immediately)
   9. In the Settings tab, change the "adminPassword" row's Value to a
      real password of your choosing — the default is a placeholder,
      not a real password. Then open yoursite.com/admin.html and log in.

   IMPORTANT — editing this file later: after pasting changes in, you
   only need to tap Save (disk icon). You do NOT need a new deployment
   or a new URL — "Manage deployments" already serves the latest saved
   code for the existing /exec link.
   ==================================================================== */

const RECIPIENTS = "imamericanyess@gmail.com,kumarabhinav7349@gmail.com";

// Paste your Google Sheet ID here (required for settings + round sheets).
const SHEET_ID = "11joNjVKYi4CrnNOBHYMX_aIer4eiKnzpg0VaWfm2WdQ";

const DRIVE_FOLDER_NAME = "BGMI Scrim Payment Screenshots";
const TIMEZONE = "Asia/Kolkata";

// Default values used to auto-create the Settings tab the first time.
// After that, EDIT THE SHEET ITSELF (or use admin.html) — these defaults
// are only a seed, changing them here won't affect an existing sheet.
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
  ["roomTime", ""],
  ["activeMap", "erangel"],
  ["erangelMax", "100"],
  ["livikMax", "50"],
  ["miramarMax", "100"],
  ["adminPassword", "CHANGE_ME_NOW"]
];

// Only these keys are ever sent to the PUBLIC website (index.html).
// adminPassword/roomId/roomTime are deliberately excluded — those stay
// server-side / admin-only and are never exposed through this endpoint.
const PUBLIC_SETTINGS_KEYS = [
  "mode", "schedule", "winPrize", "mvpPrize",
  "liveStream", "whatsappGroup", "entryFee", "upiId", "upiPayeeName"
];

const MAP_MAX_KEYS = { erangel: "erangelMax", livik: "livikMax", miramar: "miramarMax" };
const VALID_MAPS = ["erangel", "livik", "miramar"];

const REG_HEADERS = [
  "Timestamp", "Team", "Captain Name", "Captain Email", "WhatsApp",
  "P1 Name", "P1 UID", "P2 Name", "P2 UID", "P3 Name", "P3 UID",
  "P4 Name", "P4 UID", "Txn Ref", "Screenshot", "Status"
];
// Column numbers (1-indexed, matches REG_HEADERS order) used in a few places:
const COL_CAPTAIN_EMAIL = 4;
const COL_TXN_REF = 14;
const COL_STATUS = 16;

/* ====================================================================
   SETTINGS HELPERS
   ==================================================================== */
function getAllSettings() {
  const sheet = getOrCreateSettingsSheet();
  const rows = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
  }
  return settings;
}

function setSetting(key, value) {
  const sheet = getOrCreateSettingsSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getOrCreateSettingsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["Key", "Value"]);
    DEFAULT_SETTINGS.forEach(function (row) { sheet.appendRow(row); });
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
  }
  return sheet;
}

function setupSettingsSheet() {
  getOrCreateSettingsSheet();
}

/* ====================================================================
   ROUND-SHEET HELPERS (one sheet per map per date)
   ==================================================================== */
function capitalize(s) {
  s = String(s || "");
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function getTodayDateString() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
}

function getRoundSheetName(map) {
  return capitalize(map) + "_" + getTodayDateString();
}

function getOrCreateRoundSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(REG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function countTeamsInSheet(sheet) {
  return Math.max(0, sheet.getLastRow() - 1); // minus header row
}

/* ====================================================================
   doGet — public settings/capacity, or admin data dump when
   ?admin=1&password=... is supplied (used by admin.html)
   ==================================================================== */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.admin === "1") return handleAdminGet(params);

    const settings = getAllSettings();
    const publicSettings = {};
    PUBLIC_SETTINGS_KEYS.forEach(function (k) {
      if (settings[k] !== undefined) publicSettings[k] = settings[k];
    });

    const activeMap = String(settings.activeMap || "erangel").toLowerCase();
    const maxKey = MAP_MAX_KEYS[activeMap] || "erangelMax";
    const mapMax = Number(settings[maxKey] || 100);
    const roundSheet = getOrCreateRoundSheet(getRoundSheetName(activeMap));
    const playersRegistered = countTeamsInSheet(roundSheet) * 4;

    publicSettings.activeMap = activeMap;
    publicSettings.mapMax = mapMax;
    publicSettings.playersRegistered = playersRegistered;
    publicSettings.mapFull = playersRegistered >= mapMax;

    return jsonResponse({ status: "success", settings: publicSettings });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

/* ====================================================================
   doPost — registrations, FamPay payment notifications, admin actions
   ==================================================================== */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === "payment_notification") return handlePaymentNotification(data);
    if (data.type === "admin_action") return handleAdminAction(data);
    return handleRegistration(data);
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function handleRegistration(data) {
  const settings = getAllSettings();
  const activeMap = String(settings.activeMap || "erangel").toLowerCase();
  const maxKey = MAP_MAX_KEYS[activeMap] || "erangelMax";
  const mapMax = Number(settings[maxKey] || 100);
  const roundSheet = getOrCreateRoundSheet(getRoundSheetName(activeMap));

  // Server-side capacity check — never trust the client about this.
  if (countTeamsInSheet(roundSheet) * 4 >= mapMax) {
    return jsonResponse({ status: "full", message: "This map's lobby is full." });
  }

  let screenshotUrl = "Not provided";
  if (data.screenshotBase64) screenshotUrl = saveScreenshotToDrive(data);

  MailApp.sendEmail({
    to: RECIPIENTS,
    subject: "New Scrim Registration — " + (data.teamName || "Unnamed Team") + " [" + capitalize(activeMap) + "]",
    body: buildEmailBody(data, screenshotUrl, activeMap)
  });

  roundSheet.appendRow([
    new Date(),
    data.teamName, data.googleName, data.googleEmail, data.whatsapp,
    data.p1name, data.p1uid,
    data.p2name, data.p2uid,
    data.p3name, data.p3uid,
    data.p4name, data.p4uid,
    data.txnRef, screenshotUrl,
    "Pending" // verification status — approve via Sheet or admin.html
  ]);

  return jsonResponse({ status: "success" });
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

function buildEmailBody(data, screenshotUrl, map) {
  const lines = [];
  lines.push("MAP: " + capitalize(map));
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
  lines.push("⚠ Verify the payment actually landed before approving (Sheet or admin.html).");
  return lines.join("\n");
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

/* ====================================================================
   PAYMENT NOTIFICATION HANDLER (the FamPay/MacroDroid hack)
   ----------------------------------------------------------------
   MacroDroid POSTs here whenever a FamPay notification fires:
   { "type": "payment_notification", "title": "...", "text": "..." }

   Matching strategy, most to least reliable:
   1. Notification text contains our SCRIM ref code -> 100% reliable.
   2. Otherwise, match by amount among "Pending" rows in the CURRENT
      round sheet, submitted in the last 15 minutes. Only auto-approves
      if exactly ONE candidate matches — ambiguous cases email you
      instead of guessing wrong.

   Honesty note: this reads your own phone's notifications via
   MacroDroid — it's a best-effort hack, not an official bank/UPI
   integration. Keep an eye on the admin panel early on to confirm it's
   matching correctly.
   ==================================================================== */
function handlePaymentNotification(data) {
  const rawText = ((data.text || "") + " " + (data.title || "")).trim();
  const amountMatch = rawText.match(/(?:₹|rs\.?|inr)\s?([0-9]+(?:\.[0-9]{1,2})?)/i);
  const refMatch = rawText.match(/SCRIM[A-Z0-9]{6}/i);
  const amount = amountMatch ? amountMatch[1] : null;
  const ref = refMatch ? refMatch[0].toUpperCase() : null;

  const settings = getAllSettings();
  const activeMap = String(settings.activeMap || "erangel").toLowerCase();
  const sheet = getOrCreateRoundSheet(getRoundSheetName(activeMap));
  const rows = sheet.getDataRange().getValues();

  if (ref) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][COL_TXN_REF - 1]).toUpperCase() === ref) {
        approveRow(sheet, i + 1, rows[i]);
        return jsonResponse({ status: "success", matched: "ref" });
      }
    }
  }

  if (amount) {
    const now = new Date();
    const candidates = [];
    for (let i = 1; i < rows.length; i++) {
      const status = String(rows[i][COL_STATUS - 1]).trim().toLowerCase();
      const timestamp = new Date(rows[i][0]);
      const minutesAgo = (now - timestamp) / 60000;
      if (status === "pending" && minutesAgo >= 0 && minutesAgo <= 15) candidates.push(i);
    }
    if (candidates.length === 1) {
      approveRow(sheet, candidates[0] + 1, rows[candidates[0]]);
      return jsonResponse({ status: "success", matched: "amount-unique" });
    } else if (candidates.length > 1) {
      MailApp.sendEmail({
        to: RECIPIENTS,
        subject: "⚠ Payment notification — too many possible matches",
        body: "Got a FamPay notification for ₹" + amount + " but found " + candidates.length +
              " pending entries in the last 15 minutes in " + sheet.getName() +
              " — too ambiguous to auto-approve safely.\n\n" +
              "Raw notification: " + rawText + "\n\nPlease check and approve manually (Sheet or admin.html)."
      });
      return jsonResponse({ status: "ambiguous", candidates: candidates.length });
    }
  }

  // No amount detected — probably an unrelated FamPay notification (offer, promo, etc).
  return jsonResponse({ status: "ignored" });
}

function approveRow(sheet, rowNumber, rowData) {
  sheet.getRange(rowNumber, COL_STATUS).setValue("Approved");
  const teamName = rowData[1];
  const captainEmail = rowData[COL_CAPTAIN_EMAIL - 1];
  if (captainEmail) {
    MailApp.sendEmail({
      to: captainEmail,
      subject: "✅ Squad Confirmed — " + teamName,
      body: "Your squad \"" + teamName + "\" is confirmed for the scrim!\n\n" +
            "Room ID and lobby time will be sent to this email once posted, and announced in the WhatsApp group — keep notifications on.\n\n" +
            "See you in the lobby.\n— VIRGO YT 707"
    });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ====================================================================
   APPROVAL EMAIL VIA DIRECT SHEET EDIT
   ----------------------------------------------------------------
   Lets you approve a team by typing "Approved" straight into the
   Status column of a round sheet (instead of using admin.html), and
   still get the captain notified automatically.

   SETUP (once):
   1. Save this file
   2. Left sidebar -> clock icon ("Triggers") -> + Add Trigger
   3. Function: onStatusChange | Event source: From spreadsheet
      Event type: On edit -> Save -> approve permissions if asked
   ==================================================================== */
function onStatusChange(e) {
  try {
    const sheet = e.range.getSheet();
    const name = sheet.getName();
    if (!/^(Erangel|Livik|Miramar)_\d{4}-\d{2}-\d{2}$/.test(name)) return;
    if (e.range.getColumn() !== COL_STATUS) return;

    const newStatus = String(e.range.getValue()).trim().toLowerCase();
    if (newStatus !== "approved" && newStatus !== "confirmed") return;

    const row = e.range.getRow();
    const rowData = sheet.getRange(row, 1, 1, COL_STATUS).getValues()[0];
    const teamName = rowData[1];
    const captainEmail = rowData[COL_CAPTAIN_EMAIL - 1];
    if (!captainEmail) return;

    MailApp.sendEmail({
      to: captainEmail,
      subject: "✅ Squad Confirmed — " + teamName,
      body: "Your squad \"" + teamName + "\" is confirmed for the scrim!\n\n" +
            "Room ID and lobby time will be sent to this email once posted, and announced in the WhatsApp group — keep notifications on.\n\n" +
            "See you in the lobby.\n— VIRGO YT 707"
    });
  } catch (err) {
    // Don't let a failed email break the sheet edit.
  }
}

/* ====================================================================
   ROOM ID BROADCAST (manual Run-button version — admin.html has a
   one-tap version of this same thing)
   ==================================================================== */
function broadcastRoomIdInternal() {
  const settings = getAllSettings();
  const roomId = settings.roomId;
  const roomTime = settings.roomTime;
  if (!roomId) return 0;

  const activeMap = String(settings.activeMap || "erangel").toLowerCase();
  const sheet = getOrCreateRoundSheet(getRoundSheetName(activeMap));
  const rows = sheet.getDataRange().getValues();
  let sentCount = 0;

  for (let i = 1; i < rows.length; i++) {
    const status = String(rows[i][COL_STATUS - 1]).trim().toLowerCase();
    const teamName = rows[i][1];
    const captainEmail = rows[i][COL_CAPTAIN_EMAIL - 1];
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
  return sentCount;
}

// Run this manually (Run button, function dropdown -> broadcastRoomId)
// if you'd rather not use admin.html for this.
function broadcastRoomId() {
  const sentCount = broadcastRoomIdInternal();
  Logger.log(sentCount + " approved team(s) notified.");
}

/* ====================================================================
   ADMIN ENDPOINTS (used by admin.html)
   ==================================================================== */
function checkAdminPassword(password) {
  const settings = getAllSettings();
  const stored = settings.adminPassword;
  return !!stored && !!password && String(password) === String(stored) && String(stored) !== "CHANGE_ME_NOW";
}

function handleAdminGet(params) {
  if (!checkAdminPassword(params.password)) {
    return jsonResponse({ status: "unauthorized" });
  }
  const settings = getAllSettings();
  const activeMap = String(settings.activeMap || "erangel").toLowerCase();
  const roundName = params.round || getRoundSheetName(activeMap);
  const sheet = getOrCreateRoundSheet(roundName);
  const rows = sheet.getDataRange().getValues();

  const teams = [];
  for (let i = 1; i < rows.length; i++) {
    teams.push({
      timestamp: rows[i][0],
      team: rows[i][1],
      captainName: rows[i][2],
      captainEmail: rows[i][3],
      whatsapp: rows[i][4],
      players: [
        { name: rows[i][5], uid: rows[i][6] },
        { name: rows[i][7], uid: rows[i][8] },
        { name: rows[i][9], uid: rows[i][10] },
        { name: rows[i][11], uid: rows[i][12] }
      ],
      txnRef: rows[i][13],
      screenshot: rows[i][14],
      status: rows[i][15]
    });
  }

  const maxKey = MAP_MAX_KEYS[activeMap] || "erangelMax";
  return jsonResponse({
    status: "success",
    activeMap: activeMap,
    roundName: roundName,
    mapMax: Number(settings[maxKey] || 100),
    erangelMax: Number(settings.erangelMax || 100),
    livikMax: Number(settings.livikMax || 50),
    miramarMax: Number(settings.miramarMax || 100),
    roomId: settings.roomId || "",
    roomTime: settings.roomTime || "",
    teams: teams
  });
}

function handleAdminAction(data) {
  if (!checkAdminPassword(data.password)) {
    return jsonResponse({ status: "unauthorized" });
  }

  if (data.action === "setActiveMap") {
    const map = String(data.map || "").toLowerCase();
    if (VALID_MAPS.indexOf(map) === -1) {
      return jsonResponse({ status: "error", message: "Invalid map" });
    }
    setSetting("activeMap", map);
    getOrCreateRoundSheet(getRoundSheetName(map)); // ensure the round sheet exists right away
    return jsonResponse({ status: "success" });
  }

  if (data.action === "setRoom") {
    setSetting("roomId", data.roomId || "");
    setSetting("roomTime", data.roomTime || "");
    let sentCount = 0;
    if (data.broadcast) sentCount = broadcastRoomIdInternal();
    return jsonResponse({ status: "success", sentCount: sentCount });
  }

  if (data.action === "approveTeam") {
    const settings = getAllSettings();
    const activeMap = String(settings.activeMap || "erangel").toLowerCase();
    const roundName = data.round || getRoundSheetName(activeMap);
    const sheet = getOrCreateRoundSheet(roundName);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][COL_TXN_REF - 1]) === String(data.txnRef)) {
        approveRow(sheet, i + 1, rows[i]);
        return jsonResponse({ status: "success" });
      }
    }
    return jsonResponse({ status: "error", message: "Team not found" });
  }

  return jsonResponse({ status: "error", message: "Unknown action" });
}
