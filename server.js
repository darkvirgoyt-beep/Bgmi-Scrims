import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { google } from 'googleapis';
import nodemailer from 'nodemailer';

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ============ CONFIG ============
const SHEET_ID = process.env.GOOGLE_SHEET_ID || '11joNjVKYi4CrnNOBHYMX_aIer4eiKnzpg0VaWfm2WdQ';
const RECIPIENTS = process.env.RECIPIENTS || 'imamericanyess@gmail.com,kumarabhinav7349@gmail.com';
const GOOGLE_SERVICE_ACCOUNT = process.env.GOOGLE_SERVICE_ACCOUNT ? JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT) : null;
const ADMIN_KEY = process.env.ADMIN_KEY || 'Princeraj@1331';
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD || 'bqce odep pzlh ghbs';

const MAP_CAPACITIES = {
  livik: 50,
  erangel: 100,
  miramar: 100
};

// ============ GOOGLE SHEETS AUTH ============
const sheets = google.sheets('v4');
let authClient = null;

async function getAuthClient() {
  if (!authClient && GOOGLE_SERVICE_ACCOUNT) {
    authClient = new google.auth.GoogleAuth({
      credentials: GOOGLE_SERVICE_ACCOUNT,
      scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    });
  }
  return authClient;
}

// ============ EMAIL SETUP ============
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASSWORD
  }
});

// ============ HELPER FUNCTIONS ============
async function getSettingsSheet() {
  try {
    const auth = await getAuthClient();
    const response = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId: SHEET_ID,
      range: 'Settings!A:B'
    });
    
    const rows = response.data.values || [];
    const settings = {};
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
    }
    return settings;
  } catch (err) {
    console.error('Error reading settings:', err);
    return {};
  }
}

async function getOrCreateMapSheet(map, date) {
  try {
    const auth = await getAuthClient();
    const sheetName = `${map.toUpperCase()}_${date}`;
    
    // Check if sheet exists
    const spreadsheet = await sheets.spreadsheets.get({
      auth,
      spreadsheetId: SHEET_ID
    });
    
    let sheetId = null;
    for (const sheet of spreadsheet.data.sheets) {
      if (sheet.properties.title === sheetName) {
        sheetId = sheet.properties.sheetId;
        break;
      }
    }
    
    // Create sheet if it doesn't exist
    if (sheetId === null) {
      const addSheetResponse = await sheets.spreadsheets.batchUpdate({
        auth,
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{
            addSheet: {
              properties: { title: sheetName }
            }
          }]
        }
      });
      sheetId = addSheetResponse.data.replies[0].addSheet.properties.sheetId;
      
      // Add headers
      await sheets.spreadsheets.values.update({
        auth,
        spreadsheetId: SHEET_ID,
        range: `${sheetName}!A1:Q1`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [[
            'Timestamp', 'Team', 'Captain Name', 'Captain Email', 'WhatsApp',
            'P1 Name', 'P1 UID', 'P2 Name', 'P2 UID', 'P3 Name', 'P3 UID',
            'P4 Name', 'P4 UID', 'Txn Ref', 'Screenshot', 'Status', 'Map'
          ]]
        }
      });
    }
    
    return sheetName;
  } catch (err) {
    console.error('Error with map sheet:', err);
    throw err;
  }
}

async function countTeamsForMap(map, date) {
  try {
    const auth = await getAuthClient();
    const sheetName = `${map.toUpperCase()}_${date}`;
    
    const response = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!A:A`
    });
    
    return (response.data.values?.length || 1) - 1; // Subtract header row
  } catch (err) {
    console.error('Error counting teams:', err);
    return 0;
  }
}

async function logToSheet(data, map, date) {
  try {
    const auth = await getAuthClient();
    const sheetName = await getOrCreateMapSheet(map, date);
    
    await sheets.spreadsheets.values.append({
      auth,
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!A:Q`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[
          new Date().toLocaleString(),
          data.teamName,
          data.googleName,
          data.googleEmail,
          data.whatsapp,
          data.p1name, data.p1uid,
          data.p2name, data.p2uid,
          data.p3name, data.p3uid,
          data.p4name, data.p4uid,
          data.txnRef || '',
          data.screenshotUrl || 'Not provided',
          'Pending',
          map
        ]]
      }
    });
  } catch (err) {
    console.error('Error logging to sheet:', err);
    throw err;
  }
}

// ============ ROUTES ============

// Get live settings + capacity
app.get('/api/settings', async (req, res) => {
  try {
    const settings = await getSettingsSheet();
    const activeMap = (settings.activeMap || 'erangel').toLowerCase();
    const today = new Date().toISOString().split('T')[0];
    
    const teamCount = await countTeamsForMap(activeMap, today);
    const maxPlayers = MAP_CAPACITIES[activeMap] || 100;
    const playersRegistered = teamCount * 4;
    const isFull = playersRegistered >= maxPlayers;
    
    res.json({
      status: 'success',
      settings: {
        activeMap,
        mapMax: maxPlayers,
        playersRegistered,
        mapFull: isFull,
        mode: settings.mode || 'Erangel · TPP Squad',
        schedule: settings.schedule || 'Sat & Sun · 8 PM IST',
        winPrize: settings.winPrize || '₹20 / kill',
        mvpPrize: settings.mvpPrize || '₹5 / kill',
        liveStream: settings.liveStream || 'https://www.youtube.com/@VirgoYT707',
        whatsappGroup: settings.whatsappGroup || 'https://chat.whatsapp.com/E8lXkGkA4ShE9yRrnOWhcC',
        entryFee: settings.entryFee || '20',
        upiId: settings.upiId || 'princeraj21@fam',
        upiPayeeName: settings.upiPayeeName || 'PRINCE'
      }
    });
  } catch (err) {
    res.status(500).json({ status: 'error', message: err.message });
  }
});

// Update active map (admin only)
app.post('/api/admin/set-map', async (req, res) => {
  try {
    const { map, adminKey } = req.body;
    
    // Simple admin key check
    if (adminKey !== ADMIN_KEY) {
      return res.status(403).json({ status: 'error', message: 'Unauthorized' });
    }
    
    if (!MAP_CAPACITIES[map.toLowerCase()]) {
      return res.status(400).json({ status: 'error', message: 'Invalid map' });
    }
    
    const auth = await getAuthClient();
    await sheets.spreadsheets.values.update({
      auth,
      spreadsheetId: SHEET_ID,
      range: 'Settings!A:B',
      valueInputOption: 'RAW',
      requestBody: {
        values: [
          ['activeMap', map.toLowerCase()]
        ]
      }
    });
    
    res.json({ status: 'success', activeMap: map.toLowerCase() });
  } catch (err) {
    res.status(500).json({ status: 'error', message: err.message });
  }
});

// Register team
app.post('/api/register', async (req, res) => {
  try {
    const data = req.body;
    const settings = await getSettingsSheet();
    const activeMap = (settings.activeMap || 'erangel').toLowerCase();
    const today = new Date().toISOString().split('T')[0];
    
    // Check capacity
    const teamCount = await countTeamsForMap(activeMap, today);
    const maxPlayers = MAP_CAPACITIES[activeMap] || 100;
    const playersRegistered = teamCount * 4;
    
    if (playersRegistered >= maxPlayers) {
      return res.json({ status: 'full', message: 'This map lobby is full' });
    }
    
    // Log to sheet
    await logToSheet(data, activeMap, today);
    
    // Send emails
    const emailBody = `
TEAM: ${data.teamName}
CAPTAIN: ${data.googleName} <${data.googleEmail}>
WHATSAPP: ${data.whatsapp}
MAP: ${activeMap.toUpperCase()}
DATE: ${today}

PLAYERS:
Player 1: ${data.p1name} | UID: ${data.p1uid}
Player 2: ${data.p2name} | UID: ${data.p2uid}
Player 3: ${data.p3name} | UID: ${data.p3uid}
Player 4: ${data.p4name} | UID: ${data.p4uid}

TXN REF: ${data.txnRef || 'Not provided'}
Payment Screenshot: ${data.screenshotUrl || 'Not provided'}
Submitted: ${new Date().toLocaleString()}

Status: Pending (verify payment before confirming)
    `;
    
    const emailPromises = RECIPIENTS.split(',').map(email =>
      transporter.sendMail({
        from: process.env.EMAIL_USER,
        to: email.trim(),
        subject: `New Scrim Registration — ${data.teamName}`,
        text: emailBody
      })
    );
    
    await Promise.all(emailPromises);
    
    res.json({ status: 'success', message: 'Registration submitted' });
  } catch (err) {
    console.error('Registration error:', err);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// ============ START SERVER ============
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`🎮 BGMI Scrims Backend running on port ${PORT}`);
});
