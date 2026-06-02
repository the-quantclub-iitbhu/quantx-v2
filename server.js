require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());

// ── Sheet IDs ──
const SHEET_ID          = process.env.SHEET_ID;           // QuantX game sheet (has Rounds tab + Sheet1)
const IIT_SHEET_ID      = process.env.IIT_SHEET_ID;
const EXTERNAL_SHEET_ID = process.env.EXTERNAL_SHEET_ID;

if (!SHEET_ID || !IIT_SHEET_ID || !EXTERNAL_SHEET_ID) {
  console.error('Missing required env vars: SHEET_ID, IIT_SHEET_ID, EXTERNAL_SHEET_ID');
  process.exit(1);
}

// ── Google Sheets Auth ──
function getAuth() {
  const creds = process.env.GOOGLE_CREDENTIALS;
  if (creds) {
    const parsed = JSON.parse(creds);
    parsed.private_key = parsed.private_key.replace(/\\n/g, '\n');
    return new google.auth.GoogleAuth({
      credentials: parsed,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
  }
  const keyFile = require('./credentials.json');
  keyFile.private_key = keyFile.private_key.replace(/\\n/g, '\n');
  return new google.auth.GoogleAuth({
    credentials: keyFile,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheets() {
  const auth = getAuth();
  return google.sheets({ version: 'v4', auth: await auth.getClient() });
}

// ── Registration cache (refreshed every 5 mins) ──
let registrationCache = [];
let cacheLastFetched = 0;

async function fetchRegistrations() {
  const now = Date.now();
  if (now - cacheLastFetched < 5 * 60 * 1000 && registrationCache.length > 0) {
    return registrationCache; // return cache if fresh
  }

  try {
    const sheets = await getSheets();

    // Fetch IIT sheet
    const iitRes = await sheets.spreadsheets.values.get({
      spreadsheetId: IIT_SHEET_ID,
      range: 'Sheet1!A:Z'
    });
    const iitRows = iitRes.data.values || [];
    const iitHeaders = iitRows[0] || [];
    const iitData = iitRows.slice(1).map(row => {
      const obj = {};
      iitHeaders.forEach((h, i) => { obj[h.trim()] = row[i] ? String(row[i]).trim() : ''; });
      obj._source = 'iit';
      return obj;
    });

    // Fetch External sheet
    const extRes = await sheets.spreadsheets.values.get({
      spreadsheetId: EXTERNAL_SHEET_ID,
      range: 'Sheet1!A:Z'
    });
    const extRows = extRes.data.values || [];
    const extHeaders = extRows[0] || [];
    const extData = extRows.slice(1).map(row => {
      const obj = {};
      extHeaders.forEach((h, i) => { obj[h.trim()] = row[i] ? String(row[i]).trim() : ''; });
      obj._source = 'external';
      return obj;
    });

    registrationCache = [...iitData, ...extData];
    cacheLastFetched = now;
    console.log(`Registration cache refreshed: ${iitData.length} IIT + ${extData.length} external`);
    return registrationCache;
  } catch (e) {
    console.error('Failed to fetch registrations:', e.message);
    return registrationCache; // return stale cache on error
  }
}

// Helper — find participant in registration sheet by email
function findRegistration(rows, email) {
  return rows.find(r => {
    const rowEmail = (r['Email'] || r['E-mail'] || r['email'] || '').toLowerCase().trim();
    return rowEmail === email.toLowerCase().trim();
  });
}

// Helper — extract phone number from row (handles different column names)
function getPhone(row) {
  return (row['Phone Number'] || row['Phone No.'] || row['Phone No'] || row['Phone_Number'] || row['phone'] || '').replace(/\s+/g, '').trim();
}

// Helper — extract name from row
function getName(row) {
  return (row['Name'] || row['Full Name'] || row['FullName'] || '').trim();
}

// ── Rounds sheet helper ──
async function fetchRoundData(roundNumber) {
  const sheets = await getSheets();
  // Row 1 = headers, Round 1 = Row 2, Round N = Row N+1
  const rowIndex = parseInt(roundNumber) + 1;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `Rounds!A${rowIndex}:G${rowIndex}`
  });
  const row = (res.data.values || [[]])[0];
  if (!row || row.length < 7) throw new Error(`Round ${roundNumber} data not found in Rounds sheet`);
  return {
    roundNo:           parseInt(row[0]),
    stockName:         row[1],
    ticker:            row[2],
    currentPrice:      parseFloat(row[3]),
    futurePrice:       parseFloat(row[4]),
    imageUrl:          row[5],
    correctDirection:  row[6].trim().toUpperCase()
  };
}

// ── In-memory store ──
let participants = {};

// Game state
let gameState = {
  currentRound:     0,
  timerActive:      false,
  currentImage:     '',
  currentPrice:     0,
  futurePrice:      0,
  stockName:        '',
  ticker:           '',
  correctDirection: '',
  gameActive:       false
};

// ── Sheets helpers (sync game results) ──
async function syncParticipantToSheet(p) {
  try {
    const sheets = await getSheets();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A:A'
    });
    const ids = (res.data.values || []).map(r => r[0]);
    const rowIndex = ids.indexOf(p.id);
    const row = [
 p.id,
 p.name,
 p.email || '',
 10000 + p.totalPnl,
 p.position,
 p.units,
 p.entryPrice,
 p.totalPnl,
 p.round,
 new Date().toISOString()
];

    if (rowIndex === -1) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: 'Sheet1!A:J',
        valueInputOption: 'RAW',
        resource: { values: [row] }
      });
    } else {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Sheet1!A${rowIndex + 1}:J${rowIndex + 1}`,
        valueInputOption: 'RAW',
        resource: { values: [row] }
      });
    }
  } catch (e) {
    console.error('Sheet sync error:', e.message);
  }
}

async function syncAllToSheet() {
  try {
    const sheets = await getSheets();
    const allRows = Object.values(participants).map(p => [
  p.id,                               // participant_id
  p.name,                             // name
  p.email || '',                      // mail
  10000 + p.totalPnl,                 // cash_balance
  p.position || 'FLAT',               // current_position
  p.units || 0,                       // units_held
  p.entryPrice || 0,                  // entry_price
  p.totalPnl || 0,                    // total_pnl
  p.round || 0,                       // round_number
  new Date().toISOString()            // last_action
]);

if (!allRows.length) return;

await sheets.spreadsheets.values.clear({
  spreadsheetId: SHEET_ID,
  range: 'Sheet1!A2:J1000'
});
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!A2',
      valueInputOption: 'RAW',
      resource: { values: allRows }
    });
  } catch (e) {
    console.error('Full sync error:', e.message);
  }
}

// ── ADMIN: login ──
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  res.json({ success: password === process.env.ADMIN_PASSWORD });
});

// ── Page routes ──
app.get('/', (req, res) => {
  res.redirect('/participant');
});

app.get('/participant', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'participant.html'));
});

app.get('/participant.html', (req, res) => {
  res.redirect('/participant');
});

app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

app.get('/admin.html', (req, res) => {
  res.redirect('/admin');
});

app.use(express.static(path.join(__dirname, 'public')));

// ── ADMIN: start round (only needs roundNumber now) ──
app.post('/api/admin/start-round', async (req, res) => {
  const { roundNumber } = req.body;
  if (!roundNumber) return res.status(400).json({ success: false, message: 'roundNumber required' });

  try {
    const round = await fetchRoundData(roundNumber);

    gameState.currentRound     = round.roundNo;
    gameState.currentImage     = round.imageUrl;
    gameState.correctDirection = round.correctDirection;
    gameState.currentPrice     = round.currentPrice;
    gameState.futurePrice      = round.futurePrice;
    gameState.stockName        = round.stockName;
    gameState.ticker           = round.ticker;
    gameState.timerActive      = true;
    gameState.gameActive       = true;

    setTimeout(async () => {
  if (!gameState.timerActive) return;

  gameState.timerActive = false;

  const futurePrice = gameState.futurePrice;
  const currentPrice = gameState.currentPrice;
  const difference = futurePrice - currentPrice;

  Object.values(participants).forEach(p => {
    if (p.units === 0 || p.position === 'FLAT') return;

    let roundPnl = 0;

    if (p.position === 'BUY')
      roundPnl = p.units * difference;

    if (p.position === 'SELL')
      roundPnl = p.units * (-difference);

    p.totalPnl += roundPnl;
    p.position = 'FLAT';
    p.units = 0;
    p.entryPrice = 0;
  });

  await syncAllToSheet();

  console.log(`Round ${gameState.currentRound} auto-ended`);
}, 100000);

    res.json({ success: true, gameState });
  } catch (e) {
    console.error('start-round error:', e.message);
    res.status(500).json({ success: false, message: e.message });
  }
});

// ── ADMIN: end round (uses futurePrice already in gameState) ──
app.post('/api/admin/end-round', async (req, res) => {
  const futurePrice  = gameState.futurePrice;
  const currentPrice = gameState.currentPrice;
  gameState.timerActive = false;

  // PnL = units × (futurePrice - currentPrice) for BUY
  // PnL = units × (currentPrice - futurePrice) for SELL
  const difference = futurePrice - currentPrice;

  Object.values(participants).forEach(p => {
    if (p.units === 0 || p.position === 'FLAT') return;
    let roundPnl = 0;
    if (p.position === 'BUY')  roundPnl = p.units * difference;
    if (p.position === 'SELL') roundPnl = p.units * (-difference);
    p.totalPnl += roundPnl;
    p.position  = 'FLAT';
    p.units     = 0;
    p.entryPrice = 0;
  });

  await syncAllToSheet();

  res.json({
    success: true,
    currentPrice,
    futurePrice,
    difference,
    correctDirection: gameState.correctDirection
  });
});

// ── ADMIN: get all participants ──
app.get('/api/admin/participants', (req, res) => {
  res.json(Object.values(participants));
});

// ── ADMIN: force refresh registration cache ──
app.post('/api/admin/refresh-registrations', async (req, res) => {
  cacheLastFetched = 0; // force refresh
  const rows = await fetchRegistrations();
  res.json({ success: true, count: rows.length });
});

// ── PARTICIPANT: login ──
// Email + password (first 4 letters of name lowercase + last 4 digits of phone)
app.post('/api/participant/register', async (req, res) => {
  const { email, password } = req.body;

  if (!email || !email.includes('@')) return res.json({ success: false, message: 'Enter a valid email' });
  if (!password || password.trim() === '') return res.json({ success: false, message: 'Enter your password' });

  // If already in session — return existing session
  const existingKey = Object.keys(participants).find(k => k === email.toLowerCase().trim());
  if (existingKey) {
    const p = participants[existingKey];
    return res.json({
      success: true,
      participantId: p.id,
      name: p.name,
      totalPnl: p.totalPnl,
      rejoined: true
    });
  }

  // Fetch registrations (cached)
  const rows = await fetchRegistrations();
  const match = findRegistration(rows, email);

  if (!match) {
    return res.json({ success: false, message: 'Email not found. Are you registered for QuantX?' });
  }

  // Build expected password: first 4 letters of name (lowercase) + last 4 digits of phone
  const name  = getName(match);
  const phone = getPhone(match);

  if (!name || !phone) {
    return res.json({ success: false, message: 'Registration data incomplete. Contact the organizers.' });
  }

  const expectedPassword = name.slice(0, 4).toLowerCase() + phone.slice(-4);

  if (password.trim() !== expectedPassword) {
    return res.json({ success: false, message: 'Incorrect password. Hint: first 4 letters of your name + last 4 digits of your phone.' });
  }

  // Auth passed — create session
  const id = 'P' + Date.now();
  const p = {
    id,
    name,
    email:       email.trim().toLowerCase(),
    position:    'FLAT',
    units:       0,
    entryPrice:  0,
    totalPnl:    0,
    round:       0,
    registeredAt: new Date().toISOString()
  };

  participants[email.toLowerCase().trim()] = p;
  syncParticipantToSheet(p);

  res.json({ success: true, participantId: id, name: p.name, totalPnl: 0 });
});

// ── PARTICIPANT: game state ──
app.get('/api/game-state', (req, res) => {
  res.json(gameState);
});

// ── PARTICIPANT: submit order ──
app.post('/api/participant/order', async (req, res) => {
  const { participantId, direction, units } = req.body;
  const u = parseInt(units);

  if (!gameState.timerActive) return res.json({ success: false, message: 'Round closed' });
  if (u < 1 || u > 20)        return res.json({ success: false, message: 'Units must be 1–20' });

  const p = Object.values(participants).find(x => x.id === participantId);
  if (!p) return res.json({ success: false, message: 'Participant not found' });

  // Place order
  p.position   = direction;
  p.units      = u;
  p.entryPrice = gameState.currentPrice;
  p.round      = gameState.currentRound;

  res.json({ success: true });
});

// ── LEADERBOARD ──
app.get('/api/leaderboard', (req, res) => {
  const board = Object.values(participants)
    .map(p => ({ name: p.name, totalPnl: p.totalPnl }))
    .sort((a, b) => b.totalPnl - a.totalPnl);
  res.json(board);
});

// ── START ──
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  fetchRegistrations(); // warm up cache on start
});
