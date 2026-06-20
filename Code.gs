// ============================================================
// SvS Prep Buff Sign-Up — Google Apps Script Backend
// State 2679 — Created by Goldee
//
// ARCHITECTURE : Git repo → Apps Script deployed → Google Sheet
//
// SHEET TABS REQUIRED:
//   BOOKINGS  | HISTORY | SETTINGS | CURTAIN | ALLIANCES
//
// ALLIANCES tab layout (columns repeat x4 alliances):
//   A: Alliance A  B: Pseudo  C: ID  (blank col D)
//   E: Alliance B  F: Pseudo  G: ID  (blank col H)
//   I: Alliance C  J: Pseudo  K: ID  (blank col L)
//   M: Alliance D  N: Pseudo  O: ID
//
// DEPLOY:
//   Extensions > Apps Script > Deploy > New deployment
//   Execute as: Me  |  Who has access: Anyone
// ============================================================

const SS = SpreadsheetApp.openById('1NA54cMynvbVCf_7I_CEhN6zsAW2Gx4zte329hJl47Q4');

// ---- Sheet helpers ----------------------------------------

function getSheet(name) {
  return SS.getSheetByName(name);
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0a1628')
      .setFontColor('#e8b86d');
  }
}

function sheetToObjects(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ---- CORS response ----------------------------------------

function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Main entry points ------------------------------------

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getBookings')     return handleGetBookings(e.parameter.day);
    if (action === 'getHistory')      return handleGetHistory();
    if (action === 'getSetting')      return handleGetSetting(e.parameter.key);
    if (action === 'getWeekBookings') return handleGetWeekBookings();
    if (action === 'getPlayers')      return handleGetPlayers();
    return buildResponse({ error: 'Unknown GET action' });
  } catch (err) {
    return buildResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'book')             return handleBook(body);
    if (action === 'cancel')           return handleCancel(body);
    if (action === 'setSetting')       return handleSetSetting(body);
    if (action === 'addBookingAdmin')  return handleAddBookingAdmin(body);
    if (action === 'deleteBooking')    return handleDeleteBooking(body);
    return buildResponse({ error: 'Unknown POST action' });
  } catch (err) {
    return buildResponse({ error: err.message });
  }
}

// ---- GET handlers -----------------------------------------

function handleGetBookings(day) {
  const sheet = getSheet('BOOKINGS');
  ensureHeaders(sheet, ['day', 'slot', 'pseudo', 'playerId', 'alliance', 'sessionId', 'timestamp']);
  const rows = sheetToObjects(sheet).filter(r => r.day === day);
  const bookings = {};
  rows.forEach(r => {
    bookings[r.slot] = {
      pseudo: r.pseudo,
      playerId: String(r.playerId),
      alliance: r.alliance,
      sessionId: r.sessionId,
      ts: r.timestamp
    };
  });
  return buildResponse({ success: true, bookings });
}

function handleGetWeekBookings() {
  const sheet = getSheet('BOOKINGS');
  ensureHeaders(sheet, ['day', 'slot', 'pseudo', 'playerId', 'alliance', 'sessionId', 'timestamp']);
  const rows = sheetToObjects(sheet);
  const result = {};
  rows.forEach(r => {
    if (!result[r.day]) result[r.day] = {};
    result[r.day][r.slot] = {
      pseudo: r.pseudo,
      playerId: String(r.playerId),
      alliance: r.alliance,
      sessionId: r.sessionId,
      ts: r.timestamp
    };
  });
  return buildResponse({ success: true, bookings: result });
}

function handleGetHistory() {
  const sheet = getSheet('HISTORY');
  ensureHeaders(sheet, ['type', 'result', 'pseudo', 'playerId', 'alliance', 'day', 'slot', 'timestamp']);
  const TWO_WEEKS_MS = 14 * 24 * 60 * 60 * 1000;
  const cutoff = Date.now() - TWO_WEEKS_MS;
  const rows = sheetToObjects(sheet)
    .filter(r => r.timestamp > cutoff)
    .sort((a, b) => b.timestamp - a.timestamp);
  return buildResponse({ success: true, history: rows });
}

// ---- FIX: toujours retourner une string, jamais un booléen natif ----
function handleGetSetting(key) {
  const sheet = getSheet('SETTINGS');
  ensureHeaders(sheet, ['key', 'value', 'updatedAt']);
  const rows = sheetToObjects(sheet);
  const row = rows.find(r => r.key === key);
  return buildResponse({
    success: true,
    value: row != null ? String(row.value) : null
  });
}

// ---- Lit l'onglet ALLIANCES (4 blocs côte à côte) ----------
// Structure : A=alliance, B=pseudo, C=id, (D vide),
//             E=alliance, F=pseudo, G=id, (H vide), etc.
function handleGetPlayers() {
  const sheet = getSheet('ALLIANCES');
  if (!sheet) return buildResponse({ success: false, error: 'ALLIANCES sheet not found' });

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return buildResponse({ success: true, players: {} });

  // Colonnes des 4 blocs (index 0-based) : [allianceCol, pseudoCol, idCol]
  const BLOCKS = [
    [0, 1, 2],   // A, B, C
    [4, 5, 6],   // E, F, G
    [8, 9, 10],  // I, J, K
    [12, 13, 14] // M, N, O
  ];

  const players = {}; // { "CAT": [["Neo", "418346341"], ...], ... }

  // Ligne 1 = headers, on commence à la ligne 2 (index 1)
  for (let row = 1; row < data.length; row++) {
    for (const [aCol, pCol, iCol] of BLOCKS) {
      const alliance = String(data[row][aCol] || '').trim();
      const pseudo   = String(data[row][pCol] || '').trim();
      const id       = String(data[row][iCol] || '').trim();

      if (!alliance || !pseudo) continue;

      if (!players[alliance]) players[alliance] = [];
      players[alliance].push([pseudo, id]);
    }
  }

  return buildResponse({ success: true, players });
}

// ---- POST handlers ----------------------------------------

function handleBook(body) {
  const { day, slot, pseudo, playerId, alliance, sessionId } = body;
  if (!day || !slot || !pseudo || !playerId || !alliance || !sessionId) {
    return buildResponse({ success: false, error: 'Missing required fields' });
  }

  const sheet = getSheet('BOOKINGS');
  ensureHeaders(sheet, ['day', 'slot', 'pseudo', 'playerId', 'alliance', 'sessionId', 'timestamp']);
  const rows = sheetToObjects(sheet);

  // Slot déjà pris ?
  const existing = rows.find(r => r.day === day && r.slot === slot);
  if (existing) {
    logHistory('confirm', 'failed_taken', pseudo, playerId, alliance, day, slot);
    return buildResponse({ success: false, error: 'slot_taken', takenBy: existing.pseudo });
  }

  // Anti multi-compte : même playerId déjà bookée ce jour ?
  const playerAlreadyBooked = rows.find(r => r.day === day && String(r.playerId) === String(playerId));
  if (playerAlreadyBooked) {
    logHistory('confirm', 'failed_already_booked', pseudo, playerId, alliance, day, slot);
    return buildResponse({ success: false, error: 'already_booked', existingSlot: playerAlreadyBooked.slot });
  }

  sheet.appendRow([day, slot, pseudo, String(playerId), alliance, sessionId, Date.now()]);
  logHistory('confirm', 'success', pseudo, playerId, alliance, day, slot);
  return buildResponse({ success: true });
}

function handleCancel(body) {
  const { day, slot, playerId } = body;
  if (!day || !slot || !playerId) {
    return buildResponse({ success: false, error: 'Missing required fields' });
  }

  const sheet = getSheet('BOOKINGS');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dayIdx      = headers.indexOf('day');
  const slotIdx     = headers.indexOf('slot');
  const playerIdIdx = headers.indexOf('playerId');
  const pseudoIdx   = headers.indexOf('pseudo');
  const allianceIdx = headers.indexOf('alliance');

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (row[dayIdx] === day && row[slotIdx] === slot && String(row[playerIdIdx]) === String(playerId)) {
      logHistory('cancel', 'success', row[pseudoIdx], playerId, row[allianceIdx], day, slot);
      sheet.deleteRow(i + 1);
      return buildResponse({ success: true });
    }
  }

  return buildResponse({ success: false, error: 'booking_not_found' });
}

function handleAddBookingAdmin(body) {
  const { day, slot, pseudo, playerId, alliance } = body;
  if (!day || !slot || !pseudo || !playerId || !alliance) {
    return buildResponse({ success: false, error: 'Missing required fields' });
  }

  const sheet = getSheet('BOOKINGS');
  ensureHeaders(sheet, ['day', 'slot', 'pseudo', 'playerId', 'alliance', 'sessionId', 'timestamp']);
  const rows = sheetToObjects(sheet);

  const existing = rows.find(r => r.day === day && r.slot === slot);
  if (existing) return buildResponse({ success: false, error: 'slot_taken' });

  sheet.appendRow([day, slot, pseudo, String(playerId), alliance, 'admin', Date.now()]);
  return buildResponse({ success: true });
}

function handleDeleteBooking(body) {
  const { day, slot } = body;
  if (!day || !slot) {
    return buildResponse({ success: false, error: 'Missing required fields' });
  }

  const sheet = getSheet('BOOKINGS');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dayIdx      = headers.indexOf('day');
  const slotIdx     = headers.indexOf('slot');
  const pseudoIdx   = headers.indexOf('pseudo');
  const playerIdIdx = headers.indexOf('playerId');
  const allianceIdx = headers.indexOf('alliance');

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (row[dayIdx] === day && row[slotIdx] === slot) {
      logHistory('cancel', 'admin_delete', row[pseudoIdx], row[playerIdIdx], row[allianceIdx], day, slot);
      sheet.deleteRow(i + 1);
      return buildResponse({ success: true });
    }
  }

  return buildResponse({ success: false, error: 'not_found' });
}

// ---- FIX: forcer String() pour éviter les booléens natifs de Sheets ----
function handleSetSetting(body) {
  const { key, value } = body;
  if (!key) return buildResponse({ success: false, error: 'Missing key' });

  const sheet = getSheet('SETTINGS');
  ensureHeaders(sheet, ['key', 'value', 'updatedAt']);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx = headers.indexOf('key');

  for (let i = 1; i < data.length; i++) {
    if (data[i][keyIdx] === key) {
      sheet.getRange(i + 1, headers.indexOf('value') + 1).setValue(String(value));
      sheet.getRange(i + 1, headers.indexOf('updatedAt') + 1).setValue(new Date().toISOString());
      return buildResponse({ success: true });
    }
  }

  sheet.appendRow([key, String(value), new Date().toISOString()]);
  return buildResponse({ success: true });
}

// ---- History logger (ne bloque jamais le flux principal) --

function logHistory(type, result, pseudo, playerId, alliance, day, slot) {
  try {
    const sheet = getSheet('HISTORY');
    ensureHeaders(sheet, ['type', 'result', 'pseudo', 'playerId', 'alliance', 'day', 'slot', 'timestamp']);
    sheet.appendRow([type, result, pseudo, String(playerId), alliance, day, slot, Date.now()]);
  } catch (e) {
    // silencieux
  }
}
