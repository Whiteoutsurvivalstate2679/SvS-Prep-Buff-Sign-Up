// ============================================================
// SvS Prep Buff Sign-Up — Google Apps Script Backend
// State 2679 — Created by Goldee
//
// SETUP INSTRUCTIONS:
// 1. Create a Google Sheet with 4 tabs named exactly:
//    BOOKINGS | HISTORY | SETTINGS | CURTAIN
// 2. Paste this entire script in Apps Script (Extensions > Apps Script)
// 3. Deploy as Web App:
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the generated URL and paste it in the HTML file
//    where it says: const API_URL = "YOUR_APPS_SCRIPT_URL_HERE";
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

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
    if (action === 'getBookings') return handleGetBookings(e.parameter.day);
    if (action === 'getHistory') return handleGetHistory();
    if (action === 'getSetting') return handleGetSetting(e.parameter.key);
    if (action === 'getWeekBookings') return handleGetWeekBookings();
    return buildResponse({ error: 'Unknown GET action' });
  } catch (err) {
    return buildResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'book') return handleBook(body);
    if (action === 'cancel') return handleCancel(body);
    if (action === 'setSetting') return handleSetSetting(body);
    if (action === 'addBookingAdmin') return handleAddBookingAdmin(body);
    if (action === 'deleteBooking') return handleDeleteBooking(body);
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
    bookings[r.slot] = { pseudo: r.pseudo, playerId: r.playerId, alliance: r.alliance, sessionId: r.sessionId, ts: r.timestamp };
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
    result[r.day][r.slot] = { pseudo: r.pseudo, playerId: r.playerId, alliance: r.alliance, sessionId: r.sessionId, ts: r.timestamp };
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

function handleGetSetting(key) {
  const sheet = getSheet('SETTINGS');
  ensureHeaders(sheet, ['key', 'value', 'updatedAt']);
  const rows = sheetToObjects(sheet);
  const row = rows.find(r => r.key === key);
  return buildResponse({ success: true, value: row ? row.value : null });
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

  // Check slot not already taken
  const existing = rows.find(r => r.day === day && r.slot === slot);
  if (existing) {
    logHistory('confirm', 'failed_taken', pseudo, playerId, alliance, day, slot);
    return buildResponse({ success: false, error: 'slot_taken', takenBy: existing.pseudo });
  }

  // Anti multi-account: check same playerId already has a slot this day
  const playerAlreadyBooked = rows.find(r => r.day === day && String(r.playerId) === String(playerId));
  if (playerAlreadyBooked) {
    logHistory('confirm', 'failed_already_booked', pseudo, playerId, alliance, day, slot);
    return buildResponse({ success: false, error: 'already_booked', existingSlot: playerAlreadyBooked.slot });
  }

  sheet.appendRow([day, slot, pseudo, playerId, alliance, sessionId, Date.now()]);
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
  const dayIdx = headers.indexOf('day');
  const slotIdx = headers.indexOf('slot');
  const playerIdIdx = headers.indexOf('playerId');
  const pseudoIdx = headers.indexOf('pseudo');
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
  if (existing) {
    return buildResponse({ success: false, error: 'slot_taken' });
  }

  sheet.appendRow([day, slot, pseudo, playerId, alliance, 'admin', Date.now()]);
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
  const dayIdx = headers.indexOf('day');
  const slotIdx = headers.indexOf('slot');
  const pseudoIdx = headers.indexOf('pseudo');
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
      sheet.getRange(i + 1, headers.indexOf('value') + 1).setValue(value);
      sheet.getRange(i + 1, headers.indexOf('updatedAt') + 1).setValue(new Date().toISOString());
      return buildResponse({ success: true });
    }
  }

  sheet.appendRow([key, value, new Date().toISOString()]);
  return buildResponse({ success: true });
}

// ---- History logger ---------------------------------------

function logHistory(type, result, pseudo, playerId, alliance, day, slot) {
  try {
    const sheet = getSheet('HISTORY');
    ensureHeaders(sheet, ['type', 'result', 'pseudo', 'playerId', 'alliance', 'day', 'slot', 'timestamp']);
    sheet.appendRow([type, result, pseudo, playerId, alliance, day, slot, Date.now()]);
  } catch (e) {
    // History logging must never block the main flow
  }
}
