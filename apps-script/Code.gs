/**
 * ═══════════════════════════════════════════════════════════════
 *  INVENTORY TRACKER — Google Apps Script Backend
 *  Bound to a Google Sheet with tabs: users, transactions, inventory
 * ═══════════════════════════════════════════════════════════════
 *
 *  SETUP (see Setup.md for full guide):
 *  1. Create Google Sheet with 3 tabs (users, transactions, inventory).
 *  2. Extensions → Apps Script → paste this file.
 *  3. Run `setup()` once to store config + create weekly trigger.
 *  4. Deploy → New deployment → Web app → access: Anyone.
 *  5. Paste deployment URL into index.html  APPS_SCRIPT_URL.
 */

/* ── CONFIG (set these, or set via setup() / PropertiesService) ── */
var GEMINI_API_KEY   = '';  // from Google AI Studio
var GEMINI_MODEL     = 'gemini-2.0-flash';  // free-tier vision model
var TELEGRAM_BOT_TOKEN = '';  // from @BotFather
var TELEGRAM_CHAT_ID   = '';  // your chat/group ID
var LOW_STOCK_THRESHOLD = 20;
var DRIVE_FOLDER_NAME  = 'InventoryTrackerPhotos';

/* ── Sheet column indexes (0-based) ── */
// users:        id, name, pin, role, active
// transactions: ts, user, type, item, qty, notes, photo_url, gemini_raw
// inventory:    item, current_qty, last_updated

function doGet(e) {
  return jsonOut({ ok: true, message: 'Inventory Tracker backend is running.' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var result;

    switch (action) {
      case 'login':       result = handleLogin(body); break;
      case 'users':       result = handleUsers(body); break;
      case 'ocr':         result = handleOCR(body); break;
      case 'transaction': result = handleTransaction(body); break;
      case 'inventory':   result = handleInventory(body); break;
      case 'log':         result = handleLog(body); break;
      case 'report':      result = handleReport(body); break;
      default:            result = { ok: false, error: 'Unknown action: ' + action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

/* ════════════════════════════════════════════════════════════
   LOGIN
   ════════════════════════════════════════════════════════════ */

function handleLogin(body) {
  var sheet = getSheet('users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var pin = String(row[2]).trim();
    var active = row[4];
    if (pin === String(body.pin).trim() && active !== false && active !== 'FALSE') {
      return { ok: true, user: { id: row[0], name: row[1], role: row[3] } };
    }
  }
  return { ok: false, error: 'Invalid PIN' };
}

/* ════════════════════════════════════════════════════════════
   USERS CRUD
   ════════════════════════════════════════════════════════════ */

function handleUsers(body) {
  // Only admins can manage users — caller must pass their pin for verification
  if (!verifyAdmin(body.pin)) {
    return { ok: false, error: 'Admin access required' };
  }

  if (body.op === 'add') {
    if (!body.name || !/^\d{4}$/.test(String(body.pin_new || body.pin))) {
      return { ok: false, error: 'Invalid name or PIN (must be 4 digits)' };
    }
    var sheet = getSheet('users');
    var id = sheet.getLastRow(); // simple auto-increment
    var newPin = body.pin_new || body.pin;
    sheet.appendRow([id, body.name, newPin, body.role || 'worker', true]);
    return { ok: true };
  }

  if (body.op === 'delete') {
    var sheet = getSheet('users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(body.id)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'User not found' };
  }

  // Default: list users (without exposing pins to non-admins — but admin sees all)
  return { ok: true, users: getAllUsers() };
}

function getAllUsers() {
  var sheet = getSheet('users');
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({
      id: data[i][0], name: data[i][1], pin: data[i][2],
      role: data[i][3], active: data[i][4]
    });
  }
  return users;
}

function verifyAdmin(pin) {
  var users = getAllUsers();
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].pin) === String(pin) && users[i].role === 'admin') return true;
  }
  return false;
}

/* ════════════════════════════════════════════════════════════
   OCR — Gemini Vision API
   ════════════════════════════════════════════════════════════ */

function handleOCR(body) {
  var apiKey = getProp('GEMINI_API_KEY') || GEMINI_API_KEY;
  if (!apiKey) return { ok: false, error: 'Gemini API key not configured' };

  var model = GEMINI_MODEL;
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

  var prompt = 'You are an inventory assistant. Look at this photo of a stock item or box. ' +
    'Identify: 1) The item type/name (be concise, e.g. "A4 Paper Ream 80gsm"). ' +
    '2) The quantity if visible on the label/packaging (if a pack says "24 pcs", quantity is 24). ' +
    'If quantity is not visible, default to 1. ' +
    'Respond ONLY with valid JSON: {"item":"...","qty":number}';

  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'image/jpeg', data: body.photo } }
      ]
    }],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(resp.getContentText());

    if (json.error) {
      return { ok: false, error: 'Gemini: ' + json.error.message };
    }

    var text = json.candidates[0].content.parts[0].text;
    var parsed;
    try {
      // Gemini may return JSON with markdown fences
      text = text.replace(/```json\n?/g, '').replace(/```/g, '').trim();
      parsed = JSON.parse(text);
    } catch (e) {
      return { ok: false, error: 'Could not parse OCR result: ' + text };
    }

    return { ok: true, ocr: { item: parsed.item || '', qty: parseInt(parsed.qty) || 1 }, raw: text };
  } catch (err) {
    return { ok: false, error: 'Gemini request failed: ' + err.toString() };
  }
}

/* ════════════════════════════════════════════════════════════
   TRANSACTION — save + photo + Telegram
   ════════════════════════════════════════════════════════════ */

function handleTransaction(body) {
  var ts = new Date().toISOString();
  var photoUrl = '';

  // Upload photo to Drive if provided
  if (body.photo) {
    try {
      photoUrl = uploadPhotoToDrive(body.photo, body.item, body.user);
    } catch (e) {
      Logger.log('Photo upload failed: ' + e);
    }
  }

  // Save transaction row
  var sheet = getSheet('transactions');
  sheet.appendRow([
    ts, body.user, body.type, body.item, parseInt(body.qty),
    body.notes || '', photoUrl, ''
  ]);

  // Update inventory
  updateInventory(body.item, parseInt(body.qty), body.type, ts);

  // Send Telegram notification
  var telegramSent = false;
  try {
    sendTelegramTransaction(body.user, body.type, body.item, parseInt(body.qty), body.notes);
    telegramSent = true;
  } catch (e) {
    Logger.log('Telegram send failed: ' + e);
  }

  return {
    ok: true,
    transaction: { ts: ts, user: body.user, type: body.type, item: body.item, qty: parseInt(body.qty) },
    telegram_sent: telegramSent
  };
}

function updateInventory(item, qty, type, ts) {
  var sheet = getSheet('inventory');
  var data = sheet.getDataRange().getValues();
  var delta = (type === 'in') ? qty : -qty;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(item).trim()) {
      var newQty = Number(data[i][1]) + delta;
      sheet.getRange(i + 1, 2).setValue(newQty);
      sheet.getRange(i + 1, 3).setValue(ts);
      return;
    }
  }
  // New item
  sheet.appendRow([item, delta, ts]);
}

/* ════════════════════════════════════════════════════════════
   DRIVE PHOTO UPLOAD
   ════════════════════════════════════════════════════════════ */

function uploadPhotoToDrive(base64Data, itemName, user) {
  var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  var safeName = (itemName || 'item').replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
  var fileName = ts + '_' + user + '_' + safeName + '.jpg';

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function getOrCreateFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

/* ════════════════════════════════════════════════════════════
   TELEGRAM
   ════════════════════════════════════════════════════════════ */

function sendTelegramTransaction(user, type, item, qty, notes) {
  var token = getProp('TELEGRAM_BOT_TOKEN') || TELEGRAM_BOT_TOKEN;
  var chatId = getProp('TELEGRAM_CHAT_ID') || TELEGRAM_CHAT_ID;
  if (!token || !chatId) throw new Error('Telegram not configured');

  var arrow = type === 'in' ? '📥 STOCK IN' : '📤 STOCK OUT';
  var sign = type === 'in' ? '+' : '−';
  var time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  var msg = arrow + '\n' +
    '📦 ' + item + '\n' +
    '🔢 ' + sign + qty + '\n' +
    '👤 ' + user + '\n' +
    '🕐 ' + time;
  if (notes) msg += '\n💬 ' + notes;

  sendTelegram(token, chatId, msg);
}

function sendTelegram(token, chatId, text) {
  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'HTML' }),
    muteHttpExceptions: true
  });
}

/* ════════════════════════════════════════════════════════════
   INVENTORY & LOG QUERIES
   ════════════════════════════════════════════════════════════ */

function handleInventory() {
  var sheet = getSheet('inventory');
  var data = sheet.getDataRange().getValues();
  var inv = [];
  for (var i = 1; i < data.length; i++) {
    inv.push({ item: data[i][0], current_qty: Number(data[i][1]), last_updated: data[i][2] });
  }
  return { ok: true, inventory: inv };
}

function handleLog() {
  var sheet = getSheet('transactions');
  var data = sheet.getDataRange().getValues();
  var txs = [];
  for (var i = data.length - 1; i >= 1; i--) { // newest first
    txs.push({
      ts: data[i][0], user: data[i][1], type: data[i][2],
      item: data[i][3], qty: Number(data[i][4]), notes: data[i][5], photo_url: data[i][6]
    });
  }
  return { ok: true, transactions: txs };
}

/* ════════════════════════════════════════════════════════════
   WEEKLY REPORT
   ════════════════════════════════════════════════════════════ */

function handleReport() {
  var msg = generateWeeklyReport();
  var token = getProp('TELEGRAM_BOT_TOKEN') || TELEGRAM_BOT_TOKEN;
  var chatId = getProp('TELEGRAM_CHAT_ID') || TELEGRAM_CHAT_ID;
  if (!token || !chatId) return { ok: false, error: 'Telegram not configured' };

  sendTelegram(token, chatId, msg);
  return { ok: true, sent: true, message: 'Weekly report sent to Telegram' };
}

function generateWeeklyReport() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  var sheet = getSheet('transactions');
  var data = sheet.getDataRange().getValues();
  var weekTx = [];
  for (var i = 1; i < data.length; i++) {
    var ts = new Date(data[i][0]);
    if (ts >= weekAgo && ts <= now) {
      weekTx.push({
        user: data[i][1], type: data[i][2], item: data[i][3], qty: Number(data[i][4])
      });
    }
  }

  var inTx = weekTx.filter(function(t) { return t.type === 'in'; });
  var outTx = weekTx.filter(function(t) { return t.type === 'out'; });
  var inQty = inTx.reduce(function(s, t) { return s + t.qty; }, 0);
  var outQty = outTx.reduce(function(s, t) { return s + t.qty; }, 0);

  // By user
  var byUser = {};
  weekTx.forEach(function(t) { byUser[t.user] = (byUser[t.user] || 0) + 1; });
  var userLines = Object.keys(byUser).map(function(u) { return u + ' (' + byUser[u] + ')'; }).join(', ');

  // Current inventory levels
  var invSheet = getSheet('inventory');
  var invData = invSheet.getDataRange().getValues();
  var invLines = [];
  var lowStock = [];
  for (var j = 1; j < invData.length; j++) {
    var item = invData[j][0];
    var qty = Number(invData[j][1]);
    invLines.push('• ' + item + ': ' + qty);
    if (qty <= LOW_STOCK_THRESHOLD) lowStock.push('• ' + item + ' (' + qty + ' left)');
  }

  var periodStart = Utilities.formatDate(weekAgo, tz, 'd MMM');
  var periodEnd = Utilities.formatDate(now, tz, 'd MMM yyyy');

  var msg = '📊 WEEKLY INVENTORY REPORT\n';
  msg += 'Period: ' + periodStart + '–' + periodEnd + '\n\n';
  msg += 'STOCK MOVEMENT\n';
  msg += '📥 In:  ' + inQty + ' items (' + inTx.length + ' transactions)\n';
  msg += '📤 Out: ' + outQty + ' items (' + outTx.length + ' transactions)\n\n';
  msg += 'CURRENT LEVELS\n' + invLines.join('\n') + '\n';
  if (lowStock.length) {
    msg += '\n⚠️ LOW STOCK ALERTS\n' + lowStock.join('\n') + '\n';
  }
  msg += '\nTotal transactions: ' + weekTx.length + '\n';
  msg += 'By: ' + userLines;

  return msg;
}

/* ════════════════════════════════════════════════════════════
   WEEKLY TRIGGER (time-driven, every Monday 9am)
   ════════════════════════════════════════════════════════════ */

function sendWeeklyReport() {
  // Called by time-driven trigger
  var msg = generateWeeklyReport();
  var token = getProp('TELEGRAM_BOT_TOKEN') || TELEGRAM_BOT_TOKEN;
  var chatId = getProp('TELEGRAM_CHAT_ID') || TELEGRAM_CHAT_ID;
  if (token && chatId) {
    sendTelegram(token, chatId, msg);
  }
}

/* ════════════════════════════════════════════════════════════
   SETUP — run once
   ════════════════════════════════════════════════════════════ */

function setup() {
  // Store config in PropertiesService (more secure than code)
  // EDIT THESE VALUES, then run setup() once:
  var config = {
    GEMINI_API_KEY:     'PASTE_YOUR_GEMINI_KEY',
    TELEGRAM_BOT_TOKEN: 'PASTE_YOUR_BOT_TOKEN',
    TELEGRAM_CHAT_ID:   'PASTE_YOUR_CHAT_ID',
  };
  var props = PropertiesService.getScriptProperties();
  for (var key in config) {
    if (config[key] && config[key] !== 'PASTE_YOUR_' + key.split('_').slice(0,2).join('_')) {
      props.setProperty(key, config[key]);
    }
  }

  // Ensure sheet headers exist
  initSheetHeaders();

  // Seed default admin if no users
  var users = getAllUsers();
  if (users.length === 0) {
    getSheet('users').appendRow([1, 'Admin', '9999', 'admin', true]);
    getSheet('users').appendRow([2, 'Worker', '1234', 'worker', true]);
    Logger.log('Seeded default users: Admin/9999, Worker/1234');
  }

  // Create weekly trigger (Monday 9am) — remove old triggers first
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('✅ Setup complete. Config stored, trigger created, sheet initialized.');
  Logger.log('Next step: Deploy → New deployment → Web app → Access: Anyone');
}

/* ════════════════════════════════════════════════════════════
   HELPERS
   ════════════════════════════════════════════════════════════ */

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheetHeadersFor(sheet, name);
  }
  return sheet;
}

function initSheetHeaders() {
  initSheetHeadersFor(getSheet('users'), 'users');
  initSheetHeadersFor(getSheet('transactions'), 'transactions');
  initSheetHeadersFor(getSheet('inventory'), 'inventory');
}

function initSheetHeadersFor(sheet, name) {
  var headers = {
    users:        ['id', 'name', 'pin', 'role', 'active'],
    transactions: ['ts', 'user', 'type', 'item', 'qty', 'notes', 'photo_url', 'gemini_raw'],
    inventory:    ['item', 'current_qty', 'last_updated']
  };
  sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
  sheet.setFrozenRows(1);
}

function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
