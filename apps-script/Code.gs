/**
 * ═══════════════════════════════════════════════════════════════
 *  INVENTORY TRACKER — Google Apps Script Backend  (v1.8.1-codex)
 *  Tabs: users, transactions, catalogue, sites
 * ═══════════════════════════════════════════════════════════════
 */

/* ── CONFIG ── */
var GEMINI_API_KEY   = '';
var GEMINI_MODEL     = 'gemini-3.5-flash';
var TELEGRAM_BOT_TOKEN = '';
var TELEGRAM_CHAT_ID   = '';
var LOW_STOCK_THRESHOLD = 20;
var DRIVE_FOLDER_NAME  = 'InventoryTrackerPhotos';
var DAILY_REPORT_HOUR  = 18;  // 6 PM — manager can change via setSchedule()

/* ── Sheet column indexes (0-based) ── */
// users:        id, name, pin, role, active
// transactions: ts, user, type, item, qty, destination, photo_url, gemini_raw
// catalogue:    item_name, category, unit, min_qty, current_qty, last_updated, ref_photo_url
// sites:        site_name, active

function doGet(e) {
  return jsonOut({ ok: true, message: 'Inventory Tracker backend is running. v1.8.1-codex' });
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
      case 'catalogue':   result = handleCatalogue(body); break;
      case 'sites':       result = handleSites(body); break;
      case 'log':         result = handleLog(body); break;
      case 'report':      result = handleReport(body); break;
      case 'schedule':    result = handleSchedule(body); break;
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
  if (body.op === 'add' || body.op === 'delete' || body.op === 'edit') {
    if (!verifyAdmin(body.pin)) return { ok: false, error: 'Manager access required' };
  }

  if (body.op === 'add') {
    if (!body.name || !/^\d{4}$/.test(String(body.pin_new || body.pin))) {
      return { ok: false, error: 'Invalid name or PIN (must be 4 digits)' };
    }
    var sheet = getSheet('users');
    var id = sheet.getLastRow();
    sheet.appendRow([id, body.name, body.pin_new || body.pin, body.role || 'worker', true]);
    return { ok: true };
  }

  if (body.op === 'edit') {
    if (!body.id) return { ok: false, error: 'User ID required' };
    var sheet = getSheet('users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(body.id)) {
        if (body.name) sheet.getRange(i + 1, 2).setValue(body.name);
        if (body.pin_new) {
          if (!/^\d{4}$/.test(String(body.pin_new))) return { ok: false, error: 'PIN must be 4 digits' };
          sheet.getRange(i + 1, 3).setValue(body.pin_new);
        }
        if (body.role) sheet.getRange(i + 1, 4).setValue(body.role);
        return { ok: true };
      }
    }
    return { ok: false, error: 'User not found' };
  }

  if (body.op === 'delete') {
    var sheet = getSheet('users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(body.id)) { sheet.deleteRow(i + 1); return { ok: true }; }
    }
    return { ok: false, error: 'User not found' };
  }

  return { ok: true, users: getAllUsers() };
}

function getAllUsers() {
  var sheet = getSheet('users');
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({ id: data[i][0], name: data[i][1], pin: data[i][2], role: data[i][3], active: data[i][4] });
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
   SITES CRUD
   ════════════════════════════════════════════════════════════ */

function handleSites(body) {
  if (body.op === 'add' || body.op === 'edit' || body.op === 'delete') {
    if (!verifyAdmin(body.pin)) return { ok: false, error: 'Manager access required' };
  }

  if (body.op === 'edit') {
    if (!body.old_site_name || !body.site_name) return { ok: false, error: 'Current and new site names are required' };
    var editSheet = getSheet('sites');
    var editData = editSheet.getDataRange().getValues();
    for (var d = 1; d < editData.length; d++) {
      if (String(editData[d][0]).trim() === String(body.site_name).trim() && String(editData[d][0]).trim() !== String(body.old_site_name).trim()) return { ok: false, error: 'Site already exists' };
    }
    for (var e = 1; e < editData.length; e++) {
      if (String(editData[e][0]).trim() === String(body.old_site_name).trim()) { editSheet.getRange(e + 1, 1).setValue(String(body.site_name).trim()); return { ok: true }; }
    }
    return { ok: false, error: 'Site not found' };
  }

  if (body.op === 'add') {
    if (!body.site_name) return { ok: false, error: 'Site name required' };
    var sheet = getSheet('sites');
    // Check duplicate
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(body.site_name).trim()) return { ok: false, error: 'Site already exists' };
    }
    sheet.appendRow([body.site_name, true]);
    return { ok: true };
  }

  if (body.op === 'delete') {
    var sheet = getSheet('sites');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(body.site_name).trim()) { sheet.deleteRow(i + 1); return { ok: true }; }
    }
    return { ok: false, error: 'Site not found' };
  }

  // Default: list sites
  return { ok: true, sites: getAllSites() };
}

function getAllSites() {
  var sheet = getSheet('sites');
  var data = sheet.getDataRange().getValues();
  var sites = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] !== false && data[i][1] !== 'FALSE') sites.push(data[i][0]);
  }
  return sites;
}

/* ════════════════════════════════════════════════════════════
   OCR — Gemini Vision API
   ════════════════════════════════════════════════════════════ */

function handleOCR(body) {
  var apiKey = getProp('GEMINI_API_KEY') || GEMINI_API_KEY;
  if (!apiKey) return { ok: false, error: 'Gemini API key not configured' };

  var catItems = getAllCatalogueItems().map(function(c) { return c.item_name; });
  var knownItemsHint = catItems.length > 0
    ? ' Known items in this inventory (try to match if possible): ' + catItems.join(', ') + '.'
    : '';

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;

  var prompt = 'You are an inventory assistant for a CCTV installation project. ' +
    'Look at this photo of a stock item or box. ' +
    'Identify: 1) The item type/name (be concise, e.g. "Cable Trunking 25x25mm"). ' +
    '2) The quantity if visible on the label/packaging (if a pack says "24 pcs", quantity is 24). ' +
    'If quantity is not visible, default to 1.' + knownItemsHint + ' ' +
    'Respond ONLY with valid JSON: {"item":"...","qty":number}';

  var payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: body.photo } }] }],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
  };

  try {
    var resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    if (json.error) return { ok: false, error: 'Gemini: ' + json.error.message };
    var text = json.candidates[0].content.parts[0].text;
    try { text = text.replace(/```json\n?/g, '').replace(/```/g, '').trim(); var objectMatch = text.match(/\{[\s\S]*?\}/); if (!objectMatch) throw new Error('No JSON object'); var parsed = JSON.parse(objectMatch[0]); }
    catch (e) { return { ok: false, error: 'Could not parse OCR result: ' + text }; }
    return { ok: true, ocr: { item: parsed.item || '', qty: parseInt(parsed.qty) || 1 }, raw: text };
  } catch (err) { return { ok: false, error: 'Gemini request failed: ' + err.toString() }; }
}

/* ════════════════════════════════════════════════════════════
   TRANSACTION — save + photo + Telegram + update catalogue stock
   ════════════════════════════════════════════════════════════ */

function handleTransaction(body) {
  if (body.type === 'out' && !String(body.destination || '').trim()) return { ok: false, error: 'Destination is required for Stock Out' };
  var ts = new Date().toISOString();
  var photoUrl = '';
  if (body.photo) {
    try { photoUrl = uploadPhotoToDrive(body.photo, body.item, body.user); } catch (e) { Logger.log('Photo upload failed: ' + e); }
  }

  var sheet = getSheet('transactions');
  // Note: column 6 is now 'destination' (was 'notes')
  sheet.appendRow([ts, body.user, body.type, body.item, parseInt(body.qty), body.destination || '', photoUrl, '']);

  updateCatalogueStock(body.item, parseInt(body.qty), body.type, ts);

  // Per-transaction Telegram is now disabled — daily EOD summary instead
  return {
    ok: true,
    transaction: { ts: ts, user: body.user, type: body.type, item: body.item, qty: parseInt(body.qty), destination: body.destination || '' },
    telegram_sent: false
  };
}

function updateCatalogueStock(item, qty, type, ts) {
  var sheet = getSheet('catalogue');
  var data = sheet.getDataRange().getValues();
  var delta = (type === 'in') ? qty : -qty;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(item).trim()) {
      sheet.getRange(i + 1, 5).setValue(Number(data[i][4]) + delta);
      sheet.getRange(i + 1, 6).setValue(ts);
      return;
    }
  }
  sheet.appendRow([item, 'Uncategorized', 'pcs', LOW_STOCK_THRESHOLD, delta, ts, '']);
}

/* ════════════════════════════════════════════════════════════
   CATALOGUE
   ════════════════════════════════════════════════════════════ */

function handleCatalogue(body) {
  if (body.op === 'add' || body.op === 'edit' || body.op === 'delete') {
    if (!verifyAdmin(body.pin)) return { ok: false, error: 'Manager access required' };
  }
  if (body.op === 'add')    return addCatalogueItem(body);
  if (body.op === 'edit')   return editCatalogueItem(body);
  if (body.op === 'delete') return deleteCatalogueItem(body);
  return { ok: true, catalogue: getAllCatalogueItems() };
}

function getAllCatalogueItems() {
  var sheet = getSheet('catalogue');
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    items.push({
      item_name: data[i][0], category: data[i][1], unit: data[i][2],
      min_qty: Number(data[i][3]), current_qty: Number(data[i][4]),
      last_updated: data[i][5], ref_photo_url: normalizeDriveImageUrl(data[i][6] || '')
    });
  }
  return items;
}

function addCatalogueItem(body) {
  if (!body.item_name) return { ok: false, error: 'Item name required' };
  var sheet = getSheet('catalogue');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(body.item_name).trim()) return { ok: false, error: 'Item already exists' };
  }
  var ts = new Date().toISOString();
  var refPhotoUrl = '';
  if (body.ref_photo) {
    try { refPhotoUrl = uploadPhotoToDrive(body.ref_photo, 'ref_' + body.item_name, 'catalogue'); }
    catch (e) { Logger.log('Ref photo upload failed: ' + e); return { ok: false, error: 'Reference photo upload failed: ' + e.message }; }
  }
  sheet.appendRow([body.item_name, body.category || 'Uncategorized', body.unit || 'pcs',
    body.min_qty != null ? Number(body.min_qty) : LOW_STOCK_THRESHOLD,
    body.current_qty != null ? Number(body.current_qty) : 0, ts, refPhotoUrl]);
  return { ok: true, ref_photo_url: refPhotoUrl };
}

function editCatalogueItem(body) {
  if (!body.item_name) return { ok: false, error: 'Item name required' };
  var sheet = getSheet('catalogue');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(body.item_name).trim()) {
      if (body.category) sheet.getRange(i + 1, 2).setValue(body.category);
      if (body.unit) sheet.getRange(i + 1, 3).setValue(body.unit);
      if (body.min_qty != null) sheet.getRange(i + 1, 4).setValue(Number(body.min_qty));
      if (body.current_qty != null) sheet.getRange(i + 1, 5).setValue(Number(body.current_qty));
      sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
      if (body.ref_photo) {
        try { var refPhotoUrl = uploadPhotoToDrive(body.ref_photo, 'ref_' + body.item_name, 'catalogue'); sheet.getRange(i + 1, 7).setValue(refPhotoUrl); }
        catch (e) { Logger.log('Ref photo upload failed: ' + e); return { ok: false, error: 'Reference photo upload failed: ' + e.message }; }
      }
      return { ok: true, ref_photo_url: body.ref_photo ? refPhotoUrl : normalizeDriveImageUrl(data[i][6] || '') };
    }
  }
  return { ok: false, error: 'Item not found' };
}

function deleteCatalogueItem(body) {
  var sheet = getSheet('catalogue');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(body.item_name).trim()) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: 'Item not found' };
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
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
}

function normalizeDriveImageUrl(url) {
  if (!url) return '';
  var value = String(url).trim();
  var match = value.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|thumbnail\?id=|uc\?(?:export=[^&]+&)?id=)([-\w]+)/i);
  return match ? 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w400' : value;
}

function getOrCreateFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

/* ════════════════════════════════════════════════════════════
   TELEGRAM (for daily EOD report only)
   ════════════════════════════════════════════════════════════ */

function sendTelegram(token, chatId, text) {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'HTML' }),
    muteHttpExceptions: true
  });
}

/* ════════════════════════════════════════════════════════════
   LOG QUERIES
   ════════════════════════════════════════════════════════════ */

function handleLog() {
  var sheet = getSheet('transactions');
  var data = sheet.getDataRange().getValues();
  var txs = [];
  for (var i = data.length - 1; i >= 1; i--) {
    txs.push({
      ts: data[i][0], user: data[i][1], type: data[i][2],
      item: data[i][3], qty: Number(data[i][4]), destination: data[i][5], photo_url: data[i][6]
    });
  }
  return { ok: true, transactions: txs };
}

/* ════════════════════════════════════════════════════════════
   REPORT — returns text for in-app display
   ════════════════════════════════════════════════════════════ */

function handleReport(body) {
  // If op=send, also push to Telegram
  if (body.op === 'send') {
    var msg = generateDailyReport();
    var token = getProp('TELEGRAM_BOT_TOKEN') || TELEGRAM_BOT_TOKEN;
    var chatId = getProp('TELEGRAM_CHAT_ID') || TELEGRAM_CHAT_ID;
    if (!token || !chatId) return { ok: false, error: 'Telegram not configured', report: msg };
    sendTelegram(token, chatId, msg);
    return { ok: true, sent: true, message: 'Daily report sent to Telegram', report: msg };
  }
  // Default: return report text for in-app display
  return { ok: true, report: generateDailyReport() };
}

function generateDailyReport() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  var sheet = getSheet('transactions');
  var data = sheet.getDataRange().getValues();
  var todayTx = [];
  for (var i = 1; i < data.length; i++) {
    var ts = new Date(data[i][0]);
    if (ts >= todayStart && ts <= now) {
      todayTx.push({ user: data[i][1], type: data[i][2], item: data[i][3], qty: Number(data[i][4]), destination: data[i][5] });
    }
  }

  var inTx = todayTx.filter(function(t) { return t.type === 'in'; });
  var outTx = todayTx.filter(function(t) { return t.type === 'out'; });
  var inQty = inTx.reduce(function(s, t) { return s + t.qty; }, 0);
  var outQty = outTx.reduce(function(s, t) { return s + t.qty; }, 0);

  var byUser = {};
  todayTx.forEach(function(t) { byUser[t.user] = (byUser[t.user] || 0) + 1; });
  var userLines = Object.keys(byUser).map(function(u) { return u + ' (' + byUser[u] + ')'; }).join(', ');

  // By destination
  var byDest = {};
  todayTx.forEach(function(t) { if (t.destination) byDest[t.destination] = (byDest[t.destination] || 0) + t.qty; });
  var destLines = Object.keys(byDest).map(function(d) { return '• ' + d + ': ' + byDest[d] + ' items'; }).join('\n');

  var catItems = getAllCatalogueItems();
  var lowStock = [];
  catItems.forEach(function(c) {
    var threshold = c.min_qty != null ? Number(c.min_qty) : LOW_STOCK_THRESHOLD;
    if (c.current_qty <= threshold) lowStock.push('• ' + c.item_name + ' (' + c.current_qty + ' ' + (c.unit||'') + ')');
  });

  var dateStr = Utilities.formatDate(now, tz, 'd MMM yyyy');

  var msg = '📊 DAILY INVENTORY REPORT\n' + dateStr + '\n\n';
  msg += 'STOCK MOVEMENT TODAY\n';
  msg += '📥 In:  ' + inQty + ' items (' + inTx.length + ' transactions)\n';
  msg += '📤 Out: ' + outQty + ' items (' + outTx.length + ' transactions)\n';
  if (destLines) msg += '\nBY DESTINATION\n' + destLines + '\n';
  msg += '\nLOW STOCK ALERTS\n';
  msg += lowStock.length ? lowStock.join('\n') + '\n' : '✅ All items above minimum\n';
  msg += '\nTotal transactions today: ' + todayTx.length + '\n';
  if (userLines) msg += 'By: ' + userLines;

  return msg;
}

/* ════════════════════════════════════════════════════════════
   DAILY EOD TRIGGER
   ════════════════════════════════════════════════════════════ */

function sendDailyReport() {
  // Called by time-driven trigger
  var msg = generateDailyReport();
  var token = getProp('TELEGRAM_BOT_TOKEN') || TELEGRAM_BOT_TOKEN;
  var chatId = getProp('TELEGRAM_CHAT_ID') || TELEGRAM_CHAT_ID;
  if (token && chatId) sendTelegram(token, chatId, msg);
}

function handleSchedule(body) {
  // Manager can change the daily EOD hour
  if (body.op === 'set') {
    if (!verifyAdmin(body.pin)) return { ok: false, error: 'Manager access required' };
    var hour = parseInt(body.hour);
    if (isNaN(hour) || hour < 0 || hour > 23) return { ok: false, error: 'Invalid hour (0-23)' };
    PropertiesService.getScriptProperties().setProperty('DAILY_REPORT_HOUR', String(hour));
    // Recreate trigger
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t) {
      if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(hour).create();
    return { ok: true, hour: hour };
  }
  // Default: get current schedule
  var h = PropertiesService.getScriptProperties().getProperty('DAILY_REPORT_HOUR');
  return { ok: true, hour: h ? parseInt(h) : DAILY_REPORT_HOUR };
}

/* ════════════════════════════════════════════════════════════
   CHECK CONFIG
   ════════════════════════════════════════════════════════════ */

function checkConfig() {
  var p = PropertiesService.getScriptProperties().getProperties();
  Logger.log('GEMINI_API_KEY: ' + (p.GEMINI_API_KEY ? '✅ set' : '❌ missing'));
  Logger.log('TELEGRAM_BOT_TOKEN: ' + (p.TELEGRAM_BOT_TOKEN ? '✅ set' : '❌ missing'));
  Logger.log('TELEGRAM_CHAT_ID: ' + (p.TELEGRAM_CHAT_ID ? '✅ set' : '❌ missing'));
  var h = PropertiesService.getScriptProperties().getProperty('DAILY_REPORT_HOUR');
  Logger.log('DAILY_REPORT_HOUR: ' + (h || DAILY_REPORT_HOUR) + ':00');
}

/* ════════════════════════════════════════════════════════════
   SETUP
   ════════════════════════════════════════════════════════════ */

function setup() {
  var config = { GEMINI_API_KEY: '', TELEGRAM_BOT_TOKEN: '', TELEGRAM_CHAT_ID: '' };
  var props = PropertiesService.getScriptProperties();
  for (var key in config) { if (config[key]) props.setProperty(key, config[key]); }

  initSheetHeaders();

  var users = getAllUsers();
  if (users.length === 0) {
    getSheet('users').appendRow([1, 'Manager', '9999', 'admin', true]);
    getSheet('users').appendRow([2, 'User1', '1234', 'worker', true]);
    Logger.log('Seeded default users: Manager/9999, User1/1234');
  }

  var cat = getAllCatalogueItems();
  if (cat.length === 0) { seedCCTVCatalogue(); Logger.log('Seeded CCTV catalogue'); }

  var sites = getAllSites();
  if (sites.length === 0) {
    var sSheet = getSheet('sites');
    sSheet.appendRow(['Site A - Office Tower', true]);
    sSheet.appendRow(['Site B - Residential', true]);
    sSheet.appendRow(['Site C - Warehouse', true]);
    Logger.log('Seeded 3 work sites');
  }

  // Daily EOD trigger
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailyReport' || t.getHandlerFunction() === 'sendWeeklyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });
  var dailyHour = props.getProperty('DAILY_REPORT_HOUR');
  var hour = dailyHour ? parseInt(dailyHour) : DAILY_REPORT_HOUR;
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(hour).create();

  Logger.log('✅ Setup complete (v1.8.1-codex). Daily EOD trigger at ' + hour + ':00.');
}

function seedCCTVCatalogue() {
  var sheet = getSheet('catalogue');
  var ts = new Date().toISOString();
  // Using Wikimedia Commons for representative photos (reliable, permanent URLs)
  var items = [
    ['Cable Trunking 25x25mm White', 'Trunking', 'm',    50,  200, ts, 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/Cable_trunking.jpg/300px-Cable_trunking.jpg'],
    ['Cable Trunking 50x50mm White', 'Trunking', 'm',    30,  120, ts, 'https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/Cable_trunking.jpg/300px-Cable_trunking.jpg'],
    ['RG59 Coaxial Cable 100m Roll', 'Cable',    'roll', 5,   18,  ts, 'https://upload.wikimedia.org/wikipedia/commons/thumb/5/55/Coaxial_cable_cutaway.jpg/300px-Coaxial_cable_cutaway.jpg'],
    ['CCTV Camera Dome 4MP',         'Camera',   'pcs',  10,  45,  ts, 'https://upload.wikimedia.org/wikipedia/commons/thumb/c/c4/Dome_camera.jpg/300px-Dome_camera.jpg'],
    ['Power Supply 12V 10A',         'Power',    'pcs',  5,   8,   ts, 'https://upload.wikimedia.org/wikipedia/commons/thumb/8/85/AC-DC_Power_Supply.jpg/300px-AC-DC_Power_Supply.jpg'],
  ];
  items.forEach(function(row) { sheet.appendRow(row); });
}

/* ════════════════════════════════════════════════════════════
   HELPERS
   ════════════════════════════════════════════════════════════ */

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); initSheetHeadersFor(sheet, name); }
  return sheet;
}

function initSheetHeaders() {
  initSheetHeadersFor(getSheet('users'), 'users');
  initSheetHeadersFor(getSheet('transactions'), 'transactions');
  initSheetHeadersFor(getSheet('catalogue'), 'catalogue');
  initSheetHeadersFor(getSheet('sites'), 'sites');
}

function initSheetHeadersFor(sheet, name) {
  var headers = {
    users:        ['id', 'name', 'pin', 'role', 'active'],
    transactions: ['ts', 'user', 'type', 'item', 'qty', 'destination', 'photo_url', 'gemini_raw'],
    catalogue:    ['item_name', 'category', 'unit', 'min_qty', 'current_qty', 'last_updated', 'ref_photo_url'],
    sites:        ['site_name', 'active']
  };
  sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
  sheet.setFrozenRows(1);
}

function getProp(key) { return PropertiesService.getScriptProperties().getProperty(key); }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
