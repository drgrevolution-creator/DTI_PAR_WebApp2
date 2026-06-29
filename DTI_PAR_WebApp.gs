/**
 * DTI_PAR_WebApp.gs  —  v2.1
 * Google Apps Script · Web App backend cho Bảng theo dõi DTI & PAR INDEX
 * Xã Phước Hải · 2026
 *
 * ═══════════════════════════════════════════════════════════
 *  CÁCH TRIỂN KHAI (chọn 1 trong 2 phương án)
 * ═══════════════════════════════════════════════════════════
 *
 *  PHƯƠNG ÁN A — Script gắn vào Google Sheet (khuyến nghị, dễ nhất)
 *  ──────────────────────────────────────────────────────────
 *  1. Mở Google Sheet của bạn
 *  2. Vào menu: Extensions → Apps Script
 *  3. Xóa code mặc định, dán toàn bộ file này vào → Lưu (Ctrl+S)
 *  4. Để SHEET_ID = '' (trống) — script sẽ tự dùng Sheet hiện tại
 *  5. Chạy hàm setupSheets() một lần: Run → Run function → setupSheets
 *     → Bấm "Review permissions" → cho phép truy cập
 *  6. Deploy → New deployment → Web App
 *     · Execute as: Me
 *     · Who has access: Anyone
 *  7. Copy URL Web App → dán vào modal ⚙️ trong file HTML
 *
 *  PHƯƠNG ÁN B — Script standalone tại script.google.com
 *  ──────────────────────────────────────────────────────────
 *  1. Vào script.google.com → New project → dán code này vào
 *  2. Chạy hàm setupSheets() lần đầu — script sẽ TỰ TẠO Google Sheet mới
 *     và in ra SHEET_ID trong Logs (View → Logs)
 *  3. Sao chép SHEET_ID từ Logs → điền vào dòng const SHEET_ID bên dưới → Lưu
 *  4. Deploy → New deployment → Web App (Execute as: Me, Access: Anyone)
 *  5. Copy URL → dán vào modal ⚙️ trong file HTML
 * ═══════════════════════════════════════════════════════════
 */

// ===================== CẤU HÌNH =====================
const SHEET_ID = '';       // Để trống nếu dùng Phương án A; điền ID nếu dùng Phương án B
const SHEET_STATUS = 'STATUS';
const SHEET_LOG    = 'LOG';
const PROP_SHEET_ID  = 'CREATED_SHEET_ID';  // PropertiesService key lưu ID Sheet tự tạo
const PROP_ADMIN_PIN = 'ADMIN_PIN_HASH';    // SHA-256 hash của mã PIN admin
const PROP_ADMINS    = 'ADMIN_EMAILS';      // JSON array danh sách email admin

// ===================== LẤY SPREADSHEET =====================
function getSS() {
  // Ưu tiên 1: SHEET_ID được cấu hình thủ công
  if (SHEET_ID && SHEET_ID.trim() !== '') {
    return SpreadsheetApp.openById(SHEET_ID.trim());
  }

  // Ưu tiên 2: Script gắn vào Sheet (Phương án A)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) { /* standalone — không có active spreadsheet */ }

  // Ưu tiên 3: ID đã được lưu từ lần setupSheets() trước (Phương án B)
  const savedId = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  if (savedId) {
    return SpreadsheetApp.openById(savedId);
  }

  throw new Error('Chưa cấu hình Google Sheet. Hãy chạy hàm setupSheets() trước.');
}

function getSheet(name) {
  const ss = getSS();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Không tìm thấy sheet "' + name + '". Hãy chạy setupSheets() trước.');
  return sheet;
}

// ===================== ENTRY POINTS =====================

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || '';
    if (action === 'ping')     return jsonOk({ time: new Date().toISOString(), version: '2.1' });
    if (action === 'getData')     return jsonOk(handleGetData(params.sheet || 'ALL'));
    if (action === 'getStats')    return jsonOk(handleGetStats());
    if (action === 'getCriteria')  return jsonOk(handleGetCriteria(params.sheet || 'ALL'));
    if (action === 'getRole')      return jsonOk(handleGetRole(params.token || ''));
    if (action === 'load')         return jsonOk(handleGetData('ALL')); // legacy
    return jsonOk({ message: 'DTI_PAR WebApp v2.1 ready', time: new Date().toISOString() });
  } catch(err) {
    return jsonError(err.message);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Không có dữ liệu POST');
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    if (action === 'updateRow')   return jsonOk(handleUpdateRow(body));
    if (action === 'save')        return jsonOk(handleSave(body));
    if (action === 'saveCriteria') { requireAdmin(body); return jsonOk(handleSaveCriteria(body)); }
    if (action === 'verifyPin')    return jsonOk(handleVerifyPin(body));
    if (action === 'saveUpdater')  return jsonOk({ saved: true });
    return jsonError('Hành động không hợp lệ: ' + action);
  } catch(err) {
    return jsonError(err.message);
  }
}

// ===================== HANDLERS =====================

function handleGetData(sheetFilter) {
  const sheet = getSheet(SHEET_STATUS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { statusMap: {} };

  const statusMap = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    const source = String(row[7] || '');
    if (sheetFilter !== 'ALL' && source !== sheetFilter) continue;

    const rawDate = row[5];
    let updatedAt = null;
    if (rawDate) {
      try { updatedAt = new Date(rawDate).getTime(); } catch(e) {}
    }

    statusMap[id] = {
      status:    String(row[1] || 'chua'),
      progress:  Number(row[2]) || 0,
      result:    String(row[3] || ''),
      note:      String(row[4] || ''),
      updatedAt: updatedAt,
      updatedBy: String(row[6] || ''),
      source:    source
    };
  }
  return { statusMap };
}

function handleUpdateRow(body) {
  const id = String(body.id || '').trim();
  if (!id) throw new Error('Thiếu id');

  const sheet = getSheet(SHEET_STATUS);
  const rows  = sheet.getDataRange().getValues();
  const now   = new Date().toISOString();

  // Tìm dòng theo id (rows[0] là header, rows[i] là dữ liệu → Sheet row = i+1)
  let sheetRow = -1;   // 1-based row index trong Sheet
  let existingRow = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === id) {
      sheetRow = i + 1;
      existingRow = rows[i];
      break;
    }
  }

  // Dữ liệu mới — giữ giá trị cũ nếu field không được gửi lên
  const newRow = [
    id,
    body.status    !== undefined ? body.status    : (existingRow ? existingRow[1] : 'chua'),
    body.progress  !== undefined ? Number(body.progress) : (existingRow ? Number(existingRow[2]) : 0),
    body.result    !== undefined ? body.result    : (existingRow ? existingRow[3] : ''),
    body.note      !== undefined ? body.note      : (existingRow ? existingRow[4] : ''),
    now,
    body.updaterName || (existingRow ? existingRow[6] : '') || '',
    body.source    || (existingRow ? existingRow[7] : '') || ''
  ];

  if (sheetRow > 0) {
    writeLog(id, existingRow, newRow, body.updaterName);
    sheet.getRange(sheetRow, 1, 1, 8).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
    writeLog(id, null, newRow, body.updaterName);
  }

  return { id: id, updatedAt: now, action: sheetRow > 0 ? 'updated' : 'created' };
}

function handleSave(body) {
  const statusMap  = body.statusMap;
  const updaterName = body.updaterName || '';
  if (!statusMap || typeof statusMap !== 'object') throw new Error('Thiếu statusMap hợp lệ');

  const sheet = getSheet(SHEET_STATUS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();

  const now = new Date().toISOString();
  const rows = Object.entries(statusMap).map(function(entry) {
    const id = entry[0];
    const s  = entry[1];
    return [
      id,
      s.status   || 'chua',
      Number(s.progress) || 0,
      s.result   || '',
      s.note     || '',
      s.updatedAt ? new Date(s.updatedAt).toISOString() : now,
      s.updatedBy || updaterName,
      s.source   || ''
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  }
  return { saved: rows.length, updatedAt: now };
}

function handleGetStats() {
  const result  = handleGetData('ALL');
  const entries = Object.values(result.statusMap);
  const stats = { total: entries.length, hoanthanh: 0, dangtrienkhai: 0, chua: 0, khongdat: 0, updatedToday: 0 };
  const todayStart = new Date(); todayStart.setHours(0, 0, 0, 0);

  entries.forEach(function(s) {
    if (stats[s.status] !== undefined) stats[s.status]++;
    if (s.updatedAt && s.updatedAt >= todayStart.getTime()) stats.updatedToday++;
  });
  stats.pctDone = stats.total ? Math.round(stats.hoanthanh / stats.total * 1000) / 10 : 0;
  return { stats: stats };
}

// ===================== UTILITIES =====================

function writeLog(id, oldRow, newRow, updaterName) {
  try {
    const log = getSheet(SHEET_LOG);
    const now = new Date().toISOString();
    var fields = ['status', 'progress', 'result', 'note'];
    fields.forEach(function(f, i) {
      var oldVal = oldRow ? String(oldRow[i + 1] || '') : '';
      var newVal = String(newRow[i + 1] || '');
      if (oldVal !== newVal) {
        log.appendRow([now, updaterName || '', id, f, oldVal, newVal]);
      }
    });
  } catch(e) { /* Log không quan trọng, bỏ qua */ }
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== CRITERIA HANDLERS =====================

/**
 * Trả về bộ chỉ số DTI và/hoặc PAR từ Google Sheets
 * sheet = 'DTI' | 'PAR' | 'ALL'
 */
function handleGetCriteria(sheetFilter) {
  var result = {};

  if (sheetFilter === 'DTI' || sheetFilter === 'ALL') {
    try {
      var dtiSheet = getSheet('DTI_CRITERIA');
      var dtiRows  = dtiSheet.getDataRange().getValues();
      if (dtiRows.length > 1) {
        var dtiHdr = dtiRows[0];
        result.dti = dtiRows.slice(1).filter(function(r){ return r[0]; }).map(function(r){
          var obj = {};
          dtiHdr.forEach(function(h, i){ if(h) obj[h] = r[i] !== null && r[i] !== undefined ? r[i] : ''; });
          return obj;
        });
      }
    } catch(e) { /* sheet chưa tồn tại — trả về rỗng */ }
  }

  if (sheetFilter === 'PAR' || sheetFilter === 'ALL') {
    try {
      var parSheet = getSheet('PAR_CRITERIA');
      var parRows  = parSheet.getDataRange().getValues();
      if (parRows.length > 1) {
        var parHdr = parRows[0];
        result.par = parRows.slice(1).filter(function(r){ return r[0]; }).map(function(r){
          var obj = {};
          parHdr.forEach(function(h, i){ if(h) obj[h] = r[i] !== null && r[i] !== undefined ? r[i] : ''; });
          return obj;
        });
      }
    } catch(e) {}
  }

  return result;
}

/**
 * Lưu toàn bộ bộ chỉ số (DTI hoặc PAR) lên sheet tương ứng
 */
function handleSaveCriteria(body) {
  var source = body.sheet;      // 'DTI' hoặc 'PAR'
  var rows   = body.rows;       // mảng object
  if (!source || !rows || !Array.isArray(rows)) throw new Error('Thiếu sheet hoặc rows');

  var sheetName = source === 'DTI' ? 'DTI_CRITERIA' : 'PAR_CRITERIA';
  var ss = getSS();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  // Lấy hoặc tạo header từ keys của object đầu tiên
  var headers = rows.length > 0 ? Object.keys(rows[0]).filter(function(k){ return k !== '_new'; }) : [];
  if (headers.length === 0) return { saved: 0 };

  // Xóa nội dung cũ
  sheet.clearContents();

  // Ghi header
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var hdrRange = sheet.getRange('1:1');
  hdrRange.setFontWeight('bold');
  hdrRange.setBackground(source === 'DTI' ? '#1e2a44' : '#4b5774');
  hdrRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Ghi dữ liệu
  var dataRows = rows.map(function(obj){
    return headers.map(function(h){ return obj[h] !== undefined ? obj[h] : ''; });
  });
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  Logger.log('Đã lưu ' + dataRows.length + ' dòng vào sheet ' + sheetName);
  return { saved: dataRows.length, sheet: sheetName };
}

// ===================== AUTH / PHÂN QUYỀN =====================

function sha256Hash(text) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return digest.map(function(b){ return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

/**
 * Xác minh mã PIN admin. Trả về { ok: true, token } hoặc { ok: false }.
 * Token ngắn hạn (8 giờ) được lưu trong PropertiesService để xác thực các request sau.
 */
function handleVerifyPin(body) {
  var pin = String(body.pin || '').trim();
  if (!pin) return { ok: false, error: 'Thiếu mã PIN.' };

  var storedHash = PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_PIN);
  if (!storedHash) return { ok: false, error: 'PIN admin chưa được thiết lập. Hãy chạy setAdminPin("your-pin") trong Apps Script Editor.' };

  if (sha256Hash(pin) !== storedHash) return { ok: false };

  // Tạo token ngắn hạn
  var token = Utilities.getUuid();
  var tokens = [];
  try { tokens = JSON.parse(PropertiesService.getScriptProperties().getProperty('ADMIN_TOKENS') || '[]'); } catch(e) {}
  // Xoay vòng: chỉ giữ token còn hiệu lực (< 8 giờ), tối đa 20 token
  var now = Date.now();
  tokens = tokens.filter(function(t){ return now - t.created < 8 * 3600 * 1000; }).slice(-19);
  tokens.push({ token: token, created: now });
  PropertiesService.getScriptProperties().setProperty('ADMIN_TOKENS', JSON.stringify(tokens));

  return { ok: true, token: token };
}

/**
 * Kiểm tra _adminToken trong body. Ném lỗi nếu token không hợp lệ hoặc hết hạn.
 */
function requireAdmin(body) {
  var adminToken = body._adminToken;
  if (!adminToken) throw new Error('Cần quyền Admin để thực hiện thao tác này.');

  var tokens = [];
  try { tokens = JSON.parse(PropertiesService.getScriptProperties().getProperty('ADMIN_TOKENS') || '[]'); } catch(e) {}
  var now = Date.now();
  var valid = tokens.some(function(t){ return t.token === adminToken && (now - t.created < 8 * 3600 * 1000); });
  if (!valid) throw new Error('Phiên Admin đã hết hạn hoặc không hợp lệ. Vui lòng đăng nhập lại.');
}

/**
 * Xác minh Google ID Token (JWT) và trả về role của người dùng.
 * Gọi qua GET ?action=getRole&token=<id_token>
 */
function handleGetRole(idToken) {
  if (!idToken) return { role: 'viewer' };
  try {
    var url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return { role: 'editor' }; // token hợp lệ nhưng không xác minh được → coi là editor
    var payload = JSON.parse(response.getContentText());
    var email = payload.email || '';
    if (!email) return { role: 'viewer' };

    var adminEmails = [];
    try { adminEmails = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_ADMINS) || '[]'); } catch(e) {}
    var role = adminEmails.indexOf(email) >= 0 ? 'admin' : 'editor';
    return { role: role, email: email };
  } catch(e) {
    Logger.log('handleGetRole error: ' + e.message);
    return { role: 'editor' }; // fallback an toàn
  }
}

// ── Các hàm tiện ích để cấu hình — chạy thủ công trong Apps Script Editor ──

/**
 * Thiết lập mã PIN admin. Chạy 1 lần trong Editor.
 * Ví dụ: setAdminPin("1234")
 */
function setAdminPin(pin) {
  if (!pin || String(pin).trim().length < 4) {
    Logger.log('❌ PIN phải có ít nhất 4 ký tự.');
    return;
  }
  var hash = sha256Hash(String(pin).trim());
  PropertiesService.getScriptProperties().setProperty(PROP_ADMIN_PIN, hash);
  // Xóa toàn bộ token cũ khi đổi PIN
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_TOKENS');
  Logger.log('✅ Admin PIN đã được thiết lập thành công.');
}

/**
 * Thêm email vào danh sách Admin (có thể đăng nhập bằng Google và có quyền Admin).
 * Ví dụ: addAdminEmail("ten@gmail.com")
 */
function addAdminEmail(email) {
  var list = [];
  try { list = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_ADMINS) || '[]'); } catch(e) {}
  email = String(email || '').trim().toLowerCase();
  if (!email) { Logger.log('❌ Email không hợp lệ.'); return; }
  if (list.indexOf(email) < 0) {
    list.push(email);
    PropertiesService.getScriptProperties().setProperty(PROP_ADMINS, JSON.stringify(list));
  }
  Logger.log('✅ Danh sách Admin: ' + JSON.stringify(list));
}

/**
 * Xóa email khỏi danh sách Admin.
 * Ví dụ: removeAdminEmail("ten@gmail.com")
 */
function removeAdminEmail(email) {
  var list = [];
  try { list = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_ADMINS) || '[]'); } catch(e) {}
  email = String(email || '').trim().toLowerCase();
  list = list.filter(function(e){ return e !== email; });
  PropertiesService.getScriptProperties().setProperty(PROP_ADMINS, JSON.stringify(list));
  Logger.log('✅ Danh sách Admin sau khi xóa: ' + JSON.stringify(list));
}

/**
 * Xem danh sách Admin hiện tại.
 */
function listAdmins() {
  var list = [];
  try { list = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_ADMINS) || '[]'); } catch(e) {}
  var pinSet = !!PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_PIN);
  Logger.log('Admin emails: ' + JSON.stringify(list));
  Logger.log('PIN đã thiết lập: ' + (pinSet ? 'Có' : 'Chưa'));
}

// ===================== SETUP — chạy 1 lần =====================

/**
 * Chạy hàm này ĐẦU TIÊN để tạo cấu trúc sheet.
 *
 * Phương án A (script gắn vào Sheet): sẽ tạo sheet STATUS và LOG trong Spreadsheet hiện tại.
 * Phương án B (standalone):           sẽ TỰ TẠO Spreadsheet mới trên Drive, in ID vào Logs.
 *
 * Cách chạy trong Apps Script Editor:
 *   → Chọn "setupSheets" trong dropdown hàm → Nhấn nút ▶ Run
 *   → Lần đầu sẽ hỏi cấp quyền → Bấm "Review permissions" → Allow
 */
function setupSheets() {
  var ss;

  // Lấy hoặc tạo Spreadsheet
  if (SHEET_ID && SHEET_ID.trim() !== '') {
    ss = SpreadsheetApp.openById(SHEET_ID.trim());
    Logger.log('Dùng Sheet có SHEET_ID: ' + SHEET_ID);
  } else {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) throw new Error('Không có active spreadsheet');
      Logger.log('Dùng Active Spreadsheet: ' + ss.getId());
    } catch(e) {
      // Standalone — tạo Sheet mới
      var savedId = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
      if (savedId) {
        try {
          ss = SpreadsheetApp.openById(savedId);
          Logger.log('Dùng Sheet đã tạo trước: ' + savedId);
        } catch(e2) { ss = null; }
      }
      if (!ss) {
        ss = SpreadsheetApp.create('DTI_PAR_PhuocHai_2026');
        var newId = ss.getId();
        PropertiesService.getScriptProperties().setProperty(PROP_SHEET_ID, newId);
        Logger.log('═══════════════════════════════════════');
        Logger.log('✅ Đã tạo Google Sheet mới!');
        Logger.log('SHEET_ID = ' + newId);
        Logger.log('Mở Sheet: https://docs.google.com/spreadsheets/d/' + newId);
        Logger.log('═══════════════════════════════════════');
        Logger.log('→ Hãy sao chép SHEET_ID trên vào dòng const SHEET_ID rồi lưu lại.');
      }
    }
  }

  // Tạo sheet STATUS
  var statusSheet = ss.getSheetByName(SHEET_STATUS);
  if (!statusSheet) {
    statusSheet = ss.insertSheet(SHEET_STATUS);
    Logger.log('Đã tạo sheet: ' + SHEET_STATUS);
  }
  // Chỉ thêm header nếu sheet trống
  if (statusSheet.getLastRow() === 0) {
    statusSheet.appendRow(['id', 'status', 'progress', 'result', 'note', 'updatedAt', 'updatedBy', 'source']);
    var headerRange = statusSheet.getRange('1:1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e2a44');
    headerRange.setFontColor('#ffffff');
    statusSheet.setFrozenRows(1);
    statusSheet.setColumnWidth(1, 120);  // id
    statusSheet.setColumnWidth(2, 130);  // status
    statusSheet.setColumnWidth(3, 80);   // progress
    statusSheet.setColumnWidth(4, 220);  // result
    statusSheet.setColumnWidth(5, 260);  // note
    statusSheet.setColumnWidth(6, 160);  // updatedAt
    statusSheet.setColumnWidth(7, 140);  // updatedBy
    statusSheet.setColumnWidth(8, 70);   // source
    // Dropdown validation cho cột status
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['chua', 'dangtrienkhai', 'hoanthanh', 'khongdat'], true)
      .setAllowInvalid(true)
      .build();
    statusSheet.getRange('B2:B1000').setDataValidation(rule);
    Logger.log('Đã thiết lập header và định dạng cho sheet STATUS');
  } else {
    Logger.log('Sheet STATUS đã tồn tại (' + statusSheet.getLastRow() + ' dòng)');
  }

  // Tạo sheet LOG
  var logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_LOG);
    Logger.log('Đã tạo sheet: ' + SHEET_LOG);
  }
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['timestamp', 'updatedBy', 'id', 'field', 'oldValue', 'newValue']);
    var logHeader = logSheet.getRange('1:1');
    logHeader.setFontWeight('bold');
    logHeader.setBackground('#4b5774');
    logHeader.setFontColor('#ffffff');
    logSheet.setFrozenRows(1);
    Logger.log('Đã thiết lập header cho sheet LOG');
  } else {
    Logger.log('Sheet LOG đã tồn tại (' + logSheet.getLastRow() + ' dòng)');
  }

  Logger.log('');
  // Sheet DTI_CRITERIA và PAR_CRITERIA — sẽ được tạo tự động khi lưu lần đầu từ HTML
  // (không cần tạo thủ công — handleSaveCriteria() sẽ tạo khi cần)
  Logger.log('Ghi chú: Sheet DTI_CRITERIA và PAR_CRITERIA sẽ tự tạo khi nhấn "Lưu lên Sheets" từ tab Bộ chỉ số trong file HTML.');

  Logger.log('✅ Setup hoàn tất!');
  Logger.log('Bước tiếp theo: Deploy → New deployment → Web App');
  Logger.log('  · Execute as: Me');
  Logger.log('  · Who has access: Anyone');
  Logger.log('Sau khi deploy xong → copy URL → dán vào modal ⚙️ trong file HTML');

  // Hiện thông báo nếu chạy từ Sheet UI (Phương án A)
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ Setup thành công!\n\n' +
      'Sheet STATUS và LOG đã sẵn sàng.\n\n' +
      'Bước tiếp theo:\n' +
      '1. Deploy → New deployment → Web App\n' +
      '2. Execute as: Me  |  Who has access: Anyone\n' +
      '3. Nhấn Deploy → Copy URL\n' +
      '4. Dán URL vào modal ⚙️ Cài đặt trong file HTML'
    );
  } catch(e) {
    // Standalone — không có UI, chỉ log (đã log ở trên)
    Logger.log('(Chạy ở chế độ standalone — xem kết quả trong View → Logs)');
  }
}

/**
 * Hàm test nhanh — chạy trong Editor để kiểm tra kết nối và đọc dữ liệu
 */
function testConnection() {
  try {
    var result = handleGetData('ALL');
    var count = Object.keys(result.statusMap).length;
    Logger.log('✅ Kết nối thành công!');
    Logger.log('Số dòng dữ liệu trong STATUS: ' + count);
    Logger.log('Sheet URL: ' + getSS().getUrl());
  } catch(e) {
    Logger.log('❌ Lỗi: ' + e.message);
  }
}
