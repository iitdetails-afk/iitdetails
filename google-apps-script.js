/**
 * =========================================================================
 * INDIAN INFO TECH — GOOGLE SHEETS 3-SHEET DATABASE AUTO-SYNC SCRIPT
 * =========================================================================
 * 
 * Automatically manages 3 separate sheets in your linked Google Spreadsheet:
 *   - Sheet 1: "User Details" (Green header) - Student / User admissions & registrations
 *   - Sheet 2: "Staff Details" (Purple header) - Staff & employee directory
 *   - Sheet 3: "Address Details" (Blue header) - Courier & shipping records
 * 
 * Features:
 *   - Multi-device instant live sync
 *   - Single-record live upsert (append/update without clearing)
 *   - Pull / read from all sheets
 *   - Smart merge & wipe protection
 *   - Auto formatting, frozen headers & zebra stripes
 * =========================================================================
 */
// Fixed permanent Google Spreadsheet Master URL
var SPREADSHEET_URL_OR_ID = "https://docs.google.com/spreadsheets/d/1pqBKWRtoJJJX9JE5pgoEevwkthjicMLLnowEbFCPihY/edit";

function testRun() {
  var ss = getSpreadsheet();
  if (ss) {
    Logger.log('========================================================');
    Logger.log('✅ Connected to Google Sheet: ' + ss.getName());
    Logger.log('🔗 URL: ' + ss.getUrl());
    Logger.log('📊 Sheets in document: ' + ss.getSheets().map(function(s) { return s.getName(); }).join(', '));
    Logger.log('========================================================');
  } else {
    Logger.log('❌ Could not connect. Paste your Google Sheet URL into SPREADSHEET_URL_OR_ID.');
  }
}

function getSpreadsheet(payload, e) {
  var ss = null;

  // 1. Try active spreadsheet (when opened via Extensions > Apps Script)
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (err) {}

  // 2. Try the SPREADSHEET_URL_OR_ID variable at top of script
  if (typeof SPREADSHEET_URL_OR_ID === 'string' && SPREADSHEET_URL_OR_ID.trim()) {
    try {
      var topTarget = SPREADSHEET_URL_OR_ID.trim();
      ss = topTarget.indexOf('http') === 0 ? SpreadsheetApp.openByUrl(topTarget) : SpreadsheetApp.openById(topTarget);
      if (ss) {
        PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
        return ss;
      }
    } catch (err) {}
  }

  // 3. Try URL or ID passed in request payload
  var target = (payload && (payload.spreadsheetUrl || payload.spreadsheetId)) ||
               (e && e.parameter && (e.parameter.spreadsheetUrl || e.parameter.spreadsheetId));
  if (target && typeof target === 'string' && target.trim()) {
    try {
      var reqT = target.trim();
      ss = reqT.indexOf('http') === 0 ? SpreadsheetApp.openByUrl(reqT) : SpreadsheetApp.openById(reqT);
      if (ss) {
        PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
        return ss;
      }
    } catch (err) {}
  }

  // 4. Try stored ID in script properties
  try {
    var storedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (storedId) {
      ss = SpreadsheetApp.openById(storedId);
      if (ss) return ss;
    }
  } catch (err) {}

  // 5. Auto-create in Google Drive if still null
  try {
    ss = SpreadsheetApp.create('IIT Database Master (User Details, Staff Details, Address Details)');
    if (ss) {
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
      return ss;
    }
  } catch (err) {}

  return null;
}

// Column headers definitions
var USER_HEADERS = [
  'S.No', 'User / Submission ID', 'User / Student Name', 'Father / Guardian Name', 'Course / Applied Title',
  'Email Address', 'Primary Mobile', 'Secondary Mobile', 'Date of Birth', 'Gender',
  'City / District', 'State', 'Pincode', 'Full Residential Address', 'College / School',
  'Department / Year', 'Educational Qualification', 'Aadhar / ID Number', 'Fees Fixed',
  'Date of Joining', 'Certificate No', 'Date of Issue', 'Submission Date & Time'
];

var STAFF_HEADERS = [
  'S.No', 'Staff ID', 'Full Name', 'Mobile Number', 'Email Address',
  'Department', 'Designation / Role', 'Account Status', 'Date of Registration', 'Last Updated'
];

var ADDRESS_HEADERS = [
  'S.No', 'Address ID', 'Date & Time', 'To: Recipient Name', 'To: Address Line 1',
  'To: Address Line 2', 'To: Address Line 3', 'To: City / District', 'To: State',
  'To: Pincode', 'To: Primary Phone', 'To: Extra Phone', 'To: Country',
  'From: Sender / Org Name', 'From: Address Line 1', 'From: Address Line 2',
  'From: Address Line 3', 'From: City', 'From: State', 'From: Pincode',
  'From: Phone', 'From: Extra Phone', 'From: Country'
];

function doGet(e) {
  try {
    e = e || { parameter: {} };
    var ss = getSpreadsheet(null, e);
    if (!ss) return jsonResponse({ success: false, error: 'No Google Sheet found. Paste your Google Sheet URL into SPREADSHEET_URL_OR_ID.' });
    
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'status';
    if (action === 'pull' || action === 'read') {
      return jsonResponse(readAllSheetsData(ss));
    }

    return jsonResponse({
      success: true,
      status: 'connected',
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      sheets: ss.getSheets().map(function(s) { return s.getName(); }),
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: 'No payload received' });
    }
    
    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch(parseErr) {
      return jsonResponse({ success: false, error: 'Malformed JSON payload: ' + parseErr.message });
    }

    var ss = getSpreadsheet(payload, e);
    if (!ss) {
      return jsonResponse({ success: false, error: 'No Google Sheet found. Paste your Google Sheet URL into SPREADSHEET_URL_OR_ID.' });
    }

    var action = payload.action || 'sync';

    // ── Ping / Status Check ───────────────────────────────────────────
    if (action === 'ping' || action === 'status') {
      return jsonResponse({
        success: true,
        message: 'Google Sheets is online and connected',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl(),
        sheets: ss.getSheets().map(function(s) { return s.getName(); }),
        timestamp: new Date().toISOString()
      });
    }

    // ── Pull / Read Data ──────────────────────────────────────────────
    if (action === 'pull' || action === 'read') {
      return jsonResponse(readAllSheetsData(ss));
    }

    // ── Upsert Single Student / User ──────────────────────────────────
    if (action === 'upsert_user' || action === 'upsert_student') {
      var userRecord = payload.user || payload.student;
      if (!userRecord) return jsonResponse({ success: false, error: 'No user record provided in payload' });
      var resUser = upsertSingleRecord(ss, 'User Details', ['Students', 'Student Details', 'Users'], USER_HEADERS, userRecord, ['User / Submission ID', 'Email Address'], '#059669');
      cleanSheet1(ss);
      return jsonResponse({
        success: true,
        action: resUser.action,
        row: resUser.row,
        message: 'Student record ' + resUser.action + ' in User Details sheet',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    // ── Upsert Single Staff Member ────────────────────────────────────
    if (action === 'upsert_staff') {
      var staffRecord = payload.staff || payload.staffMember;
      if (!staffRecord) return jsonResponse({ success: false, error: 'No staff record provided in payload' });
      var resStaff = upsertSingleRecord(ss, 'Staff Details', ['Staff'], STAFF_HEADERS, staffRecord, ['Staff ID', 'Email Address'], '#4f46e5');
      cleanSheet1(ss);
      return jsonResponse({
        success: true,
        action: resStaff.action,
        row: resStaff.row,
        message: 'Staff record ' + resStaff.action + ' in Staff Details sheet',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    // ── Upsert Single Address ─────────────────────────────────────────
    if (action === 'upsert_address') {
      var addrRecord = payload.address;
      if (!addrRecord) return jsonResponse({ success: false, error: 'No address record provided in payload' });
      var resAddr = upsertSingleRecord(ss, 'Address Details', ['Addresses', 'Address'], ADDRESS_HEADERS, addrRecord, ['Address ID'], '#2563eb');
      cleanSheet1(ss);
      return jsonResponse({
        success: true,
        action: resAddr.action,
        row: resAddr.row,
        message: 'Address record ' + resAddr.action + ' in Address Details sheet',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    // ── Delete Single Student ─────────────────────────────────────────
    if (action === 'delete_user' || action === 'delete_student') {
      var subId = payload.submissionId || payload.id;
      if (!subId) return jsonResponse({ success: false, error: 'No submissionId provided for deletion' });
      var delRes = deleteRecordFromSheet(ss, 'User Details', ['Students', 'Student Details', 'Users'], 'User / Submission ID', subId);
      return jsonResponse({
        success: true,
        deleted: delRes.deleted,
        message: delRes.deleted ? 'User removed from Google Sheet' : 'User not found in Google Sheet',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    // ── Full Sync (All 3 Sheets) ──────────────────────────────────────
    if (action === 'sync') {
      var staffList = payload.staff;
      var userList = payload.users || payload.students;
      var addressList = payload.addresses;
      var force = !!payload.force;

      var syncedCounts = {};

      // 1. Separate Sheet 1: User Details (Green header)
      if (userList !== undefined) {
        syncSpecificSheet(ss, 'User Details', ['Students', 'Student Details', 'Users'], USER_HEADERS, userList, '#059669', force);
        syncedCounts.users = userList.length;
      }

      // 2. Separate Sheet 2: Staff Details (Purple header)
      if (staffList !== undefined) {
        syncSpecificSheet(ss, 'Staff Details', ['Staff'], STAFF_HEADERS, staffList, '#4f46e5', force);
        syncedCounts.staff = staffList.length;
      }

      // 3. Separate Sheet 3: Address Details (Blue header)
      if (addressList !== undefined) {
        syncSpecificSheet(ss, 'Address Details', ['Addresses', 'Address'], ADDRESS_HEADERS, addressList, '#2563eb', force);
        syncedCounts.addresses = addressList.length;
      }

      cleanSheet1(ss);

      return jsonResponse({
        success: true,
        message: 'All sheets synchronized successfully with Google Drive',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl(),
        counts: syncedCounts,
        syncedAt: new Date().toISOString()
      });
    }

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });

  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Gets or creates a sheet styled with headers
 */
function getOrCreateStyledSheet(ss, targetName, oldAliases, headers, headerColorHex) {
  var sheet = ss.getSheetByName(targetName);
  if (!sheet && oldAliases && oldAliases.length > 0) {
    for (var i = 0; i < oldAliases.length; i++) {
      var oldSheet = ss.getSheetByName(oldAliases[i]);
      if (oldSheet) {
        oldSheet.setName(targetName);
        sheet = oldSheet;
        break;
      }
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(targetName);
  }
  
  // Verify header row exists
  if (sheet.getLastRow() === 0) {
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers]);
    hRange.setBackground(headerColorHex || '#1e293b');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    hRange.setFontSize(10);
    hRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Inserts or updates a single record in a sheet based on key columns
 */
function upsertSingleRecord(ss, targetName, oldAliases, headers, record, keyHeaders, headerColorHex) {
  var sheet = getOrCreateStyledSheet(ss, targetName, oldAliases, headers, headerColorHex);
  var data = sheet.getDataRange().getValues();
  var sheetHeaders = data[0] || headers;

  var keyIndices = [];
  for (var k = 0; k < keyHeaders.length; k++) {
    var idx = sheetHeaders.indexOf(keyHeaders[k]);
    if (idx !== -1) keyIndices.push(idx);
  }

  var targetRow = -1;
  if (data.length > 1 && keyIndices.length > 0) {
    for (var r = 1; r < data.length; r++) {
      for (var ki = 0; ki < keyIndices.length; ki++) {
        var hName = sheetHeaders[keyIndices[ki]];
        var targetVal = record[hName];
        if (targetVal === undefined || targetVal === null) {
          if (hName === 'User / Submission ID') targetVal = record['Submission ID'] || record['submissionId'] || record['User ID'];
          else if (hName === 'Staff ID') targetVal = record['staffId'];
          else if (hName === 'Address ID') targetVal = record['id'];
          else if (hName === 'Email Address') targetVal = record['email'];
        }
        if (targetVal !== undefined && targetVal !== null && String(targetVal).trim() !== '') {
          if (String(data[r][keyIndices[ki]]).trim().toLowerCase() === String(targetVal).trim().toLowerCase()) {
            targetRow = r + 1;
            break;
          }
        }
      }
      if (targetRow !== -1) break;
    }
  }

  // Format row
  var rowValues = headers.map(function(h) {
    var val = record[h];
    if (val === undefined || val === null) {
      if (h === 'User / Submission ID') val = record['Submission ID'] || record['submissionId'] || record['User ID'];
      else if (h === 'User / Student Name') val = record['Student Name'] || record['name'];
      else if (h === 'Staff ID') val = record['staffId'];
      else if (h === 'Full Name') val = record['name'];
      else if (h === 'Address ID') val = record['id'];
      else if (h === 'S.No') val = targetRow !== -1 ? data[targetRow - 1][0] : data.length;
    }
    return val !== undefined && val !== null ? String(val) : '';
  });

  if (targetRow !== -1) {
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowValues]);
    return { action: 'updated', row: targetRow };
  } else {
    var newRowIdx = data.length + 1;
    if (rowValues[0] === '' || rowValues[0] === '0') {
      rowValues[0] = String(data.length);
    }
    sheet.appendRow(rowValues);
    var appRange = sheet.getRange(newRowIdx, 1, 1, headers.length);
    appRange.setFontSize(9);
    appRange.setVerticalAlignment('middle');
    if (newRowIdx % 2 === 1) {
      appRange.setBackground('#f8fafc');
    }
    return { action: 'inserted', row: newRowIdx };
  }
}

/**
 * Deletes a row matching keyHeader == keyValue
 */
function deleteRecordFromSheet(ss, sheetName, oldAliases, keyHeader, keyValue) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet && oldAliases) {
    for (var i = 0; i < oldAliases.length; i++) {
      sheet = ss.getSheetByName(oldAliases[i]);
      if (sheet) break;
    }
  }
  if (!sheet) return { deleted: false };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { deleted: false };
  var colIdx = data[0].indexOf(keyHeader);
  if (colIdx === -1) {
    // Try alternate header
    for (var c = 0; c < data[0].length; c++) {
      var h = String(data[0][c]).toLowerCase();
      if (h.indexOf('submission id') !== -1 || h.indexOf('user id') !== -1) {
        colIdx = c;
        break;
      }
    }
  }
  if (colIdx === -1) return { deleted: false };

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][colIdx]).trim().toLowerCase() === String(keyValue).trim().toLowerCase()) {
      sheet.deleteRow(r + 1);
      return { deleted: true, row: r + 1 };
    }
  }
  return { deleted: false };
}

/**
 * Creates, renames from old alias, and styles each specific sheet
 */
function syncSpecificSheet(ss, targetName, oldAliases, headers, records, headerColorHex, force) {
  var sheet = ss.getSheetByName(targetName);

  if (!sheet && oldAliases && oldAliases.length > 0) {
    for (var i = 0; i < oldAliases.length; i++) {
      var oldSheet = ss.getSheetByName(oldAliases[i]);
      if (oldSheet) {
        oldSheet.setName(targetName);
        sheet = oldSheet;
        break;
      }
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(targetName);
  }

  // Safety check: protect existing rows if records array is empty and force is not set
  if ((!records || records.length === 0) && !force && sheet.getLastRow() > 1) {
    Logger.log('[Sync] Protected existing sheet data from being wiped: ' + targetName);
    return;
  }

  // Clear previous data and formatting
  sheet.clear();

  // Write headers
  var hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setBackground(headerColorHex || '#1e293b');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(10);
  hRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Write rows
  if (records && records.length > 0) {
    var rows = records.map(function(item) {
      return headers.map(function(h) {
        var val = item[h];
        if (val === undefined || val === null) {
          if (h === 'User / Submission ID') val = item['Submission ID'] || item['submissionId'] || item['User ID'];
          else if (h === 'User / Student Name') val = item['Student Name'] || item['name'];
          else if (h === 'Staff ID') val = item['staffId'];
          else if (h === 'Full Name') val = item['name'];
          else if (h === 'Address ID') val = item['id'];
        }
        return val !== undefined && val !== null ? String(val) : '';
      });
    });

    var dRange = sheet.getRange(2, 1, rows.length, headers.length);
    dRange.setValues(rows);
    dRange.setFontSize(9);
    dRange.setVerticalAlignment('middle');

    // Alternating zebra row colors
    for (var r = 0; r < rows.length; r++) {
      if (r % 2 === 1) sheet.getRange(r + 2, 1, 1, headers.length).setBackground('#f8fafc');
    }
  }

  // Auto column widths
  for (var col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
    var w = sheet.getColumnWidth(col);
    if (w < 90) sheet.setColumnWidth(col, 90);
    if (w > 350) sheet.setColumnWidth(col, 350);
  }
}

function cleanSheet1(ss) {
  try {
    var dSheet = ss.getSheetByName('Sheet1');
    if (dSheet && ss.getSheets().length > 3) ss.deleteSheet(dSheet);
  } catch(err) {}
}

function readAllSheetsData(ss) {
  return {
    success: true,
    users: findAndReadSheet(ss, ['User Details', 'Students', 'Users']),
    students: findAndReadSheet(ss, ['User Details', 'Students', 'Users']),
    staff: findAndReadSheet(ss, ['Staff Details', 'Staff']),
    addresses: findAndReadSheet(ss, ['Address Details', 'Addresses', 'Address']),
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    pulledAt: new Date().toISOString()
  };
}

function findAndReadSheet(ss, candidateNames) {
  for (var i = 0; i < candidateNames.length; i++) {
    var sheet = ss.getSheetByName(candidateNames[i]);
    if (sheet) {
      return readSheetRows(sheet);
    }
  }
  return [];
}

function readSheetRows(sheet) {
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) row[headers[j]] = data[i][j];
    rows.push(row);
  }
  return rows;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
