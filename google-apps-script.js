/**
 * =========================================================================
 * INDIAN INFO TECH — GOOGLE SHEETS 3-SHEET DATABASE AUTO-SYNC SCRIPT
 * =========================================================================
 * 
 * Automatically manages 3 separate sheets in your linked Google Spreadsheet:
 *   - Sheet 1: "User Details" (Green header) - Student / User admissions & registrations
 *   - Sheet 2: "Staff Details" (Purple header) - Staff & employee directory
 *   - Sheet 3: "Address Details" (Blue header) - Courier & shipping records
 * =========================================================================
 */
var SPREADSHEET_URL_OR_ID = "";

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

  // 4. Try stored ID
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

function doGet(e) {
  try {
    e = e || { parameter: {} };
    var ss = getSpreadsheet(null, e);
    if (!ss) return jsonResponse({ success: false, error: 'No Google Sheet found. Paste your Google Sheet URL into SPREADSHEET_URL_OR_ID.' });
    
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'status';
    if (action === 'pull' || action === 'read') return jsonResponse(readAllSheetsData(ss));

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
    if (!e || !e.postData || !e.postData.contents) return jsonResponse({ success: false, error: 'No payload received' });
    var payload = JSON.parse(e.postData.contents);
    var ss = getSpreadsheet(payload, e);
    if (!ss) return jsonResponse({ success: false, error: 'No Google Sheet found. Paste your Google Sheet URL into SPREADSHEET_URL_OR_ID.' });

    if (payload.action === 'ping') {
      return jsonResponse({
        success: true,
        message: 'Google Sheets is online and connected',
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl(),
        sheets: ss.getSheets().map(function(s) { return s.getName(); })
      });
    }

    var staffList = payload.staff || [];
    var userList = payload.users || payload.students || [];
    var addressList = payload.addresses || [];

    // 1. Separate Sheet 1: User Details (Green header)
    var userHeaders = [
      'S.No', 'User / Submission ID', 'User / Student Name', 'Father / Guardian Name', 'Course / Applied Title',
      'Email Address', 'Primary Mobile', 'Secondary Mobile', 'Date of Birth', 'Gender',
      'City / District', 'State', 'Pincode', 'Full Residential Address', 'College / School',
      'Department / Year', 'Educational Qualification', 'Aadhar / ID Number', 'Fees Fixed',
      'Date of Joining', 'Certificate No', 'Date of Issue', 'Submission Date & Time'
    ];
    syncSpecificSheet(ss, 'User Details', ['Students', 'Student Details', 'Users'], userHeaders, userList, '#059669');

    // 2. Separate Sheet 2: Staff Details (Purple header)
    var staffHeaders = [
      'S.No', 'Staff ID', 'Full Name', 'Mobile Number', 'Email Address',
      'Department', 'Designation / Role', 'Account Status', 'Date of Registration', 'Last Updated'
    ];
    syncSpecificSheet(ss, 'Staff Details', ['Staff'], staffHeaders, staffList, '#4f46e5');

    // 3. Separate Sheet 3: Address Details (Blue header)
    var addressHeaders = [
      'S.No', 'Address ID', 'Date & Time', 'To: Recipient Name', 'To: Address Line 1',
      'To: Address Line 2', 'To: Address Line 3', 'To: City / District', 'To: State',
      'To: Pincode', 'To: Primary Phone', 'To: Extra Phone', 'To: Country',
      'From: Sender / Org Name', 'From: Address Line 1', 'From: Address Line 2',
      'From: Address Line 3', 'From: City', 'From: State', 'From: Pincode',
      'From: Phone', 'From: Extra Phone', 'From: Country'
    ];
    syncSpecificSheet(ss, 'Address Details', ['Addresses', 'Address'], addressHeaders, addressList, '#2563eb');

    // Clean up empty default Sheet1 if all 3 separate sheets exist
    try {
      var dSheet = ss.getSheetByName('Sheet1');
      if (dSheet && ss.getSheets().length > 3) ss.deleteSheet(dSheet);
    } catch(err) {}

    return jsonResponse({
      success: true,
      message: 'All 3 separate sheets (User Details, Staff Details, Address Details) synchronized successfully',
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      counts: {
        users: userList.length,
        staff: staffList.length,
        addresses: addressList.length
      },
      syncedAt: new Date().toISOString()
    });

  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Creates, renames from old alias, and styles each specific sheet
 */
function syncSpecificSheet(ss, targetName, oldAliases, headers, records, headerColorHex) {
  var sheet = ss.getSheetByName(targetName);

  // If target does not exist, check if an old alias exists (e.g. Students -> User Details) and rename it
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
          // Check alternate alias
          if (h === 'User / Submission ID') val = item['Submission ID'] || item['User ID'];
          else if (h === 'User / Student Name') val = item['Student Name'] || item['Name'];
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
