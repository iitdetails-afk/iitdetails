const fs = require('fs');
const path = require('path');
const https = require('https');
const http = require('http');
const excelDb = require('./excel-db');

/**
 * HTTP helper that follows redirects (Google Apps Script always responds to POST with 302)
 */
function requestWithRedirect(url, options = {}, postData = null, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) {
      return reject(new Error('Too many HTTP redirects'));
    }

    try {
      const parsedUrl = new URL(url);
      const isHttps = parsedUrl.protocol === 'https:';
      const lib = isHttps ? https : http;

      const reqOptions = {
        method: options.method || 'GET',
        headers: { ...(options.headers || {}) },
        timeout: options.timeout || 35000
      };

      const req = lib.request(parsedUrl, reqOptions, (res) => {
        // Handle 301, 302, 303, 307, 308 redirects
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          const redirectUrl = new URL(res.headers.location, url).href;
          const nextHeaders = { ...options.headers };
          delete nextHeaders['content-length'];
          delete nextHeaders['Content-Length'];
          delete nextHeaders['content-type'];
          delete nextHeaders['Content-Type'];
          const nextOptions = { ...options, method: 'GET', headers: nextHeaders };
          return resolve(requestWithRedirect(redirectUrl, nextOptions, null, maxRedirects - 1));
        }

        let body = '';
        res.on('data', chunk => { body += chunk; });
        res.on('end', () => {
          resolve({
            statusCode: res.statusCode,
            headers: res.headers,
            body: body
          });
        });
      });

      req.on('error', (err) => reject(err));
      req.on('timeout', () => {
        req.destroy();
        reject(new Error('Request timed out'));
      });

      if (postData) {
        req.write(postData);
      }
      req.end();

    } catch (err) {
      reject(err);
    }
  });
}

// Fixed permanent Master Database URLs - auto-connected on ALL devices and deployments
const MASTER_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbx_luyKLX-eXPnOibeVsWLxaZps0nC_K_FGEEFKKFEE8FhIt-YKaxmD29fgbYccXdcM/exec';
const MASTER_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1pqBKWRtoJJJX9JE5pgoEevwkthjicMLLnowEbFCPihY/edit';

class GoogleSheetsSync {
  constructor(configFilePath) {
    this.configFilePath = configFilePath;
    this.isSyncing = false;
    this.lastPullTime = 0;
    this.cacheTTLMs = 25000; // 25 seconds
  }

  getConfig() {
    let cfg = {
      enabled: true,
      webAppUrl: MASTER_WEB_APP_URL,
      spreadsheetUrl: MASTER_SPREADSHEET_URL,
      spreadsheetName: 'IIT Database Master',
      autoSync: true,
      lastSyncTime: null,
      lastSyncStatus: null,
      lastSyncCounts: null
    };

    try {
      if (fs.existsSync(this.configFilePath)) {
        const fileContent = JSON.parse(fs.readFileSync(this.configFilePath, 'utf8'));
        cfg = { ...cfg, ...fileContent };
      }
    } catch (e) {
      console.error('[GoogleSheets] Error reading config file:', e);
    }

    // Check environment variables as production / cloud override (e.g. Vercel)
    const envUrl = process.env.GOOGLE_SHEETS_WEB_APP_URL || process.env.GSHEET_WEB_APP_URL || '';
    if (envUrl.trim()) {
      cfg.webAppUrl = envUrl.trim();
    }
    const envSheet = process.env.GOOGLE_SHEET_URL || process.env.SPREADSHEET_URL || '';
    if (envSheet.trim()) {
      cfg.spreadsheetUrl = envSheet.trim();
    }

    // Always ensure fixed master URL is active if empty
    if (!cfg.webAppUrl || !cfg.webAppUrl.trim()) {
      cfg.webAppUrl = MASTER_WEB_APP_URL;
    }
    if (!cfg.spreadsheetUrl || !cfg.spreadsheetUrl.trim()) {
      cfg.spreadsheetUrl = MASTER_SPREADSHEET_URL;
    }
    cfg.enabled = true;
    cfg.autoSync = true;

    return cfg;
  }

  isConfigured() {
    const config = this.getConfig();
    return Boolean(config.webAppUrl && config.webAppUrl.trim());
  }

  shouldRefresh() {
    return Date.now() - this.lastPullTime > this.cacheTTLMs;
  }

  saveConfig(newConfig) {
    try {
      const current = this.getConfig();
      const updated = { ...current, ...newConfig };
      if (updated.webAppUrl && updated.webAppUrl.trim()) {
        updated.enabled = true;
      }
      const dir = path.dirname(this.configFilePath);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(this.configFilePath, JSON.stringify(updated, null, 2), 'utf8');
      return updated;
    } catch (e) {
      console.error('[GoogleSheets] Error saving config:', e);
      throw e;
    }
  }

  getScriptCode() {
    try {
      return fs.readFileSync(path.join(__dirname, 'google-apps-script.js'), 'utf8');
    } catch (e) {
      return '';
    }
  }

  /**
   * Unified request sender using native fetch or requestWithRedirect
   */
  async sendRequest(url, payload) {
    const postData = JSON.stringify(payload);

    // Prefer native global fetch (Node 18+)
    if (typeof fetch === 'function') {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: postData,
        redirect: 'follow'
      });

      const text = await response.text();
      try {
        return JSON.parse(text);
      } catch (e) {
        throw new Error(`Invalid response from Google Sheets: ${text.slice(0, 300)}`);
      }
    }

    // Fallback using requestWithRedirect
    const response = await requestWithRedirect(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
        'Content-Length': Buffer.byteLength(postData)
      },
      timeout: 35000
    }, postData);

    try {
      return JSON.parse(response.body);
    } catch (e) {
      throw new Error(`Invalid response from Google Sheets: ${response.body.slice(0, 300)}`);
    }
  }

  /**
   * Tests connection to Google Apps Script Web App
   */
  async testConnection(targetUrl, targetSpreadsheetUrl) {
    const config = this.getConfig();
    const url = (targetUrl || config.webAppUrl || '').trim();
    const sheetUrl = (targetSpreadsheetUrl !== undefined ? targetSpreadsheetUrl : config.spreadsheetUrl) || '';

    if (!url) {
      return { success: false, error: 'Google Apps Script Web App URL is required' };
    }

    try {
      const payload = {
        action: 'ping',
        spreadsheetUrl: sheetUrl
      };

      const parsed = await this.sendRequest(url, payload);

      if (parsed.success) {
        return {
          success: true,
          message: 'Connected successfully to Google Sheet!',
          spreadsheetName: parsed.spreadsheetName || 'IIT Database Master',
          spreadsheetUrl: parsed.spreadsheetUrl || sheetUrl,
          sheets: parsed.sheets || []
        };
      } else {
        return { success: false, error: parsed.error || 'Unknown error from Google Apps Script' };
      }
    } catch (err) {
      return { success: false, error: err.message || 'Connection failed' };
    }
  }

  /**
   * Pushes all 3 entities (Staff, Students, Addresses) to Google Sheets
   */
  async pushToGoogleSheets(submissions = [], staff = [], addresses = [], options = {}) {
    const config = this.getConfig();
    if (!config.webAppUrl || !config.webAppUrl.trim()) {
      return { success: false, error: 'Google Apps Script Web App URL is not configured' };
    }

    if (this.isSyncing) {
      return { success: false, error: 'Sync is currently in progress' };
    }

    this.isSyncing = true;
    try {
      const formattedStaff = excelDb.formatStaffForSheet(staff);
      const formattedStudents = excelDb.formatStudentsForSheet(submissions);
      const formattedAddresses = excelDb.formatAddressesForSheet(addresses);

      const payload = {
        action: 'sync',
        users: formattedStudents,
        students: formattedStudents,
        staff: formattedStaff,
        addresses: formattedAddresses,
        force: !!options.force,
        spreadsheetUrl: config.spreadsheetUrl || '',
        clientTimestamp: new Date().toISOString()
      };

      const result = await this.sendRequest(config.webAppUrl, payload);

      if (result.success) {
        this.saveConfig({
          lastSyncTime: new Date().toISOString(),
          lastSyncStatus: 'Success',
          spreadsheetName: result.spreadsheetName || config.spreadsheetName || 'IIT Database Master',
          spreadsheetUrl: result.spreadsheetUrl || config.spreadsheetUrl,
          lastSyncCounts: {
            staff: formattedStaff.length,
            students: formattedStudents.length,
            addresses: formattedAddresses.length
          }
        });

        console.log(`[GoogleSheets] Synced successfully: ${formattedStaff.length} Staff, ${formattedStudents.length} Students, ${formattedAddresses.length} Addresses`);
        return {
          success: true,
          message: 'All 3 sheets synchronized successfully with Google Drive',
          spreadsheetName: result.spreadsheetName,
          spreadsheetUrl: result.spreadsheetUrl,
          counts: result.counts || {
            staff: formattedStaff.length,
            students: formattedStudents.length,
            addresses: formattedAddresses.length
          }
        };
      } else {
        throw new Error(result.error || 'Failed to update Google Sheet');
      }
    } catch (err) {
      console.error('[GoogleSheets] Sync error:', err);
      this.saveConfig({
        lastSyncTime: new Date().toISOString(),
        lastSyncStatus: 'Failed: ' + err.message
      });
      return { success: false, error: err.message };
    } finally {
      this.isSyncing = false;
    }
  }

  /**
   * Fast single-record upsert for a student submission
   */
  async upsertStudent(submission) {
    const config = this.getConfig();
    if (!config.webAppUrl || !config.webAppUrl.trim()) return { success: false };

    try {
      const formattedList = excelDb.formatStudentsForSheet([submission]);
      const formattedUser = formattedList[0];
      const payload = {
        action: 'upsert_user',
        user: formattedUser,
        student: formattedUser,
        spreadsheetUrl: config.spreadsheetUrl || ''
      };

      const result = await this.sendRequest(config.webAppUrl, payload);
      if (result.success) {
        this.lastPullTime = 0; // invalidate cache
        console.log(`[GoogleSheets] Student ${submission.submissionId || submission.name} permanently saved to live sheet.`);
        return { success: true, result };
      }
      return { success: false, error: result.error };
    } catch (err) {
      console.warn('[GoogleSheets] Upsert student failed:', err.message);
      return { success: false, error: err.message };
    }
  }

  /**
   * Fast single-record upsert for staff
   */
  async upsertStaff(staffMember) {
    const config = this.getConfig();
    if (!config.webAppUrl || !config.webAppUrl.trim()) return { success: false };

    try {
      const formattedList = excelDb.formatStaffForSheet([staffMember]);
      const formattedStaff = formattedList[0];
      const payload = {
        action: 'upsert_staff',
        staff: formattedStaff,
        staffMember: formattedStaff,
        spreadsheetUrl: config.spreadsheetUrl || ''
      };

      const result = await this.sendRequest(config.webAppUrl, payload);
      if (result.success) {
        this.lastPullTime = 0;
        return { success: true, result };
      }
      return { success: false, error: result.error };
    } catch (err) {
      console.warn('[GoogleSheets] Upsert staff failed:', err.message);
      return { success: false, error: err.message };
    }
  }

  /**
   * Fast single-record upsert for address
   */
  async upsertAddress(address) {
    const config = this.getConfig();
    if (!config.webAppUrl || !config.webAppUrl.trim()) return { success: false };

    try {
      const formattedList = excelDb.formatAddressesForSheet([address]);
      const formattedAddress = formattedList[0];
      const payload = {
        action: 'upsert_address',
        address: formattedAddress,
        spreadsheetUrl: config.spreadsheetUrl || ''
      };

      const result = await this.sendRequest(config.webAppUrl, payload);
      if (result.success) {
        this.lastPullTime = 0;
        return { success: true, result };
      }
      return { success: false, error: result.error };
    } catch (err) {
      console.warn('[GoogleSheets] Upsert address failed:', err.message);
      return { success: false, error: err.message };
    }
  }

  /**
   * Delete student from Google Sheet
   */
  async deleteStudent(submissionId) {
    const config = this.getConfig();
    if (!config.webAppUrl || !config.webAppUrl.trim()) return { success: false };

    try {
      const payload = {
        action: 'delete_user',
        submissionId: submissionId,
        spreadsheetUrl: config.spreadsheetUrl || ''
      };

      const result = await this.sendRequest(config.webAppUrl, payload);
      this.lastPullTime = 0;
      return result;
    } catch (err) {
      console.warn('[GoogleSheets] Delete student failed:', err.message);
      return { success: false, error: err.message };
    }
  }

  /**
   * Pulls rows from Google Sheet to restore / import into database
   */
  async pullFromGoogleSheets() {
    const config = this.getConfig();
    if (!config.webAppUrl || !config.webAppUrl.trim()) {
      return { success: false, error: 'Google Apps Script Web App URL is not configured' };
    }

    try {
      // Send pull request via POST for maximum reliability
      const payload = {
        action: 'pull',
        spreadsheetUrl: config.spreadsheetUrl || ''
      };

      const data = await this.sendRequest(config.webAppUrl, payload);

      if (!data.success) {
        throw new Error(data.error || 'Failed to pull data from Google Sheets');
      }

      this.lastPullTime = Date.now();

      // 1. Staff conversion
      const rawStaff = data.staff || [];
      const staffList = rawStaff.map(row => ({
        staffId: row['Staff ID'] || row['staffId'] || '',
        name: row['Full Name'] || row['Name'] || row['name'] || '',
        mobile: String(row['Mobile Number'] || row['Mobile'] || row['mobile'] || ''),
        email: row['Email Address'] || row['Email'] || row['email'] || '',
        department: row['Department'] || row['department'] || '',
        role: row['Designation / Role'] || row['Role'] || row['role'] || '',
        status: row['Account Status'] || row['Status'] || row['status'] || 'Active',
        createdAt: row['Date of Registration'] || row['createdAt'] || new Date().toLocaleString(),
        updatedAt: row['Last Updated'] || row['updatedAt'] || new Date().toLocaleString()
      })).filter(s => s.staffId || s.name);

      // 2. Student conversion
      const rawStudents = data.students || data.users || [];
      const studentList = rawStudents.map(row => ({
        submissionId: row['User / Submission ID'] || row['Submission ID'] || row['submissionId'] || row['User ID'] || '',
        name: row['User / Student Name'] || row['Student Name'] || row['Name'] || row['name'] || '',
        fatherName: row['Father / Guardian Name'] || row['fatherName'] || '',
        title: row['Course / Applied Title'] || row['Course'] || row['title'] || '',
        email: row['Email Address'] || row['Email'] || row['email'] || '',
        mobile1: String(row['Primary Mobile'] || row['Mobile 1'] || row['mobile1'] || ''),
        mobile2: String(row['Secondary Mobile'] || row['Mobile 2'] || row['mobile2'] || ''),
        dob: row['Date of Birth'] || row['dob'] || '',
        gender: row['Gender'] || row['gender'] || '',
        city: row['City / District'] || row['city'] || '',
        state: row['State'] || row['state'] || '',
        pincode: String(row['Pincode'] || row['pincode'] || ''),
        address: row['Full Residential Address'] || row['Address'] || row['address'] || '',
        collegeSchool: row['College / School'] || row['collegeSchool'] || '',
        yearDept: row['Department / Year'] || row['yearDept'] || '',
        qualification: row['Educational Qualification'] || row['qualification'] || '',
        aadhar: String(row['Aadhar / ID Number'] || row['aadhar'] || ''),
        feesFixed: row['Fees Fixed'] || row['feesFixed'] || '',
        dateJoining: row['Date of Joining'] || row['dateJoining'] || '',
        certNo: row['Certificate No'] || row['certNo'] || '',
        dateIssue: row['Date of Issue'] || row['dateIssue'] || '',
        submittedAt: row['Submission Date & Time'] || row['submittedAt'] || new Date().toLocaleString()
      })).filter(s => s.submissionId || s.name);

      // 3. Address conversion
      const rawAddresses = data.addresses || [];
      const addressList = rawAddresses.map(row => ({
        id: row['Address ID'] || row['id'] || '',
        dateTime: row['Date & Time'] || row['dateTime'] || new Date().toISOString(),
        to: {
          name: row['To: Recipient Name'] || '',
          line1: row['To: Address Line 1'] || '',
          line2: row['To: Address Line 2'] || '',
          line3: row['To: Address Line 3'] || '',
          city: row['To: City / District'] || '',
          state: row['To: State'] || '',
          pincode: String(row['To: Pincode'] || ''),
          phone: String(row['To: Primary Phone'] || ''),
          extraPhone: String(row['To: Extra Phone'] || ''),
          country: row['To: Country'] || 'India'
        },
        from: {
          name: row['From: Sender / Org Name'] || 'Indian Info Tech',
          line1: row['From: Address Line 1'] || '101-C, KM Complex,',
          line2: row['From: Address Line 2'] || 'Amma Unavagam Back Side,',
          line3: row['From: Address Line 3'] || '',
          city: row['From: City'] || 'Namakkal',
          state: row['From: State'] || 'Tamilnadu',
          pincode: String(row['From: Pincode'] || '637001'),
          phone: String(row['From: Phone'] || '9003446612'),
          extraPhone: String(row['From: Extra Phone'] || ''),
          country: row['From: Country'] || 'India'
        }
      })).filter(a => a.id || (a.to && a.to.name));

      // Update cached info
      this.saveConfig({
        spreadsheetName: data.spreadsheetName || config.spreadsheetName || 'IIT Database Master',
        spreadsheetUrl: data.spreadsheetUrl || config.spreadsheetUrl
      });

      return {
        success: true,
        spreadsheetName: data.spreadsheetName,
        spreadsheetUrl: data.spreadsheetUrl,
        staff: staffList,
        students: studentList,
        addresses: addressList
      };

    } catch (err) {
      console.error('[GoogleSheets] Pull error:', err);
      return { success: false, error: err.message };
    }
  }

  /**
   * Background auto-sync trigger
   */
  triggerAutoSync(submissionsGetter, staffGetter, addressesGetter) {
    const config = this.getConfig();
    if (!config.enabled || !config.webAppUrl) {
      return;
    }

    setImmediate(async () => {
      try {
        const subs = typeof submissionsGetter === 'function' ? submissionsGetter() : submissionsGetter;
        const stf = typeof staffGetter === 'function' ? staffGetter() : staffGetter;
        const addr = typeof addressesGetter === 'function' ? addressesGetter() : addressesGetter;
        await this.pushToGoogleSheets(subs, stf, addr);
      } catch (err) {
        console.warn('[GoogleSheets] Background auto-sync failed:', err.message);
      }
    });
  }
}

module.exports = GoogleSheetsSync;
