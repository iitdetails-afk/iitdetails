const fs = require('fs');
const path = require('path');

let XLSX;
try {
  XLSX = require('xlsx');
} catch (e) {
  try {
    XLSX = require(path.join(__dirname, 'node_modules', 'xlsx'));
  } catch (err) {
    console.warn('[ExcelDB] Warning: xlsx module not found, fallback will be used');
  }
}

/**
 * Format Student (Submissions) records for the Excel sheet
 */
function formatStudentsForSheet(submissions = []) {
  return submissions.map((s, idx) => ({
    'S.No': idx + 1,
    'User / Submission ID': s.submissionId || `ID-${String(idx + 1).padStart(4, '0')}`,
    'User / Student Name': s.name || '',
    'Father / Guardian Name': s.fatherName || '',
    'Course / Applied Title': s.title || '',
    'Email Address': s.email || '',
    'Primary Mobile': s.mobile1 || '',
    'Secondary Mobile': s.mobile2 || '',
    'Date of Birth': s.dob || '',
    'Gender': s.gender || '',
    'City / District': s.city || '',
    'State': s.state || '',
    'Pincode': s.pincode || '',
    'Full Residential Address': s.address || '',
    'College / School': s.collegeSchool || '',
    'Department / Year': s.yearDept || '',
    'Educational Qualification': s.qualification || '',
    'Aadhar / ID Number': s.aadhar || '',
    'Fees Fixed': s.feesFixed || '',
    'Date of Joining': s.dateJoining || '',
    'Certificate No': s.certNo || '',
    'Date of Issue': s.dateIssue || '',
    'Submission Date & Time': s.submittedAt || ''
  }));
}

/**
 * Format Staff records for the Excel sheet
 */
function formatStaffForSheet(staffList = []) {
  return staffList.map((st, idx) => ({
    'S.No': idx + 1,
    'Staff ID': st.staffId || `STF-${idx + 1}`,
    'Full Name': st.name || '',
    'Mobile Number': st.mobile || '',
    'Email Address': st.email || '',
    'Department': st.department || '',
    'Designation / Role': st.role || '',
    'Account Status': st.status || 'Active',
    'Date of Registration': st.createdAt || '',
    'Last Updated': st.updatedAt || ''
  }));
}

/**
 * Format Address records for the Excel sheet
 */
function formatAddressesForSheet(addresses = []) {
  return addresses.map((addr, idx) => {
    // Handle both nested to/from format and flat format
    const to = addr.to || {};
    const from = addr.from || {};
    return {
      'S.No': idx + 1,
      'Address ID': addr.id || `ADDR-${idx + 1}`,
      'Date & Time': addr.dateTime || addr.createdAt || '',
      // TO Section
      'To: Recipient Name': to.name || addr.name || '',
      'To: Address Line 1': to.line1 || addr.addressLine1 || '',
      'To: Address Line 2': to.line2 || addr.addressLine2 || '',
      'To: Address Line 3': to.line3 || '',
      'To: City / District': to.city || addr.city || '',
      'To: State': to.state || addr.state || '',
      'To: Pincode': to.pincode || addr.pincode || '',
      'To: Primary Phone': to.phone || addr.phone || '',
      'To: Extra Phone': to.extraPhone || '',
      'To: Country': to.country || 'India',
      // FROM Section
      'From: Sender / Org Name': from.name || 'Indian Info Tech',
      'From: Address Line 1': from.line1 || '101-C, KM Complex,',
      'From: Address Line 2': from.line2 || 'Amma Unavagam Back Side,',
      'From: Address Line 3': from.line3 || '',
      'From: City': from.city || 'Namakkal',
      'From: State': from.state || 'Tamilnadu',
      'From: Pincode': from.pincode || '637001',
      'From: Phone': from.phone || '9003446612',
      'From: Extra Phone': from.extraPhone || '',
      'From: Country': from.country || 'India'
    };
  });
}

/**
 * Compute auto column widths for a worksheet based on data
 */
function autoColumnWidths(data, defaultWidth = 15) {
  if (!data || !data.length) return [];
  const keys = Object.keys(data[0]);
  return keys.map(key => {
    let maxLen = key.length;
    for (let i = 0; i < Math.min(data.length, 100); i++) {
      const val = data[i][key];
      if (val !== undefined && val !== null) {
        const len = String(val).length;
        if (len > maxLen) maxLen = len;
      }
    }
    return { wch: Math.min(Math.max(maxLen + 3, defaultWidth), 45) };
  });
}

/**
 * Builds the 3-sheet workbook in memory
 */
function buildWorkbook(submissions = [], staffList = [], addresses = []) {
  if (!XLSX) {
    throw new Error('XLSX module is not available');
  }

  const wb = XLSX.utils.book_new();

  // 1. Sheet 1: User Details
  const studentData = formatStudentsForSheet(submissions);
  const studentSheet = studentData.length
    ? XLSX.utils.json_to_sheet(studentData)
    : XLSX.utils.aoa_to_sheet([['S.No', 'User / Submission ID', 'User / Student Name', 'Father / Guardian Name', 'Course / Applied Title', 'Email Address', 'Primary Mobile', 'Secondary Mobile', 'Date of Birth', 'Gender', 'City / District', 'State', 'Pincode', 'Full Residential Address', 'College / School', 'Department / Year', 'Educational Qualification', 'Aadhar / ID Number', 'Fees Fixed', 'Date of Joining', 'Certificate No', 'Date of Issue', 'Submission Date & Time']]);
  studentSheet['!cols'] = autoColumnWidths(studentData);
  XLSX.utils.book_append_sheet(wb, studentSheet, 'User Details');

  // 2. Sheet 2: Staff Details
  const staffData = formatStaffForSheet(staffList);
  const staffSheet = staffData.length 
    ? XLSX.utils.json_to_sheet(staffData)
    : XLSX.utils.aoa_to_sheet([['S.No', 'Staff ID', 'Full Name', 'Mobile Number', 'Email Address', 'Department', 'Designation / Role', 'Account Status', 'Date of Registration', 'Last Updated']]);
  staffSheet['!cols'] = autoColumnWidths(staffData);
  XLSX.utils.book_append_sheet(wb, staffSheet, 'Staff Details');

  // 3. Sheet 3: Address Details
  const addressData = formatAddressesForSheet(addresses);
  const addressSheet = addressData.length
    ? XLSX.utils.json_to_sheet(addressData)
    : XLSX.utils.aoa_to_sheet([['S.No', 'Address ID', 'Date & Time', 'To: Recipient Name', 'To: Address Line 1', 'To: Address Line 2', 'To: Address Line 3', 'To: City / District', 'To: State', 'To: Pincode', 'To: Primary Phone', 'To: Extra Phone', 'To: Country', 'From: Sender / Org Name', 'From: Address Line 1', 'From: Address Line 2', 'From: Address Line 3', 'From: City', 'From: State', 'From: Pincode', 'From: Phone', 'From: Extra Phone', 'From: Country']]);
  addressSheet['!cols'] = autoColumnWidths(addressData);
  XLSX.utils.book_append_sheet(wb, addressSheet, 'Address Details');

  return wb;
}

/**
 * Saves all 3 data arrays to the permanent Excel file
 */
function saveToExcelFile(filePath, submissions = [], staffList = [], addresses = []) {
  try {
    if (!XLSX) return false;
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    const wb = buildWorkbook(submissions, staffList, addresses);
    XLSX.writeFile(wb, filePath);
    console.log(`[ExcelDB] Master Excel database saved successfully at ${filePath} (Staff: ${staffList.length}, Students: ${submissions.length}, Addresses: ${addresses.length})`);
    return true;
  } catch (err) {
    console.error('[ExcelDB] Error saving Excel database:', err);
    return false;
  }
}

/**
 * Generates Excel buffer for HTTP download
 */
function getExcelBuffer(submissions = [], staffList = [], addresses = []) {
  if (!XLSX) return null;
  const wb = buildWorkbook(submissions, staffList, addresses);
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

/**
 * Parses an uploaded/existing Excel database file into 3 data arrays
 */
function readFromExcelFile(filePathOrBuffer) {
  if (!XLSX) return null;
  try {
    let wb;
    if (Buffer.isBuffer(filePathOrBuffer)) {
      wb = XLSX.read(filePathOrBuffer, { type: 'buffer' });
    } else {
      if (!fs.existsSync(filePathOrBuffer)) return null;
      wb = XLSX.readFile(filePathOrBuffer);
    }

    const result = {
      staff: [],
      submissions: [],
      addresses: []
    };

    // Find sheets case-insensitively
    const sheetNames = wb.SheetNames;
    const staffSheetName = sheetNames.find(n => n.toLowerCase().includes('staff'));
    const studentSheetName = sheetNames.find(n => n.toLowerCase().includes('user') || n.toLowerCase().includes('student') || n.toLowerCase().includes('submission'));
    const addressSheetName = sheetNames.find(n => n.toLowerCase().includes('address'));

    // 1. Parse Staff
    if (staffSheetName) {
      const rawStaff = XLSX.utils.sheet_to_json(wb.Sheets[staffSheetName]);
      result.staff = rawStaff.map(row => ({
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
    }

    // 2. Parse User / Student details
    if (studentSheetName) {
      const rawStudents = XLSX.utils.sheet_to_json(wb.Sheets[studentSheetName]);
      result.submissions = rawStudents.map(row => ({
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
    }

    // 3. Parse Addresses
    if (addressSheetName) {
      const rawAddresses = XLSX.utils.sheet_to_json(wb.Sheets[addressSheetName]);
      result.addresses = rawAddresses.map(row => ({
        id: row['Address ID'] || row['id'] || `ADDR-${Date.now()}-${Math.floor(Math.random() * 900)}`,
        dateTime: row['Date & Time'] || row['dateTime'] || new Date().toLocaleString(),
        to: {
          name: row['To: Recipient Name'] || row['To Name'] || row['name'] || '',
          line1: row['To: Address Line 1'] || row['To Line 1'] || row['addressLine1'] || '',
          line2: row['To: Address Line 2'] || row['To Line 2'] || row['addressLine2'] || '',
          line3: row['To: Address Line 3'] || row['To Line 3'] || '',
          city: row['To: City / District'] || row['To City'] || row['city'] || '',
          state: row['To: State'] || row['To State'] || row['state'] || '',
          pincode: String(row['To: Pincode'] || row['To Pincode'] || row['pincode'] || ''),
          phone: String(row['To: Primary Phone'] || row['To Phone'] || row['phone'] || ''),
          extraPhone: String(row['To: Extra Phone'] || row['To Extra Phone'] || ''),
          country: row['To: Country'] || 'India'
        },
        from: {
          name: row['From: Sender / Org Name'] || row['From Name'] || 'Indian Info Tech',
          line1: row['From: Address Line 1'] || row['From Line 1'] || '101-C, KM Complex,',
          line2: row['From: Address Line 2'] || row['From Line 2'] || 'Amma Unavagam Back Side,',
          line3: row['From: Address Line 3'] || row['From Line 3'] || '',
          city: row['From: City'] || 'Namakkal',
          state: row['From: State'] || 'Tamilnadu',
          pincode: String(row['From: Pincode'] || '637001'),
          phone: String(row['From: Phone'] || '9003446612'),
          extraPhone: String(row['From: Extra Phone'] || ''),
          country: row['From: Country'] || 'India'
        }
      })).filter(a => a.id || a.to.name);
    }

    return result;
  } catch (err) {
    console.error('[ExcelDB] Error reading Excel file:', err);
    return null;
  }
}

module.exports = {
  formatStudentsForSheet,
  formatStaffForSheet,
  formatAddressesForSheet,
  buildWorkbook,
  saveToExcelFile,
  getExcelBuffer,
  readFromExcelFile
};
