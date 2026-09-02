// ============================================================
// BEC PLACEMENT PORTAL - GOOGLE APPS SCRIPT BACKEND
// ============================================================

const SPREADSHEET_ID = '10cWHoio8nwJJJep6Gng3yCVft0mkV6rEs9Gkm5tSdOc';

function doGet() {
  return ContentService
    .createTextOutput("BEC Placement Portal API is running. Use POST requests.")
    .setMimeType(ContentService.MimeType.TEXT);
}
// ============================================================
// SETUP DATABASE (Run once to create sheets)
// ============================================================
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = {
    Students: ['USN', 'Name', 'Email', 'Branch', 'Semester', 'CGPA', 'Backlogs', 'Resume', 'Frozen', 'ApprovedByHOD', 'Category', 'Mobile', 'Skills'],
    Companies: ['CompanyID', 'CompanyName', 'Industry', 'Location', 'Website', 'Contact', 'Email', 'Status'],
    Drives: ['DriveID', 'CompanyID', 'Company', 'Role', 'Package', 'Deadline', 'Location', 'Branches', 'Status'],
    Applications: ['ApplicationID', 'USN', 'DriveID', 'Company', 'Role', 'AppliedDate', 'Status'],
    Placements: ['PlacementID', 'USN', 'StudentName', 'Company', 'Role', 'Package', 'Date'],
    Notifications: ['NotificationID', 'Title', 'Message', 'TargetRole', 'Date', 'Status'],
    PlacementResults: ['ResultID', 'USN', 'StudentName', 'Company', 'Role', 'Package', 'Status', 'Date'],
    Config: ['Key', 'Value'],
    Users: ['UserId', 'Username', 'Email', 'Password', 'Role', 'Department', 'Status', 'CreatedAt']
  };

  Object.keys(sheets).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clearContents();
    }
    sheet.getRange(1, 1, 1, sheets[sheetName].length).setValues([sheets[sheetName]]);
    sheet.getRange(1, 1, 1, sheets[sheetName].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, sheets[sheetName].length);
  });

  // Insert default config
  const configSheet = ss.getSheetByName('Config');
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['CollegeName', 'Basaveshwar Engineering College, Bagalkote']);
    configSheet.appendRow(['AcademicYear', '2026-27']);
  }

  // Insert default admin user (if not exists)
  const userSheet = ss.getSheetByName('Users');
  const users = userSheet.getDataRange().getValues();
  let adminExists = false;
  for (let i = 1; i < users.length; i++) {
    if (users[i][3] === 'mmbbec@gmail.com') { // email column
      adminExists = true;
      break;
    }
  }
  if (!adminExists) {
    userSheet.appendRow([
      'ADMIN001',
      'Admin',
      'mmbbec@gmail.com',
      'Mmb@1980',
      'admin',
      '',
      'Active',
      new Date().toISOString()
    ]);
  }

  Logger.log('Database setup completed.');
}

// ============================================================
// DO POST - Main API endpoint
// ============================================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, message: 'No data received' });
    }
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    switch (action) {
      // Student operations
      case 'addStudent':        return addStudent(data);
      case 'getStudent':        return getStudent(data);
      case 'getStudents':       return getStudents(data);
      case 'updateStudent':     return updateStudent(data);
      // Company operations
      case 'addCompany':        return addCompany(data);
      case 'getCompanies':      return getCompanies(data);
      // Drive operations
      case 'addDrive':          return addDrive(data);
      case 'getDrives':         return getDrives(data);
      // Application operations
      case 'addApplication':    return addApplication(data);
      case 'getApplications':   return getApplications(data);
      case 'updateApplication': return updateApplication(data);
      // Notification operations
      case 'addNotification':   return addNotification(data);
      case 'getNotifications':  return getNotifications(data);
      // Placement result operations
      case 'addPlacementResult':return addPlacementResult(data);
      case 'getPlacementResults':return getPlacementResults(data);
      // Config operations
      case 'getConfig':         return getConfig(data);
      case 'updateConfig':      return updateConfig(data);
      // Student approvals (legacy, kept for compatibility)
      case 'getPendingStudents':return getPendingStudents(data);
      case 'approveStudent':    return approveStudent(data);
      // User Management (NEW)
      case 'getUsers':          return getUsers(data);
      case 'addUser':           return addUser(data);
      case 'updateUser':        return updateUser(data);
      case 'deleteUser':        return deleteUser(data);
      default:
        return jsonResponse({ success: false, message: 'Unknown action: ' + action });
    }
  } catch (error) {
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ============================================================
// CONFIG OPERATIONS
// ============================================================
function getConfig(data) {
  const sheet = getSheet('Config');
  const values = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < values.length; i++) {
    config[values[i][0]] = values[i][1];
  }
  return jsonResponse({ success: true, config });
}

function updateConfig(data) {
  const sheet = getSheet('Config');
  const key = data.key;
  const value = data.value;
  if (!key) return jsonResponse({ success: false, message: 'Key required' });
  const rows = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i+1, 2).setValue(value);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([key, value]);
  }
  return jsonResponse({ success: true, message: 'Config updated' });
}

// ============================================================
// STUDENT OPERATIONS (unchanged)
// ============================================================
function addStudent(data) {
  const sheet = getSheet('Students');
  const usn = (data.usn || '').trim();
  if (!usn) return jsonResponse({ success: false, message: 'USN is required' });
  if (findRow(sheet, 1, usn)) {
    return jsonResponse({ success: false, message: 'Student already exists' });
  }
  sheet.appendRow([
    usn,
    data.name || '',
    data.email || '',
    data.branch || '',
    data.semester || '',
    data.cgpa || '',
    data.backlogs || '',
    data.resume || '',
    data.frozen || false,
    data.approvedByHod || false,
    data.category || '',
    data.mobile || '',
    data.skills || ''
  ]);
  return jsonResponse({ success: true, message: 'Student saved', usn });
}

function getStudent(data) {
  const sheet = getSheet('Students');
  const usn = (data.usn || '').trim();
  if (!usn) return jsonResponse({ success: false, message: 'USN required' });
  const row = findRow(sheet, 1, usn);
  if (!row) return jsonResponse({ success: false, message: 'Student not found' });
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return jsonResponse({
    success: true,
    student: {
      usn: values[0],
      name: values[1],
      email: values[2],
      branch: values[3],
      semester: values[4],
      cgpa: values[5],
      backlogs: values[6],
      resume: values[7],
      frozen: values[8],
      approvedByHod: values[9],
      category: values[10],
      mobile: values[11],
      skills: values[12]
    }
  });
}

function getStudents(data) {
  const sheet = getSheet('Students');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length < 2) return jsonResponse({ success: true, students: [] });
  const students = [];
  for (let i = 1; i < values.length; i++) {
    students.push({
      usn: values[i][0],
      name: values[i][1],
      email: values[i][2],
      branch: values[i][3],
      semester: values[i][4],
      cgpa: values[i][5],
      backlogs: values[i][6],
      resume: values[i][7],
      frozen: values[i][8],
      approvedByHod: values[i][9],
      category: values[i][10],
      mobile: values[i][11],
      skills: values[i][12],
      placed: false // will be determined by placement records
    });
  }
  return jsonResponse({ success: true, students });
}

function updateStudent(data) {
  const sheet = getSheet('Students');
  const usn = (data.usn || '').trim();
  if (!usn) return jsonResponse({ success: false, message: 'USN required' });
  const row = findRow(sheet, 1, usn);
  if (!row) return jsonResponse({ success: false, message: 'Student not found' });
  const fields = ['name','email','branch','semester','cgpa','backlogs','resume','frozen','approvedByHod','category','mobile','skills'];
  const colMap = { name:2, email:3, branch:4, semester:5, cgpa:6, backlogs:7, resume:8, frozen:9, approvedByHod:10, category:11, mobile:12, skills:13 };
  for (let key in data) {
    if (colMap[key]) {
      sheet.getRange(row, colMap[key]).setValue(data[key]);
    }
  }
  return jsonResponse({ success: true, message: 'Student updated' });
}

function getPendingStudents(data) {
  const sheet = getSheet('Students');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, students: [] });
  const pending = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][8] === true && values[i][9] !== true) {
      pending.push({
        usn: values[i][0],
        name: values[i][1],
        email: values[i][2],
        branch: values[i][3],
        semester: values[i][4],
        cgpa: values[i][5],
        backlogs: values[i][6],
        category: values[i][10],
        mobile: values[i][11],
        skills: values[i][12]
      });
    }
  }
  return jsonResponse({ success: true, students: pending });
}

function approveStudent(data) {
  const sheet = getSheet('Students');
  const usn = (data.usn || '').trim();
  if (!usn) return jsonResponse({ success: false, message: 'USN required' });
  const row = findRow(sheet, 1, usn);
  if (!row) return jsonResponse({ success: false, message: 'Student not found' });
  sheet.getRange(row, 10).setValue(true);
  return jsonResponse({ success: true, message: 'Student approved' });
}

// ============================================================
// COMPANY, DRIVE, APPLICATION, PLACEMENT, NOTIFICATION, RESULT
// (unchanged from original)
// ============================================================
function addCompany(data) {
  const sheet = getSheet('Companies');
  const companyId = data.companyId || 'COMP-' + Date.now();
  sheet.appendRow([
    companyId,
    data.companyName || '',
    data.industry || '',
    data.location || '',
    data.website || '',
    data.contact || '',
    data.email || '',
    data.status || 'Active'
  ]);
  return jsonResponse({ success: true, message: 'Company saved', companyId });
}

function getCompanies(data) {
  const sheet = getSheet('Companies');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, companies: [] });
  const companies = [];
  for (let i = 1; i < values.length; i++) {
    companies.push({
      companyId: values[i][0],
      companyName: values[i][1],
      industry: values[i][2],
      location: values[i][3],
      website: values[i][4],
      contact: values[i][5],
      email: values[i][6],
      status: values[i][7]
    });
  }
  return jsonResponse({ success: true, companies });
}

function addDrive(data) {
  const sheet = getSheet('Drives');
  const driveId = data.driveId || 'DRIVE-' + Date.now();
  sheet.appendRow([
    driveId,
    data.companyId || '',
    data.company || '',
    data.role || '',
    data.package || '',
    data.deadline || '',
    data.location || '',
    data.branches || '',
    data.status || 'Registration Open'
  ]);
  return jsonResponse({ success: true, message: 'Drive saved', driveId });
}

function getDrives(data) {
  const sheet = getSheet('Drives');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, drives: [] });
  const drives = [];
  for (let i = 1; i < values.length; i++) {
    drives.push({
      driveId: values[i][0],
      companyId: values[i][1],
      company: values[i][2],
      role: values[i][3],
      package: values[i][4],
      deadline: values[i][5],
      location: values[i][6],
      branches: values[i][7],
      status: values[i][8]
    });
  }
  return jsonResponse({ success: true, drives });
}

function addApplication(data) {
  const sheet = getSheet('Applications');
  const usn = (data.usn || '').trim();
  const driveId = (data.driveId || '').trim();
  if (!usn || !driveId) return jsonResponse({ success: false, message: 'USN and Drive ID required' });
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === usn && String(values[i][2]) === driveId) {
      return jsonResponse({ success: false, message: 'Already applied' });
    }
  }
  const applicationId = 'APP-' + Date.now();
  sheet.appendRow([
    applicationId,
    usn,
    driveId,
    data.company || '',
    data.role || '',
    new Date().toISOString().split('T')[0],
    'Applied'
  ]);
  return jsonResponse({ success: true, message: 'Application submitted', applicationId });
}

function getApplications(data) {
  const sheet = getSheet('Applications');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, applications: [] });
  const usnFilter = (data.usn || '').trim();
  const apps = [];
  for (let i = 1; i < values.length; i++) {
    const rowUsn = String(values[i][1]).trim();
    if (usnFilter && rowUsn !== usnFilter) continue;
    apps.push({
      applicationId: values[i][0],
      usn: values[i][1],
      driveId: values[i][2],
      company: values[i][3],
      role: values[i][4],
      appliedDate: values[i][5],
      status: values[i][6]
    });
  }
  return jsonResponse({ success: true, applications: apps });
}

function updateApplication(data) {
  const sheet = getSheet('Applications');
  const appId = (data.applicationId || '').trim();
  if (!appId) return jsonResponse({ success: false, message: 'Application ID required' });
  const row = findRow(sheet, 1, appId);
  if (!row) return jsonResponse({ success: false, message: 'Application not found' });
  if (data.status) sheet.getRange(row, 7).setValue(data.status);
  return jsonResponse({ success: true, message: 'Application updated' });
}

function addNotification(data) {
  const sheet = getSheet('Notifications');
  const notificationId = 'NOTIF-' + Date.now();
  sheet.appendRow([
    notificationId,
    data.title || '',
    data.message || '',
    data.targetRole || 'all',
    new Date().toISOString().split('T')[0],
    data.status || 'Active'
  ]);
  return jsonResponse({ success: true, message: 'Notification saved', notificationId });
}

function getNotifications(data) {
  const sheet = getSheet('Notifications');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, notifications: [] });
  const notifs = [];
  for (let i = 1; i < values.length; i++) {
    notifs.push({
      notificationId: values[i][0],
      title: values[i][1],
      message: values[i][2],
      targetRole: values[i][3],
      date: values[i][4],
      status: values[i][5]
    });
  }
  return jsonResponse({ success: true, notifications: notifs });
}

function addPlacementResult(data) {
  const sheet = getSheet('PlacementResults');
  const resultId = 'RES-' + Date.now();
  const usn = data.usn || '';
  let studentName = data.studentName || '';
  if (!studentName && usn) {
    const studentSheet = getSheet('Students');
    const row = findRow(studentSheet, 1, usn);
    if (row) {
      const vals = studentSheet.getRange(row, 1, 1, 2).getValues()[0];
      studentName = vals[1] || '';
    }
  }
  sheet.appendRow([
    resultId,
    usn,
    studentName,
    data.company || '',
    data.role || '',
    data.package || '',
    data.status || 'Selected',
    data.date || new Date().toISOString().split('T')[0]
  ]);
  return jsonResponse({ success: true, message: 'Placement result saved', resultId });
}

function getPlacementResults(data) {
  const sheet = getSheet('PlacementResults');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, results: [] });
  const results = [];
  for (let i = 1; i < values.length; i++) {
    results.push({
      resultId: values[i][0],
      usn: values[i][1],
      studentName: values[i][2],
      company: values[i][3],
      role: values[i][4],
      package: values[i][5],
      status: values[i][6],
      date: values[i][7]
    });
  }
  return jsonResponse({ success: true, results });
}

// ============================================================
// USER MANAGEMENT (NEW)
// ============================================================

// Get all users
function getUsers(data) {
  const sheet = getSheet('Users');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse({ success: true, users: [] });
  const users = [];
  for (let i = 1; i < values.length; i++) {
    users.push({
      userId: values[i][0],
      username: values[i][1],
      email: values[i][2],
      password: values[i][3], // In production, you'd hash this; for demo we keep plain
      role: values[i][4],
      department: values[i][5] || '',
      status: values[i][6] || 'Active',
      createdAt: values[i][7] || ''
    });
  }
  return jsonResponse({ success: true, users });
}

// Add a new user
function addUser(data) {
  const sheet = getSheet('Users');
  const userId = (data.userId || '').trim();
  const email = (data.email || '').trim();
  if (!userId || !email) return jsonResponse({ success: false, message: 'UserId and Email are required' });
  // Check duplicates
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === userId) {
      return jsonResponse({ success: false, message: 'UserId already exists' });
    }
    if (String(rows[i][2]) === email) {
      return jsonResponse({ success: false, message: 'Email already registered' });
    }
  }
  // Default password if not provided
  const password = data.password || 'password123';
  const role = data.role || 'student';
  const department = data.department || '';
  const status = data.status || 'Active';
  const createdAt = new Date().toISOString();
  sheet.appendRow([userId, data.username || '', email, password, role, department, status, createdAt]);
  return jsonResponse({ success: true, message: 'User added successfully', userId });
}

// Update an existing user
function updateUser(data) {
  const sheet = getSheet('Users');
  const userId = (data.userId || '').trim();
  if (!userId) return jsonResponse({ success: false, message: 'UserId required' });
  const row = findRow(sheet, 1, userId); // column 1 = UserId
  if (!row) return jsonResponse({ success: false, message: 'User not found' });
  // Update fields: username, email, password, role, department, status
  const updates = {
    username: 2,
    email: 3,
    password: 4,
    role: 5,
    department: 6,
    status: 7
  };
  for (let key in updates) {
    if (data[key] !== undefined && data[key] !== null) {
      if (key === 'password' && data[key].trim() === '') continue; // skip empty password
      sheet.getRange(row, updates[key]).setValue(data[key]);
    }
  }
  return jsonResponse({ success: true, message: 'User updated successfully' });
}

// Delete a user
function deleteUser(data) {
  const sheet = getSheet('Users');
  const userId = (data.userId || '').trim();
  if (!userId) return jsonResponse({ success: false, message: 'UserId required' });
  const row = findRow(sheet, 1, userId);
  if (!row) return jsonResponse({ success: false, message: 'User not found' });
  // Prevent deletion of the default admin (optional)
  const email = sheet.getRange(row, 3).getValue();
  if (email === 'mmbbec@gmail.com') {
    return jsonResponse({ success: false, message: 'Cannot delete the primary admin account' });
  }
  sheet.deleteRow(row);
  return jsonResponse({ success: true, message: 'User deleted successfully' });
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  return sheet;
}

function findRow(sheet, columnIndex, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
  const values = range.getValues();
  const searchVal = String(value).trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === searchVal) {
      return i + 2;
    }
  }
  return null;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
