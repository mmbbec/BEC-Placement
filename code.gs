const REGISTRY_SPREADSHEET_ID = "YOUR_USER_REGISTRY_SPREADSHEET_ID"; // Paste your Sheet ID
const ADMIN_EMAIL = "mmbbec@gmail.com";
const ADMIN_PASS = "admin123";

// Pre-configured Mock Test Profiles (No real emails needed to test!)
const MOCK_PROFILES = {
  officer: {
    name: "Dr. S G Kambalimath",
    role: "Placement Officer",
    email: "test.officer@example.com"
  },
  principal: {
    name: "Dr. Baswaraj Hiremath",
    role: "Principal",
    email: "test.principal@example.com"
  },
  staff: {
    name: "Prof. Ramesh (Coordinator)",
    role: "Placement Staff",
    email: "test.staff@example.com"
  },
  student_eligible: {
    name: "Rahul Patil",
    role: "Student",
    email: "test.student1@example.com",
    usn: "2BA23CS045",
    branch: "CSE",
    cgpa: 8.20,
    backlogs: 0,
    currentPkg: 0
  },
  student_placed: {
    name: "Pooja Kulkarni",
    role: "Student",
    email: "test.student2@example.com",
    usn: "2BA23ME012",
    branch: "MECH",
    cgpa: 6.20,
    backlogs: 1,
    currentPkg: 4.0
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BEC Placement Portal [TEST MODE]')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Return selected test identity
function getTestSession(profileKey) {
  return { success: true, profile: MOCK_PROFILES[profileKey] || MOCK_PROFILES.officer };
}

// 2. Admin Fixed Login
function loginAsAdmin(email, password) {
  if (email.trim().toLowerCase() === ADMIN_EMAIL && password === ADMIN_PASS) {
    return { success: true, name: "System Admin", role: "Super Admin", email: ADMIN_EMAIL };
  }
  return { success: false, error: "Invalid Admin Email or Password." };
}

// 3. Post a Drive
function createRecruitmentDrive(driveData) {
  try {
    const ss = SpreadsheetApp.openById(REGISTRY_SPREADSHEET_ID);
    let driveSheet = ss.getSheetByName("Recruitment_Drives");
    if (!driveSheet) {
      driveSheet = ss.insertSheet("Recruitment_Drives");
      driveSheet.appendRow(["DriveID", "Company", "CTC_LPA", "EligibleBranches", "MinCGPA", "MaxCGPA", "MaxBacklogs", "PlacedPolicy", "MinPackageDiff", "Status", "CreatedDate"]);
      driveSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e2e8f0");
    }

    const driveId = "DRV-" + Utilities.getUuid().slice(0, 5).toUpperCase();
    driveSheet.appendRow([
      driveId,
      driveData.company.trim(),
      parseFloat(driveData.ctc) || 0,
      driveData.branches.join(", "),
      parseFloat(driveData.minCgpa) || 0,
      parseFloat(driveData.maxCgpa) || 10,
      parseInt(driveData.maxBacklogs) || 0,
      driveData.placedPolicy,
      parseFloat(driveData.minDiff) || 0,
      "ACTIVE",
      new Date().toLocaleString('en-IN')
    ]);

    return { success: true, message: `Drive for ${driveData.company} (${driveId}) published!` };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// 4. Fetch Drives Evaluated Against Current Selected Student
function getEligibleDrivesForStudent(profileKey) {
  try {
    const student = MOCK_PROFILES[profileKey] || MOCK_PROFILES.student_eligible;
    const ss = SpreadsheetApp.openById(REGISTRY_SPREADSHEET_ID);
    const driveSheet = ss.getSheetByName("Recruitment_Drives");
    const appSheet = ss.getSheetByName("Drive_Applications");

    if (!driveSheet || driveSheet.getLastRow() <= 1) {
      return { success: true, drives: [], student: student };
    }

    // Check applied list
    const appliedDrives = new Set();
    if (appSheet && appSheet.getLastRow() > 1) {
      const appData = appSheet.getDataRange().getValues();
      for (let j = 1; j < appData.length; j++) {
        if (String(appData[j][3]).toUpperCase() === String(student.usn).toUpperCase()) {
          appliedDrives.add(String(appData[j][1]));
        }
      }
    }

    const drivesData = driveSheet.getDataRange().getValues();
    const driveList = [];

    for (let k = 1; k < drivesData.length; k++) {
      const d = drivesData[k];
      const driveId = d[0];
      const company = d[1];
      const ctc = parseFloat(d[2]) || 0;
      const branches = String(d[3]).split(",").map(b => b.trim().toUpperCase());
      const minCgpa = parseFloat(d[4]) || 0;
      const maxCgpa = parseFloat(d[5]) || 10;
      const maxBacklogs = parseInt(d[6]) || 0;
      const placedPolicy = d[7];
      const minDiff = parseFloat(d[8]) || 0;
      const status = d[9];

      if (status !== "ACTIVE") continue;

      let eligible = true;
      let reasons = [];

      // Criteria checks
      if (!branches.includes(student.branch) && !branches.includes("ALL")) {
        eligible = false;
        reasons.push(`Branch ${student.branch} not allowed`);
      }
      if (student.cgpa < minCgpa || student.cgpa > maxCgpa) {
        eligible = false;
        reasons.push(`CGPA ${student.cgpa} outside ${minCgpa}-${maxCgpa}`);
      }
      if (student.backlogs > maxBacklogs) {
        eligible = false;
        reasons.push(`Backlogs (${student.backlogs}) > limit (${maxBacklogs})`);
      }
      if (student.currentPkg > 0) {
        if (placedPolicy === "UNPLACED_ONLY") {
          eligible = false;
          reasons.push(`Unplaced only (You have ${student.currentPkg} LPA)`);
        } else if (placedPolicy === "UPGRADE_ONLY") {
          const req = student.currentPkg + minDiff;
          if (ctc < req) {
            eligible = false;
            reasons.push(`Needs >= ${req} LPA (${student.currentPkg} + ${minDiff} hike)`);
          }
        }
      }

      driveList.push({
        driveId: driveId,
        company: company,
        ctc: ctc,
        branches: branches.join(", "),
        cgpaRange: `${minCgpa} - ${maxCgpa}`,
        maxBacklogs: maxBacklogs,
        isEligible: eligible,
        reasons: reasons,
        hasApplied: appliedDrives.has(driveId)
      });
    }

    return { success: true, drives: driveList, student: student };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// 5. Apply for Drive + Send Email to Custom Address
function applyForDrive(driveId, profileKey, customEmail) {
  try {
    const student = MOCK_PROFILES[profileKey] || MOCK_PROFILES.student_eligible;
    const ss = SpreadsheetApp.openById(REGISTRY_SPREADSHEET_ID);
    let appSheet = ss.getSheetByName("Drive_Applications");
    if (!appSheet) {
      appSheet = ss.insertSheet("Drive_Applications");
      appSheet.appendRow(["AppID", "DriveID", "Company", "USN", "StudentName", "Branch", "CGPA", "AppliedDate", "Status"]);
      appSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#e2e8f0");
    }

    const driveSheet = ss.getSheetByName("Recruitment_Drives");
    const dData = driveSheet.getDataRange().getValues();
    let company = "Recruiter";
    let ctc = "N/A";
    for (let i = 1; i < dData.length; i++) {
      if (dData[i][0] === driveId) {
        company = dData[i][1];
        ctc = dData[i][2];
        break;
      }
    }

    const appId = "APP-" + Utilities.getUuid().slice(0, 6).toUpperCase();
    const timestamp = new Date().toLocaleString('en-IN');

    appSheet.appendRow([
      appId, driveId, company, student.usn, student.name, student.branch, student.cgpa, timestamp, "APPLIED"
    ]);

    // Send confirmation email to test recipient
    const recipient = (customEmail && customEmail.includes("@")) ? customEmail.trim() : Session.getActiveUser().getEmail();
    if (recipient) {
      MailApp.sendEmail({
        to: recipient,
        subject: `[TEST MODE] Application Confirmed: ${company} [${appId}]`,
        htmlBody: `
          <div style="font-family: sans-serif; padding: 16px; border: 1px solid #cbd5e1; border-radius: 8px;">
            <h3 style="color:#0284c7; margin-top:0;">BEC Placement Cell (Test Run)</h3>
            <p>Registration successful for <b>${student.name}</b> (${student.usn})!</p>
            <ul>
              <li><b>Company:</b> ${company} (${ctc} LPA)</li>
              <li><b>Branch:</b> ${student.branch} | <b>CGPA:</b> ${student.cgpa}</li>
              <li><b>Application ID:</b> ${appId}</li>
              <li><b>Timestamp:</b> ${timestamp}</li>
            </ul>
            <p style="font-size:0.8rem; color:#64748b;">This test email was routed to: ${recipient}</p>
          </div>
        `
      });
    }

    return { success: true, message: `Application submitted! Email sent to ${recipient}` };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
