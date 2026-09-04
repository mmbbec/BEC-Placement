// =============== NEW FUNCTIONS ===============

function createRecruitmentDrive(driveData) {
  // driveData: { company, ctc, branches, minCgpa, maxCgpa, maxBacklogs, placedPolicy, minDiff }
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

function getEligibleDrivesForStudent(studentUsn) {
  try {
    // Fetch student data from your existing student sheet (or mock)
    // Assuming you have a function getStudentByUsn(usn) that returns the student object
    // For simplicity, we'll fetch from the existing data source.
    // In your current setup, you probably have a sheet "Students".
    // We'll implement a helper.
    const student = getStudentByUsn(studentUsn);
    if (!student) return { success: false, error: "Student not found" };

    const ss = SpreadsheetApp.openById(REGISTRY_SPREADSHEET_ID);
    const driveSheet = ss.getSheetByName("Recruitment_Drives");
    const appSheet = ss.getSheetByName("Drive_Applications");

    if (!driveSheet || driveSheet.getLastRow() <= 1) {
      return { success: true, drives: [], student: student };
    }

    // Get applied drives
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

function applyForDrive(driveId, studentUsn, customEmail) {
  try {
    const student = getStudentByUsn(studentUsn);
    if (!student) return { success: false, error: "Student not found" };

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

    // Send email to custom address
    const recipient = (customEmail && customEmail.includes("@")) ? customEmail.trim() : Session.getActiveUser().getEmail();
    if (recipient) {
      MailApp.sendEmail({
        to: recipient,
        subject: `[Placement] Application Confirmed: ${company} [${appId}]`,
        htmlBody: `
          <div style="font-family: sans-serif; padding: 16px; border: 1px solid #cbd5e1; border-radius: 8px;">
            <h3 style="color:#0284c7; margin-top:0;">BEC Placement Cell</h3>
            <p>Registration successful for <b>${student.name}</b> (${student.usn})!</p>
            <ul>
              <li><b>Company:</b> ${company} (${ctc} LPA)</li>
              <li><b>Branch:</b> ${student.branch} | <b>CGPA:</b> ${student.cgpa}</li>
              <li><b>Application ID:</b> ${appId}</li>
              <li><b>Timestamp:</b> ${timestamp}</li>
            </ul>
            <p style="font-size:0.8rem; color:#64748b;">This confirmation was sent to: ${recipient}</p>
          </div>
        `
      });
    }

    return { success: true, message: `Application submitted! Email sent to ${recipient}` };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Helper to get student from sheet (adapt to your existing sheet structure)
function getStudentByUsn(usn) {
  // This is a placeholder – you likely have a getStudents() function already.
  // Implement it to fetch from your "Students" sheet.
  // Example:
  const ss = SpreadsheetApp.openById(REGISTRY_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === String(usn).toUpperCase()) {
      return {
        usn: data[i][0],
        name: data[i][1],
        branch: data[i][2],
        cgpa: parseFloat(data[i][3]) || 0,
        backlogs: parseInt(data[i][4]) || 0,
        currentPkg: parseFloat(data[i][5]) || 0
      };
    }
  }
  return null;
}
