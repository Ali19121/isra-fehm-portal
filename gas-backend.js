// ============================================================
//  FEHM-E-QURAN COURSE MONITORING PORTAL
//  Google Apps Script Backend — Paste this in script.google.com
//  Version 2.0 | Complete Backend
// ============================================================
//
//  SETUP INSTRUCTIONS:
//  1. Go to https://script.google.com → New Project
//  2. Paste this entire file
//  3. Change SPREADSHEET_ID below to your Google Sheet ID
//  4. Run setupSheets() ONCE to create all sheets
//  5. Run seedInitialData() ONCE to add default admin account
//  6. Deploy → New Deployment → Web App
//     - Execute as: Me
//     - Who has access: Anyone
//  7. Copy the Web App URL → paste in portal Settings
//
// ============================================================

const SPREADSHEET_ID = '1_QM4_LxH6q9zIGVyQPL3w-v4UcWGPhs0_QOO79y0cvE'; // ← CHANGE THIS
const API_SECRET     = 'fqportal2026secret';         // ← Change to something private

// ── Sheet Names ──
const SH = {
  USERS:       'Users',
  DEPARTMENTS: 'Departments',
  COURSES:     'Courses',
  ASSIGNMENTS: 'Assignments',
  WEEKLY_PLAN: 'WeeklyPlan',
  REPORTS:     'WeeklyReports',
  OBSERVATIONS:'Observations',
  LOGS:        'AdminLogs',
  SESSIONS:    'Sessions',
  TIMETABLE:   'Timetable',
};

// ── CORS Headers ──
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Content-Type':                 'application/json',
  };
}

// ── Main Router ──
function doPost(e) {
  try {
    // Support both JSON body and form-encoded body (data=<JSON>)
    // Form-encoded avoids CORS preflight — required when calling from local HTML file
    let body;
    try {
      if (e.postData.type === 'application/x-www-form-urlencoded') {
        body = JSON.parse(decodeURIComponent(e.parameter.data || e.postData.contents));
      } else {
        body = JSON.parse(e.postData.contents);
      }
    } catch(parseErr) {
      // Last fallback — try raw contents
      body = JSON.parse(e.postData.contents);
    }
    const action = body.action;

    // Public routes (no auth needed)
    if (action === 'login')  return respond(handleLogin(body));
    if (action === 'ping')   return respond({ ok: true, ts: new Date().toISOString() });

    // Protected routes — validate token
    const authResult = validateToken(body.token);
    if (!authResult.ok) return respond({ ok: false, error: 'Unauthorized' }, 401);
    const caller = authResult.user;

    switch (action) {
      // ── Auth ──
      case 'logout':              return respond(handleLogout(body.token));

      // ── Users ──
      case 'getUsers':            return respond(getUsers(caller, body));
      case 'addUser':             return respond(addUser(caller, body));
      case 'updateUser':          return respond(updateUser(caller, body));
      case 'deleteUser':          return respond(deleteUser(caller, body));
      case 'bulkAddUsers':        return respond(bulkAddUsers(caller, body));
      case 'changePassword':      return respond(changePassword(caller, body));

      // ── Departments ──
      case 'getDepartments':      return respond(getDepartments());
      case 'addDepartment':       return respond(addDepartment(caller, body));
      case 'updateDepartment':    return respond(updateDepartment(caller, body));
      case 'deleteDepartment':    return respond(deleteDepartment(caller, body));

      // ── Courses ──
      case 'getCourses':          return respond(getCourses(body));
      case 'addCourse':           return respond(addCourse(caller, body));

      // ── Weekly Plan ──
      case 'getWeeklyPlan':       return respond(getWeeklyPlan(body));

      // ── Assignments ──
      case 'getAssignments':      return respond(getAssignments(caller, body));
      case 'assignLecturer':      return respond(assignLecturer(caller, body));

      // ── Reports ──
      case 'submitReport':        return respond(submitReport(caller, body));
      case 'getReports':          return respond(getReports(caller, body));
      case 'updateReportStatus':  return respond(updateReportStatus(caller, body));
      case 'getMyReports':        return respond(getMyReports(caller, body));

      // ── Dashboard ──
      case 'getDashboardStats':   return respond(getDashboardStats(caller));

      // ── Logs ──
      case 'getLogs':             return respond(getLogs(caller));

      // ── Timetable ──
      case 'getTimetable':        return respond(getTimetable(caller, body));
      case 'addTimetableSlot':    return respond(addTimetableSlot(caller, body));
      case 'updateTimetableSlot': return respond(updateTimetableSlot(caller, body));
      case 'deleteTimetableSlot': return respond(deleteTimetableSlot(caller, body));

      default: return respond({ ok: false, error: 'Unknown action: ' + action }, 400);
    }
  } catch (err) {
    console.error('doPost error:', err);
    return respond({ ok: false, error: err.message }, 500);
  }
}

function doGet(e) {
  return respond({ ok: true, message: 'Fehm-e-Quran API is running', version: '2.0' });
}

function respond(data, code) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Handle CORS preflight OPTIONS request
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
//  SETUP — Run once
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const schemas = {
    [SH.USERS]: [
      'UserID','FullName','Username','PasswordHash','Role',
      'DepartmentID','CourseID','Section','CoordinatorID',
      'Status','Email','CreatedAt','UpdatedAt'
    ],
    [SH.DEPARTMENTS]: [
      'DeptID','DeptName','DeptCode','HeadName','Status','CreatedAt'
    ],
    [SH.COURSES]: [
      'CourseID','CourseName','CourseCode','DeptID',
      'TotalWeeks','Semester','Description','Status','CreatedAt'
    ],
    [SH.ASSIGNMENTS]: [
      'AssignID','LecturerID','CourseID','DeptID',
      'Section','CoordinatorID','AcademicYear','Status','CreatedAt'
    ],
    [SH.WEEKLY_PLAN]: [
      'PlanID','CourseID','WeekNum','BookUnit',
      'LessonsFrom','LessonsTo','LessonsList','TajweedTopic',
      'TranslationTopic','Notes'
    ],
    [SH.REPORTS]: [
      'ReportID','LecturerID','AssignID','CourseID','DeptID',
      'WeekNum','ClassDate','LessonsFrom','LessonsTo','LessonsList',
      'TajweedTopic','TajweedStatus','TranslationStatus','SurahPracticed',
      'Attendance','TotalStudents','UnderstandingLevel',
      'Observations','Issues','Suggestions',
      'SubmittedAt','Status','CoordRemarks','CoordID','ApprovedAt'
    ],
    [SH.OBSERVATIONS]: [
      'ObsID','ReportID','LecturerID','Type','Description','CreatedAt'
    ],
    [SH.LOGS]: [
      'LogID','UserID','Username','Role','Action','Details','Timestamp','IP'
    ],
    [SH.SESSIONS]: [
      'Token','UserID','Username','Role','CreatedAt','ExpiresAt','Active'
    ],
    [SH.TIMETABLE]: [
      'SlotID','LecturerID','LecturerName','DeptID','DeptName',
      'CourseID','CourseName','Section','Subject',
      'Day','StartTime','EndTime','Mode','Room','Link','Notes','CreatedAt','UpdatedAt'
    ],
  };

  Object.entries(schemas).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1a2744')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });

  console.log('✅ All sheets created successfully!');
}

function seedInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Seed Departments
  const deptSheet = ss.getSheetByName(SH.DEPARTMENTS);
  if (deptSheet.getLastRow() <= 1) {
    deptSheet.appendRow(['DEPT001','Islamic Studies Department','ISD','Dr. Admin','active', now()]);
    deptSheet.appendRow(['DEPT002','Quran Sciences Department','QSD','Dr. Coordinator','active', now()]);
  }

  // Seed Courses
  const courseSheet = ss.getSheetByName(SH.COURSES);
  if (courseSheet.getLastRow() <= 1) {
    courseSheet.appendRow(['CRS001','Fehm-e-Quran Course 1','FEQ-1','DEPT001',16,'Semester 1','Understanding of Holy Quran - 1','active', now()]);
    courseSheet.appendRow(['CRS002','Fehm-e-Quran Course 2','FEQ-2','DEPT001',16,'Semester 2','Understanding of Holy Quran - 2','active', now()]);
  }

  // Seed Weekly Plan Course 1
  const planSheet = ss.getSheetByName(SH.WEEKLY_PLAN);
  if (planSheet.getLastRow() <= 1) {
    const c1Plan = [
      ['PLN001','CRS001',1,'Vol. 1 – Unit 1',1,4,'1,2,3,4','Importance of Tajweed','—',''],
      ['PLN002','CRS001',2,'Unit 1',5,9,'5,6,7,8,9','Importance of Tajweed','—',''],
      ['PLN003','CRS001',3,'Unit 1',10,14,'10,11,12,13,14','Makharij (Pronunciation Points)','—',''],
      ['PLN004','CRS001',4,'Unit 1',15,19,'15,16,17,18,19','Makharij (Pronunciation Points)','—',''],
      ['PLN005','CRS001',5,'Vol. 1 – Unit 2',1,4,'1,2,3,4','Makharij (Pronunciation Points)','—',''],
      ['PLN006','CRS001',6,'Unit 2',5,8,'5,6,7,8','Ghunna','—',''],
      ['PLN007','CRS001',7,'Unit 2',9,13,'9,10,11,12,13','Madd','—',''],
      ['PLN008','CRS001',8,'Unit 3',1,3,'1,2,3','Heavy & Light Letters','—',''],
      ['PLN009','CRS001',9,'Unit 3',4,6,'4,5,6','Heavy & Light Letters','—',''],
      ['PLN010','CRS001',10,'Unit 3',7,10,'7,8,9,10','Waqf (Stopping Rules)','—',''],
      ['PLN011','CRS001',11,'Vol. 2 – Unit 4',1,2,'1,2','Recitation of Surah Al-Fatiha & Last 10 Surahs (Practice)','—',''],
      ['PLN012','CRS001',12,'Unit 4',3,4,'3,4','Recitation of Surah Al-Fatiha & Last 10 Surahs (Practice)','—',''],
      ['PLN013','CRS001',13,'Unit 4',5,8,'5,6,7,8','Recitation of Surah Al-Fatiha & Last 10 Surahs (Practice)','—',''],
      ['PLN014','CRS001',14,'Unit 4',9,12,'9,10,11,12','Recitation of Surah Al-Fatiha & Last 10 Surahs (Practice)','—',''],
      ['PLN015','CRS001',15,'Unit 4',13,16,'13,14,15,16','Recitation of Surah Al-Fatiha & Last 10 Surahs (Practice)','—',''],
      ['PLN016','CRS001',16,'Unit 4',17,18,'17,18','Recitation of Surah Al-Fatiha & Last 10 Surahs (Practice)','—',''],
    ];
    const c2Plan = [
      ['PLN017','CRS002',1,'Vol. 2 – Unit 5',1,4,'1,2,3,4','Translation of Surah Al-Fatiha (1)','Surah Al-Fatiha',''],
      ['PLN018','CRS002',2,'Unit 5',5,12,'5,6,7,8,9,10,11,12','Translation of Surah Al-Baqarah (2) – Verse 164','Surah Al-Baqarah',''],
      ['PLN019','CRS002',3,'Unit 5',13,20,'13,14,15,16,17,18,19,20','Translation of Surah Al-Baqarah (2) – Verses 284-286','Surah Al-Baqarah',''],
      ['PLN020','CRS002',4,'Unit 5',21,23,'21,22,23','Translation of Surah Al-Imran (3) – Verses 190-191','Surah Al-Imran',''],
      ['PLN021','CRS002',5,'Vol. 3 – Unit 6',1,5,'1,2,3,4,5','Translation of Surah Al-An\'am (6) – Verses 151-153','Surah Al-Anam',''],
      ['PLN022','CRS002',6,'Unit 6',6,10,'6,7,8,9,10','Translation of Surah Al-Isra (17) – Verses 23-24','Surah Al-Isra',''],
      ['PLN023','CRS002',7,'Unit 6',11,15,'11,12,13,14,15','Translation of Surah An-Nur (24) – Verse 19','Surah An-Nur',''],
      ['PLN024','CRS002',8,'Unit 6',16,21,'16,17,18,19,20,21','Translation of Surah An-Nur (24) – Verses 30-31','Surah An-Nur',''],
      ['PLN025','CRS002',9,'Vol. 4 – Unit 7',1,2,'1,2','Translation of Surah Ar-Rum (30) – Verses 20-27','Surah Ar-Rum',''],
      ['PLN026','CRS002',10,'Unit 7',3,4,'3,4','Translation of Surah Al-Ahzab (33) – Verse 35','Surah Al-Ahzab',''],
      ['PLN027','CRS002',11,'Unit 7',4,6,'4,5,6','Translation of Surah Ha-Meem Sajdah (41) – Verses 33-36','Surah Ha-Meem',''],
      ['PLN028','CRS002',12,'Unit 7',7,9,'7,8,9','Translation of Surah Al-Hujurat (49) – Verses 12-13','Surah Al-Hujurat',''],
      ['PLN029','CRS002',13,'Vol. 5 – Unit 8',1,2,'1,2','Translation of Surah Adh-Dhariyat (51) – Verses 15-19','Surah Adh-Dhariyat',''],
      ['PLN030','CRS002',14,'Unit 8',3,4,'3,4','Translation of Surah Al-Hashr (59) – Verses 22-24','Surah Al-Hashr',''],
      ['PLN031','CRS002',15,'Unit 8',5,6,'5,6','Translation of Surah Ash-Shams (91)','Surah Ash-Shams',''],
      ['PLN032','CRS002',16,'Unit 8',7,8,'7,8','Translation of Surah Al-Ikhlas (112), Al-Falaq (113), An-Nas (114)','Short Surahs',''],
    ];
    [...c1Plan, ...c2Plan].forEach(row => planSheet.appendRow(row));
  }

  // Seed default Admin user
  const userSheet = ss.getSheetByName(SH.USERS);
  if (userSheet.getLastRow() <= 1) {
    userSheet.appendRow([
      'USR001','Admin User','admin', hashPassword('admin123'),
      'admin','DEPT001','','','','active','admin@university.edu', now(), now()
    ]);
    userSheet.appendRow([
      'USR002','Dr. Coordinator','coord1', hashPassword('coord123'),
      'coordinator','DEPT001','CRS001','','USR001','active','coord@university.edu', now(), now()
    ]);
  }

  console.log('✅ Seed data inserted!');
}

// ============================================================
//  AUTH FUNCTIONS
// ============================================================
function handleLogin(body) {
  const { username, password } = body;
  if (!username || !password) return { ok: false, error: 'Username and password required' };

  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet   = ss.getSheetByName(SH.USERS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const idx = {
    id:       headers.indexOf('UserID'),
    name:     headers.indexOf('FullName'),
    uname:    headers.indexOf('Username'),
    pass:     headers.indexOf('PasswordHash'),
    role:     headers.indexOf('Role'),
    deptId:   headers.indexOf('DepartmentID'),
    courseId: headers.indexOf('CourseID'),
    section:  headers.indexOf('Section'),
    coordId:  headers.indexOf('CoordinatorID'),
    status:   headers.indexOf('Status'),
    email:    headers.indexOf('Email'),
  };

  const row = data.slice(1).find(r =>
    r[idx.uname].toString().toLowerCase() === username.toLowerCase()
  );

  if (!row) return { ok: false, error: 'User not found' };
  if (row[idx.status] !== 'active') return { ok: false, error: 'Account is inactive' };
  if (!verifyPassword(password, row[idx.pass].toString())) return { ok: false, error: 'Incorrect password' };

  const token   = generateToken();
  const expires = new Date(Date.now() + 8 * 60 * 60 * 1000).toISOString(); // 8 hours

  const sessSheet = ss.getSheetByName(SH.SESSIONS);
  sessSheet.appendRow([token, row[idx.id], row[idx.uname], row[idx.role], now(), expires, true]);

  writeLog(ss, row[idx.id], row[idx.uname], row[idx.role], 'Login', 'Successful login');

  // Get department info
  const deptName = getDeptName(ss, row[idx.deptId]);
  const courseName = getCourseName(ss, row[idx.courseId]);

  return {
    ok: true,
    token,
    user: {
      id:         row[idx.id],
      name:       row[idx.name],
      username:   row[idx.uname],
      role:       row[idx.role],
      deptId:     row[idx.deptId],
      deptName,
      courseId:   row[idx.courseId],
      courseName,
      section:    row[idx.section],
      coordId:    row[idx.coordId],
      email:      row[idx.email],
    }
  };
}

function handleLogout(token) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.SESSIONS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.getRange(i + 1, 7).setValue(false);
      break;
    }
  }
  return { ok: true };
}

function validateToken(token) {
  if (!token) return { ok: false };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.SESSIONS);
  const data  = sheet.getDataRange().getValues();
  const row   = data.slice(1).find(r => r[0] === token && r[6] === true);
  if (!row) return { ok: false };
  if (new Date(row[5]) < new Date()) return { ok: false };
  return {
    ok: true,
    user: { id: row[1], username: row[2], role: row[3] }
  };
}

// ============================================================
//  USER MANAGEMENT
// ============================================================
function getUsers(caller, body) {
  if (!['admin','coordinator'].includes(caller.role)) return { ok:false, error:'Forbidden' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  let users = data.slice(1).map(r => rowToObj(headers, r));

  // Coordinators see only their lecturers
  if (caller.role === 'coordinator') {
    users = users.filter(u => u.CoordinatorID === caller.id || u.UserID === caller.id);
  }

  // Filter by role if requested
  if (body.role) users = users.filter(u => u.Role === body.role);
  if (body.deptId) users = users.filter(u => u.DepartmentID === body.deptId);
  if (body.courseId) users = users.filter(u => u.CourseID === body.courseId);

  // Remove password hashes
  users.forEach(u => delete u.PasswordHash);

  // Enrich with dept/course names
  users.forEach(u => {
    u.DeptName   = getDeptName(ss, u.DepartmentID);
    u.CourseName = getCourseName(ss, u.CourseID);
  });

  return { ok: true, users };
}

function addUser(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Only admin can add users' };
  const { fullName, username, password, role, deptId, courseId, section, coordId, email } = body;

  if (!fullName || !username || !password || !role) return { ok:false, error:'Required fields missing' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  const data  = sheet.getDataRange().getValues();

  // Check username unique
  const exists = data.slice(1).find(r => r[2].toString().toLowerCase() === username.toLowerCase());
  if (exists) return { ok:false, error:'Username already exists' };

  const userId = 'USR' + String(Date.now()).slice(-6);
  sheet.appendRow([
    userId, fullName, username, hashPassword(password),
    role, deptId||'', courseId||'', section||'', coordId||'',
    'active', email||'', now(), now()
  ]);

  writeLog(ss, caller.id, caller.username, caller.role, 'Add User', `Added ${role}: ${fullName} (${username})`);
  return { ok:true, userId, message: 'User added successfully' };
}

function bulkAddUsers(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Only admin can bulk add users' };
  const { users } = body; // Array of user objects
  if (!Array.isArray(users) || users.length === 0) return { ok:false, error:'No users provided' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  const existingData = sheet.getDataRange().getValues();
  const existingUsernames = existingData.slice(1).map(r => r[2].toString().toLowerCase());

  const results = { added:0, skipped:0, errors:[] };

  users.forEach((u, i) => {
    if (!u.username || !u.password || !u.fullName) {
      results.errors.push(`Row ${i+2}: Missing required fields`);
      results.skipped++;
      return;
    }
    if (existingUsernames.includes(u.username.toLowerCase())) {
      results.errors.push(`Row ${i+2}: Username "${u.username}" already exists`);
      results.skipped++;
      return;
    }
    const uid = 'USR' + String(Date.now()).slice(-6) + i;
    sheet.appendRow([
      uid, u.fullName, u.username, hashPassword(u.password),
      u.role||'teacher', u.deptId||'', u.courseId||'', u.section||'',
      u.coordId||'', 'active', u.email||'', now(), now()
    ]);
    existingUsernames.push(u.username.toLowerCase());
    results.added++;
  });

  writeLog(ss, caller.id, caller.username, caller.role, 'Bulk Add Users', `Added ${results.added}, Skipped ${results.skipped}`);
  return { ok:true, ...results };
}

function updateUser(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.userId) {
      if (body.fullName)   sheet.getRange(i+1, 2).setValue(body.fullName);
      if (body.role)       sheet.getRange(i+1, 5).setValue(body.role);
      if (body.deptId)     sheet.getRange(i+1, 6).setValue(body.deptId);
      if (body.courseId)   sheet.getRange(i+1, 7).setValue(body.courseId);
      if (body.section)    sheet.getRange(i+1, 8).setValue(body.section);
      if (body.coordId)    sheet.getRange(i+1, 9).setValue(body.coordId);
      if (body.status)     sheet.getRange(i+1, 10).setValue(body.status);
      if (body.email)      sheet.getRange(i+1, 11).setValue(body.email);
      sheet.getRange(i+1, 13).setValue(now());
      writeLog(ss, caller.id, caller.username, caller.role, 'Update User', `Updated: ${data[i][2]}`);
      return { ok:true };
    }
  }
  return { ok:false, error:'User not found' };
}

function deleteUser(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.userId) {
      // Soft delete — set status inactive
      sheet.getRange(i+1, 10).setValue('inactive');
      writeLog(ss, caller.id, caller.username, caller.role, 'Delete User', `Deactivated: ${data[i][2]}`);
      return { ok:true };
    }
  }
  return { ok:false, error:'User not found' };
}

function changePassword(caller, body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.USERS);
  const data  = sheet.getDataRange().getValues();

  const targetId = caller.role === 'admin' ? body.userId : caller.id;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetId) {
      if (caller.role !== 'admin') {
        // Verify old password
        if (!verifyPassword(body.oldPassword, data[i][3].toString())) {
          return { ok:false, error:'Current password incorrect' };
        }
      }
      sheet.getRange(i+1, 4).setValue(hashPassword(body.newPassword));
      sheet.getRange(i+1, 13).setValue(now());
      return { ok:true, message:'Password changed successfully' };
    }
  }
  return { ok:false, error:'User not found' };
}

// ============================================================
//  DEPARTMENTS
// ============================================================
function getDepartments() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.DEPARTMENTS);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const depts = data.slice(1)
    .filter(r => r[4] === 'active')
    .map(r => rowToObj(headers, r));
  return { ok:true, departments: depts };
}

function addDepartment(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.DEPARTMENTS);
  const deptId = 'DEPT' + String(Date.now()).slice(-6);
  sheet.appendRow([deptId, body.deptName, body.deptCode||'', body.headName||'', 'active', now()]);
  writeLog(ss, caller.id, caller.username, caller.role, 'Add Department', body.deptName);
  return { ok:true, deptId };
}

function updateDepartment(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.DEPARTMENTS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.deptId) {
      if (body.deptName) sheet.getRange(i+1, 2).setValue(body.deptName);
      if (body.headName) sheet.getRange(i+1, 4).setValue(body.headName);
      if (body.status)   sheet.getRange(i+1, 5).setValue(body.status);
      return { ok:true };
    }
  }
  return { ok:false, error:'Department not found' };
}

function deleteDepartment(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.DEPARTMENTS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.deptId) {
      sheet.getRange(i+1, 5).setValue('inactive');
      return { ok:true };
    }
  }
  return { ok:false, error:'Not found' };
}

// ============================================================
//  COURSES
// ============================================================
function getCourses(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.COURSES);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  let courses = data.slice(1)
    .filter(r => r[7] === 'active')
    .map(r => rowToObj(headers, r));
  if (body && body.deptId) courses = courses.filter(c => c.DeptID === body.deptId);
  return { ok:true, courses };
}

function addCourse(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.COURSES);
  const cId   = 'CRS' + String(Date.now()).slice(-6);
  sheet.appendRow([cId, body.courseName, body.courseCode||'', body.deptId,
    body.totalWeeks||16, body.semester||'', body.description||'', 'active', now()]);
  return { ok:true, courseId: cId };
}

// ============================================================
//  WEEKLY PLAN
// ============================================================
function getWeeklyPlan(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.WEEKLY_PLAN);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  let plan = data.slice(1).map(r => rowToObj(headers, r));
  if (body.courseId) plan = plan.filter(p => p.CourseID === body.courseId);
  if (body.weekNum)  plan = plan.filter(p => String(p.WeekNum) === String(body.weekNum));

  // Parse lesson list to array
  plan.forEach(p => {
    if (p.LessonsList) {
      p.LessonsArray = p.LessonsList.toString().split(',').map(x => parseInt(x.trim())).filter(Boolean);
    } else {
      p.LessonsArray = [];
    }
  });

  return { ok:true, plan };
}

// ============================================================
//  REPORTS
// ============================================================
function submitReport(caller, body) {
  if (caller.role !== 'teacher') return { ok:false, error:'Only teachers can submit reports' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.REPORTS);
  const data  = sheet.getDataRange().getValues();

  // Check if report for this week already exists
  const existing = data.slice(1).find(r =>
    r[1] === caller.id && String(r[4]) === String(body.weekNum) && r[2] === body.assignId
  );
  if (existing && existing[21] === 'submitted') {
    return { ok:false, error:'Report for this week already submitted. Contact coordinator to edit.' };
  }

  const rId = 'RPT' + String(Date.now()).slice(-8);
  sheet.appendRow([
    rId, caller.id, body.assignId||'', body.courseId||'', body.deptId||'',
    body.weekNum, body.classDate||'', body.lessonsFrom||'', body.lessonsTo||'',
    body.lessonsList||'',
    body.tajweedTopic||'', body.tajweedStatus||'', body.translationStatus||'',
    body.surahPracticed||'',
    body.attendance||'', body.totalStudents||'', body.understandingLevel||'',
    body.observations||'', body.issues||'', body.suggestions||'',
    now(), 'pending', '', '', ''
  ]);

  writeLog(ss, caller.id, caller.username, caller.role, 'Submit Report', `Week ${body.weekNum} report submitted`);
  return { ok:true, reportId: rId, message:'Report submitted successfully' };
}

function getReports(caller, body) {
  if (!['admin','coordinator'].includes(caller.role)) return { ok:false, error:'Forbidden' };

  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rSheet  = ss.getSheetByName(SH.REPORTS);
  const rData   = rSheet.getDataRange().getValues();
  const rHeaders= rData[0];

  let reports = rData.slice(1).map(r => rowToObj(rHeaders, r));

  if (body.courseId)   reports = reports.filter(r => r.CourseID    === body.courseId);
  if (body.deptId)     reports = reports.filter(r => r.DeptID      === body.deptId);
  if (body.weekNum)    reports = reports.filter(r => String(r.WeekNum) === String(body.weekNum));
  if (body.status)     reports = reports.filter(r => r.Status      === body.status);

  // Enrich with lecturer names
  const uSheet  = ss.getSheetByName(SH.USERS);
  const uData   = uSheet.getDataRange().getValues();
  const uHeaders= uData[0];
  const usersMap = {};
  uData.slice(1).forEach(r => {
    const u = rowToObj(uHeaders, r);
    usersMap[u.UserID] = u.FullName;
  });

  reports.forEach(r => {
    r.LecturerName = usersMap[r.LecturerID] || r.LecturerID;
    r.CourseName   = getCourseName(ss, r.CourseID);
    r.DeptName     = getDeptName(ss, r.DeptID);
  });

  // Coordinator: only their lecturers
  if (caller.role === 'coordinator') {
    const myLecturers = uData.slice(1)
      .filter(r => r[uHeaders.indexOf('CoordinatorID')] === caller.id)
      .map(r => r[uHeaders.indexOf('UserID')]);
    reports = reports.filter(r => myLecturers.includes(r.LecturerID));
  }

  return { ok:true, reports };
}

function getMyReports(caller, body) {
  if (caller.role !== 'teacher') return { ok:false, error:'Forbidden' };

  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet  = ss.getSheetByName(SH.REPORTS);
  const data   = sheet.getDataRange().getValues();
  const headers= data[0];

  let reports = data.slice(1)
    .map(r => rowToObj(headers, r))
    .filter(r => r.LecturerID === caller.id);

  reports.forEach(r => {
    r.CourseName = getCourseName(ss, r.CourseID);
  });

  return { ok:true, reports };
}

function updateReportStatus(caller, body) {
  if (!['admin','coordinator'].includes(caller.role)) return { ok:false, error:'Forbidden' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.REPORTS);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.reportId) {
      if (body.status)      sheet.getRange(i+1, 22).setValue(body.status);
      if (body.coordRemarks)sheet.getRange(i+1, 23).setValue(body.coordRemarks);
      if (body.status === 'approved') {
        sheet.getRange(i+1, 24).setValue(caller.id);
        sheet.getRange(i+1, 25).setValue(now());
      }
      writeLog(ss, caller.id, caller.username, caller.role, 'Update Report', `Report ${body.reportId} → ${body.status}`);
      return { ok:true };
    }
  }
  return { ok:false, error:'Report not found' };
}

// ============================================================
//  DASHBOARD STATS
// ============================================================
function getDashboardStats(caller) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const uData = ss.getSheetByName(SH.USERS).getDataRange().getValues();
  const uH    = uData[0];
  const users = uData.slice(1).map(r => rowToObj(uH, r)).filter(u => u.Status === 'active');

  const rData = ss.getSheetByName(SH.REPORTS).getDataRange().getValues();
  const rH    = rData[0];
  const reports = rData.slice(1).map(r => rowToObj(rH, r));

  const teachers     = users.filter(u => u.Role === 'teacher');
  const coordinators = users.filter(u => u.Role === 'coordinator');

  // Current week (rough calculation from semester start)
  const semStart  = new Date('2026-01-15');
  const currentWk = Math.min(Math.max(Math.floor((Date.now() - semStart)/604800000) + 1, 1), 16);

  const thisWeekReports = reports.filter(r => String(r.WeekNum) === String(currentWk));
  const pendingReports  = reports.filter(r => r.Status === 'pending');

  // Calculate delays
  const submittedTeachers = [...new Set(thisWeekReports.map(r => r.LecturerID))];
  const delayedTeachers   = teachers.filter(t => !submittedTeachers.includes(t.UserID));

  return {
    ok: true,
    stats: {
      totalTeachers:     teachers.length,
      totalCoordinators: coordinators.length,
      currentWeek:       currentWk,
      thisWeekSubmitted: thisWeekReports.length,
      pendingApproval:   pendingReports.length,
      delayedTeachers:   delayedTeachers.length,
      totalReports:      reports.length,
    }
  };
}

// ============================================================
//  LOGS
// ============================================================
function getLogs(caller) {
  if (caller.role !== 'admin') return { ok:false, error:'Forbidden' };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.LOGS);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const logs  = data.slice(1)
    .map(r => rowToObj(headers, r))
    .reverse()
    .slice(0, 200); // last 200
  return { ok:true, logs };
}


// ============================================================
//  TIMETABLE FUNCTIONS
// ============================================================
function getTimetable(caller, body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.TIMETABLE);
  if (!sheet) return { ok:false, error:'Timetable sheet not found. Run setupSheets() first.' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  let slots = data.slice(1).map(r => rowToObj(headers, r)).filter(s => s.SlotID);

  // Lecturers see only their own slots
  if (caller.role === 'teacher') {
    slots = slots.filter(s => s.LecturerID === caller.id);
  }
  // Coordinators see their dept
  if (caller.role === 'coordinator') {
    const uSheet  = ss.getSheetByName(SH.USERS);
    const uData   = uSheet.getDataRange().getValues();
    const uH      = uData[0];
    const myLecIds = uData.slice(1)
      .filter(r => r[uH.indexOf('CoordinatorID')] === caller.id)
      .map(r => r[uH.indexOf('UserID')]);
    slots = slots.filter(s => myLecIds.includes(s.LecturerID));
  }

  // Filters
  if (body.deptId)      slots = slots.filter(s => s.DeptID     === body.deptId);
  if (body.courseId)    slots = slots.filter(s => s.CourseID   === body.courseId);
  if (body.day)         slots = slots.filter(s => s.Day        === body.day);
  if (body.mode)        slots = slots.filter(s => s.Mode       === body.mode);
  if (body.lecturerId)  slots = slots.filter(s => s.LecturerID === body.lecturerId);

  return { ok:true, slots };
}

function addTimetableSlot(caller, body) {
  if (!['admin','teacher','coordinator'].includes(caller.role)) return { ok:false, error:'Forbidden' };
  // Teachers can only add for themselves
  if (caller.role === 'teacher' && body.lecturerId !== caller.id) {
    body.lecturerId = caller.id;
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.TIMETABLE);
  if (!sheet) return { ok:false, error:'Timetable sheet not found. Run setupSheets() first.' };

  const slotId = 'SLT' + String(Date.now()).slice(-8);
  sheet.appendRow([
    slotId,
    body.lecturerId   || '',
    body.lecturerName || '',
    body.deptId       || '',
    body.deptName     || '',
    body.courseId     || '',
    body.courseName   || '',
    body.section      || '',
    body.subject      || '',
    body.day          || '',
    body.startTime    || '',
    body.endTime      || '',
    body.mode         || 'Physical',
    body.room         || '',
    body.link         || '',
    body.notes        || '',
    now(), now()
  ]);

  writeLog(ss, caller.id, caller.username, caller.role, 'Add Timetable Slot',
    `${body.lecturerName||caller.username} — ${body.day} ${body.startTime}–${body.endTime}`);
  return { ok:true, slotId };
}

function updateTimetableSlot(caller, body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.TIMETABLE);
  if (!sheet) return { ok:false, error:'Timetable sheet not found.' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.slotId) {
      // Teachers can only edit their own slots
      if (caller.role === 'teacher' && data[i][1] !== caller.id) {
        return { ok:false, error:'Ap sirf apna slot edit kar sakte hain.' };
      }
      const row = i + 1;
      const setCol = (col, val) => { if (val !== undefined) sheet.getRange(row, headers.indexOf(col)+1).setValue(val); };
      setCol('LecturerName', body.lecturerName);
      setCol('DeptID',       body.deptId);
      setCol('DeptName',     body.deptName);
      setCol('CourseID',     body.courseId);
      setCol('CourseName',   body.courseName);
      setCol('Section',      body.section);
      setCol('Subject',      body.subject);
      setCol('Day',          body.day);
      setCol('StartTime',    body.startTime);
      setCol('EndTime',      body.endTime);
      setCol('Mode',         body.mode);
      setCol('Room',         body.room);
      setCol('Link',         body.link);
      setCol('Notes',        body.notes);
      sheet.getRange(row, headers.indexOf('UpdatedAt')+1).setValue(now());
      writeLog(ss, caller.id, caller.username, caller.role, 'Update Timetable Slot', body.slotId);
      return { ok:true };
    }
  }
  return { ok:false, error:'Slot not found' };
}

function deleteTimetableSlot(caller, body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SH.TIMETABLE);
  if (!sheet) return { ok:false, error:'Timetable sheet not found.' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.slotId) {
      // Teachers can only delete their own slots
      if (caller.role === 'teacher' && data[i][1] !== caller.id) {
        return { ok:false, error:'Ap sirf apna slot delete kar sakte hain.' };
      }
      sheet.deleteRow(i + 1);
      writeLog(ss, caller.id, caller.username, caller.role, 'Delete Timetable Slot', body.slotId);
      return { ok:true };
    }
  }
  return { ok:false, error:'Slot not found' };
}

// ============================================================
//  HELPER FUNCTIONS
// ============================================================
function hashPassword(plain) {
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plain + 'fqsalt2026',
    Utilities.Charset.UTF_8
  );
  return hash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0')).join('');
}

function verifyPassword(plain, hash) {
  return hashPassword(plain) === hash;
}

function generateToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t = '';
  for (let i = 0; i < 48; i++) t += chars[Math.floor(Math.random() * chars.length)];
  return t;
}

function now() {
  return new Date().toISOString();
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function getDeptName(ss, deptId) {
  if (!deptId) return '';
  const sheet = ss.getSheetByName(SH.DEPARTMENTS);
  const data  = sheet.getDataRange().getValues();
  const row   = data.slice(1).find(r => r[0] === deptId);
  return row ? row[1] : deptId;
}

function getCourseName(ss, courseId) {
  if (!courseId) return '';
  const sheet = ss.getSheetByName(SH.COURSES);
  const data  = sheet.getDataRange().getValues();
  const row   = data.slice(1).find(r => r[0] === courseId);
  return row ? row[1] : courseId;
}

function writeLog(ss, userId, username, role, action, details) {
  const sheet = ss.getSheetByName(SH.LOGS);
  const logId = 'LOG' + String(Date.now()).slice(-8);
  sheet.appendRow([logId, userId, username, role, action, details, now(), '']);
}
