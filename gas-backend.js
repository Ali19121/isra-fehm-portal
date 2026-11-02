// ============================================================
//  FEHM-E-QURAN COURSE MONITORING PORTAL
//  Google Apps Script Backend — Paste this in script.google.com
//  Version 3.0 | Complete Backend (additive upgrade)
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

      case 'sync':                return respond(syncAction(caller, body));
      case 'getComparison':       return respond(getComparison(caller, body));
      case 'aiChat':              return respond(aiChat(caller, body));

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
      'Status','Email','CreatedAt','UpdatedAt','LastSeen'
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
      'SubmittedAt','Status','CoordRemarks','CoordID','ApprovedAt',
      'BookUnit','StudentYear','Semester','UpdatedAt'
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
  CacheService.getScriptCache().remove('tk_' + token);
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
  const cache = CacheService.getScriptCache();
  const hit = cache.get('tk_' + token);
  if (hit) return { ok: true, user: JSON.parse(hit) };
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data  = ss.getSheetByName(SH.SESSIONS).getDataRange().getValues();
  const row   = data.slice(1).find(r => r[0] === token && r[6] === true);
  if (!row) return { ok: false };
  if (new Date(row[5]) < new Date()) return { ok: false };
  const user = { id: row[1], username: row[2], role: row[3] };
  cache.put('tk_' + token, JSON.stringify(user), 600);
  return { ok: true, user };
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

function getReports(caller, body) {
  if (!['admin','coordinator','teacher'].includes(caller.role)) return { ok:false, error:'Forbidden' };

  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rSheet  = ss.getSheetByName(SH.REPORTS);
  const rData   = rSheet.getDataRange().getValues();
  const rHeaders= rData[0];

  let reports = rData.slice(1).map(r => rowToObj(rHeaders, r));
  if (caller.role === 'teacher') reports = reports.filter(r => r.LecturerID === caller.id);

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

let _DN = null, _CN = null;
function getDeptName(ss, deptId) {
  if (!deptId) return '';
  if (!_DN) { _DN = {}; ss.getSheetByName(SH.DEPARTMENTS).getDataRange().getValues().slice(1).forEach(r => _DN[r[0]] = r[1]); }
  return _DN[deptId] || deptId;
}

function getCourseName(ss, courseId) {
  if (!courseId) return '';
  if (!_CN) { _CN = {}; ss.getSheetByName(SH.COURSES).getDataRange().getValues().slice(1).forEach(r => _CN[r[0]] = r[1]); }
  return _CN[courseId] || courseId;
}

function writeLog(ss, userId, username, role, action, details) {
  bumpVer();
  const sheet = ss.getSheetByName(SH.LOGS);
  const logId = 'LOG' + String(Date.now()).slice(-8);
  sheet.appendRow([logId, userId, username, role, action, details, now(), '']);
}

// ============================================================
//  V3 ADDITIONS — additive only (no existing data is touched)
//  1) Run migrateSheets() ONCE (adds new columns at the END of sheets)
//  2) Project Settings → Script properties → GEMINI_KEY = <your Google AI Studio key>
//     (optional GEMINI_MODEL, default gemini-2.5-flash)
// ============================================================
// AI model fallback chain (fast first). Override with Script Property GEMINI_MODELS (comma separated).
const AI_MODELS = ['gemini-3-flash-preview','gemini-3.1-flash-lite','gemini-flash-latest','gemini-flash-lite-latest'];
const NEW_COLS = { REPORTS:['BookUnit','StudentYear','Semester','UpdatedAt'], USERS:['LastSeen'] };

function ensureCols(sheet, cols) {
  const lc = Math.max(sheet.getLastColumn(), 1);
  const h  = sheet.getRange(1, 1, 1, lc).getValues()[0];
  const miss = cols.filter(c => h.indexOf(c) < 0);
  if (miss.length) {
    sheet.getRange(1, lc + 1, 1, miss.length).setValues([miss])
      .setFontWeight('bold').setBackground('#1a2744').setFontColor('#ffffff');
    return h.concat(miss);
  }
  return h;
}
function migrateSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCols(ss.getSheetByName(SH.REPORTS), NEW_COLS.REPORTS);
  ensureCols(ss.getSheetByName(SH.USERS),   NEW_COLS.USERS);
  console.log('✅ Migration done — existing data untouched.');
}

// ── Realtime: version counter + heartbeat ──
function bumpVer() { try { CacheService.getScriptCache().put('ver', String(Date.now()), 21600); } catch (e) {} }
function getVer()  { return CacheService.getScriptCache().get('ver') || '0'; }

function syncAction(caller, body) {
  const c = CacheService.getScriptCache();
  if (!c.get('hb_' + caller.id)) {
    c.put('hb_' + caller.id, '1', 50);
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SH.USERS);
    const h  = ensureCols(sh, NEW_COLS.USERS);
    const ids = sh.getRange(2, h.indexOf('UserID') + 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === caller.id) { sh.getRange(i + 2, h.indexOf('LastSeen') + 1).setValue(now()); break; }
    }
  }
  const out = { ok: true, ver: getVer() };
  if (caller.role === 'admin') out.live = getLive();
  return out;
}

function getLive() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const d  = ss.getSheetByName(SH.USERS).getDataRange().getValues();
  const h  = d[0], li = h.indexOf('LastSeen');
  if (li < 0) return [];
  const cut = Date.now() - 3 * 60 * 1000, out = [];
  d.slice(1).forEach(r => {
    const t = r[li] ? new Date(r[li]).getTime() : 0;
    if (t > cut && r[h.indexOf('Status')] === 'active') out.push({
      id: r[h.indexOf('UserID')], name: r[h.indexOf('FullName')], role: r[h.indexOf('Role')],
      deptName: getDeptName(ss, r[h.indexOf('DepartmentID')]), lastSeen: r[li]
    });
  });
  return out;
}

// ── Reports: submit / edit (upsert) ──
function submitReport(caller, body) {
  if (caller.role !== 'teacher') return { ok:false, error:'Only teachers can submit reports' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SH.REPORTS);
    const h  = ensureCols(sh, NEW_COLS.REPORTS);
    const data = sh.getDataRange().getValues();
    const ix = n => h.indexOf(n);
    let row = -1;
    if (body.reportId) {
      row = data.findIndex((r, i) => i > 0 && r[ix('ReportID')] === body.reportId);
      if (row < 0) return { ok:false, error:'Report not found' };
      if (data[row][ix('LecturerID')] !== caller.id) return { ok:false, error:'Forbidden' };
    } else {
      row = data.findIndex((r, i) => i > 0 && r[ix('LecturerID')] === caller.id &&
        String(r[ix('WeekNum')]) === String(body.weekNum) && String(r[ix('CourseID')]) === String(body.courseId || ''));
    }
    if (row > 0 && data[row][ix('Status')] === 'approved')
      return { ok:false, error:'Approved report locked. Coordinator se return karwa ke edit karen.' };

    const rec = {
      LecturerID:caller.id, AssignID:body.assignId||'', CourseID:body.courseId||'', DeptID:body.deptId||'',
      WeekNum:body.weekNum, ClassDate:body.classDate||'', LessonsFrom:body.lessonsFrom||'', LessonsTo:body.lessonsTo||'',
      LessonsList:body.lessonsList||'', TajweedTopic:body.tajweedTopic||'', TajweedStatus:body.tajweedStatus||'',
      TranslationStatus:body.translationStatus||'', SurahPracticed:body.surahPracticed||'', Attendance:body.attendance||'',
      TotalStudents:body.totalStudents||'', UnderstandingLevel:body.understandingLevel||'', Observations:body.observations||'',
      Issues:body.issues||'', Suggestions:body.suggestions||'', BookUnit:body.bookUnit||'', StudentYear:body.studentYear||'',
      Semester:body.semester||'', UpdatedAt:now(), Status:'pending'
    };
    let id;
    if (row > 0) {
      id = data[row][ix('ReportID')];
      const vals = h.map((c, i) => rec[c] !== undefined ? rec[c] : (data[row][i] === undefined ? '' : data[row][i]));
      sh.getRange(row + 1, 1, 1, h.length).setValues([vals]);
      writeLog(ss, caller.id, caller.username, caller.role, 'Edit Report', `Week ${body.weekNum} report updated`);
      return { ok:true, reportId:id, updated:true, message:'Report updated' };
    }
    id = 'RPT' + String(Date.now()).slice(-8);
    rec.ReportID = id; rec.SubmittedAt = now(); rec.CoordRemarks = '';
    sh.appendRow(h.map(c => rec[c] !== undefined ? rec[c] : ''));
    writeLog(ss, caller.id, caller.username, caller.role, 'Submit Report', `Week ${body.weekNum} report submitted`);
    return { ok:true, reportId:id, message:'Report submitted successfully' };
  } finally { lock.releaseLock(); }
}

// ── Planned vs Actual ──
function unitNo(s) { const m = String(s || '').match(/unit\s*(\d+)/i); return m ? +m[1] : 0; }

function getComparison(caller, body) {
  if (!['admin','coordinator'].includes(caller.role)) return { ok:false, error:'Forbidden' };
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rep = getReports(caller, body || {});
  if (!rep.ok) return rep;
  const pd = ss.getSheetByName(SH.WEEKLY_PLAN).getDataRange().getValues(), ph = pd[0];
  const byC = {};
  pd.slice(1).map(r => rowToObj(ph, r)).forEach(p => { (byC[p.CourseID] = byC[p.CourseID] || []).push(p); });
  const idx = {};
  Object.keys(byC).forEach(c => {
    let n = 0; const m = {};
    byC[c].sort((a, b) => a.WeekNum - b.WeekNum).forEach(p => {
      const u = unitNo(p.BookUnit);
      String(p.LessonsList).split(',').forEach(l => { const k = u + '|' + (+l); if (l !== '' && !(k in m)) m[k] = ++n; });
    });
    idx[c] = m;
  });
  const pos = (m, u, l) => {
    if (m[u + '|' + l]) return m[u + '|' + l];
    let best = null;
    Object.keys(m).forEach(k => { const [uu, ll] = k.split('|').map(Number); if (uu === u && ll <= l && (best === null || m[k] > best)) best = m[k]; });
    return best;
  };
  const rows = rep.reports.map(r => {
    const plist = byC[r.CourseID] || [];
    const p = plist.find(x => +x.WeekNum === +r.WeekNum);
    const m = idx[r.CourseID] || {};
    const pu = p ? unitNo(p.BookUnit) : 0;
    const au = unitNo(r.BookUnit) || pu;
    const pEnd = p ? pos(m, pu, +p.LessonsTo) : null;
    const aEnd = r.LessonsTo !== '' ? pos(m, au, +r.LessonsTo) : null;
    const gap  = (pEnd !== null && aEnd !== null) ? aEnd - pEnd : null;
    return {
      ReportID:r.ReportID, LecturerID:r.LecturerID, LecturerName:r.LecturerName, DeptName:r.DeptName, CourseID:r.CourseID,
      CourseName:r.CourseName, WeekNum:r.WeekNum, Status:r.Status, StudentYear:r.StudentYear||'', Semester:r.Semester||'',
      Attendance:r.Attendance, TotalStudents:r.TotalStudents, Understanding:r.UnderstandingLevel,
      planned: p ? { unit:'Unit ' + pu, from:p.LessonsFrom, to:p.LessonsTo, topic:p.TajweedTopic } : null,
      actual:  { unit: au ? 'Unit ' + au : '—', from:r.LessonsFrom, to:r.LessonsTo, topic:r.TajweedTopic, known: !!unitNo(r.BookUnit) },
      gap, state: gap === null ? 'unknown' : gap > 0 ? 'ahead' : gap < 0 ? 'behind' : 'ontrack'
    };
  });
  return { ok:true, rows };
}

// ── Admin AI assistant (Gemini). Key stays in Script Properties, never in the browser ──
function aiChat(caller, body) {
  if (caller.role !== 'admin') return { ok:false, error:'Admin only' };
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_KEY');
  if (!key) return { ok:false, error:'GEMINI_KEY Script Properties me set nahi hai (Project Settings → Script properties).' };
  const models = (props.getProperty('GEMINI_MODELS') || AI_MODELS.join(',')).split(',').map(s => s.trim()).filter(Boolean);
  const cmp = getComparison(caller, {});
  const tt  = getTimetable(caller, {});
  const ctx = {
    today: now().slice(0, 10), week: body.week || '',
    reports: (cmp.rows || []).slice(-350).map(r => ({
      t:r.LecturerName, dept:r.DeptName, course:r.CourseName, wk:r.WeekNum, st:r.Status, yr:r.StudentYear, sem:r.Semester,
      att:r.Attendance, tot:r.TotalStudents, und:r.Understanding,
      plan: r.planned ? `${r.planned.unit} L${r.planned.from}-${r.planned.to}` : '', act:`${r.actual.unit} L${r.actual.from}-${r.actual.to}`,
      gap:r.gap, state:r.state })),
    timetable: ((tt && tt.slots) || []).slice(0, 250).map(s => ({
      t:s.LecturerName, dept:s.DeptName, course:s.CourseName, sec:s.Section, day:s.Day, mode:s.Mode }))
  };
  const SYS = 'You are the analytics assistant for Isra University Fehm-e-Quran portal (admin only). ' +
    'Answer in the language of the question (Roman Urdu/Urdu/English). Use ONLY the supplied DATA. ' +
    'Reply as JSON: {"answer":string,"charts":[{"type":"bar|line|pie|doughnut","title":string,"labels":[string],"data":[number],"label":string}],' +
    '"slides":[{"title":string,"bullets":[string]}]}. Add charts only when useful; add slides only when asked for a presentation/slides/report. ' +
    'gap = actual lesson position minus planned (positive = ahead, negative = behind).';
  const hist = (body.history || []).slice(-6).map(m => ({ role: m.role === 'user' ? 'user' : 'model', parts:[{ text:String(m.text).slice(0, 1500) }] }));
  const payload = {
    systemInstruction: { parts:[{ text:SYS }] },
    contents: hist.concat([{ role:'user', parts:[{ text:'DATA:\n' + JSON.stringify(ctx) + '\n\nQUESTION: ' + String(body.message || '').slice(0, 2000) }] }]),
    generationConfig: { responseMimeType:'application/json', temperature:0.3 }
  };
  let lastErr = '';
  for (const model of models) {
    let res;
    try {
      res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent',
        { method:'post', contentType:'application/json', headers:{ 'x-goog-api-key': key }, payload:JSON.stringify(payload), muteHttpExceptions:true });
    } catch (e) { lastErr = model + ': ' + e.message; continue; }
    const code = res.getResponseCode();
    let j = {}; try { j = JSON.parse(res.getContentText()); } catch (e) {}
    const emsg = (j.error && j.error.message) || code;
    if (code === 401 || code === 403) return { ok:false, error:'API key reject hui: ' + emsg };
    if (code !== 200) { lastErr = model + ' (' + code + '): ' + emsg; continue; }   // 404/429/5xx → next model
    const parts = (((j.candidates || [])[0] || {}).content || {}).parts;
    const txt = parts ? parts.map(p => p.text || '').join('') : '';
    if (!txt) { lastErr = model + ': empty response'; continue; }
    let out; try { out = JSON.parse(txt); } catch (e) { out = { answer: txt }; }
    return { ok:true, answer: out.answer || '', charts: out.charts || [], slides: out.slides || [], model };
  }
  return { ok:false, error:'Sab AI models fail hue. Last: ' + lastErr };
}

// Run once from the editor to check which models/key work (see View → Logs)
function testAI() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  AI_MODELS.forEach(m => {
    const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent',
      { method:'post', contentType:'application/json', headers:{ 'x-goog-api-key': key },
        payload: JSON.stringify({ contents:[{ parts:[{ text:'Say OK' }] }] }), muteHttpExceptions:true });
    console.log(m + ' → ' + r.getResponseCode());
  });
}
