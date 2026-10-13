/**
 * ATTENDANCE SYSTEM — Google Sheets + Apps Script
 * Setup:
 * 1. Open the target Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Replace Code.gs with this file.
 * 4. Add an HTML file named exactly "Index" and paste Index.html.
 * 5. Save, select setup, click Run, and authorize.
 * 6. Deploy > New deployment > Web app; Execute as Me.
 *
 * IMPORTANT:
 * - Default Admin PIN is 2468. Change it in the Settings sheet after setup.
 * - Employee demo PINs are 1111. Replace demo employees/PINs before real use.
 * - For office geofencing, set Require Location = Yes and enter Office Location
 *   as "latitude, longitude" in Settings.
 */

const TIME_ZONE = 'Asia/Kolkata';
const SHEET_EMPLOYEES = 'Employees';
const SHEET_LOG = 'Attendance_Log';
const SHEET_SETTINGS = 'Settings';
const SHEET_HOLIDAYS = 'Holidays';
const SHEET_REGULARIZATION = 'Regularization_Requests';
const SHEET_AUDIT = 'Admin_Audit_Log';
const SHEET_WORK_SCHEDULES = 'Work_Schedules';
const SHEET_WFH_REQUESTS = 'WFH_Requests';
const SHEET_LEAVE_REQUESTS = 'Leave_Requests';
const SHEET_NOTIFICATIONS = 'Notifications';
const LEAVE_REQUEST_HEADERS = ['Request ID', 'Submitted At', 'Employee ID', 'Employee Name', 'Leave Type', 'Start Date', 'End Date', 'Days', 'Reason', 'Status', 'Reviewed At', 'Reviewed By', 'Admin Remarks'];
const NOTIFICATION_HEADERS = ['Notification ID', 'Created At', 'Audience', 'Recipient ID', 'Title', 'Message', 'Read', 'Related ID'];
const WORK_SCHEDULE_HEADERS = ['Date', 'Employee ID', 'Work Mode', 'Approved By', 'Approved At', 'Notes'];
const WFH_REQUEST_HEADERS = ['Request ID', 'Submitted At', 'Employee ID', 'Employee Name', 'Start Date', 'End Date', 'Reason', 'Status', 'Reviewed At', 'Reviewed By', 'Admin Remarks'];
const REGULARIZATION_HEADERS = ['Request ID', 'Submitted At', 'Employee ID', 'Employee Name', 'Attendance Date', 'Type', 'Requested Login Time', 'Requested Logout Time', 'Reason', 'Status', 'Reviewed At', 'Reviewed By', 'Admin Remarks'];

const EMPLOYEE_HEADERS = [
  'Employee ID', 'Employee Name', 'Designation', 'Department',
  'Email', 'Phone', 'PIN', 'Status', 'Role', 'Manager ID'
];
const LOG_HEADERS = [
  'Date', 'Employee ID', 'Employee Name', 'Designation', 'Department',
  'Login Time', 'Logout Time', 'Working Hours', 'Status',
  'Login Distance (m)', 'Logout Distance (m)', 'Remarks', 'Timestamp'
];
const SETTINGS_HEADERS = ['Setting', 'Value', 'Help'];
const HOLIDAY_HEADERS = ['Date', 'Holiday Name'];

const DEMO_EMPLOYEES = [
  ['EMP001', 'Aarav Mehta', 'Sales Executive', 'Sales', '', '', '1111', 'Active'],
  ['EMP002', 'Priya Sharma', 'HR Manager', 'HR', '', '', '1111', 'Active'],
  ['EMP003', 'Rohan Patel', 'Accountant', 'Accounts', '', '', '1111', 'Active'],
  ['EMP004', 'Sneha Iyer', 'Graphic Designer', 'Marketing', '', '', '1111', 'Active'],
  ['EMP005', 'Vikram Singh', 'Office Assistant', 'Admin', '', '', '1111', 'Active']
];

const DEFAULT_SETTINGS = [
  ['Business Name', 'Attendance System', 'Shown in the portal and dashboard'],
  ['Logo URL', '', 'Optional public HTTPS image URL'],
  ['Brand Color', '#4F46E5', 'Hex colour code'],
  ['Office Start Time', '10:00', '24-hour time, for example 09:30'],
  ['Late After (minutes)', '15', 'Minutes after office start time before marking Late'],
  ['Half Day Below (hours)', '4', 'Hours below which a completed day is Half Day'],
  ['Weekly Off Days', 'Sunday', 'Comma-separated weekdays'],
  ['Allow Logout', 'Yes', 'Yes enables employee logout'],
  ['Require PIN', 'Yes', 'Employee PIN verification'],
  ['Require Location', 'No', 'Enable only after setting Office Location'],
  ['Office Location', '', 'Latitude, longitude; for example 19.0760, 72.8777'],
  ['Allowed Radius (meters)', '200', 'Allowed distance from office'],
  ['Admin PIN', '2468', 'Legacy PIN field; cleared after secure hash migration'],
  ['Admin Email', '', 'Registered email for Admin PIN recovery OTP']
];

const ADMIN_SESSION_MINUTES = 30;
const ADMIN_TOKEN_PREFIX = 'att_admin_';
const MAX_WRONG_PIN = 5;
const LOCKOUT_MINUTES = 15;

function doGet(e) {
  const requestedPage = e && e.parameter && e.parameter.page;
  const page = requestedPage === 'admin' ? 'admin' : (requestedPage === 'manager' ? 'manager' : 'employee');
  const settings = getSettingsSafe_();
  const config = {
    page: page,
    webAppUrl: ScriptApp.getService().getUrl(),
    businessName: settings.businessName || 'Attendance System',
    logoUrl: settings.logoUrl || '',
    brandColor: settings.brandColor || '#4F46E5',
    allowLogout: settings.allowLogout !== false,
    requirePin: settings.requirePin !== false,
    requireLocation: settings.requireLocation === true,
    timeZone: TIME_ZONE
  };
  const template = HtmlService.createTemplateFromFile('Index');
  template.config = JSON.stringify(config).replace(/</g, '\\u003c');
  return template.evaluate()
    .setTitle(config.businessName + (page === 'admin' ? ' — Admin Dashboard' : (page === 'manager' ? ' — Manager Dashboard' : '')))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once from the Apps Script editor while bound to the target spreadsheet. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet. Open the target Google Sheet, then use Extensions > Apps Script.');
  }
  ensureSheet_(ss, SHEET_EMPLOYEES, EMPLOYEE_HEADERS, DEMO_EMPLOYEES);
  ensureSheet_(ss, SHEET_LOG, LOG_HEADERS, []);
  ensureSheet_(ss, SHEET_SETTINGS, SETTINGS_HEADERS, DEFAULT_SETTINGS);
  ensureSheet_(ss, SHEET_HOLIDAYS, HOLIDAY_HEADERS, []);
  ensureSheet_(ss, SHEET_REGULARIZATION, REGULARIZATION_HEADERS, []);
  ensureSheet_(ss, SHEET_WORK_SCHEDULES, WORK_SCHEDULE_HEADERS, []);
  ensureSheet_(ss, SHEET_WFH_REQUESTS, WFH_REQUEST_HEADERS, []);
  ensureSheet_(ss, SHEET_LEAVE_REQUESTS, LEAVE_REQUEST_HEADERS, []);
  ensureSheet_(ss, SHEET_NOTIFICATIONS, NOTIFICATION_HEADERS, []);
  ensureSheet_(ss, SHEET_AUDIT, ['Timestamp', 'Actor', 'Event', 'Details'], []);
  ensureAttendanceModeColumns_();
  formatSheets_(ss);
  SpreadsheetApp.getUi().alert(
    'Setup completed.\n\nTabs created/checked: Employees, Attendance_Log, Settings, Holidays, Regularization_Requests, Work_Schedules, WFH_Requests, Leave_Requests, Notifications, Admin_Audit_Log.\nDefault Admin PIN: 2468\nDemo employee PIN: 1111\nChange these before using the system.'
  );
}

/**
 * SAFE, additive migration. Run manually after making a spreadsheet backup.
 * Adds only missing columns/settings; never clears, deletes, or rewrites existing rows.
 */
function upgradeSchemaSafely_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the target spreadsheet before running this migration.');
  const emp = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!emp) throw new Error('Employees sheet not found. No changes made.');
  const headers = emp.getRange(1, 1, 1, Math.max(emp.getLastColumn(), 1)).getDisplayValues()[0].map(clean_);
  ['Role', 'Manager ID'].forEach(function(name) {
    if (headers.indexOf(name) < 0) {
      const col = emp.getLastColumn() + 1;
      emp.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#EEF2FF');
    }
  });
  saveSetting_('Admin Email', getSettingsMap_()['Admin Email'] || '');
  ensureSheet_(ss, SHEET_WORK_SCHEDULES, WORK_SCHEDULE_HEADERS, []);
  ensureSheet_(ss, SHEET_WFH_REQUESTS, WFH_REQUEST_HEADERS, []);
  ensureSheet_(ss, SHEET_LEAVE_REQUESTS, LEAVE_REQUEST_HEADERS, []);
  ensureSheet_(ss, SHEET_NOTIFICATIONS, NOTIFICATION_HEADERS, []);
  ensureAttendanceModeColumns_();
  if (!ss.getSheetByName(SHEET_AUDIT)) { const audit = ss.insertSheet(SHEET_AUDIT); audit.getRange(1,1,1,4).setValues([['Timestamp','Actor','Event','Details']]).setFontWeight('bold').setBackground('#EEF2FF'); audit.setFrozenRows(1); }
  SpreadsheetApp.getUi().alert('Safe schema upgrade completed. Existing records were not deleted or cleared.\nSet Admin Email in Settings before using OTP recovery.');
}

function ensureSheet_(ss, name, headers, defaultRows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  // Fill headers only if the tab is completely empty.
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    if (defaultRows && defaultRows.length) {
      sheet.getRange(2, 1, defaultRows.length, defaultRows[0].length).setValues(defaultRows);
    }
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatSheets_(ss) {
  const emp = ss.getSheetByName(SHEET_EMPLOYEES);
  const log = ss.getSheetByName(SHEET_LOG);
  const settings = ss.getSheetByName(SHEET_SETTINGS);
  const holidays = ss.getSheetByName(SHEET_HOLIDAYS);
  [emp, log, settings, holidays].forEach(function(s) {
    if (s) {
      s.getRange(1, 1, 1, s.getLastColumn() || 1).setFontWeight('bold').setBackground('#EEF2FF');
      s.setFrozenRows(1);
      s.autoResizeColumns(1, Math.max(s.getLastColumn(), 1));
    }
  });
  if (emp && emp.getLastRow() > 1) emp.getRange(2, 1, emp.getLastRow() - 1, 1).setNumberFormat('@');
  if (log && log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 2).setNumberFormat('@');
}

function getRequiredSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Make sure this project is bound to your Google Sheet.');
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" is missing. Run setup from the Apps Script editor.');
  return sheet;
}

function getSettingsMap_() {
  const sheet = getRequiredSheet_(SHEET_SETTINGS);
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const key = clean_(values[i][0]);
    if (key) map[key] = clean_(values[i][1]);
  }
  return map;
}

function getSettingsSafe_() {
  try {
    const m = getSettingsMap_();
    const office = parseLocation_(m['Office Location'] || '');
    return {
      businessName: m['Business Name'] || 'Attendance System',
      logoUrl: m['Logo URL'] || '',
      brandColor: validHex_(m['Brand Color']) ? m['Brand Color'] : '#4F46E5',
      officeStartTime: m['Office Start Time'] || '10:00',
      lateAfterMinutes: numberOr_(m['Late After (minutes)'], 15),
      halfDayBelowHours: numberOr_(m['Half Day Below (hours)'], 4),
      weeklyOffDays: (m['Weekly Off Days'] || 'Sunday').split(',').map(function(x) { return x.trim().toLowerCase(); }),
      allowLogout: yes_(m['Allow Logout'] || 'Yes'),
      requirePin: yes_(m['Require PIN'] || 'Yes'),
      requireLocation: yes_(m['Require Location'] || 'No'),
      officeLocationText: m['Office Location'] || '',
      officeLocation: office,
      radiusMeters: numberOr_(m['Allowed Radius (meters)'], 200),
      adminPin: m['Admin PIN'] || '2468'
    };
  } catch (err) {
    return {
      businessName: 'Attendance System', logoUrl: '', brandColor: '#4F46E5',
      officeStartTime: '10:00', lateAfterMinutes: 15, halfDayBelowHours: 4,
      weeklyOffDays: ['sunday'], allowLogout: true, requirePin: true,
      requireLocation: false, officeLocationText: '', officeLocation: null,
      radiusMeters: 200, adminPin: '2468'
    };
  }
}

function clean_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}
function yes_(v) {
  return ['yes', 'true', '1', 'on'].indexOf(String(v || '').toLowerCase()) >= 0;
}
function numberOr_(v, fallback) {
  const n = Number(v);
  return isFinite(n) && clean_(v) !== '' ? n : fallback;
}
function validHex_(v) {
  return /^#[0-9a-f]{6}$/i.test(String(v || ''));
}
function parseLocation_(text) {
  const p = String(text || '').split(',');
  if (p.length < 2) return null;
  const lat = Number(p[0].trim()), lng = Number(p[1].trim());
  return isFinite(lat) && isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180
    ? { lat: lat, lng: lng } : null;
}
function today_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
}
function nowTime_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'HH:mm:ss');
}
function timestamp_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
}
function toSeconds_(time) {
  const p = String(time || '').split(':');
  return p.length >= 2 ? Number(p[0]) * 3600 + Number(p[1]) * 60 + Number(p[2] || 0) : null;
}
function hoursBetween_(start, end) {
  const a = toSeconds_(start), b = toSeconds_(end);
  if (a === null || b === null || b <= a) return 0;
  return Math.round(((b - a) / 3600) * 100) / 100;
}
function distanceMeters_(a, b) {
  const rad = Math.PI / 180;
  const dLat = (b.lat - a.lat) * rad;
  const dLng = (b.lng - a.lng) * rad;
  const x = Math.sin(dLat / 2) ** 2 +
    Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLng / 2) ** 2;
  return 6371000 * 2 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
}

function readEmployees_() {
  const sheet = getRequiredSheet_(SHEET_EMPLOYEES);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = (values[0] || []).map(clean_);
  const col = function(name, fallback) { const ix = headers.indexOf(name); return ix >= 0 ? ix : fallback; };
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!clean_(r[0])) continue;
    out.push({
      rowNumber: i + 1, id: clean_(r[0]), name: clean_(r[1]),
      designation: clean_(r[2]), department: clean_(r[3]),
      email: clean_(r[4]), phone: clean_(r[5]), pin: clean_(r[6]),
      status: clean_(r[7]) || 'Active',
      role: clean_(r[col('Role', 8)]) || 'Employee',
      managerId: clean_(r[col('Manager ID', 9)])
    });
  }
  return out;
}
function findEmployee_(id) {
  const wanted = clean_(id).toLowerCase();
  return readEmployees_().find(function(e) { return e.id.toLowerCase() === wanted; }) || null;
}
function getLogSheet_() {
  return getRequiredSheet_(SHEET_LOG);
}
function readLogs_(date) {
  const values = getLogSheet_().getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!clean_(r[0])) continue;
    if (date && clean_(r[0]) !== date) continue;
    out.push({
      rowNumber: i + 1, date: clean_(r[0]), employeeId: clean_(r[1]),
      name: clean_(r[2]), designation: clean_(r[3]), department: clean_(r[4]),
      loginTime: clean_(r[5]), logoutTime: clean_(r[6]),
      hours: Number(r[7]) || 0, status: clean_(r[8]),
      loginDistance: clean_(r[9]), logoutDistance: clean_(r[10]),
      remarks: clean_(r[11]), timestamp: clean_(r[12])
    });
  }
  return out;
}
function findTodayLog_(employeeId) {
  const id = clean_(employeeId).toLowerCase();
  return readLogs_(today_()).find(function(r) { return r.employeeId.toLowerCase() === id; }) || null;
}
function classifyStatus_(loginTime, settings) {
  const actual = toSeconds_(loginTime);
  const start = toSeconds_(settings.officeStartTime);
  if (actual !== null && start !== null && actual > start + settings.lateAfterMinutes * 60) return 'Late';
  return 'Present';
}
function getOrCreateTodayRow_(employee) {
  const existing = findTodayLog_(employee.id);
  if (existing) return existing;
  const sheet = getLogSheet_();
  const row = [
    today_(), employee.id, employee.name, employee.designation, employee.department,
    '', '', '', '', '', '', '', timestamp_()
  ];
  sheet.appendRow(row);
  return {
    rowNumber: sheet.getLastRow(), date: row[0], employeeId: employee.id,
    name: employee.name, designation: employee.designation, department: employee.department,
    loginTime: '', logoutTime: '', hours: 0, status: '', loginDistance: '', logoutDistance: '',
    remarks: '', timestamp: row[12]
  };
}
function employeeLogin(employeeId, pin) {
  try {
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return { ok: false, message: 'Employee ID not found or employee is inactive.' };
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return { ok: false, message: 'Incorrect employee PIN.' };
    const row = getOrCreateTodayRow_(employee);
    if (row.loginTime) return { ok: false, message: 'Attendance login already recorded today at ' + row.loginTime + '.' };
    return { ok: true, message: 'Employee verified. You can now mark attendance.', employee: { id: employee.id, name: employee.name, designation: employee.designation, department: employee.department } };
  } catch (err) {
    return { ok: false, message: err.message || 'Unable to verify employee.' };
  }
}

/** Called by the employee page. PIN is checked on every attendance action. */
function markAttendance(req) {
  const attendanceLock = LockService.getScriptLock();
  let attendanceLockHeld = false;
  try {
    req = req || {};
    const employee = findEmployee_(req.employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return { ok: false, message: 'Employee ID not found or employee is inactive.' };
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(req.pin) !== employee.pin) return { ok: false, message: 'Incorrect employee PIN.' };

    // Work mode is checked on the server. A client cannot self-approve WFH.
    const requestedMode = clean_(req.workMode || 'Office').toLowerCase() === 'wfh' ? 'WFH' : 'Office';
    if (requestedMode === 'WFH' && !hasApprovedWFHSchedule_(employee.id, today_())) {
      return { ok: false, message: 'WFH is not approved for you today. Request WFH and wait for Admin approval.' };
    }

    let distance = '';
    if (requestedMode === 'Office' && settings.requireLocation) {
      if (!settings.officeLocation) return { ok: false, message: 'Admin must set Office Location in Settings before location verification can be used.' };
      if (!req.location || !isFinite(Number(req.location.lat)) || !isFinite(Number(req.location.lng))) return { ok: false, message: 'Please allow location access and try again.' };
      distance = Math.round(distanceMeters_(settings.officeLocation, { lat: Number(req.location.lat), lng: Number(req.location.lng) }));
      if (distance > settings.radiusMeters) return { ok: false, message: 'You are outside the allowed office radius (' + distance + ' m away).' };
    }

    attendanceLock.waitLock(10000);
    attendanceLockHeld = true;
    const action = clean_(req.action).toLowerCase();
    const row = getOrCreateTodayRow_(employee);
    const sheet = getLogSheet_();
    if (clean_(row.status).toLowerCase() === 'leave') return { ok: false, message: 'Your approved leave is recorded for today. Contact Admin if attendance needs to be reopened.' };
    const existingMode = getLogField_(row.rowNumber, 'Work Mode');
    if (existingMode && existingMode !== requestedMode) {
      return { ok: false, message: 'Your attendance is already started in ' + existingMode + ' mode. Contact Admin if this needs correction.' };
    }
    if (action === 'login') {
      if (row.loginTime) return { ok: false, message: 'Login already recorded at ' + row.loginTime + '.' };
      const login = nowTime_();
      sheet.getRange(row.rowNumber, 6).setValue(login);
      sheet.getRange(row.rowNumber, 9).setValue(classifyStatus_(login, settings));
      sheet.getRange(row.rowNumber, 10).setValue(distance);
      sheet.getRange(row.rowNumber, 13).setValue(timestamp_());
      setLogField_(row.rowNumber, 'Work Mode', requestedMode);
      setLogField_(row.rowNumber, 'WFH Approval ID', requestedMode === 'WFH' ? getWFHApprovalId_(employee.id, today_()) : '');
      setLogField_(row.rowNumber, 'Verification Method', requestedMode === 'WFH' ? 'Approved WFH schedule + employee PIN' : (settings.requireLocation ? 'Employee PIN + browser location check' : 'Employee PIN'));
      return { ok: true, message: 'Attendance login saved at ' + login + ' (' + requestedMode + ').', record: { loginTime: login, status: classifyStatus_(login, settings), workMode: requestedMode } };
    }
    if (action === 'logout') {
      if (!settings.allowLogout) return { ok: false, message: 'Logout is disabled by the administrator.' };
      if (!row.loginTime) return { ok: false, message: 'Please record login before logout.' };
      if (row.logoutTime) return { ok: false, message: 'Logout already recorded at ' + row.logoutTime + '.' };
      const logout = nowTime_();
      const hours = hoursBetween_(row.loginTime, logout);
      let status = row.status || 'Present';
      if (settings.halfDayBelowHours > 0 && hours < settings.halfDayBelowHours) status = 'Half Day';
      sheet.getRange(row.rowNumber, 7).setValue(logout);
      sheet.getRange(row.rowNumber, 8).setValue(hours);
      sheet.getRange(row.rowNumber, 9).setValue(status);
      sheet.getRange(row.rowNumber, 11).setValue(distance);
      sheet.getRange(row.rowNumber, 13).setValue(timestamp_());
      setLogField_(row.rowNumber, 'Work Mode', requestedMode);
      setLogField_(row.rowNumber, 'WFH Approval ID', requestedMode === 'WFH' ? getWFHApprovalId_(employee.id, today_()) : '');
      setLogField_(row.rowNumber, 'Verification Method', requestedMode === 'WFH' ? 'Approved WFH schedule + employee PIN' : (settings.requireLocation ? 'Employee PIN + browser location check' : 'Employee PIN'));
      return { ok: true, message: 'Attendance logout saved at ' + logout + '. Working hours: ' + hours + '.', record: { logoutTime: logout, hours: hours, status: status, workMode: requestedMode } };
    }
    return { ok: false, message: 'Unknown attendance action.' };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not save attendance.' };
  } finally {
    if (attendanceLockHeld) {
      try { attendanceLock.releaseLock(); } catch (e) {}
    }
  }
}

function bytesToHex_(bytes) {
  return bytes.map(function(b) { const v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
}
function pinHash_(pin, salt) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(pin), Utilities.Charset.UTF_8));
}
function getAdminPinCredentials_() {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('ADMIN_PIN_SALT');
  let hash = props.getProperty('ADMIN_PIN_HASH');
  if (salt && hash) return { salt: salt, hash: hash, migrated: true };
  const settings = getSettingsSafe_();
  const legacy = settings.adminPin || '2468';
  salt = Utilities.getUuid() + Utilities.getUuid();
  hash = pinHash_(legacy, salt);
  props.setProperty('ADMIN_PIN_SALT', salt);
  props.setProperty('ADMIN_PIN_HASH', hash);
  props.setProperty('ADMIN_SESSION_VERSION', props.getProperty('ADMIN_SESSION_VERSION') || '1');
  // Remove the plain-text PIN from the spreadsheet after preserving it as a hash.
  saveSetting_('Admin PIN', '');
  return { salt: salt, hash: hash, migrated: false };
}
function writeAuditLog_(actor, eventName, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    let sheet = ss.getSheetByName(SHEET_AUDIT);
    if (!sheet) { sheet = ss.insertSheet(SHEET_AUDIT); sheet.appendRow(['Timestamp','Actor','Event','Details']); sheet.setFrozenRows(1); }
    sheet.appendRow([timestamp_(), clean_(actor) || 'System', clean_(eventName), clean_(details)]);
  } catch (e) { /* audit logging must not expose secrets or block safe authentication */ }
}

function adminLogin(pin) {
  try {
    const cache = CacheService.getScriptCache();
    const failKey = 'admin_pin_fail';
    const lockKey = 'admin_pin_lock';
    if (cache.get(lockKey)) return { ok: false, message: 'Too many failed attempts. Try again in 15 minutes.' };
    const creds = getAdminPinCredentials_();
    if (pinHash_(clean_(pin), creds.salt) !== creds.hash) {
      const count = Number(cache.get(failKey) || 0) + 1;
      if (count >= MAX_WRONG_PIN) {
        cache.put(lockKey, '1', LOCKOUT_MINUTES * 60); cache.remove(failKey);
        return { ok: false, message: 'Too many failed attempts. Try again in 15 minutes.' };
      }
      cache.put(failKey, String(count), LOCKOUT_MINUTES * 60);
      return { ok: false, message: 'Incorrect Admin PIN. Attempts remaining: ' + (MAX_WRONG_PIN - count) };
    }
    cache.remove(failKey); cache.remove(lockKey);
    writeAuditLog_('Admin', 'Admin login', 'Successful PIN authentication');
    const token = Utilities.getUuid();
    const version = PropertiesService.getScriptProperties().getProperty('ADMIN_SESSION_VERSION') || '1';
    cache.put(ADMIN_TOKEN_PREFIX + token, version, ADMIN_SESSION_MINUTES * 60);
    return { ok: true, token: token };
  } catch (err) { return { ok: false, message: err.message || 'Admin login failed.' }; }
}

/** Run manually once in the Apps Script editor to grant MailApp authorization before enabling OTP recovery. */
function authorizeAdminOtpEmail_() { return MailApp.getRemainingDailyQuota(); }

/** Sends a short-lived OTP only to the configured Admin Email. Response is generic to avoid email enumeration. */
function requestAdminPinReset(email) {
  const generic = { ok: true, message: 'If the email matches the registered admin email, an OTP will be sent shortly.' };
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const normalized = clean_(email).toLowerCase();
  const lastSent = Number(props.getProperty('ADMIN_OTP_LAST_SENT') || 0);
  if (now - lastSent < 60000) return { ok: false, message: 'Please wait 60 seconds before requesting another OTP.' };
  const settings = getSettingsMap_();
  const adminEmail = clean_(settings['Admin Email']).toLowerCase();
  if (!adminEmail || normalized !== adminEmail) return generic;
  if (!/^\S+@\S+\.\S+$/.test(normalized)) return generic;
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid();
  props.setProperty('ADMIN_OTP_SALT', salt);
  props.setProperty('ADMIN_OTP_HASH', pinHash_(otp, salt));
  props.setProperty('ADMIN_OTP_EMAIL', normalized);
  props.setProperty('ADMIN_OTP_EXPIRES', String(now + 10 * 60 * 1000));
  props.setProperty('ADMIN_OTP_ATTEMPTS', '0');
  props.setProperty('ADMIN_OTP_LAST_SENT', String(now));
  try {
    MailApp.sendEmail({to: adminEmail, subject:'Attendance Portal Admin PIN reset code', body:'Your Attendance Portal Admin PIN reset code is: ' + otp + '\n\nThis code expires in 10 minutes and can be used once. If you did not request this, ignore this email.'});
  } catch (mailError) {
    ['ADMIN_OTP_SALT','ADMIN_OTP_HASH','ADMIN_OTP_EMAIL','ADMIN_OTP_EXPIRES','ADMIN_OTP_ATTEMPTS','ADMIN_OTP_LAST_SENT'].forEach(function(key) { props.deleteProperty(key); });
    return {ok:false, message:'Email could not be sent. Open Apps Script, select authorizeAdminOtpEmail_, click Run, and approve the requested email permission. Then try again.'};
  }
  writeAuditLog_('Admin Email', 'Admin PIN OTP requested', 'OTP issued; code not logged');
  return generic;
}
function verifyAdminPinReset(email, otp, newPin) {
  const props = PropertiesService.getScriptProperties();
  const normalized = clean_(email).toLowerCase();
  const attempts = Number(props.getProperty('ADMIN_OTP_ATTEMPTS') || 0);
  if (attempts >= 5) return {ok:false,message:'Too many incorrect OTP attempts. Request a new code.'};
  const settings = getSettingsMap_();
  const salt = props.getProperty('ADMIN_OTP_SALT') || '';
  const valid = normalized && normalized === clean_(settings['Admin Email']).toLowerCase() &&
    normalized === (props.getProperty('ADMIN_OTP_EMAIL') || '') &&
    Date.now() <= Number(props.getProperty('ADMIN_OTP_EXPIRES') || 0) &&
    salt && pinHash_(clean_(otp), salt) === (props.getProperty('ADMIN_OTP_HASH') || '');
  if (!valid) {
    props.setProperty('ADMIN_OTP_ATTEMPTS', String(attempts + 1));
    return {ok:false,message:'OTP is invalid or expired. Check the code or request a new one.'};
  }
  newPin = clean_(newPin);
  if (!/^[0-9A-Za-z!@#$%^&*()_+\-=]{8,64}$/.test(newPin)) return {ok:false,message:'New PIN must be 8–64 characters and may include letters, numbers or symbols.'};
  const newSalt = Utilities.getUuid() + Utilities.getUuid();
  props.setProperty('ADMIN_PIN_SALT', newSalt);
  props.setProperty('ADMIN_PIN_HASH', pinHash_(newPin, newSalt));
  props.setProperty('ADMIN_SESSION_VERSION', Utilities.getUuid()); // revokes every existing admin session
  ['ADMIN_OTP_SALT','ADMIN_OTP_HASH','ADMIN_OTP_EMAIL','ADMIN_OTP_EXPIRES','ADMIN_OTP_ATTEMPTS','ADMIN_OTP_LAST_SENT'].forEach(function(k){props.deleteProperty(k);});
  saveSetting_('Admin PIN', '');
  writeAuditLog_('Admin Email', 'Admin PIN reset', 'OTP verified; previous sessions revoked');
  return {ok:true,message:'Admin PIN updated. Please sign in using your new PIN.'};
}

/** Owner-controlled emergency reset. Run only from the Apps Script editor, never from the web client.
 * To use: create a temporary editor-only wrapper calling ownerResetAdminPin_('YourNewStrongPIN'), run it,
 * then delete the wrapper. This underscore-suffixed function is not exposed to google.script.run.
 */
function ownerResetAdminPin_(newPin) {
  newPin = clean_(newPin);
  if (!/^[0-9A-Za-z!@#$%^&*()_+\-=]{8,64}$/.test(newPin)) throw new Error('Supply a strong 8–64 character PIN/password from the Apps Script editor.');
  const props = PropertiesService.getScriptProperties();
  const salt = Utilities.getUuid() + Utilities.getUuid();
  props.setProperty('ADMIN_PIN_SALT', salt);
  props.setProperty('ADMIN_PIN_HASH', pinHash_(newPin, salt));
  props.setProperty('ADMIN_SESSION_VERSION', Utilities.getUuid());
  saveSetting_('Admin PIN', '');
  ['ADMIN_OTP_SALT','ADMIN_OTP_HASH','ADMIN_OTP_EMAIL','ADMIN_OTP_EXPIRES','ADMIN_OTP_ATTEMPTS','ADMIN_OTP_LAST_SENT'].forEach(function(k){props.deleteProperty(k);});
  writeAuditLog_('Script owner', 'Emergency Admin PIN reset', 'Owner-controlled reset; previous sessions revoked');
  return 'Admin PIN reset completed; all existing sessions were revoked.';
}

function isAdmin_(token) {
  if (!token) return false;
  const cache = CacheService.getScriptCache();
  const key = ADMIN_TOKEN_PREFIX + token;
  const currentVersion = PropertiesService.getScriptProperties().getProperty('ADMIN_SESSION_VERSION') || '1';
  if (cache.get(key) !== currentVersion) return false;
  // Sliding inactivity timeout: authenticated requests renew the 30-minute session.
  cache.put(key, currentVersion, ADMIN_SESSION_MINUTES * 60);
  return true;
}
function validateAdminSession(token) {
  return isAdmin_(token)
    ? { ok: true }
    : { ok: false, authFailed: true, message: 'Admin session expired. Sign in again.' };
}
function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove(ADMIN_TOKEN_PREFIX + token);
  return { ok: true };
}
function getDashboard(token, date) {
  if (!isAdmin_(token)) return { ok: false, authFailed: true, message: 'Admin session expired. Sign in again.' };
  try {
    const selectedDate = /^\d{4}-\d{2}-\d{2}$/.test(clean_(date)) ? clean_(date) : today_();
    const employees = readEmployees_().filter(function(e) { return String(e.status).toLowerCase() !== 'inactive'; });
    const logs = readLogs_(selectedDate);
    const byId = {};
    logs.forEach(function(r) { byId[r.employeeId.toLowerCase()] = r; });
    const daily = employees.map(function(e) {
      const r = byId[e.id.toLowerCase()];
      return {
        employeeId: e.id, name: e.name, designation: e.designation, department: e.department,
        loginTime: r ? r.loginTime : '', logoutTime: r ? r.logoutTime : '',
        hours: r ? r.hours : 0, workMode: r ? (getLogField_(r.rowNumber, 'Work Mode') || 'Office') : '—', status: r ? (r.status || (r.loginTime ? 'Present' : 'Absent')) : 'Absent'
      };
    });
    const counts = { Present: 0, Late: 0, 'Half Day': 0, Leave: 0, Absent: 0 };
    daily.forEach(function(r) { counts[counts[r.status] === undefined ? 'Absent' : r.status]++; });
    const total = daily.length;
    let wfhToday = 0, pendingWfh = 0, pendingLeave = 0, pendingRegularization = 0;
    try {
      const scheduleRows = getRequiredSheet_(SHEET_WORK_SCHEDULES).getDataRange().getDisplayValues();
      for (let i=1;i<scheduleRows.length;i++) if (clean_(scheduleRows[i][0])===selectedDate && clean_(scheduleRows[i][2]).toLowerCase()==='wfh') wfhToday++;
    } catch (e) {}
    try { const v=getRequiredSheet_(SHEET_WFH_REQUESTS).getDataRange().getDisplayValues(); for(let i=1;i<v.length;i++) if(clean_(v[i][7])==='Pending') pendingWfh++; } catch(e) {}
    try { const v=getRequiredSheet_(SHEET_LEAVE_REQUESTS).getDataRange().getDisplayValues(); for(let i=1;i<v.length;i++) if(clean_(v[i][9])==='Pending') pendingLeave++; } catch(e) {}
    try { const v=getRequiredSheet_(SHEET_REGULARIZATION).getDataRange().getDisplayValues(); for(let i=1;i<v.length;i++) if(clean_(v[i][9])==='Pending') pendingRegularization++; } catch(e) {}
    return {
      ok: true, date: selectedDate, daily: daily,
      kpi: {
        total: total, present: counts.Present, late: counts.Late,
        halfDay: counts['Half Day'], leave: counts.Leave, absent: counts.Absent,
        wfh: wfhToday, pendingWfh: pendingWfh, pendingLeave: pendingLeave, pendingRegularization: pendingRegularization,
        attendanceRate: total ? Math.round(((counts.Present + counts.Late + counts['Half Day']) / total) * 100) : 0
      }
    };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not load dashboard.' };
  }
}
function getAdminData(token) {
  if (!isAdmin_(token)) return { ok: false, authFailed: true, message: 'Admin session expired. Sign in again.' };
  try {
    const settingsMap = getSettingsMap_();
    const safeSettings = {};
    Object.keys(settingsMap).forEach(function(k) {
      // Never send the stored Admin PIN to the client.
      if (k !== 'Admin PIN') safeSettings[k] = settingsMap[k];
    });
    return {
      ok: true,
      employees: readEmployees_().map(function(e) {
        return { id: e.id, name: e.name, designation: e.designation, department: e.department, email: e.email, phone: e.phone, status: e.status, role: e.role, managerId: e.managerId };
      }),
      settings: safeSettings
    };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not load admin data.' };
  }
}
function saveSettings(token, vals) {
  if (!isAdmin_(token)) return { ok: false, authFailed: true, message: 'Admin session expired. Sign in again.' };
  try {
    vals = vals || {};
    const allowed = [
      'Business Name', 'Logo URL', 'Brand Color', 'Office Start Time',
      'Late After (minutes)', 'Half Day Below (hours)', 'Weekly Off Days',
      'Allow Logout', 'Require PIN', 'Require Location', 'Office Location',
      'Allowed Radius (meters)', 'Admin PIN', 'Admin Email'
    ];
    allowed.forEach(function(key) {
      if (Object.prototype.hasOwnProperty.call(vals, key)) {
        const value = clean_(vals[key]);
        if (key === 'Brand Color' && value && !validHex_(value)) throw new Error('Brand Color must be a hex colour such as #4F46E5.');
        if (key === 'Admin PIN' && value && value.length < 8) throw new Error('For security, new Admin PIN/password must contain at least 8 characters. Use Forgot PIN to change it after hash migration.');
        if (key === 'Admin Email' && value && !/^\S+@\S+\.\S+$/.test(value)) throw new Error('Enter a valid Admin Email address.');
        if (key === 'Admin PIN' && value) {
          const salt = Utilities.getUuid()+Utilities.getUuid();
          const props = PropertiesService.getScriptProperties();
          props.setProperty('ADMIN_PIN_SALT', salt); props.setProperty('ADMIN_PIN_HASH', pinHash_(value, salt));
          props.setProperty('ADMIN_SESSION_VERSION', Utilities.getUuid()); saveSetting_('Admin PIN','');
        } else {
          if (key === 'Office Location' && value && !parseLocation_(value)) throw new Error('Office Location must be latitude, longitude.');
          saveSetting_(key, value);
        }
      }
    });
    return { ok: true, message: 'Settings saved.' };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not save settings.' };
  }
}
function saveSetting_(name, value) {
  const sheet = getRequiredSheet_(SHEET_SETTINGS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (clean_(values[i][0]) === name) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([name, value, '']);
}
function addEmployee(token, emp) {
  if (!isAdmin_(token)) return { ok: false, authFailed: true, message: 'Admin session expired. Sign in again.' };
  try {
    emp = emp || {};
    const id = clean_(emp.id), name = clean_(emp.name), pin = clean_(emp.pin);
    if (!id || !name || !pin) return { ok: false, message: 'Employee ID, name and PIN are required.' };
    if (findEmployee_(id)) return { ok: false, message: 'That Employee ID already exists.' };
    getRequiredSheet_(SHEET_EMPLOYEES).appendRow([
      id, name, clean_(emp.designation), clean_(emp.department), clean_(emp.email),
      clean_(emp.phone), pin, clean_(emp.status) || 'Active', clean_(emp.role) === 'Manager' ? 'Manager' : 'Employee', clean_(emp.managerId)
    ]);
    return { ok: true, message: 'Employee added.' };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not add employee.' };
  }
}
/** Manager-only team dashboard. Manager role and team scope are read from the Employees sheet, never from client input. */
function getManagerDashboard(employeeId, pin, date) {
  try {
    const manager = findEmployee_(employeeId);
    if (!manager || String(manager.status).toLowerCase() === 'inactive' || String(manager.role).toLowerCase() !== 'manager') return {ok:false,authFailed:true,message:'Manager access is not enabled for this account.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== manager.pin) return {ok:false,authFailed:true,message:'Incorrect employee PIN.'};
    const selectedDate = /^\d{4}-\d{2}-\d{2}$/.test(clean_(date)) ? clean_(date) : today_();
    const team = readEmployees_().filter(function(e){ return String(e.status).toLowerCase() !== 'inactive' && clean_(e.managerId).toLowerCase() === manager.id.toLowerCase(); });
    const logs = readLogs_(selectedDate); const byId = {};
    logs.forEach(function(r){byId[r.employeeId.toLowerCase()]=r;});
    const daily = team.map(function(e){const r=byId[e.id.toLowerCase()];return {employeeId:e.id,name:e.name,designation:e.designation,department:e.department,loginTime:r?r.loginTime:'',logoutTime:r?r.logoutTime:'',hours:r?r.hours:0,status:r?(r.status||(r.loginTime?'Present':'Absent')):'Absent'};});
    const counts={Present:0,Late:0,'Half Day':0,Leave:0,Absent:0};
    daily.forEach(function(r){counts[counts[r.status]===undefined?'Absent':r.status]++;});
    return {ok:true,manager:{id:manager.id,name:manager.name},date:selectedDate,daily:daily,kpi:{total:daily.length,present:counts.Present,late:counts.Late,halfDay:counts['Half Day'],leave:counts.Leave,absent:counts.Absent}};
  } catch(err){return {ok:false,message:err.message||'Could not load manager dashboard.'};}
}

function getEmployeeToday(employeeId, pin) {
  try {
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return { ok: false, message: 'Employee not found or inactive.' };
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return { ok: false, message: 'Incorrect employee PIN.' };
    const row = findTodayLog_(employee.id);
    return {
      ok: true,
      employee: { id: employee.id, name: employee.name, designation: employee.designation, department: employee.department },
      record: row ? { loginTime: row.loginTime, logoutTime: row.logoutTime, hours: row.hours, status: row.status } : null,
      allowLogout: settings.allowLogout
    };
  } catch (err) {
    return { ok: false, message: err.message || 'Could not load attendance status.' };
  }
}



/** Returns the signed-in employee's recent attendance rows. */
function getEmployeeAttendanceHistory(employeeId, pin) {
  try {
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const rows = readLogs_().filter(function(r){return r.employeeId.toLowerCase() === employee.id.toLowerCase();});
    rows.sort(function(a,b){return b.date.localeCompare(a.date);});
    return {ok:true,records:rows.slice(0,30).map(function(r){return {date:r.date,loginTime:r.loginTime,logoutTime:r.logoutTime,hours:r.hours,status:r.status || (r.loginTime?'Present':'Absent')};})};
  } catch(err) { return {ok:false,message:err.message || 'Could not load attendance history.'}; }
}

/** Returns only the signed-in employee's own regularization requests. */
function getEmployeeRequests(employeeId, pin) {
  try {
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const values = getRequiredSheet_(SHEET_REGULARIZATION).getDataRange().getDisplayValues();
    const requests=[];
    for(let i=1;i<values.length;i++) {
      const r=values[i];
      if(clean_(r[2]).toLowerCase() !== employee.id.toLowerCase()) continue;
      requests.push({requestId:r[0],submittedAt:r[1],date:r[4],type:r[5],loginTime:r[6],logoutTime:r[7],reason:r[8],status:r[9],adminRemarks:r[12]});
    }
    requests.reverse();
    return {ok:true,requests:requests};
  } catch(err) { return {ok:false,message:err.message || 'Could not load your requests.'}; }
}

/** Returns published holidays from the Holidays sheet. */
function getHolidays() {
  try {
    const values=getRequiredSheet_(SHEET_HOLIDAYS).getDataRange().getDisplayValues();
    const holidays=[];
    for(let i=1;i<values.length;i++) {
      if(!clean_(values[i][0]) && !clean_(values[i][1])) continue;
      holidays.push({date:clean_(values[i][0]),name:clean_(values[i][1]) || 'Holiday'});
    }
    holidays.sort(function(a,b){return a.date.localeCompare(b.date);});
    return {ok:true,holidays:holidays};
  } catch(err) { return {ok:false,message:err.message || 'Could not load holidays.'}; }
}

/** Department listing is scoped to the authenticated employee's own department. */
function getDepartmentMembers(employeeId, pin) {
  try {
    const employee=findEmployee_(employeeId);
    if(!employee || String(employee.status).toLowerCase()==='inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings=getSettingsSafe_();
    if(settings.requirePin && clean_(pin)!==employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const department=clean_(employee.department);
    const members=readEmployees_().filter(function(e){
      return String(e.status).toLowerCase()!=='inactive' && clean_(e.department).toLowerCase()===department.toLowerCase();
    }).map(function(e){return {id:e.id,name:e.name,designation:e.designation};});
    return {ok:true,department:department,members:members};
  } catch(err) { return {ok:false,message:err.message || 'Could not load department members.'}; }
}

/** Employee submits a regularization request; this does not change attendance until approved. */
function submitRegularizationRequest(employeeId, pin, payload) {
  const lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return {ok:false, message:'Employee not found or inactive.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return {ok:false, message:'Incorrect employee PIN.'};
    const date = clean_(payload.date);
    const type = clean_(payload.type);
    const login = clean_(payload.loginTime);
    const logout = clean_(payload.logoutTime);
    const reason = clean_(payload.reason);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return {ok:false,message:'Choose a valid attendance date.'};
    if (date > today_()) return {ok:false,message:'Future dates cannot be regularized.'};
    const allowed = ['Missed Login / Check-in','Missed Logout / Check-out','Incorrect Attendance Time','Absent but Worked'];
    if (allowed.indexOf(type) < 0) return {ok:false,message:'Choose a valid regularization type.'};
    if (!reason) return {ok:false,message:'Please enter a reason for the request.'};
    const validTime = t => !t || /^([01]\d|2[0-3]):[0-5]\d$/.test(t);
    if (!validTime(login) || !validTime(logout)) return {ok:false,message:'Enter times in HH:MM 24-hour format.'};
    if ((type === 'Missed Login / Check-in' || type === 'Incorrect Attendance Time' || type === 'Absent but Worked') && !login) return {ok:false,message:'Requested login time is required for this request type.'};
    if ((type === 'Missed Logout / Check-out' || type === 'Incorrect Attendance Time' || type === 'Absent but Worked') && !logout) return {ok:false,message:'Requested logout time is required for this request type.'};
    lock.waitLock(10000);
    const sheet = getRequiredSheet_(SHEET_REGULARIZATION);
    const values = sheet.getDataRange().getDisplayValues();
    for (let i=1;i<values.length;i++) {
      if (clean_(values[i][2]).toLowerCase() === employee.id.toLowerCase() && clean_(values[i][4]) === date && clean_(values[i][9]) === 'Pending') {
        return {ok:false,message:'You already have a pending regularization request for this date.'};
      }
    }
    const requestId = 'REG-' + Utilities.getUuid().slice(0,8).toUpperCase();
    sheet.appendRow([requestId,timestamp_(),employee.id,employee.name,date,type,login,logout,reason,'Pending','','','']);
    return {ok:true,message:'Regularization request submitted to Admin for review.',requestId:requestId};
  } catch (err) { return {ok:false,message:err.message || 'Could not submit request.'}; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function getRegularizationRequests(token) {
  if (!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    const sheet = getRequiredSheet_(SHEET_REGULARIZATION);
    const values = sheet.getDataRange().getDisplayValues();
    const requests = [];
    for (let i=1;i<values.length;i++) {
      const r=values[i]; if (!clean_(r[0])) continue;
      requests.push({requestId:r[0],submittedAt:r[1],employeeId:r[2],employeeName:r[3],date:r[4],type:r[5],loginTime:r[6],logoutTime:r[7],reason:r[8],status:r[9],reviewedAt:r[10],reviewedBy:r[11],adminRemarks:r[12]});
    }
    requests.reverse();
    return {ok:true,requests:requests};
  } catch(err) { return {ok:false,message:err.message || 'Could not load regularization requests.'}; }
}

function reviewRegularizationRequest(token, requestId, decision, remarks) {
  if (!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  const lock=LockService.getScriptLock();
  try {
    decision=clean_(decision); remarks=clean_(remarks);
    if (decision !== 'Approved' && decision !== 'Rejected') return {ok:false,message:'Invalid review decision.'};
    lock.waitLock(10000);
    const reqSheet=getRequiredSheet_(SHEET_REGULARIZATION);
    const values=reqSheet.getDataRange().getDisplayValues();
    let rowNum=0, r=null;
    for(let i=1;i<values.length;i++){ if(clean_(values[i][0])===clean_(requestId)){rowNum=i+1;r=values[i];break;} }
    if(!r) return {ok:false,message:'Regularization request not found.'};
    if(clean_(r[9])!=='Pending') return {ok:false,message:'This request has already been reviewed.'};
    if(decision==='Approved') {
      const logSheet=getLogSheet_();
      const logs=logSheet.getDataRange().getDisplayValues();
      let logRow=0;
      for(let i=1;i<logs.length;i++){if(clean_(logs[i][0])===clean_(r[4]) && clean_(logs[i][1]).toLowerCase()===clean_(r[2]).toLowerCase()){logRow=i+1;break;}}
      if(!logRow){
        const emp=findEmployee_(r[2]); if(!emp) return {ok:false,message:'Employee record no longer exists.'};
        logSheet.appendRow([r[4],emp.id,emp.name,emp.designation,emp.department,'','','','', '', '', 'Regularization approved: '+r[0],timestamp_()]);
        logRow=logSheet.getLastRow();
      }
      const type=clean_(r[5]), requestedLogin=clean_(r[6]), requestedLogout=clean_(r[7]);
      const current=logSheet.getRange(logRow,1,1,13).getDisplayValues()[0];
      if ((type==='Missed Login / Check-in' || type==='Incorrect Attendance Time' || type==='Absent but Worked') && requestedLogin) logSheet.getRange(logRow,6).setValue(requestedLogin+':00');
      if ((type==='Missed Logout / Check-out' || type==='Incorrect Attendance Time' || type==='Absent but Worked') && requestedLogout) logSheet.getRange(logRow,7).setValue(requestedLogout+':00');
      const newLogin=(type==='Missed Login / Check-in' || type==='Incorrect Attendance Time' || type==='Absent but Worked') && requestedLogin ? requestedLogin+':00' : clean_(current[5]);
      const newLogout=(type==='Missed Logout / Check-out' || type==='Incorrect Attendance Time' || type==='Absent but Worked') && requestedLogout ? requestedLogout+':00' : clean_(current[6]);
      if(newLogin) logSheet.getRange(logRow,9).setValue(classifyStatus_(newLogin,getSettingsSafe_()));
      if(newLogin && newLogout) logSheet.getRange(logRow,8).setValue(hoursBetween_(newLogin,newLogout));
      logSheet.getRange(logRow,12).setValue((clean_(current[11]) ? clean_(current[11])+' | ' : '')+'Regularization approved: '+r[0]);
      logSheet.getRange(logRow,13).setValue(timestamp_());
    }
    reqSheet.getRange(rowNum,10).setValue(decision);
    reqSheet.getRange(rowNum,11).setValue(timestamp_());
    reqSheet.getRange(rowNum,12).setValue('Admin');
    reqSheet.getRange(rowNum,13).setValue(remarks);
    return {ok:true,message:'Request '+decision.toLowerCase()+'.'+(decision==='Approved'?' Attendance record updated.':'')};
  } catch(err) { return {ok:false,message:err.message || 'Could not review request.'}; }
  finally { try {lock.releaseLock();} catch(e){} }
}


/* ========================= OFFICE / WORK-FROM-HOME CONTROLS ========================= */

/** Additive migration for attendance mode fields and WFH workflow tabs. Safe to run repeatedly. */
function ensureAttendanceModeColumns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the bound attendance spreadsheet first.');
  const log = ss.getSheetByName(SHEET_LOG);
  if (!log) throw new Error('Attendance_Log sheet not found. Run setup first.');
  const wanted = ['Work Mode', 'WFH Approval ID', 'Verification Method'];
  let headers = log.getRange(1, 1, 1, Math.max(log.getLastColumn(), 1)).getDisplayValues()[0].map(clean_);
  wanted.forEach(function(name) {
    if (headers.indexOf(name) < 0) {
      const col = log.getLastColumn() + 1;
      log.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#EEF2FF');
      headers.push(name);
    }
  });
  ensureSheet_(ss, SHEET_WORK_SCHEDULES, WORK_SCHEDULE_HEADERS, []);
  ensureSheet_(ss, SHEET_WFH_REQUESTS, WFH_REQUEST_HEADERS, []);
}

function logHeaderIndex_(headerName) {
  const sheet = getLogSheet_();
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0].map(clean_);
  return headers.indexOf(headerName) + 1;
}
function getLogField_(rowNumber, headerName) {
  const col = logHeaderIndex_(headerName);
  return col > 0 ? clean_(getLogSheet_().getRange(rowNumber, col).getDisplayValue()) : '';
}
function setLogField_(rowNumber, headerName, value) {
  let col = logHeaderIndex_(headerName);
  if (col < 1) {
    ensureAttendanceModeColumns_();
    col = logHeaderIndex_(headerName);
  }
  if (col > 0) getLogSheet_().getRange(rowNumber, col).setValue(value);
}
function validIsoDate_(value) { return /^\d{4}-\d{2}-\d{2}$/.test(clean_(value)); }
function hasApprovedWFHSchedule_(employeeId, date) {
  const sheet = getRequiredSheet_(SHEET_WORK_SCHEDULES);
  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (clean_(r[0]) === clean_(date) && clean_(r[1]).toLowerCase() === clean_(employeeId).toLowerCase() && clean_(r[2]).toLowerCase() === 'wfh') return true;
  }
  return false;
}
function getWFHApprovalId_(employeeId, date) {
  const values = getRequiredSheet_(SHEET_WORK_SCHEDULES).getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (clean_(r[0]) === clean_(date) && clean_(r[1]).toLowerCase() === clean_(employeeId).toLowerCase() && clean_(r[2]).toLowerCase() === 'wfh') return 'WFH-' + clean_(r[0]) + '-' + clean_(r[1]);
  }
  return '';
}

/** Employee WFH request. Requests do not grant permission until an admin approves them. */
function requestWFH(employeeId, pin, payload) {
  const lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const start = clean_(payload.startDate), end = clean_(payload.endDate), reason = clean_(payload.reason);
    if (!validIsoDate_(start) || !validIsoDate_(end) || end < start) return {ok:false,message:'Choose a valid WFH date range.'};
    if (start < today_()) return {ok:false,message:'WFH requests cannot start in the past.'};
    if (!reason || reason.length < 5) return {ok:false,message:'Please provide a reason (at least 5 characters).'};
    const daySpan = Math.round((Date.parse(end + 'T00:00:00Z') - Date.parse(start + 'T00:00:00Z')) / 86400000) + 1;
    if (daySpan > 31) return {ok:false,message:'A single WFH request cannot exceed 31 calendar days.'};
    lock.waitLock(10000);
    const sheet = getRequiredSheet_(SHEET_WFH_REQUESTS);
    const values = sheet.getDataRange().getDisplayValues();
    for (let i=1;i<values.length;i++) {
      const r=values[i];
      if (clean_(r[2]).toLowerCase() === employee.id.toLowerCase() && clean_(r[7]) === 'Pending' && clean_(r[4]) <= end && clean_(r[5]) >= start) {
        return {ok:false,message:'You already have a pending WFH request overlapping these dates.'};
      }
    }
    const id = 'WFHR-' + Utilities.getUuid().slice(0,8).toUpperCase();
    sheet.appendRow([id,timestamp_(),employee.id,employee.name,start,end,reason,'Pending','','','']);
    createNotification_('Admin','', 'New WFH request', employee.name+' requested WFH from '+start+' to '+end+'.', id);
    createNotification_('Employee',employee.id,'WFH request submitted','Your WFH request ('+start+' to '+end+') is pending Admin review.',id);
    writeAuditLog_(employee.id,'WFH request submitted',id+' '+start+' to '+end);
    return {ok:true,message:'WFH request submitted. Attendance mode will be enabled only after Admin approval.',requestId:id};
  } catch (err) { return {ok:false,message:err.message || 'Could not submit WFH request.'}; }
  finally { try { lock.releaseLock(); } catch(e) {} }
}

/** Employee can see only their own WFH requests and today's server-side approved mode. */
function getEmployeeWFHStatus(employeeId, pin) {
  try {
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const values = getRequiredSheet_(SHEET_WFH_REQUESTS).getDataRange().getDisplayValues();
    const requests=[];
    for (let i=1;i<values.length;i++) {
      const r=values[i]; if (clean_(r[2]).toLowerCase() !== employee.id.toLowerCase()) continue;      requests.push({requestId:r[0],submittedAt:r[1],startDate:r[4],endDate:r[5],reason:r[6],status:r[7],adminRemarks:r[10]});
    }
    requests.reverse();
    return {ok:true,today:today_(),wfhApprovedToday:hasApprovedWFHSchedule_(employee.id,today_()),requests:requests};
  } catch(err) { return {ok:false,message:err.message || 'Could not load WFH status.'}; }
}

/** Admin-only list of WFH requests. */
function getWFHRequests(token) {
  if (!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    const values=getRequiredSheet_(SHEET_WFH_REQUESTS).getDataRange().getDisplayValues();
    const requests=[];
    for(let i=1;i<values.length;i++) { const r=values[i]; if(!clean_(r[0])) continue; requests.push({requestId:r[0],submittedAt:r[1],employeeId:r[2],employeeName:r[3],startDate:r[4],endDate:r[5],reason:r[6],status:r[7],reviewedAt:r[8],reviewedBy:r[9],adminRemarks:r[10]}); }
    requests.reverse(); return {ok:true,requests:requests};
  } catch(err) { return {ok:false,message:err.message || 'Could not load WFH requests.'}; }
}

/** Admin approves/rejects a request. Approval creates one server-side schedule row per date. */
function reviewWFHRequest(token, requestId, decision, remarks) {
  if (!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  const lock=LockService.getScriptLock();
  try {
    decision=clean_(decision); remarks=clean_(remarks);
    if (decision !== 'Approved' && decision !== 'Rejected') return {ok:false,message:'Invalid review decision.'};
    lock.waitLock(10000);
    const sheet=getRequiredSheet_(SHEET_WFH_REQUESTS), values=sheet.getDataRange().getDisplayValues();
    let rowNum=0,r=null;
    for(let i=1;i<values.length;i++) if(clean_(values[i][0])===clean_(requestId)){rowNum=i+1;r=values[i];break;}
    if(!r) return {ok:false,message:'WFH request not found.'};
    if(clean_(r[7])!=='Pending') return {ok:false,message:'This WFH request has already been reviewed.'};
    const employee=findEmployee_(r[2]);
    if(!employee) return {ok:false,message:'Employee record no longer exists.'};
    if(decision==='Approved') {
      const start=clean_(r[4]), end=clean_(r[5]);
      if(!validIsoDate_(start)||!validIsoDate_(end)||end<start) return {ok:false,message:'The saved date range is invalid.'};
      const schedule=getRequiredSheet_(SHEET_WORK_SCHEDULES), existing=schedule.getDataRange().getDisplayValues();
      const existingByDate={};
      for(let i=1;i<existing.length;i++) if(clean_(existing[i][1]).toLowerCase()===employee.id.toLowerCase()) existingByDate[clean_(existing[i][0])]=i+1;
      for(let ms=Date.parse(start+'T00:00:00Z');ms<=Date.parse(end+'T00:00:00Z');ms+=86400000) {
        const date=new Date(ms).toISOString().slice(0,10);
        const targetRow=existingByDate[date];
        const row=[date,employee.id,'WFH','Admin',timestamp_(),'Approved request '+clean_(r[0])];
        if(targetRow) schedule.getRange(targetRow,1,1,row.length).setValues([row]); else schedule.appendRow(row);
      }
    } else {
      // A rejection never changes previously assigned schedules.
    }
    sheet.getRange(rowNum,8).setValue(decision);
    sheet.getRange(rowNum,9).setValue(timestamp_());
    sheet.getRange(rowNum,10).setValue('Admin');
    sheet.getRange(rowNum,11).setValue(remarks);
    createNotification_('Employee',employee.id,'WFH request '+decision.toLowerCase(),'Your WFH request ('+clean_(r[4])+' to '+clean_(r[5])+') was '+decision.toLowerCase()+'.'+(remarks?' Admin note: '+remarks:''),clean_(requestId));
    writeAuditLog_('Admin','WFH request '+decision.toLowerCase(),clean_(requestId)+' for '+employee.id+'; no secrets logged');
    return {ok:true,message:'WFH request '+decision.toLowerCase()+'.'+(decision==='Approved'?' Approved dates are now available for WFH attendance.':'')};
  } catch(err) { return {ok:false,message:err.message || 'Could not review WFH request.'}; }
  finally { try {lock.releaseLock();} catch(e){} }
}

/** Admin can directly assign or revoke one employee's WFH day. */
function setWorkSchedule(token, employeeId, date, mode, notes) {
  if (!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  const lock=LockService.getScriptLock();
  try {
    employeeId=clean_(employeeId); date=clean_(date); mode=clean_(mode); notes=clean_(notes);
    const employee=findEmployee_(employeeId);
    if(!employee || String(employee.status).toLowerCase()==='inactive') return {ok:false,message:'Active employee not found.'};
    if(!validIsoDate_(date)) return {ok:false,message:'Choose a valid schedule date.'};
    if(['Office','WFH','Off'].indexOf(mode)<0) return {ok:false,message:'Mode must be Office, WFH or Off.'};
    lock.waitLock(10000);
    const sheet=getRequiredSheet_(SHEET_WORK_SCHEDULES), values=sheet.getDataRange().getDisplayValues();
    let rowNum=0;
    for(let i=1;i<values.length;i++) if(clean_(values[i][0])===date && clean_(values[i][1]).toLowerCase()===employee.id.toLowerCase()){rowNum=i+1;break;}
    if(mode==='Off') { if(rowNum) sheet.deleteRow(rowNum); return {ok:true,message:'Custom work schedule removed for '+date+'.'}; }
    const row=[date,employee.id,mode,'Admin',timestamp_(),notes];
    if(rowNum) sheet.getRange(rowNum,1,1,row.length).setValues([row]); else sheet.appendRow(row);
    writeAuditLog_('Admin','Work schedule updated',employee.id+' '+date+' '+mode);
    return {ok:true,message:mode+' schedule saved for '+employee.name+' on '+date+'.'};
  } catch(err) { return {ok:false,message:err.message || 'Could not save work schedule.'}; }
  finally { try {lock.releaseLock();} catch(e){} }
}

/** Admin-only schedule list for a date range. */
function getWorkSchedules(token, startDate, endDate) {
  if(!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    const start=validIsoDate_(startDate)?clean_(startDate):today_();
    const end=validIsoDate_(endDate)?clean_(endDate):start;
    if(end<start) return {ok:false,message:'End date must be on or after start date.'};
    const values=getRequiredSheet_(SHEET_WORK_SCHEDULES).getDataRange().getDisplayValues();
    const schedules=[];
    for(let i=1;i<values.length;i++){const r=values[i];if(clean_(r[0])>=start&&clean_(r[0])<=end)schedules.push({date:r[0],employeeId:r[1],workMode:r[2],approvedBy:r[3],approvedAt:r[4],notes:r[5]});}
    return {ok:true,schedules:schedules};
  } catch(err) { return {ok:false,message:err.message || 'Could not load work schedules.'}; }
}


/* ========================= LEAVE, REPORTS & IN-APP NOTIFICATIONS ========================= */

function createNotification_(audience, recipientId, title, message, relatedId) {
  try {
    const sheet = getRequiredSheet_(SHEET_NOTIFICATIONS);
    const id = 'NTF-' + Utilities.getUuid().slice(0, 10).toUpperCase();
    sheet.appendRow([id, timestamp_(), clean_(audience), clean_(recipientId), clean_(title), clean_(message), 'No', clean_(relatedId)]);
    return id;
  } catch (e) { return ''; }
}

/** Submit a leave request. A request never marks attendance as Leave until approved. */
function requestLeave(employeeId, pin, payload) {
  const lock = LockService.getScriptLock();
  try {
    payload = payload || {};
    const employee = findEmployee_(employeeId);
    if (!employee || String(employee.status).toLowerCase() === 'inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings = getSettingsSafe_();
    if (settings.requirePin && clean_(pin) !== employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const type = clean_(payload.leaveType || 'Casual Leave');
    const allowed = ['Casual Leave','Sick Leave','Earned Leave','Unpaid Leave','Other'];
    const start = clean_(payload.startDate), end = clean_(payload.endDate), reason = clean_(payload.reason);
    if (allowed.indexOf(type) < 0) return {ok:false,message:'Choose a valid leave type.'};
    if (!validIsoDate_(start) || !validIsoDate_(end) || end < start) return {ok:false,message:'Choose a valid leave date range.'};
    if (start < today_()) return {ok:false,message:'Leave requests must start today or in the future.'};
    if (!reason || reason.length < 5) return {ok:false,message:'Please provide a reason (at least 5 characters).'};
    const days = Math.round((Date.parse(end+'T00:00:00Z') - Date.parse(start+'T00:00:00Z')) / 86400000) + 1;
    if (days > 31) return {ok:false,message:'A single leave request cannot exceed 31 calendar days.'};
    lock.waitLock(10000);
    const sheet = getRequiredSheet_(SHEET_LEAVE_REQUESTS), values = sheet.getDataRange().getDisplayValues();
    for (let i=1;i<values.length;i++) {
      const r=values[i];
      if (clean_(r[2]).toLowerCase()===employee.id.toLowerCase() && clean_(r[9])==='Pending' && clean_(r[5])<=end && clean_(r[6])>=start) return {ok:false,message:'You already have a pending leave request overlapping these dates.'};
    }
    // Do not allow leave to overlap days already attended.
    const logs = readLogs_();
    const attended = logs.some(function(r){return r.employeeId.toLowerCase()===employee.id.toLowerCase() && r.date>=start && r.date<=end && !!r.loginTime;});
    if (attended) return {ok:false,message:'This date range overlaps an attendance check-in. Submit a separate request for dates without attendance.'};
    const id='LEAVE-' + Utilities.getUuid().slice(0,8).toUpperCase();
    sheet.appendRow([id,timestamp_(),employee.id,employee.name,type,start,end,days,reason,'Pending','','','']);
    createNotification_('Admin','', 'New leave request', employee.name+' submitted '+type+' for '+start+' to '+end+'.', id);
    createNotification_('Employee',employee.id,'Leave request submitted','Your '+type+' request ('+start+' to '+end+') is pending Admin review.',id);
    writeAuditLog_(employee.id,'Leave request submitted',id+' '+start+' to '+end);
    return {ok:true,message:'Leave request submitted. It will affect attendance only after Admin approval.',requestId:id};
  } catch(err) { return {ok:false,message:err.message || 'Could not submit leave request.'}; }
  finally { try {lock.releaseLock();} catch(e){} }
}

function getEmployeeLeaveStatus(employeeId, pin) {
  try {
    const employee=findEmployee_(employeeId);
    if(!employee || String(employee.status).toLowerCase()==='inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings=getSettingsSafe_();
    if(settings.requirePin && clean_(pin)!==employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const values=getRequiredSheet_(SHEET_LEAVE_REQUESTS).getDataRange().getDisplayValues(), requests=[];
    for(let i=1;i<values.length;i++) { const r=values[i]; if(clean_(r[2]).toLowerCase()!==employee.id.toLowerCase()) continue; requests.push({requestId:r[0],submittedAt:r[1],leaveType:r[4],startDate:r[5],endDate:r[6],days:Number(r[7])||0,reason:r[8],status:r[9],adminRemarks:r[12]}); }
    requests.reverse();
    return {ok:true,requests:requests};
  } catch(err) { return {ok:false,message:err.message || 'Could not load leave requests.'}; }
}

function getLeaveRequests(token) {
  if(!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    const values=getRequiredSheet_(SHEET_LEAVE_REQUESTS).getDataRange().getDisplayValues(), requests=[];
    for(let i=1;i<values.length;i++) { const r=values[i]; if(!clean_(r[0])) continue; requests.push({requestId:r[0],submittedAt:r[1],employeeId:r[2],employeeName:r[3],leaveType:r[4],startDate:r[5],endDate:r[6],days:Number(r[7])||0,reason:r[8],status:r[9],reviewedAt:r[10],reviewedBy:r[11],adminRemarks:r[12]}); }
    requests.reverse(); return {ok:true,requests:requests};
  } catch(err) { return {ok:false,message:err.message || 'Could not load leave requests.'}; }
}

function reviewLeaveRequest(token, requestId, decision, remarks) {
  if(!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  const lock=LockService.getScriptLock();
  try {
    decision=clean_(decision); remarks=clean_(remarks);
    if(decision!=='Approved' && decision!=='Rejected') return {ok:false,message:'Invalid review decision.'};
    lock.waitLock(10000);
    const sheet=getRequiredSheet_(SHEET_LEAVE_REQUESTS), values=sheet.getDataRange().getDisplayValues();
    let rowNum=0,r=null;
    for(let i=1;i<values.length;i++) if(clean_(values[i][0])===clean_(requestId)){rowNum=i+1;r=values[i];break;}
    if(!r) return {ok:false,message:'Leave request not found.'};
    if(clean_(r[9])!=='Pending') return {ok:false,message:'This leave request has already been reviewed.'};
    const employee=findEmployee_(r[2]); if(!employee) return {ok:false,message:'Employee record no longer exists.'};
    if(decision==='Approved') {
      const start=clean_(r[5]), end=clean_(r[6]);
      if(!validIsoDate_(start)||!validIsoDate_(end)||end<start) return {ok:false,message:'Saved leave date range is invalid.'};
      const logSheet=getLogSheet_(), existing=logSheet.getDataRange().getDisplayValues();
      const existingRows={};
      for(let i=1;i<existing.length;i++) if(clean_(existing[i][1]).toLowerCase()===employee.id.toLowerCase()) existingRows[clean_(existing[i][0])]=i+1;
      const holidays={};
      try { const h=getRequiredSheet_(SHEET_HOLIDAYS).getDataRange().getDisplayValues(); for(let i=1;i<h.length;i++) holidays[clean_(h[i][0])]=true; } catch(e){}
      const settings=getSettingsSafe_();
      for(let ms=Date.parse(start+'T00:00:00Z');ms<=Date.parse(end+'T00:00:00Z');ms+=86400000) {
        const date=new Date(ms).toISOString().slice(0,10), dayName=new Date(ms).toLocaleDateString('en-US',{weekday:'long',timeZone:'UTC'}).toLowerCase();
        if(settings.weeklyOffDays.indexOf(dayName)>=0 || holidays[date]) continue;
        const targetRow=existingRows[date];
        if(targetRow) {
          const row=existing[targetRow-1];
          if(clean_(row[5])) continue; // Never overwrite a recorded check-in.
          logSheet.getRange(targetRow,9).setValue('Leave');
          logSheet.getRange(targetRow,12).setValue((clean_(row[11])?clean_(row[11])+' | ':'')+'Approved leave '+clean_(requestId));
          logSheet.getRange(targetRow,13).setValue(timestamp_());
        } else {
          logSheet.appendRow([date,employee.id,employee.name,employee.designation,employee.department,'','','','Leave','','','Approved leave '+clean_(requestId),timestamp_()]);
        }
      }
    }
    sheet.getRange(rowNum,10).setValue(decision); sheet.getRange(rowNum,11).setValue(timestamp_()); sheet.getRange(rowNum,12).setValue('Admin'); sheet.getRange(rowNum,13).setValue(remarks);
    createNotification_('Employee',employee.id,'Leave request '+decision.toLowerCase(),'Your '+clean_(r[4])+' request ('+clean_(r[5])+' to '+clean_(r[6])+') was '+decision.toLowerCase()+'.'+(remarks?' Admin note: '+remarks:''),clean_(requestId));
    writeAuditLog_('Admin','Leave request '+decision.toLowerCase(),clean_(requestId)+' for '+employee.id);
    return {ok:true,message:'Leave request '+decision.toLowerCase()+'.'+(decision==='Approved'?' Eligible weekdays were marked Leave where no check-in existed.':'')};
  } catch(err) { return {ok:false,message:err.message || 'Could not review leave request.'}; }
  finally { try {lock.releaseLock();} catch(e){} }
}

/** Month-level attendance export source. Returns only records in the requested month. */
function getMonthlyAttendanceReport(token, month) {
  if(!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    month=clean_(month);
    if(!/^\d{4}-\d{2}$/.test(month)) month=today_().slice(0,7);
    const start=month+'-01', nextDate=new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),1));
    const end=new Date(nextDate.getTime()-86400000).toISOString().slice(0,10);
    const employees=readEmployees_().filter(function(e){return String(e.status).toLowerCase()!=='inactive';});
    const logs=readLogs_().filter(function(r){return r.date>=start && r.date<=end;});
    const records=logs.map(function(r){
      let workMode='Office';
      try { const vals=getLogSheet_().getRange(r.rowNumber,1,1,getLogSheet_().getLastColumn()).getDisplayValues()[0]; const headers=getLogSheet_().getRange(1,1,1,getLogSheet_().getLastColumn()).getDisplayValues()[0]; const idx=headers.indexOf('Work Mode'); if(idx>=0 && clean_(vals[idx])) workMode=clean_(vals[idx]); } catch(e){}
      return {date:r.date,employeeId:r.employeeId,name:r.name,department:r.department,loginTime:r.loginTime,logoutTime:r.logoutTime,hours:r.hours,status:r.status || (r.loginTime?'Present':'Absent'),workMode:workMode};
    }).sort(function(a,b){return a.date.localeCompare(b.date)||a.name.localeCompare(b.name);});
    const summary={month:month,employeeCount:employees.length,recordCount:records.length,totalHours:0,present:0,late:0,halfDay:0,leave:0,absent:0,wfh:0};
    records.forEach(function(r){summary.totalHours+=Number(r.hours)||0;if(r.status==='Present')summary.present++;if(r.status==='Late')summary.late++;if(r.status==='Half Day')summary.halfDay++;if(r.status==='Leave')summary.leave++;if(r.status==='Absent')summary.absent++;if(r.workMode==='WFH')summary.wfh++;});
    summary.totalHours=Math.round(summary.totalHours*100)/100;
    return {ok:true,month:month,startDate:start,endDate:end,summary:summary,records:records};
  } catch(err) { return {ok:false,message:err.message || 'Could not build monthly report.'}; }
}

function getEmployeeNotifications(employeeId, pin) {
  try {
    const employee=findEmployee_(employeeId); if(!employee || String(employee.status).toLowerCase()==='inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings=getSettingsSafe_(); if(settings.requirePin && clean_(pin)!==employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const values=getRequiredSheet_(SHEET_NOTIFICATIONS).getDataRange().getDisplayValues(), items=[];
    for(let i=1;i<values.length;i++){const r=values[i];if(clean_(r[2])!=='Employee'||clean_(r[3]).toLowerCase()!==employee.id.toLowerCase())continue;items.push({id:r[0],createdAt:r[1],title:r[4],message:r[5],read:clean_(r[6])==='Yes'});}
    items.reverse(); return {ok:true,notifications:items.slice(0,50),unread:items.filter(function(x){return !x.read;}).length};
  } catch(err) { return {ok:false,message:err.message || 'Could not load notifications.'}; }
}

function getAdminNotifications(token) {
  if(!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    const values=getRequiredSheet_(SHEET_NOTIFICATIONS).getDataRange().getDisplayValues(), items=[];
    for(let i=1;i<values.length;i++){const r=values[i];if(clean_(r[2])!=='Admin')continue;items.push({id:r[0],createdAt:r[1],title:r[4],message:r[5],relatedId:r[7],read:clean_(r[6])==='Yes'});}
    items.reverse(); return {ok:true,notifications:items.slice(0,100),unread:items.filter(function(x){return !x.read;}).length};
  } catch(err) { return {ok:false,message:err.message || 'Could not load notifications.'}; }
}

function markNotificationRead(employeeId, pin, notificationId) {
  try {
    const employee=findEmployee_(employeeId); if(!employee || String(employee.status).toLowerCase()==='inactive') return {ok:false,message:'Employee not found or inactive.'};
    const settings=getSettingsSafe_(); if(settings.requirePin && clean_(pin)!==employee.pin) return {ok:false,message:'Incorrect employee PIN.'};
    const sheet=getRequiredSheet_(SHEET_NOTIFICATIONS), values=sheet.getDataRange().getDisplayValues();
    for(let i=1;i<values.length;i++){const r=values[i];if(clean_(r[0])===clean_(notificationId)&&clean_(r[2])==='Employee'&&clean_(r[3]).toLowerCase()===employee.id.toLowerCase()){sheet.getRange(i+1,7).setValue('Yes');return {ok:true};}}
    return {ok:false,message:'Notification not found.'};
  } catch(err) { return {ok:false,message:err.message || 'Could not update notification.'}; }
}


function markAdminNotificationRead(token, notificationId) {
  if(!isAdmin_(token)) return {ok:false,authFailed:true,message:'Admin session expired. Sign in again.'};
  try {
    const sheet=getRequiredSheet_(SHEET_NOTIFICATIONS), values=sheet.getDataRange().getDisplayValues();
    for(let i=1;i<values.length;i++) if(clean_(values[i][0])===clean_(notificationId) && clean_(values[i][2])==='Admin') { sheet.getRange(i+1,7).setValue('Yes'); return {ok:true}; }
    return {ok:false,message:'Notification not found.'};
  } catch(err) { return {ok:false,message:err.message || 'Could not update notification.'}; }
}