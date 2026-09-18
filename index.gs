const SPREADSHEET_ID =
  '1MNgTyd1sh6UPysVXb8jATcRH2ec0P8arPasi1OYbxS0';

const EMPLOYEE_SHEET = 'Data Employee';
const DOCTOR_SHEET = 'Doctor';
const CONTACT_SHEET = 'Contact';
const BRANCH_SHEET = 'Branch';
const PROMOTION_SHEET = 'Promotion';
const SERVICE_SHEET = 'Service';
const CALENDAR_SHEET = 'Calendar';
const GOOGLE_CALENDAR_ID =
  'ffee9a7d9cac44c0736cb1dfeb973727616a08c82370d7fbefd7a6f12512b89b@group.calendar.google.com';
const GOOGLE_CALENDAR_TIMEZONE = 'Asia/Bangkok';
const GOOGLE_CALENDAR_DEFAULT_DURATION_MINUTES = 60;
const GOOGLE_CALENDAR_IMPORT_PAST_DAYS = 365;
const GOOGLE_CALENDAR_IMPORT_FUTURE_DAYS = 730;
const GOOGLE_CALENDAR_TRIGGER_HANDLER = 'calendarTwoWaySyncTrigger';

const MASTER_BOOTSTRAP_ID = '202600041';
const SESSION_HOURS = 8;
const PASSWORD_MIN_LENGTH = 8;
const DEFAULT_PASSWORD = '1234';


const ALLOWED_LEVELS = [
  'Master',
  'Executive',
  'Manager',
  'Employee',
  'Outsource',
  'Outsource-Content',
  'Outsource-Graphic',
  'Outsource-Ad Optimize',
  'Outsource-All'
];


function hasWebsiteAccessLevel_(
  level
) {
  const target =
    String(
      level || ''
    )
      .trim()
      .toLowerCase();

  if (!target) {
    return false;
  }

  return ALLOWED_LEVELS
    .some(value =>
      String(value)
        .trim()
        .toLowerCase()
      === target
    );
}



/* =========================
   TEST CONNECTION
========================= */

function testConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Connected to: ' + ss.getName());
}


/* =========================
   GET API
========================= */

function doGet(e) {
  try {
    const action =
      String(
        (e && e.parameter && e.parameter.action) || 'health'
      ).trim();

    if (action === 'health') {
      return jsonResponse({
        success: true,
        message: 'Lumie API Online',
        version: 'V31-GitHub-Live-Google-Calendar',
        timestamp: new Date().toISOString()
      });
    }

    const token =
      String(
        (e && e.parameter && e.parameter.authToken) || ''
      ).trim();

    if (action === 'session') {
      const session = requireSessionToken_(token);

      return jsonResponse({
        success: true,
        employee: session
      });
    }

    if (action === 'employees') {
      requireSessionToken_(token);
      return getEmployees();
    }

    if (action === 'doctors') {
      requireSessionToken_(token);
      return getDoctors();
    }

    if (action === 'branches') {
      requireSessionToken_(token);
      return getBranches();
    }

    if (action === 'services') {
      requireSessionToken_(token);
      return getServices();
    }

    if (action === 'promotions') {
      requireSessionToken_(token);
      return getPromotions();
    }

    if (action === 'calendar') {
      const session =
        requireSessionToken_(token);

      requireCalendarViewer_(
        session
      );

      return getCalendar(e);
    }

    if (action === 'doctorCouncil') {
      requireSessionToken_(token);
      return getDoctorCouncil(e);
    }

    if (action === 'doctorImage') {
      requireSessionToken_(token);
      return getDoctorImage(e);
    }

    if (action === 'contacts') {
      requireMasterSession_(token);
      return getContacts();
    }

    if (action === 'all') {
      requireMasterSession_(token);
      return getAllData();
    }

    return jsonResponse({
      success: false,
      message: 'Unknown action'
    });

  } catch (err) {
    return errorResponse(err);
  }
}


/* =========================
   POST API
========================= */

function doPost(e) {
  try {
    const data =
      JSON.parse(
        (e && e.postData && e.postData.contents) || '{}'
      );

    const action =
      String(data.action || '').trim();

    if (action === 'login') {
      return loginUser(data);
    }

    const token =
      String(
        data.authToken || ''
      ).trim();

    if (action === 'changePassword') {
      const session =
        requireSessionToken_(token);

      return changePassword(
        data,
        session
      );
    }

    if (action === 'addContact') {
      const session =
        requireSessionToken_(token);

      return addContact(
        data,
        session
      );
    }

    if (action === 'updateLevel') {
      const session =
        requireMasterSession_(token);

      return updateEmployeeLevel(
        data,
        session
      );
    }

    if (action === 'updateEmployee') {
      const session =
        requireMasterSession_(token);

      return updateEmployee(
        data,
        session
      );
    }

    if (action === 'addEmployee') {
      const session =
        requireMasterSession_(token);

      return addEmployee(
        data,
        session
      );
    }

    if (action === 'deleteEmployee') {
      const session =
        requireMasterSession_(token);

      return deleteEmployee(
        data,
        session
      );
    }

    if (action === 'setUserPassword') {
      const session =
        requireMasterSession_(token);

      return setUserPassword(
        data,
        session
      );
    }

    if (action === 'resetUserPassword') {
      const session =
        requireMasterSession_(token);

      return resetUserPassword(
        data,
        session
      );
    }

    if (action === 'addDoctor') {
      requireMasterSession_(token);
      return addDoctor(data);
    }

    if (action === 'updateDoctor') {
      requireMasterSession_(token);
      return updateDoctor(data);
    }

    if (action === 'uploadDoctorImage') {
      requireMasterSession_(token);
      return uploadDoctorImage(data);
    }

    if (action === 'addCalendar') {
      const session =
        requireSessionToken_(token);

      return addCalendar(
        data,
        session
      );
    }

    if (action === 'updateCalendar') {
      const session =
        requireSessionToken_(token);

      return updateCalendar(
        data,
        session
      );
    }

    if (action === 'deleteCalendar') {
      const session =
        requireSessionToken_(token);

      return deleteCalendar(
        data,
        session
      );
    }

    if (action === 'addBranch') {
      requireMasterSession_(token);
      return addBranch(data);
    }

    if (action === 'updateBranch') {
      requireMasterSession_(token);
      return updateBranch(data);
    }

    return jsonResponse({
      success: false,
      message: 'Unknown action'
    });

  } catch (err) {
    return errorResponse(err);
  }
}


/* =========================
   SECURITY INITIALIZATION
========================= */

/*
  Run initializeSecurity() after deploying this version.
  It creates Password Hash / Password Salt columns, creates
  server-side secrets, and gives DEFAULT_PASSWORD only to
  accounts that do not have a password yet.

  It DOES NOT overwrite users who already changed their password.
*/
function initializeSecurity() {
  ensureEmployeeSecurityColumns_();

  getSecuritySecret_('PASSWORD_PEPPER');
  getSecuritySecret_('SESSION_SIGNING_SECRET');

  const initialized =
    initializeMissingPasswords_();

  Logger.log(
    'Security ready. Initialized default password for ' +
    initialized +
    ' account(s).'
  );

  Logger.log(
    'Default password for new/uninitialized accounts: ' +
    DEFAULT_PASSWORD
  );
}


/*
  One-time migration helper.
  Run this ONLY if you intentionally want every current account
  to be reset to the default password 1234.
  Do not run this again after users have changed their passwords.
*/
function resetAllPasswordsToDefaultFromEditor() {
  const sheet =
    getSheet(EMPLOYEE_SHEET);

  ensureEmployeeSecurityColumns_();

  const lastRow =
    sheet.getLastRow();

  let changed = 0;

  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {
    const employeeId =
      String(
        sheet
          .getRange(row, 1)
          .getDisplayValue() || ''
      ).trim();

    if (!employeeId) {
      continue;
    }

    const level =
      String(
        sheet
          .getRange(row, 2)
          .getDisplayValue() || ''
      ).trim();

    if (
      !hasWebsiteAccessLevel_(
        level
      )
    ) {
      continue;
    }

    setPasswordForEmployee_(
      employeeId,
      DEFAULT_PASSWORD,
      true
    );

    changed++;
  }

  Logger.log(
    'Reset ' +
    changed +
    ' account(s) to default password: ' +
    DEFAULT_PASSWORD
  );
}


/*
  Emergency Master recovery.
  This resets only the Master account to 1234.
*/
function resetMasterPasswordFromEditor() {
  ensureEmployeeSecurityColumns_();

  setPasswordForEmployee_(
    MASTER_BOOTSTRAP_ID,
    DEFAULT_PASSWORD,
    true
  );

  Logger.log(
    'Master password reset to default: ' +
    DEFAULT_PASSWORD
  );
}


function initializeMissingPasswords_() {
  const sheet =
    getSheet(EMPLOYEE_SHEET);

  const columns =
    ensureEmployeeSecurityColumns_();

  const lastRow =
    sheet.getLastRow();

  let initialized = 0;

  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {
    const employeeId =
      String(
        sheet
          .getRange(row, 1)
          .getDisplayValue() || ''
      ).trim();

    if (!employeeId) {
      continue;
    }

    const level =
      String(
        sheet
          .getRange(row, 2)
          .getDisplayValue() || ''
      ).trim();

    if (
      !hasWebsiteAccessLevel_(
        level
      )
    ) {
      continue;
    }

    const hash =
      String(
        sheet
          .getRange(
            row,
            columns.passwordHash
          )
          .getDisplayValue() || ''
      ).trim();

    const salt =
      String(
        sheet
          .getRange(
            row,
            columns.passwordSalt
          )
          .getDisplayValue() || ''
      ).trim();

    if (
      hash &&
      salt
    ) {
      continue;
    }

    setPasswordForEmployee_(
      employeeId,
      DEFAULT_PASSWORD,
      true
    );

    initialized++;
  }

  return initialized;
}


/* =========================
   LOGIN
========================= */

function loginUser(data) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  const password =
    String(
      data.password || ''
    );

  if (!employeeId) {
    throw new Error(
      'กรุณากรอกรหัสพนักงาน'
    );
  }

  if (!password) {
    throw new Error(
      'กรุณากรอกรหัสผ่าน'
    );
  }

  let employee =
    getEmployeeRecordById_(
      employeeId
    );

  if (!employee) {
    throw new Error(
      'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง'
    );
  }

  if (
    !hasWebsiteAccessLevel_(
      employee.level
    )
  ) {
    throw new Error(
      'บัญชีนี้ไม่มีสิทธิ์เข้าใช้งานเว็บไซต์ กรุณาติดต่อผู้ดูแลระบบ'
    );
  }

  /*
    Backward compatibility for existing employees:
    if the account has never had a password, its first password
    is automatically 1234.
  */
  if (
    !employee.passwordHash ||
    !employee.passwordSalt
  ) {
    if (
      password !==
      DEFAULT_PASSWORD
    ) {
      throw new Error(
        'บัญชีนี้ยังใช้รหัสเริ่มต้น กรุณาเข้าสู่ระบบด้วย 1234'
      );
    }

    setPasswordForEmployee_(
      employeeId,
      DEFAULT_PASSWORD,
      true
    );

    employee =
      getEmployeeRecordById_(
        employeeId
      );
  }

  const valid =
    verifyPassword_(
      password,
      employee.passwordSalt,
      employee.passwordHash
    );

  if (!valid) {
    throw new Error(
      'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง'
    );
  }

  const publicEmployee =
    publicEmployee_(
      employee
    );

  return jsonResponse({
    success: true,
    message: 'Login success',
    employee: publicEmployee,
    authToken:
      createSessionToken_(
        employee.id
      ),
    expiresHours:
      SESSION_HOURS
  });
}


/* =========================
   PASSWORD MANAGEMENT
========================= */

function setUserPassword(
  data,
  session
) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  const newPassword =
    String(
      data.newPassword || ''
    );

  validatePassword_(
    newPassword
  );

  const targetEmployee =
    getEmployeeRecordById_(
      employeeId
    );

  if (!targetEmployee) {
    throw new Error(
      'Employee not found'
    );
  }

  assertManagerCanManageEmployee_(
    session,
    targetEmployee
  );

  if (
    !hasWebsiteAccessLevel_(
      targetEmployee.level
    )
  ) {
    throw new Error(
      'This employee has no website access'
    );
  }

  setPasswordForEmployee_(
    employeeId,
    newPassword
  );

  return jsonResponse({
    success: true,
    message:
      'Password updated',
    employeeId:
      employeeId
  });
}


function resetUserPassword(
  data,
  session
) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  if (!employeeId) {
    throw new Error(
      'Employee ID is required'
    );
  }

  const targetEmployee =
    getEmployeeRecordById_(
      employeeId
    );

  if (!targetEmployee) {
    throw new Error(
      'Employee not found'
    );
  }

  assertManagerCanManageEmployee_(
    session,
    targetEmployee
  );

  if (
    !hasWebsiteAccessLevel_(
      targetEmployee.level
    )
  ) {
    throw new Error(
      'This employee has no website access'
    );
  }

  setPasswordForEmployee_(
    employeeId,
    DEFAULT_PASSWORD,
    true
  );

  return jsonResponse({
    success: true,
    message:
      'Password reset to default',
    employeeId:
      employeeId
  });
}


function changePassword(
  data,
  session
) {
  const currentPassword =
    String(
      data.currentPassword || ''
    );

  const newPassword =
    String(
      data.newPassword || ''
    );

  validatePassword_(
    newPassword
  );

  const employee =
    getEmployeeRecordById_(
      session.id
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  if (
    !employee.passwordHash ||
    !employee.passwordSalt
  ) {
    throw new Error(
      'Password is not configured'
    );
  }

  if (
    !verifyPassword_(
      currentPassword,
      employee.passwordSalt,
      employee.passwordHash
    )
  ) {
    throw new Error(
      'รหัสผ่านปัจจุบันไม่ถูกต้อง'
    );
  }

  setPasswordForEmployee_(
    employee.id,
    newPassword
  );

  return jsonResponse({
    success: true,
    message:
      'Password changed'
  });
}


/* =========================
   SESSION TOKEN
========================= */

function createSessionToken_(
  employeeId
) {
  const payload = {
    employeeId:
      String(employeeId),

    exp:
      Date.now() +
      SESSION_HOURS *
      60 * 60 * 1000,

    nonce:
      Utilities
        .getUuid()
  };

  const payloadText =
    JSON.stringify(
      payload
    );

  const payloadEncoded =
    Utilities
      .base64EncodeWebSafe(
        payloadText,
        Utilities.Charset.UTF_8
      )
      .replace(/=+$/g, '');

  const secret =
    getSecuritySecret_(
      'SESSION_SIGNING_SECRET'
    );

  const signatureBytes =
    Utilities
      .computeHmacSha256Signature(
        payloadEncoded,
        secret,
        Utilities.Charset.UTF_8
      );

  const signature =
    Utilities
      .base64EncodeWebSafe(
        signatureBytes
      )
      .replace(/=+$/g, '');

  return (
    payloadEncoded +
    '.' +
    signature
  );
}


function verifySessionToken_(
  token
) {
  const parts =
    String(token || '')
      .split('.');

  if (parts.length !== 2) {
    throw new Error(
      'Session expired. Please login again.'
    );
  }

  const payloadEncoded =
    parts[0];

  const providedSignature =
    parts[1];

  const secret =
    getSecuritySecret_(
      'SESSION_SIGNING_SECRET'
    );

  const expectedBytes =
    Utilities
      .computeHmacSha256Signature(
        payloadEncoded,
        secret,
        Utilities.Charset.UTF_8
      );

  const expectedSignature =
    Utilities
      .base64EncodeWebSafe(
        expectedBytes
      )
      .replace(/=+$/g, '');

  if (
    !constantTimeEqual_(
      providedSignature,
      expectedSignature
    )
  ) {
    throw new Error(
      'Invalid session'
    );
  }

  let payload;

  try {
    const json =
      Utilities
        .newBlob(
          Utilities
            .base64DecodeWebSafe(
              padBase64_(
                payloadEncoded
              )
            )
        )
        .getDataAsString();

    payload =
      JSON.parse(
        json
      );
  }
  catch (err) {
    throw new Error(
      'Invalid session'
    );
  }

  if (
    !payload.exp ||
    Date.now() >
      Number(payload.exp)
  ) {
    throw new Error(
      'Session expired. Please login again.'
    );
  }

  const employee =
    getEmployeeRecordById_(
      String(
        payload.employeeId || ''
      )
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  if (
    !hasWebsiteAccessLevel_(
      employee.level
    )
  ) {
    throw new Error(
      'Website access has been disabled for this account'
    );
  }

  return publicEmployee_(
    employee
  );
}


function requireSessionToken_(
  token
) {
  return verifySessionToken_(
    token
  );
}


function requireMasterSession_(
  token
) {
  const employee =
    verifySessionToken_(
      token
    );

  if (
    String(
      employee.level || ''
    ).trim()
    !== 'Master'
  ) {
    throw new Error(
      'Master permission required'
    );
  }

  return employee;
}


function requireManagementSession_(
  token
) {
  const employee =
    verifySessionToken_(
      token
    );

  const level =
    String(
      employee.level || ''
    )
      .trim()
      .toLowerCase();

  if (
    ![
      'master',
      'manager'
    ].includes(
      level
    )
  ) {
    throw new Error(
      'Manager or Master permission required'
    );
  }

  return employee;
}


function isManagerSession_(
  session
) {
  return String(
    (session && session.level) || ''
  )
    .trim()
    .toLowerCase()
    === 'manager';
}


function assertManagerCanManageEmployee_(
  session,
  targetEmployee,
  requestedLevel
) {
  if (
    !isManagerSession_(
      session
    )
  ) {
    return;
  }

  if (
    targetEmployee &&
    String(
      targetEmployee.level || ''
    )
      .trim()
      .toLowerCase()
      === 'master'
  ) {
    throw new Error(
      'Manager cannot modify a Master account'
    );
  }

  if (
    requestedLevel !== undefined &&
    String(
      requestedLevel || ''
    )
      .trim()
      .toLowerCase()
      === 'master'
  ) {
    throw new Error(
      'Manager cannot grant Master permission'
    );
  }
}


/* =========================
   GET EMPLOYEES
========================= */

function getEmployees() {
  const employees =
    getEmployeesData_();

  return jsonResponse({
    success: true,
    count:
      employees.length,
    data:
      employees
  });
}


function getEmployeesData_() {
  const sheet =
    getSheet(
      EMPLOYEE_SHEET
    );

  const columns =
    ensureEmployeeSecurityColumns_();

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return [];
  }

  const data = [];

  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {
    const record =
      getEmployeeRecordFromRow_(
        sheet,
        row,
        columns
      );

    if (
      !record.id &&
      !record.name &&
      !record.position &&
      !record.tel &&
      !record.email
    ) {
      continue;
    }

    data.push(
      publicEmployee_(
        record
      )
    );
  }

  return data;
}


function getEmployeeRecordById_(
  employeeId
) {
  const target =
    String(
      employeeId || ''
    ).trim();

  if (!target) {
    return null;
  }

  const sheet =
    getSheet(
      EMPLOYEE_SHEET
    );

  const columns =
    ensureEmployeeSecurityColumns_();

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return null;
  }

  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {
    if (
      String(
        ids[i][0] || ''
      ).trim()
      === target
    ) {
      return getEmployeeRecordFromRow_(
        sheet,
        i + 2,
        columns
      );
    }
  }

  return null;
}


function getEmployeeRecordFromRow_(
  sheet,
  row,
  columns
) {
  return {
    row: row,

    id:
      String(
        sheet
          .getRange(row, 1)
          .getDisplayValue() || ''
      ).trim(),

    level:
      String(
        sheet
          .getRange(row, 2)
          .getDisplayValue() || ''
      ).trim(),

    position:
      String(
        sheet
          .getRange(row, 3)
          .getDisplayValue() || ''
      ).trim(),

    name:
      String(
        sheet
          .getRange(row, 4)
          .getDisplayValue() || ''
      ).trim(),

    tel:
      String(
        sheet
          .getRange(row, 5)
          .getDisplayValue() || ''
      ).trim(),

    email:
      String(
        sheet
          .getRange(row, 6)
          .getDisplayValue() || ''
      ).trim(),

    passwordHash:
      String(
        sheet
          .getRange(
            row,
            columns.passwordHash
          )
          .getDisplayValue() || ''
      ).trim(),

    passwordSalt:
      String(
        sheet
          .getRange(
            row,
            columns.passwordSalt
          )
          .getDisplayValue() || ''
      ).trim()
  };
}


function publicEmployee_(
  employee
) {
  return {
    employeeId:
      employee.id,

    id:
      employee.id,

    level:
      employee.level,

    position:
      employee.position,

    name:
      employee.name,

    tel:
      employee.tel,

    email:
      employee.email,

    passwordSet:
      !!(
        employee.passwordHash &&
        employee.passwordSalt
      )
  };
}


function ensureEmployeeSecurityColumns_() {
  const sheet =
    getSheet(
      EMPLOYEE_SHEET
    );

  const width =
    Math.max(
      sheet.getLastColumn(),
      8
    );

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        width
      )
      .getDisplayValues()[0]
      .map(v =>
        String(
          v || ''
        ).trim()
      );

  const findHeader =
    target => {
      const wanted =
        String(target)
          .toLowerCase();

      for (
        let i = 0;
        i < headers.length;
        i++
      ) {
        if (
          headers[i]
            .toLowerCase()
          === wanted
        ) {
          return i + 1;
        }
      }

      return 0;
    };

  let passwordHash =
    findHeader(
      'Password Hash'
    );

  let passwordSalt =
    findHeader(
      'Password Salt'
    );

  if (!passwordHash) {
    passwordHash =
      sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        passwordHash
      )
      .setValue(
        'Password Hash'
      );
  }

  if (!passwordSalt) {
    passwordSalt =
      sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        passwordSalt
      )
      .setValue(
        'Password Salt'
      );
  }

  SpreadsheetApp.flush();

  return {
    passwordHash:
      passwordHash,

    passwordSalt:
      passwordSalt
  };
}


function setPasswordForEmployee_(
  employeeId,
  password,
  allowDefaultPassword
) {
  const value =
    String(
      password || ''
    );

  if (
    !(
      allowDefaultPassword === true &&
      value === DEFAULT_PASSWORD
    )
  ) {
    validatePassword_(
      value
    );
  }

  const employee =
    getEmployeeRecordById_(
      employeeId
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  const sheet =
    getSheet(
      EMPLOYEE_SHEET
    );

  const columns =
    ensureEmployeeSecurityColumns_();

  const salt =
    Utilities
      .getUuid()
      .replace(/-/g, '');

  const hash =
    hashPassword_(
      value,
      salt
    );

  sheet
    .getRange(
      employee.row,
      columns.passwordHash
    )
    .setValue(
      hash
    );

  sheet
    .getRange(
      employee.row,
      columns.passwordSalt
    )
    .setValue(
      salt
    );

  SpreadsheetApp.flush();
}


function validatePassword_(
  password
) {
  const value =
    String(
      password || ''
    );

  if (
    value.length <
    PASSWORD_MIN_LENGTH
  ) {
    throw new Error(
      'รหัสผ่านต้องมีอย่างน้อย ' +
      PASSWORD_MIN_LENGTH +
      ' ตัวอักษร'
    );
  }

  if (
    !/[A-Za-zก-๙]/.test(
      value
    ) ||
    !/[0-9]/.test(
      value
    )
  ) {
    throw new Error(
      'รหัสผ่านต้องมีทั้งตัวอักษรและตัวเลข'
    );
  }
}


function hashPassword_(
  password,
  salt
) {
  const pepper =
    getSecuritySecret_(
      'PASSWORD_PEPPER'
    );

  let value =
    String(salt) +
    '|' +
    String(password) +
    '|' +
    pepper;

  let digest = [];

  /*
    Iterative SHA-256 + server-side pepper.
    Password itself is never stored in Google Sheet.
  */
  for (
    let i = 0;
    i < 1200;
    i++
  ) {
    digest =
      Utilities
        .computeDigest(
          Utilities.DigestAlgorithm.SHA_256,
          value,
          Utilities.Charset.UTF_8
        );

    value =
      Utilities
        .base64Encode(
          digest
        ) +
      '|' +
      salt +
      '|' +
      pepper;
  }

  return Utilities
    .base64Encode(
      digest
    );
}


function verifyPassword_(
  password,
  salt,
  expectedHash
) {
  const actual =
    hashPassword_(
      password,
      salt
    );

  return constantTimeEqual_(
    actual,
    expectedHash
  );
}


function getSecuritySecret_(
  key
) {
  const properties =
    PropertiesService
      .getScriptProperties();

  let secret =
    properties
      .getProperty(
        key
      );

  if (!secret) {
    secret =
      Utilities
        .getUuid()
        .replace(/-/g, '') +
      Utilities
        .getUuid()
        .replace(/-/g, '');

    properties
      .setProperty(
        key,
        secret
      );
  }

  return secret;
}


function constantTimeEqual_(
  a,
  b
) {
  const left =
    String(a || '');

  const right =
    String(b || '');

  let diff =
    left.length ^
    right.length;

  const length =
    Math.max(
      left.length,
      right.length
    );

  for (
    let i = 0;
    i < length;
    i++
  ) {
    diff |=
      (
        left.charCodeAt(i) || 0
      ) ^
      (
        right.charCodeAt(i) || 0
      );
  }

  return diff === 0;
}


function padBase64_(
  value
) {
  let text =
    String(value || '');

  while (
    text.length % 4
  ) {
    text += '=';
  }

  return text;
}


function generateTemporaryPassword_() {
  const raw =
    Utilities
      .getUuid()
      .replace(/-/g, '');

  return (
    'Lm' +
    raw.slice(0, 10) +
    '7!'
  );
}


/* =========================
   GET DOCTORS
========================= */

function getDoctors() {
  const doctors = getDoctorsData_();

  return jsonResponse({
    success: true,
    count: doctors.length,
    data: doctors
  });
}


/*
  Doctor sheet uses headers rather than fixed image columns.

  Required:
  - ชื่อหมอ / ขื่อหมอ
  - เลข ว

  Image fields:
  - Profile
  - แพทยสภา

  If Profile or แพทยสภา does not exist yet, the API creates
  the missing header in the next empty column automatically.
*/

function getDoctorsData_() {
  const sheet = getSheet(DOCTOR_SHEET);
  const columns = ensureDoctorColumns_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const doctors = [];

  for (let row = 2; row <= lastRow; row++) {
    const name =
      String(
        sheet
          .getRange(row, columns.name)
          .getDisplayValue() || ''
      ).trim();

    const license =
      String(
        sheet
          .getRange(row, columns.license)
          .getDisplayValue() || ''
      ).trim();

    if (!name) continue;

    const profileInfo =
      getDoctorImageInfo_(
        sheet.getRange(
          row,
          columns.profile
        )
      );

    const councilInfo =
      getDoctorImageInfo_(
        sheet.getRange(
          row,
          columns.council
        )
      );

    doctors.push({
      name: name,
      license: license,

      hasProfileImage:
        profileInfo.exists,

      profileImageUrl:
        profileInfo.url || '',

      hasCouncilImage:
        councilInfo.exists
    });
  }

  return doctors;
}


/* =========================
   GET DOCTOR IMAGE
========================= */

function getDoctorCouncil(e) {
  const license =
    String(
      (e && e.parameter &&
       e.parameter.license) || ''
    ).trim();

  return getDoctorImageByType_(
    license,
    'council'
  );
}


function getDoctorImage(e) {
  const license =
    String(
      (e && e.parameter &&
       e.parameter.license) || ''
    ).trim();

  const type =
    String(
      (e && e.parameter &&
       e.parameter.type) || 'profile'
    ).trim()
    .toLowerCase();

  return getDoctorImageByType_(
    license,
    type
  );
}


function getDoctorImageByType_(
  license,
  type
) {
  if (!license) {
    throw new Error(
      'Doctor license is required'
    );
  }

  if (
    type !== 'profile' &&
    type !== 'council'
  ) {
    throw new Error(
      'Invalid doctor image type'
    );
  }

  const sheet =
    getSheet(DOCTOR_SHEET);

  const columns =
    ensureDoctorColumns_();

  const row =
    findDoctorRowByLicense_(
      sheet,
      columns,
      license
    );

  if (!row) {
    throw new Error(
      'Doctor not found'
    );
  }

  const name =
    String(
      sheet
        .getRange(
          row,
          columns.name
        )
        .getDisplayValue() || ''
    ).trim();

  const targetColumn =
    type === 'profile'
      ? columns.profile
      : columns.council;

  const info =
    getDoctorImageInfo_(
      sheet.getRange(
        row,
        targetColumn
      )
    );

  if (!info.exists || !info.url) {
    return jsonResponse({
      success: false,
      message:
        type === 'profile'
          ? 'ยังไม่มีรูป Profile ของแพทย์ท่านนี้'
          : 'ยังไม่มีภาพใบแพทยสภาของแพทย์ท่านนี้',
      name: name,
      license: license,
      type: type
    });
  }

  return jsonResponse({
    success: true,
    name: name,
    license: license,
    type: type,
    imageUrl: info.url,
    temporaryUrl: true,
    timestamp:
      new Date().toISOString()
  });
}


/* =========================
   ADD DOCTOR
========================= */

function addDoctor(data) {
  const name =
    String(
      data.name || ''
    ).trim();

  const license =
    String(
      data.license || ''
    ).trim();

  if (!name) {
    throw new Error(
      'Doctor name is required'
    );
  }

  if (!license) {
    throw new Error(
      'Doctor license is required'
    );
  }

  const sheet =
    getSheet(DOCTOR_SHEET);

  const columns =
    ensureDoctorColumns_();

  if (
    findDoctorRowByLicense_(
      sheet,
      columns,
      license
    )
  ) {
    throw new Error(
      'Doctor license already exists'
    );
  }

  const row =
    Math.max(
      2,
      sheet.getLastRow() + 1
    );

  sheet
    .getRange(
      row,
      columns.name
    )
    .setValue(name);

  sheet
    .getRange(
      row,
      columns.license
    )
    .setValue(license);

  SpreadsheetApp.flush();

  return jsonResponse({
    success: true,
    message:
      'Doctor added',
    name: name,
    license: license
  });
}


/* =========================
   UPDATE DOCTOR
========================= */

function updateDoctor(data) {
  const originalLicense =
    String(
      data.originalLicense ||
      data.license ||
      ''
    ).trim();

  const name =
    String(
      data.name || ''
    ).trim();

  const newLicense =
    String(
      data.license || ''
    ).trim();

  if (!originalLicense) {
    throw new Error(
      'Original doctor license is required'
    );
  }

  if (!name) {
    throw new Error(
      'Doctor name is required'
    );
  }

  if (!newLicense) {
    throw new Error(
      'Doctor license is required'
    );
  }

  const sheet =
    getSheet(DOCTOR_SHEET);

  const columns =
    ensureDoctorColumns_();

  const row =
    findDoctorRowByLicense_(
      sheet,
      columns,
      originalLicense
    );

  if (!row) {
    throw new Error(
      'Doctor not found'
    );
  }

  if (
    newLicense !==
    originalLicense
  ) {
    const duplicatedRow =
      findDoctorRowByLicense_(
        sheet,
        columns,
        newLicense
      );

    if (
      duplicatedRow &&
      duplicatedRow !== row
    ) {
      throw new Error(
        'New doctor license already exists'
      );
    }
  }

  sheet
    .getRange(
      row,
      columns.name
    )
    .setValue(name);

  sheet
    .getRange(
      row,
      columns.license
    )
    .setValue(newLicense);

  SpreadsheetApp.flush();

  return jsonResponse({
    success: true,
    message:
      'Doctor updated',
    originalLicense:
      originalLicense,
    name: name,
    license: newLicense
  });
}


/* =========================
   UPLOAD DOCTOR IMAGE
========================= */

function uploadDoctorImage(data) {
  const license =
    String(
      data.license || ''
    ).trim();

  const imageType =
    String(
      data.imageType || ''
    ).trim()
    .toLowerCase();

  const fileName =
    String(
      data.fileName || 'doctor-image'
    ).trim();

  const mimeType =
    String(
      data.mimeType || ''
    ).trim()
    .toLowerCase();

  const base64 =
    String(
      data.base64 || ''
    ).trim();

  if (!license) {
    throw new Error(
      'Doctor license is required'
    );
  }

  if (
    imageType !== 'profile' &&
    imageType !== 'council'
  ) {
    throw new Error(
      'Invalid imageType'
    );
  }

  const allowedMimeTypes = [
    'image/jpeg',
    'image/png',
    'image/webp'
  ];

  if (
    !allowedMimeTypes.includes(
      mimeType
    )
  ) {
    throw new Error(
      'Only JPG, PNG and WEBP are supported'
    );
  }

  if (!base64) {
    throw new Error(
      'Image data is required'
    );
  }

  /*
    Limit base64 text to about 8 MB.
    This keeps upload requests practical for Apps Script.
  */
  if (
    base64.length >
    8 * 1024 * 1024
  ) {
    throw new Error(
      'Image is too large'
    );
  }

  const sheet =
    getSheet(DOCTOR_SHEET);

  const columns =
    ensureDoctorColumns_();

  const row =
    findDoctorRowByLicense_(
      sheet,
      columns,
      license
    );

  if (!row) {
    throw new Error(
      'Doctor not found'
    );
  }

  const name =
    String(
      sheet
        .getRange(
          row,
          columns.name
        )
        .getDisplayValue() || ''
    ).trim();

  const targetColumn =
    imageType === 'profile'
      ? columns.profile
      : columns.council;

  const targetRange =
    sheet.getRange(
      row,
      targetColumn
    );

  const oldFileId =
    getDriveFileIdFromNote_(
      targetRange.getNote()
    );

  const bytes =
    Utilities.base64Decode(
      base64
    );

  const cleanName =
    sanitizeFileName_(
      license +
      '-' +
      imageType +
      '-' +
      fileName
    );

  const blob =
    Utilities.newBlob(
      bytes,
      mimeType,
      cleanName
    );

  const folder =
    getDoctorAssetsFolder_();

  const file =
    folder.createFile(
      blob
    );

  try {
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  }
  catch (sharingError) {
    file.setTrashed(true);

    throw new Error(
      'Google Workspace ไม่อนุญาตให้แชร์รูปแบบ Anyone with the link กรุณาตรวจสิทธิ์ Drive ของบัญชีที่ Deploy Apps Script'
    );
  }

  /*
    Public source URL used by CellImageBuilder.
    The Drive file itself remains owned by the Apps Script account.
  */
  const sourceUrl =
    'https://drive.google.com/uc?export=view&id=' +
    encodeURIComponent(
      file.getId()
    );

  const cellImage =
    SpreadsheetApp
      .newCellImage()
      .setSourceUrl(
        sourceUrl
      )
      .setAltTextTitle(
        imageType === 'profile'
          ? 'Profile - ' + name
          : 'แพทยสภา - ' + name
      )
      .setAltTextDescription(
        license
      )
      .build();

  targetRange
    .setValue(
      cellImage
    );

  targetRange
    .setNote(
      'LUMIE_DRIVE_FILE_ID=' +
      file.getId()
    );

  SpreadsheetApp.flush();

  /*
    Trash previous website-uploaded file only after
    the new image is safely written to the sheet.
  */
  if (
    oldFileId &&
    oldFileId !== file.getId()
  ) {
    try {
      DriveApp
        .getFileById(
          oldFileId
        )
        .setTrashed(true);
    }
    catch (ignoreOldFileError) {}
  }

  const newInfo =
    getDoctorImageInfo_(
      targetRange
    );

  return jsonResponse({
    success: true,
    message:
      imageType === 'profile'
        ? 'Doctor profile updated'
        : 'Doctor council image updated',
    name: name,
    license: license,
    imageType: imageType,
    imageUrl:
      newInfo.url || '',
    fileId:
      file.getId()
  });
}


/* =========================
   DOCTOR HELPERS
========================= */

function ensureDoctorColumns_() {
  const sheet =
    getSheet(DOCTOR_SHEET);

  const readWidth =
    Math.max(
      sheet.getLastColumn(),
      4
    );

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        readWidth
      )
      .getDisplayValues()[0]
      .map(v =>
        String(v || '').trim()
      );

  const findHeader =
    names => {
      const wanted =
        names.map(v =>
          String(v).trim().toLowerCase()
        );

      for (
        let i = 0;
        i < headers.length;
        i++
      ) {
        if (
          wanted.includes(
            headers[i].toLowerCase()
          )
        ) {
          return i + 1;
        }
      }

      return 0;
    };

  let nameColumn =
    findHeader([
      'ชื่อหมอ',
      'ขื่อหมอ',
      'ชื่อแพทย์'
    ]);

  let licenseColumn =
    findHeader([
      'เลข ว',
      'เลข ว.',
      'เลขใบประกอบวิชาชีพ'
    ]);

  let profileColumn =
    findHeader([
      'Profile',
      'โปรไฟล์'
    ]);

  let councilColumn =
    findHeader([
      'แพทยสภา'
    ]);

  /*
    Backward compatibility:
    old Doctor sheet already used A=name and B=license.
  */
  if (!nameColumn) {
    nameColumn = 1;

    if (
      !String(
        sheet
          .getRange(1, 1)
          .getDisplayValue() || ''
      ).trim()
    ) {
      sheet
        .getRange(1, 1)
        .setValue('ชื่อหมอ');
    }
  }

  if (!licenseColumn) {
    licenseColumn = 2;

    if (
      !String(
        sheet
          .getRange(1, 2)
          .getDisplayValue() || ''
      ).trim()
    ) {
      sheet
        .getRange(1, 2)
        .setValue('เลข ว');
    }
  }

  if (!profileColumn) {
    profileColumn =
      sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        profileColumn
      )
      .setValue(
        'Profile'
      );
  }

  if (!councilColumn) {
    councilColumn =
      sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        councilColumn
      )
      .setValue(
        'แพทยสภา'
      );
  }

  SpreadsheetApp.flush();

  return {
    name: nameColumn,
    license: licenseColumn,
    profile: profileColumn,
    council: councilColumn
  };
}


function findDoctorRowByLicense_(
  sheet,
  columns,
  license
) {
  const lastRow =
    sheet.getLastRow();

  if (lastRow <= 1) {
    return 0;
  }

  const licenses =
    sheet
      .getRange(
        2,
        columns.license,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < licenses.length;
    i++
  ) {
    if (
      String(
        licenses[i][0] || ''
      ).trim()
      === license
    ) {
      return i + 2;
    }
  }

  return 0;
}


function getDoctorImageInfo_(
  range
) {
  const value =
    range.getValue();

  if (
    isCellImage_(value)
  ) {
    try {
      return {
        exists: true,
        url:
          value.getContentUrl() || ''
      };
    }
    catch (cellImageError) {
      return {
        exists: true,
        url: ''
      };
    }
  }

  if (
    isHttpUrl_(value)
  ) {
    return {
      exists: true,
      url:
        String(value).trim()
    };
  }

  const fileId =
    getDriveFileIdFromNote_(
      range.getNote()
    );

  if (fileId) {
    try {
      DriveApp.getFileById(
        fileId
      );

      return {
        exists: true,
        url:
          'https://drive.google.com/uc?export=view&id=' +
          encodeURIComponent(
            fileId
          )
      };
    }
    catch (fileError) {}
  }

  return {
    exists: false,
    url: ''
  };
}


function getDriveFileIdFromNote_(
  note
) {
  const match =
    String(note || '')
      .match(
        /LUMIE_DRIVE_FILE_ID=([A-Za-z0-9_-]+)/
      );

  return match
    ? match[1]
    : '';
}


function getDoctorAssetsFolder_() {
  const folderName =
    'Lumie Website - Doctor Assets';

  const folders =
    DriveApp
      .getFoldersByName(
        folderName
      );

  if (
    folders.hasNext()
  ) {
    return folders.next();
  }

  return DriveApp
    .createFolder(
      folderName
    );
}


function sanitizeFileName_(
  fileName
) {
  return String(
    fileName || 'doctor-image'
  )
    .replace(
      /[\\/:*?"<>|]+/g,
      '-'
    )
    .replace(
      /\s+/g,
      '-'
    )
    .slice(
      0,
      140
    );
}



/* =========================
   GET BRANCHES
========================= */

function getBranches() {
  const branches =
    getBranchesData_();

  return jsonResponse({
    success: true,
    count:
      branches.length,
    data:
      branches
  });
}


function getBranchesData_() {
  const sheet =
    getSheet(
      BRANCH_SHEET
    );

  const columns =
    ensureBranchColumns_();

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return [];
  }

  const data = [];

  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {
    const branchId =
      String(
        sheet
          .getRange(
            row,
            columns.id
          )
          .getDisplayValue() || ''
      ).trim();

    const name =
      String(
        sheet
          .getRange(
            row,
            columns.name
          )
          .getDisplayValue() || ''
      ).trim();

    if (
      !branchId &&
      !name
    ) {
      continue;
    }

    data.push({
      branchId:
        branchId,
      id:
        branchId,

      name:
        name,

      address:
        String(
          sheet
            .getRange(
              row,
              columns.address
            )
            .getDisplayValue() || ''
        ).trim(),

      phone:
        String(
          sheet
            .getRange(
              row,
              columns.phone
            )
            .getDisplayValue() || ''
        ).trim(),

      hours:
        String(
          sheet
            .getRange(
              row,
              columns.hours
            )
            .getDisplayValue() || ''
        ).trim(),

      map:
        String(
          sheet
            .getRange(
              row,
              columns.map
            )
            .getDisplayValue() || ''
        ).trim(),

      status:
        String(
          sheet
            .getRange(
              row,
              columns.status
            )
            .getDisplayValue() || ''
        ).trim()
    });
  }

  return data;
}


/* =========================
   ADD BRANCH
========================= */

function addBranch(data) {
  const branchId =
    String(
      data.branchId ||
      data.id ||
      ''
    ).trim();

  const name =
    String(
      data.name ||
      data.branchName ||
      ''
    ).trim();

  if (!branchId) {
    throw new Error(
      'Branch ID is required'
    );
  }

  if (!name) {
    throw new Error(
      'Branch name is required'
    );
  }

  const sheet =
    getSheet(
      BRANCH_SHEET
    );

  const columns =
    ensureBranchColumns_();

  if (
    findBranchRowById_(
      sheet,
      columns,
      branchId
    )
  ) {
    throw new Error(
      'Branch ID already exists'
    );
  }

  const row =
    Math.max(
      2,
      sheet.getLastRow() + 1
    );

  writeBranchRow_(
    sheet,
    row,
    columns,
    {
      branchId:
        branchId,
      name:
        name,
      address:
        data.address || '',
      phone:
        data.phone ||
        data.tel ||
        '',
      hours:
        data.hours ||
        data.openingHours ||
        '',
      map:
        data.map ||
        data.googleMap ||
        '',
      status:
        data.status || ''
    }
  );

  SpreadsheetApp.flush();

  return jsonResponse({
    success: true,
    message:
      'Branch added',
    branchId:
      branchId
  });
}


/* =========================
   UPDATE BRANCH
========================= */

function updateBranch(data) {
  const originalBranchId =
    String(
      data.originalBranchId ||
      data.branchId ||
      data.id ||
      ''
    ).trim();

  const newBranchId =
    String(
      data.branchId ||
      data.id ||
      originalBranchId
    ).trim();

  const name =
    String(
      data.name ||
      data.branchName ||
      ''
    ).trim();

  if (!originalBranchId) {
    throw new Error(
      'Original Branch ID is required'
    );
  }

  if (!newBranchId) {
    throw new Error(
      'Branch ID is required'
    );
  }

  if (!name) {
    throw new Error(
      'Branch name is required'
    );
  }

  const sheet =
    getSheet(
      BRANCH_SHEET
    );

  const columns =
    ensureBranchColumns_();

  const row =
    findBranchRowById_(
      sheet,
      columns,
      originalBranchId
    );

  if (!row) {
    throw new Error(
      'Branch not found'
    );
  }

  if (
    newBranchId !==
    originalBranchId
  ) {
    const duplicatedRow =
      findBranchRowById_(
        sheet,
        columns,
        newBranchId
      );

    if (
      duplicatedRow &&
      duplicatedRow !== row
    ) {
      throw new Error(
        'New Branch ID already exists'
      );
    }
  }

  writeBranchRow_(
    sheet,
    row,
    columns,
    {
      branchId:
        newBranchId,
      name:
        name,
      address:
        data.address || '',
      phone:
        data.phone ||
        data.tel ||
        '',
      hours:
        data.hours ||
        data.openingHours ||
        '',
      map:
        data.map ||
        data.googleMap ||
        '',
      status:
        data.status || ''
    }
  );

  SpreadsheetApp.flush();

  return jsonResponse({
    success: true,
    message:
      'Branch updated',
    originalBranchId:
      originalBranchId,
    branchId:
      newBranchId
  });
}


/* =========================
   BRANCH HELPERS
========================= */

function ensureBranchColumns_() {
  const sheet =
    getSheet(
      BRANCH_SHEET
    );

  const readWidth =
    Math.max(
      sheet.getLastColumn(),
      7
    );

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        readWidth
      )
      .getDisplayValues()[0]
      .map(v =>
        String(
          v || ''
        ).trim()
      );

  const findHeader =
    names => {
      const wanted =
        names.map(v =>
          String(v)
            .trim()
            .toLowerCase()
        );

      for (
        let i = 0;
        i < headers.length;
        i++
      ) {
        if (
          wanted.includes(
            headers[i]
              .toLowerCase()
          )
        ) {
          return i + 1;
        }
      }

      return 0;
    };

  const required = [
    {
      key: 'id',
      names: [
        'Branch ID',
        'รหัสสาขา'
      ],
      header: 'Branch ID'
    },
    {
      key: 'name',
      names: [
        'ชื่อสาขา',
        'Branch Name'
      ],
      header: 'ชื่อสาขา'
    },
    {
      key: 'address',
      names: [
        'ที่อยู่',
        'Address'
      ],
      header: 'ที่อยู่'
    },
    {
      key: 'phone',
      names: [
        'เบอร์โทร',
        'Phone',
        'Tel'
      ],
      header: 'เบอร์โทร'
    },
    {
      key: 'hours',
      names: [
        'เวลาเปิด',
        'Opening Hours',
        'Hours'
      ],
      header: 'เวลาเปิด'
    },
    {
      key: 'map',
      names: [
        'Google Map',
        'Google Maps',
        'Map'
      ],
      header: 'Google Map'
    },
    {
      key: 'status',
      names: [
        'Status',
        'สถานะ'
      ],
      header: 'Status'
    }
  ];

  const columns = {};

  required.forEach(item => {
    let column =
      findHeader(
        item.names
      );

    if (!column) {
      column =
        sheet.getLastColumn() + 1;

      sheet
        .getRange(
          1,
          column
        )
        .setValue(
          item.header
        );

      headers[
        column - 1
      ] =
        item.header;
    }

    columns[
      item.key
    ] =
      column;
  });

  SpreadsheetApp.flush();

  return columns;
}


function findBranchRowById_(
  sheet,
  columns,
  branchId
) {
  const target =
    String(
      branchId || ''
    ).trim();

  const lastRow =
    sheet.getLastRow();

  if (
    !target ||
    lastRow <= 1
  ) {
    return 0;
  }

  const ids =
    sheet
      .getRange(
        2,
        columns.id,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {
    if (
      String(
        ids[i][0] || ''
      ).trim()
      === target
    ) {
      return i + 2;
    }
  }

  return 0;
}


function writeBranchRow_(
  sheet,
  row,
  columns,
  branch
) {
  sheet
    .getRange(
      row,
      columns.id
    )
    .setValue(
      branch.branchId
    );

  sheet
    .getRange(
      row,
      columns.name
    )
    .setValue(
      branch.name
    );

  sheet
    .getRange(
      row,
      columns.address
    )
    .setValue(
      branch.address
    );

  sheet
    .getRange(
      row,
      columns.phone
    )
    .setValue(
      branch.phone
    );

  sheet
    .getRange(
      row,
      columns.hours
    )
    .setValue(
      branch.hours
    );

  sheet
    .getRange(
      row,
      columns.map
    )
    .setValue(
      branch.map
    );

  sheet
    .getRange(
      row,
      columns.status
    )
    .setValue(
      branch.status
    );
}




/* =========================
   CALENDAR
========================= */

/*
  Visible business columns:
  Date | Time | Location | Topic | Relate | Create by

  Internal ownership columns:
  Event ID | Created by ID

  Access:
  - Master / Executive / Employee can read.
  - They can create events.
  - Update/delete is allowed only when Created by ID matches the current Session.
  - Outsource levels cannot read or modify Calendar.
*/

function getCalendar(e) {
  const params =
    (
      e &&
      e.parameter
    )
      ? e.parameter
      : {};

  const calendar =
    getMainGoogleCalendar_();

  const timeZone =
    GOOGLE_CALENDAR_TIMEZONE;

  const today =
    new Date();

  const defaultFrom =
    new Date(
      today.getFullYear(),
      today.getMonth(),
      1
    );

  const defaultTo =
    new Date(
      today.getFullYear(),
      today.getMonth() + 1,
      8
    );

  const from =
    parseCalendarFeedDate_(
      params.from,
      defaultFrom
    );

  const to =
    parseCalendarFeedDate_(
      params.to,
      defaultTo
    );

  if (
    to.getTime() <=
    from.getTime()
  ) {
    throw new Error(
      'Calendar range is invalid'
    );
  }

  const maxRangeMs =
    75 *
    24 *
    60 *
    60 *
    1000;

  if (
    to.getTime() -
      from.getTime()
    >
    maxRangeMs
  ) {
    throw new Error(
      'Calendar range cannot exceed 75 days'
    );
  }

  const events =
    calendar
      .getEvents(
        from,
        to
      )
      .map(event =>
        serializeGoogleCalendarEvent_(
          event,
          calendar,
          timeZone
        )
      )
      .sort(
        (
          a,
          b
        ) => {
          if (
            a.startDate !==
            b.startDate
          ) {
            return String(
              a.startDate
            ).localeCompare(
              String(
                b.startDate
              )
            );
          }

          if (
            Boolean(
              a.allDay
            ) !==
            Boolean(
              b.allDay
            )
          ) {
            return a.allDay
              ? -1
              : 1;
          }

          return String(
            a.startTime || ''
          ).localeCompare(
            String(
              b.startTime || ''
            )
          );
        }
      );

  return jsonResponse({
    success:
      true,

    source:
      'Google Calendar',

    calendarId:
      GOOGLE_CALENDAR_ID,

    from:
      Utilities.formatDate(
        from,
        timeZone,
        'yyyy-MM-dd'
      ),

    to:
      Utilities.formatDate(
        to,
        timeZone,
        'yyyy-MM-dd'
      ),

    updatedAt:
      new Date()
        .toISOString(),

    count:
      events.length,

    data:
      events
  });
}


function parseCalendarFeedDate_(
  value,
  fallback
) {
  const text =
    String(
      value || ''
    )
      .trim();

  if (
    /^\d{4}-\d{2}-\d{2}$/
      .test(
        text
      )
  ) {
    try {
      return Utilities.parseDate(
        text,
        GOOGLE_CALENDAR_TIMEZONE,
        'yyyy-MM-dd'
      );
    }
    catch (err) {
      Logger.log(
        'Calendar date parse fallback: ' +
        err
      );
    }
  }

  return new Date(
    fallback
  );
}


function googleCalendarColorMap_() {
  const map = {};

  const add =
    (
      enumValue,
      hex
    ) => {
      if (
        enumValue ===
        undefined ||
        enumValue ===
        null
      ) {
        return;
      }

      map[
        String(
          enumValue
        )
      ] = hex;
    };

  add(
    CalendarApp.EventColor.PALE_BLUE,
    '#A4BDFC'
  );

  add(
    CalendarApp.EventColor.PALE_GREEN,
    '#7AE7BF'
  );

  add(
    CalendarApp.EventColor.MAUVE,
    '#DBADFF'
  );

  add(
    CalendarApp.EventColor.PALE_RED,
    '#FF887C'
  );

  add(
    CalendarApp.EventColor.YELLOW,
    '#FBD75B'
  );

  add(
    CalendarApp.EventColor.ORANGE,
    '#FFB878'
  );

  add(
    CalendarApp.EventColor.CYAN,
    '#46D6DB'
  );

  add(
    CalendarApp.EventColor.GRAY,
    '#E1E1E1'
  );

  add(
    CalendarApp.EventColor.BLUE,
    '#5484ED'
  );

  add(
    CalendarApp.EventColor.GREEN,
    '#51B749'
  );

  add(
    CalendarApp.EventColor.RED,
    '#DC2127'
  );

  return map;
}


function googleCalendarEventColor_(
  event,
  calendar
) {
  let rawColor = '';

  try {
    rawColor =
      String(
        event.getColor() || ''
      ).trim();
  }
  catch (err) {
    rawColor = '';
  }

  if (
    /^#[0-9A-F]{6}$/i
      .test(
        rawColor
      )
  ) {
    return {
      color:
        rawColor,
      colorKey:
        rawColor,
      colorSource:
        'event'
    };
  }

  const colorMap =
    googleCalendarColorMap_();

  if (
    rawColor &&
    colorMap[
      rawColor
    ]
  ) {
    return {
      color:
        colorMap[
          rawColor
        ],
      colorKey:
        rawColor,
      colorSource:
        'event'
    };
  }

  let calendarColor = '';

  try {
    calendarColor =
      String(
        calendar.getColor() || ''
      ).trim();
  }
  catch (err) {
    calendarColor = '';
  }

  if (
    /^#[0-9A-F]{6}$/i
      .test(
        calendarColor
      )
  ) {
    return {
      color:
        calendarColor,
      colorKey:
        rawColor,
      colorSource:
        'calendar'
    };
  }

  return {
    color:
      '#FDE696',
    colorKey:
      rawColor,
    colorSource:
      'fallback'
  };
}


function serializeGoogleCalendarEvent_(
  event,
  calendar,
  timeZone
) {
  const allDay =
    Boolean(
      event.isAllDayEvent()
    );

  const start =
    event.getStartTime();

  const end =
    event.getEndTime();

  const colorInfo =
    googleCalendarEventColor_(
      event,
      calendar
    );

  return {
    id:
      String(
        event.getId() || ''
      ),

    title:
      String(
        event.getTitle() || ''
      ),

    description:
      String(
        event.getDescription() || ''
      ),

    location:
      String(
        event.getLocation() || ''
      ),

    allDay:
      allDay,

    startDate:
      Utilities.formatDate(
        start,
        timeZone,
        'yyyy-MM-dd'
      ),

    endDate:
      Utilities.formatDate(
        end,
        timeZone,
        'yyyy-MM-dd'
      ),

    startTime:
      allDay
        ? ''
        : Utilities.formatDate(
            start,
            timeZone,
            'HH:mm'
          ),

    endTime:
      allDay
        ? ''
        : Utilities.formatDate(
            end,
            timeZone,
            'HH:mm'
          ),

    color:
      colorInfo.color,

    colorKey:
      colorInfo.colorKey,

    colorSource:
      colorInfo.colorSource
  };
}

function calendarTimeToMinutes_(
  value
) {
  const text =
    String(
      value || ''
    ).trim();

  const match =
    text.match(
      /^(\d{2}):(\d{2})$/
    );

  if (!match) {
    return null;
  }

  const hour =
    Number(
      match[1]
    );

  const minute =
    Number(
      match[2]
    );

  if (
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    return null;
  }

  return (
    hour * 60 +
    minute
  );
}


function defaultCalendarEndTime_(
  startTime
) {
  const start =
    calendarTimeToMinutes_(
      startTime
    );

  if (start === null) {
    return '';
  }

  const end =
    Math.min(
      start +
      GOOGLE_CALENDAR_DEFAULT_DURATION_MINUTES,
      (23 * 60) + 59
    );

  return (
    String(
      Math.floor(
        end / 60
      )
    ).padStart(
      2,
      '0'
    ) +
    ':' +
    String(
      end % 60
    ).padStart(
      2,
      '0'
    )
  );
}


function normalizeCalendarStatus_(
  value
) {
  const raw =
    String(
      value || ''
    )
      .trim()
      .toLowerCase();

  if (
    [
      'completed',
      'complete',
      'done',
      'true',
      '1',
      'yes'
    ].includes(
      raw
    )
  ) {
    return 'Completed';
  }

  return 'Pending';
}


function calendarTitleHasCompletedPrefix_(
  title
) {
  return /^(?:✅|☑️?|✔️?)\s*/.test(
    String(
      title || ''
    ).trim()
  );
}


function stripCalendarCompletedPrefix_(
  title
) {
  return String(
    title || ''
  )
    .trim()
    .replace(
      /^(?:✅|☑️?|✔️?)\s*/,
      ''
    )
    .trim();
}


function googleCalendarTitle_(
  topic,
  status
) {
  const clean =
    stripCalendarCompletedPrefix_(
      topic
    );

  return (
    normalizeCalendarStatus_(
      status
    ) === 'Completed'
      ? '✅ ' + clean
      : clean
  );
}


function getCalendarData_() {
  const sheet =
    getSheet(
      CALENDAR_SHEET
    );

  const columns =
    ensureCalendarColumns_();

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return [];
  }

  const creatorLookup =
    buildCalendarCreatorLookup_();

  const data = [];
  let identityChanged = false;

  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {
    const date =
      String(
        sheet
          .getRange(
            row,
            columns.date
          )
          .getDisplayValue() || ''
      ).trim();

    const time =
      String(
        sheet
          .getRange(
            row,
            columns.time
          )
          .getDisplayValue() || ''
      ).trim();

    let endTime =
      String(
        sheet
          .getRange(
            row,
            columns.endTime
          )
          .getDisplayValue() || ''
      ).trim();

    if (
      !endTime &&
      time
    ) {
      endTime =
        defaultCalendarEndTime_(
          time
        );

      if (endTime) {
        sheet
          .getRange(
            row,
            columns.endTime
          )
          .setNumberFormat('@')
          .setValue(
            endTime
          );

        identityChanged = true;
      }
    }

    const location =
      String(
        sheet
          .getRange(
            row,
            columns.location
          )
          .getDisplayValue() || ''
      ).trim();

    let topic =
      String(
        sheet
          .getRange(
            row,
            columns.topic
          )
          .getDisplayValue() || ''
      ).trim();

    const description =
      String(
        sheet
          .getRange(
            row,
            columns.description
          )
          .getDisplayValue() || ''
      ).trim();

    const relate =
      String(
        sheet
          .getRange(
            row,
            columns.relate
          )
          .getDisplayValue() || ''
      ).trim();

    const department =
      normalizeCalendarDepartment_(
        sheet
          .getRange(
            row,
            columns.department
          )
          .getDisplayValue()
      );

    const rawStatus =
      String(
        sheet
          .getRange(
            row,
            columns.status
          )
          .getDisplayValue() || ''
      ).trim();

    let status =
      normalizeCalendarStatus_(
        rawStatus
      );

    if (
      calendarTitleHasCompletedPrefix_(
        topic
      )
    ) {
      topic =
        stripCalendarCompletedPrefix_(
          topic
        );

      status =
        'Completed';

      sheet
        .getRange(
          row,
          columns.topic
        )
        .setValue(
          topic
        );

      identityChanged = true;
    }

    if (
      !rawStatus ||
      rawStatus !== status
    ) {
      sheet
        .getRange(
          row,
          columns.status
        )
        .setValue(
          status
        );

      identityChanged = true;
    }

    const createBy =
      String(
        sheet
          .getRange(
            row,
            columns.createBy
          )
          .getDisplayValue() || ''
      ).trim();

    if (
      !date &&
      !time &&
      !location &&
      !topic &&
      !description &&
      !relate &&
      !department &&
      !createBy
    ) {
      continue;
    }

    let eventId =
      String(
        sheet
          .getRange(
            row,
            columns.eventId
          )
          .getDisplayValue() || ''
      ).trim();

    let createdById =
      String(
        sheet
          .getRange(
            row,
            columns.createdById
          )
          .getDisplayValue() || ''
      ).trim();

    const googleEventId =
      String(
        sheet
          .getRange(
            row,
            columns.googleEventId
          )
          .getDisplayValue() || ''
      ).trim();

    if (!eventId) {
      eventId =
        Utilities.getUuid();

      sheet
        .getRange(
          row,
          columns.eventId
        )
        .setValue(
          eventId
        );

      identityChanged = true;
    }

    if (
      !createdById &&
      createBy
    ) {
      const inferred =
        creatorLookup[
          normalizeCalendarCreatorKey_(
            createBy
          )
        ];

      if (inferred) {
        createdById =
          inferred;

        sheet
          .getRange(
            row,
            columns.createdById
          )
          .setNumberFormat('@')
          .setValue(
            createdById
          );

        identityChanged = true;
      }
    }

    data.push({
      'Event ID':
        eventId,

      'Created by ID':
        createdById,

      'Google Event ID':
        googleEventId,

      Date:
        date,

      Time:
        time,

      'End Time':
        endTime,

      Location:
        location,

      Topic:
        topic,

      Description:
        description,

      Relate:
        relate,

      Department:
        department,

      Status:
        status,

      'Create by':
        createBy
    });
  }

  if (identityChanged) {
    SpreadsheetApp.flush();
  }

  return data;
}



function normalizeCalendarDepartment_(
  value
) {
  const raw =
    String(
      value || ''
    )
      .trim()
      .toLowerCase();

  if (raw === 'content') {
    return 'Content';
  }

  if (raw === 'marketing') {
    return 'Marketing';
  }

  if (raw === 'production') {
    return 'Production';
  }

  if (raw === 'kol') {
    return 'KOL';
  }

  return '';
}


function calendarGoogleColorForDepartment_(
  department
) {
  const value =
    normalizeCalendarDepartment_(
      department
    );

  if (value === 'Content') {
    return CalendarApp.EventColor.PALE_RED;
  }

  if (value === 'Marketing') {
    return CalendarApp.EventColor.YELLOW;
  }

  if (value === 'Production') {
    return CalendarApp.EventColor.CYAN;
  }

  if (value === 'KOL') {
    return CalendarApp.EventColor.MAUVE;
  }

  return '';
}


function departmentFromGoogleCalendarColor_(
  color
) {
  const value =
    String(
      color || ''
    );

  if (
    value ===
    String(
      CalendarApp.EventColor.PALE_RED
    )
  ) {
    return 'Content';
  }

  if (
    value ===
    String(
      CalendarApp.EventColor.YELLOW
    )
  ) {
    return 'Marketing';
  }

  if (
    value ===
      String(
        CalendarApp.EventColor.CYAN
      ) ||
    value ===
      String(
        CalendarApp.EventColor.PALE_BLUE
      )
  ) {
    return 'Production';
  }

  if (
    value ===
    String(
      CalendarApp.EventColor.MAUVE
    )
  ) {
    return 'KOL';
  }

  return '';
}


function applyGoogleCalendarDepartmentColor_(
  googleEvent,
  department
) {
  if (!googleEvent) {
    return;
  }

  const color =
    calendarGoogleColorForDepartment_(
      department
    );

  if (!color) {
    return;
  }

  googleEvent.setColor(
    color
  );
}


function parseGoogleCalendarDescription_(
  description
) {
  const raw =
    String(
      description || ''
    )
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n');

  const result = {
    description: '',
    relate: '',
    department: '',
    status: '',
    createBy: ''
  };

  const startMarker =
    '--- LUMIE SYNC ---';

  const endMarker =
    '--- END LUMIE SYNC ---';

  const startIndex =
    raw.indexOf(
      startMarker
    );

  let metadataText = '';

  if (
    startIndex >= 0
  ) {
    result.description =
      raw
        .slice(
          0,
          startIndex
        )
        .trim();

    const metadataStart =
      startIndex +
      startMarker.length;

    const endIndex =
      raw.indexOf(
        endMarker,
        metadataStart
      );

    metadataText =
      raw
        .slice(
          metadataStart,
          endIndex >= 0
            ? endIndex
            : raw.length
        )
        .trim();
  }
  else {
    /*
      Backward compatibility for events created before V27.
      Known LUMIE metadata lines are extracted; all other lines
      remain as the user's Google Calendar Description.
    */
    const descriptionLines = [];
    const metadataLines = [];

    raw
      .split('\n')
      .forEach(line => {
        if (
          /^(Relate|Department|Status|Create by|Source):\s*/i.test(
            line
          )
        ) {
          metadataLines.push(
            line
          );
        }
        else {
          descriptionLines.push(
            line
          );
        }
      });

    result.description =
      descriptionLines
        .join('\n')
        .trim();

    metadataText =
      metadataLines
        .join('\n');
  }

  metadataText
    .split(/\r?\n/)
    .forEach(line => {
      const relate =
        line.match(
          /^Relate:\s*(.*)$/i
        );

      if (relate) {
        result.relate =
          String(
            relate[1] || ''
          ).trim();
      }

      const department =
        line.match(
          /^Department:\s*(.*)$/i
        );

      if (department) {
        result.department =
          normalizeCalendarDepartment_(
            department[1]
          );
      }

      const status =
        line.match(
          /^Status:\s*(.*)$/i
        );

      if (status) {
        result.status =
          normalizeCalendarStatus_(
            status[1]
          );
      }

      const createBy =
        line.match(
          /^Create by:\s*(.*)$/i
        );

      if (createBy) {
        result.createBy =
          String(
            createBy[1] || ''
          ).trim();
      }
    });

  return result;
}

function formatGoogleCalendarDate_(
  date
) {
  return Utilities.formatDate(
    date,
    GOOGLE_CALENDAR_TIMEZONE,
    'yyyy-MM-dd'
  );
}


function formatGoogleCalendarTime_(
  date
) {
  return Utilities.formatDate(
    date,
    GOOGLE_CALENDAR_TIMEZONE,
    'HH:mm'
  );
}


function googleCalendarEventTimes_(
  googleEvent
) {
  if (
    googleEvent.isAllDayEvent &&
    googleEvent.isAllDayEvent()
  ) {
    return {
      date:
        formatGoogleCalendarDate_(
          googleEvent.getStartTime()
        ),
      time:
        '00:00',
      endTime:
        '23:59'
    };
  }

  const start =
    googleEvent.getStartTime();

  const end =
    googleEvent.getEndTime();

  const startDate =
    formatGoogleCalendarDate_(
      start
    );

  const endDate =
    formatGoogleCalendarDate_(
      end
    );

  let endTime =
    formatGoogleCalendarTime_(
      end
    );

  if (
    endDate !==
    startDate
  ) {
    endTime =
      '23:59';
  }

  return {
    date:
      startDate,
    time:
      formatGoogleCalendarTime_(
        start
      ),
    endTime:
      endTime
  };
}


function calendarEmployeeFromGoogleCreators_(
  googleEvent
) {
  let creators = [];

  try {
    if (
      googleEvent &&
      typeof googleEvent.getCreators ===
        'function'
    ) {
      creators =
        googleEvent.getCreators() || [];
    }
  }
  catch (err) {
    creators = [];
  }

  if (!creators.length) {
    return null;
  }

  const creatorEmails =
    creators
      .map(value =>
        String(
          value || ''
        )
          .trim()
          .toLowerCase()
      )
      .filter(Boolean);

  if (!creatorEmails.length) {
    return null;
  }

  const employees =
    getEmployeesData_();

  for (
    let i = 0;
    i < employees.length;
    i++
  ) {
    const employee =
      employees[i];

    const email =
      String(
        employee.email ||
        employee.Email ||
        ''
      )
        .trim()
        .toLowerCase();

    if (
      email &&
      creatorEmails.includes(
        email
      )
    ) {
      return employee;
    }
  }

  return null;
}


function calendarSheetRowRecord_(
  sheet,
  row,
  columns
) {
  return {
    row:
      row,

    eventId:
      String(
        sheet.getRange(
          row,
          columns.eventId
        ).getDisplayValue() || ''
      ).trim(),

    createdById:
      String(
        sheet.getRange(
          row,
          columns.createdById
        ).getDisplayValue() || ''
      ).trim(),

    googleEventId:
      String(
        sheet.getRange(
          row,
          columns.googleEventId
        ).getDisplayValue() || ''
      ).trim(),

    date:
      String(
        sheet.getRange(
          row,
          columns.date
        ).getDisplayValue() || ''
      ).trim(),

    time:
      String(
        sheet.getRange(
          row,
          columns.time
        ).getDisplayValue() || ''
      ).trim(),

    endTime:
      String(
        sheet.getRange(
          row,
          columns.endTime
        ).getDisplayValue() || ''
      ).trim(),

    location:
      String(
        sheet.getRange(
          row,
          columns.location
        ).getDisplayValue() || ''
      ).trim(),

    topic:
      String(
        sheet.getRange(
          row,
          columns.topic
        ).getDisplayValue() || ''
      ).trim(),

    description:
      String(
        sheet.getRange(
          row,
          columns.description
        ).getDisplayValue() || ''
      ).trim(),

    relate:
      String(
        sheet.getRange(
          row,
          columns.relate
        ).getDisplayValue() || ''
      ).trim(),

    department:
      normalizeCalendarDepartment_(
        sheet.getRange(
          row,
          columns.department
        ).getDisplayValue()
      ),

    status:
      normalizeCalendarStatus_(
        sheet.getRange(
          row,
          columns.status
        ).getDisplayValue()
      ),

    createBy:
      String(
        sheet.getRange(
          row,
          columns.createBy
        ).getDisplayValue() || ''
      ).trim()
  };
}


function findMatchingGoogleEventForSheetRow_(
  calendar,
  record
) {
  if (
    !record.date ||
    !record.time ||
    !record.topic
  ) {
    return null;
  }

  const dayStart =
    new Date(
      record.date +
      'T00:00:00+07:00'
    );

  const dayEnd =
    new Date(
      dayStart.getTime() +
      24 * 60 * 60 * 1000
    );

  const candidates =
    calendar.getEvents(
      dayStart,
      dayEnd
    );

  const wantedTitle =
    stripCalendarCompletedPrefix_(
      record.topic
    );

  const wantedTime =
    String(
      record.time || ''
    ).trim();

  let matches =
    candidates.filter(event =>
      stripCalendarCompletedPrefix_(
        event.getTitle()
      ) ===
        wantedTitle &&
      formatGoogleCalendarTime_(
        event.getStartTime()
      ) ===
        wantedTime
    );

  if (
    matches.length > 1 &&
    record.location
  ) {
    const locationMatches =
      matches.filter(event =>
        String(
          event.getLocation() || ''
        ).trim() ===
          record.location
      );

    if (
      locationMatches.length === 1
    ) {
      matches =
        locationMatches;
    }
  }

  return (
    matches.length === 1
      ? matches[0]
      : null
  );
}


function googleCalendarEventToSheetEvent_(
  googleEvent,
  existingRecord
) {
  const parsed =
    parseGoogleCalendarDescription_(
      googleEvent.getDescription()
    );

  const times =
    googleCalendarEventTimes_(
      googleEvent
    );

  const colorDepartment =
    departmentFromGoogleCalendarColor_(
      googleEvent.getColor()
    );

  const employee =
    existingRecord &&
    existingRecord.createdById
      ? null
      : calendarEmployeeFromGoogleCreators_(
          googleEvent
        );

  let createdById =
    existingRecord
      ? String(
          existingRecord.createdById || ''
        ).trim()
      : '';

  let createBy =
    existingRecord
      ? String(
          existingRecord.createBy || ''
        ).trim()
      : '';

  if (!createdById && employee) {
    createdById =
      String(
        employee.id || ''
      ).trim();
  }

  if (
    parsed.createBy
  ) {
    createBy =
      parsed.createBy;
  }
  else if (
    !createBy &&
    employee
  ) {
    createBy =
      calendarCreatorLabelFromEmployee_(
        employee
      );
  }
  else if (!createBy) {
    createBy =
      'Google Calendar';
  }

  const department =
    colorDepartment ||
    parsed.department ||
    (
      existingRecord
        ? normalizeCalendarDepartment_(
            existingRecord.department
          )
        : ''
    );

  const description =
    parsed.description ||
    (
      existingRecord
        ? String(
            existingRecord.description || ''
          ).trim()
        : ''
    );

  const relate =
    parsed.relate ||
    (
      existingRecord
        ? String(
            existingRecord.relate || ''
          ).trim()
        : ''
    );

  const rawTitle =
    String(
      googleEvent.getTitle() || ''
    ).trim();

  const status =
    calendarTitleHasCompletedPrefix_(
      rawTitle
    )
      ? 'Completed'
      : 'Pending';

  return {
    eventId:
      existingRecord &&
      existingRecord.eventId
        ? existingRecord.eventId
        : Utilities.getUuid(),

    createdById:
      createdById,

    googleEventId:
      String(
        googleEvent.getId() || ''
      ).trim(),

    date:
      times.date,

    time:
      times.time,

    endTime:
      times.endTime,

    location:
      String(
        googleEvent.getLocation() || ''
      ).trim(),

    topic:
      stripCalendarCompletedPrefix_(
        rawTitle
      ) ||
      'Untitled Event',

    description:
      description,

    relate:
      relate,

    department:
      department,

    status:
      status,

    createBy:
      createBy
  };
}


function syncGoogleCalendarToSheet_() {
  const lock =
    LockService.getScriptLock();

  lock.waitLock(
    20000
  );

  try {
    const calendar =
      getMainGoogleCalendar_();

    const sheet =
      getSheet(
        CALENDAR_SHEET
      );

    const columns =
      ensureCalendarColumns_();

    let linked = 0;
    let updated = 0;
    let imported = 0;
    let deleted = 0;

    /*
      Step 1:
      Existing Sheet rows without Google Event ID
      -> link to an exact Google event if possible,
      otherwise create one.
    */
    for (
      let row = 2;
      row <= sheet.getLastRow();
      row++
    ) {
      const record =
        calendarSheetRowRecord_(
          sheet,
          row,
          columns
        );

      if (
        record.googleEventId ||
        !record.date ||
        !record.time ||
        !record.topic
      ) {
        continue;
      }

      let googleEvent =
        findMatchingGoogleEventForSheetRow_(
          calendar,
          record
        );

      if (!googleEvent) {
        const event =
          validateCalendarPayload_({
            date:
              record.date,
            time:
              record.time,
            endTime:
              record.endTime,
            location:
              record.location,
            topic:
              record.topic,
            description:
              record.description,
            relate:
              record.relate ||
              '-',
            department:
              record.department,
            status:
              record.status
          });

        googleEvent =
          createGoogleCalendarEvent_(
            event,
            record.createBy
          );
      }
      else {
        if (
          record.department
        ) {
          applyGoogleCalendarDepartmentColor_(
            googleEvent,
            record.department
          );
        }
      }

      sheet
        .getRange(
          row,
          columns.googleEventId
        )
        .setNumberFormat('@')
        .setValue(
          googleEvent.getId()
        );

      linked++;
    }

    SpreadsheetApp.flush();

    /*
      Step 2:
      Google is authoritative for rows already linked
      when this function runs because it is invoked by
      Google Calendar trigger / website pull refresh.
      Missing Google event = deleted from Google -> remove Sheet row.
    */
    for (
      let row =
        sheet.getLastRow();
      row >= 2;
      row--
    ) {
      const record =
        calendarSheetRowRecord_(
          sheet,
          row,
          columns
        );

      if (
        !record.googleEventId
      ) {
        continue;
      }

      let googleEvent = null;

      try {
        googleEvent =
          calendar.getEventById(
            record.googleEventId
          );
      }
      catch (err) {
        Logger.log(
          'getEventById failed for ' +
          record.googleEventId +
          ': ' +
          err
        );

        continue;
      }

      if (!googleEvent) {
        sheet.deleteRow(
          row
        );

        deleted++;
        continue;
      }

      const synced =
        googleCalendarEventToSheetEvent_(
          googleEvent,
          record
        );

      writeCalendarRow_(
        sheet,
        row,
        columns,
        synced
      );

      updated++;
    }

    SpreadsheetApp.flush();

    /*
      Step 3:
      Import Google Calendar events that do not exist
      in Sheet yet.
    */
    const knownIds = {};

    for (
      let row = 2;
      row <= sheet.getLastRow();
      row++
    ) {
      const googleEventId =
        String(
          sheet.getRange(
            row,
            columns.googleEventId
          ).getDisplayValue() || ''
        ).trim();

      if (googleEventId) {
        knownIds[
          googleEventId
        ] = true;
      }
    }

    const now =
      new Date();

    const scanStart =
      new Date(
        now.getTime() -
        GOOGLE_CALENDAR_IMPORT_PAST_DAYS *
        24 * 60 * 60 * 1000
      );

    const scanEnd =
      new Date(
        now.getTime() +
        GOOGLE_CALENDAR_IMPORT_FUTURE_DAYS *
        24 * 60 * 60 * 1000
      );

    const googleEvents =
      calendar.getEvents(
        scanStart,
        scanEnd
      );

    googleEvents.forEach(
      googleEvent => {
        const googleEventId =
          String(
            googleEvent.getId() || ''
          ).trim();

        if (
          !googleEventId ||
          knownIds[
            googleEventId
          ]
        ) {
          return;
        }

        const synced =
          googleCalendarEventToSheetEvent_(
            googleEvent,
            null
          );

        const row =
          Math.max(
            2,
            sheet.getLastRow() + 1
          );

        writeCalendarRow_(
          sheet,
          row,
          columns,
          synced
        );

        knownIds[
          googleEventId
        ] = true;

        imported++;
      }
    );

    SpreadsheetApp.flush();

    return {
      success:
        true,
      linked:
        linked,
      updated:
        updated,
      imported:
        imported,
      deleted:
        deleted
    };
  }
  finally {
    lock.releaseLock();
  }
}


function calendarTwoWaySyncTrigger(
  e
) {
  /*
    V31:
    The website Calendar reads Google Calendar directly.
    Legacy Google Calendar -> Sheet synchronization is no longer
    required for the website view.

    Keeping this handler as a no-op prevents an older installed
    trigger from causing errors until it is removed manually.
  */
  Logger.log(
    'Legacy Calendar sync trigger skipped: V31 read-only Google Calendar feed is active.'
  );

  return {
    success:
      true,
    skipped:
      true
  };
}

function removeLegacyCalendarSyncTriggers() {
  let removed = 0;

  ScriptApp
    .getProjectTriggers()
    .forEach(
      trigger => {
        if (
          trigger.getHandlerFunction() ===
          GOOGLE_CALENDAR_TRIGGER_HANDLER
        ) {
          ScriptApp.deleteTrigger(
            trigger
          );

          removed++;
        }
      }
    );

  Logger.log(
    'Removed legacy Calendar triggers: ' +
    removed
  );

  return {
    success:
      true,
    removed:
      removed
  };
}


function syncCalendarTwoWayNow() {
  const result =
    syncGoogleCalendarToSheet_();

  Logger.log(
    JSON.stringify(
      result
    )
  );

  return result;
}


function setupCalendarTwoWaySync() {
  /*
    Run manually ONCE from Apps Script editor.
    It:
    1) requests Calendar permission,
    2) links/imports current events,
    3) installs an event-updated trigger,
    4) installs an hourly safety sync.
  */
  getMainGoogleCalendar_();

  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(
    trigger => {
      if (
        trigger.getHandlerFunction() ===
        GOOGLE_CALENDAR_TRIGGER_HANDLER
      ) {
        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );

  const initial =
    syncGoogleCalendarToSheet_();

  ScriptApp
    .newTrigger(
      GOOGLE_CALENDAR_TRIGGER_HANDLER
    )
    .forUserCalendar(
      GOOGLE_CALENDAR_ID
    )
    .onEventUpdated()
    .create();

  ScriptApp
    .newTrigger(
      GOOGLE_CALENDAR_TRIGGER_HANDLER
    )
    .timeBased()
    .everyHours(1)
    .create();

  const result = {
    success:
      true,
    initial:
      initial,
    triggers:
      2
  };

  Logger.log(
    JSON.stringify(
      result
    )
  );

  return result;
}


function getMainGoogleCalendar_() {
  const calendar =
    CalendarApp.getCalendarById(
      GOOGLE_CALENDAR_ID
    );

  if (!calendar) {
    throw new Error(
      'Google Calendar not found or Apps Script has no access'
    );
  }

  return calendar;
}


function calendarGoogleStart_(
  date,
  time
) {
  const value =
    new Date(
      String(date) +
      'T' +
      String(time) +
      ':00+07:00'
    );

  if (isNaN(value.getTime())) {
    throw new Error(
      'Invalid Calendar date/time'
    );
  }

  return value;
}


function calendarGoogleEnd_(
  date,
  endTime
) {
  return calendarGoogleStart_(
    date,
    endTime
  );
}


function calendarGoogleDescription_(
  event,
  creator
) {
  const userDescription =
    String(
      event.description || ''
    ).trim();

  const metadata = [
    '--- LUMIE SYNC ---',
    event.relate
      ? 'Relate: ' + event.relate
      : 'Relate:',
    event.department
      ? 'Department: ' + normalizeCalendarDepartment_(event.department)
      : 'Department:',
    'Status: ' + normalizeCalendarStatus_(event.status),
    creator
      ? 'Create by: ' + creator
      : 'Create by:',
    'Source: LUMIE Internal Data Hub',
    '--- END LUMIE SYNC ---'
  ]
    .join('\n');

  return [
    userDescription,
    metadata
  ]
    .filter(Boolean)
    .join('\n\n');
}

function createGoogleCalendarEvent_(
  event,
  creator
) {
  const calendar =
    getMainGoogleCalendar_();

  const start =
    calendarGoogleStart_(
      event.date,
      event.time
    );

  const end =
    calendarGoogleEnd_(
      event.date,
      event.endTime
    );

  const googleEvent =
    calendar.createEvent(
      googleCalendarTitle_(
        event.topic,
        event.status
      ),
      start,
      end,
      {
        location:
          event.location,
        description:
          calendarGoogleDescription_(
            event,
            creator
          )
      }
    );

  applyGoogleCalendarDepartmentColor_(
    googleEvent,
    event.department
  );

  return googleEvent;
}


function updateGoogleCalendarEvent_(
  googleEventId,
  event,
  creator
) {
  const calendar =
    getMainGoogleCalendar_();

  let googleEvent = null;

  if (googleEventId) {
    googleEvent =
      calendar.getEventById(
        googleEventId
      );
  }

  if (!googleEvent) {
    return createGoogleCalendarEvent_(
      event,
      creator
    );
  }

  const start =
    calendarGoogleStart_(
      event.date,
      event.time
    );

  const end =
    calendarGoogleEnd_(
      event.date,
      event.endTime
    );

  googleEvent
    .setTitle(
      googleCalendarTitle_(
        event.topic,
        event.status
      )
    )
    .setTime(
      start,
      end
    )
    .setLocation(
      event.location
    )
    .setDescription(
      calendarGoogleDescription_(
        event,
        creator
      )
    );

  applyGoogleCalendarDepartmentColor_(
    googleEvent,
    event.department
  );

  return googleEvent;
}


function deleteGoogleCalendarEvent_(
  googleEventId
) {
  if (!googleEventId) {
    return false;
  }

  const calendar =
    getMainGoogleCalendar_();

  const googleEvent =
    calendar.getEventById(
      googleEventId
    );

  if (!googleEvent) {
    return false;
  }

  googleEvent.deleteEvent();

  return true;
}


function getCalendarGoogleEventIdForRow_(
  sheet,
  row,
  columns
) {
  return String(
    sheet
      .getRange(
        row,
        columns.googleEventId
      )
      .getDisplayValue() || ''
  ).trim();
}


function addCalendar(
  data,
  session
) {
  requireCalendarViewer_(
    session
  );

  const employee =
    getEmployeeRecordById_(
      session.id
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  const event =
    validateCalendarPayload_(
      data
    );

  const creator =
    calendarCreatorLabelFromEmployee_(
      employee
    );

  const sheet =
    getSheet(
      CALENDAR_SHEET
    );

  const columns =
    ensureCalendarColumns_();

  const lock =
    LockService
      .getScriptLock();

  lock.waitLock(10000);

  let googleEvent = null;

  try {
    const row =
      Math.max(
        2,
        sheet.getLastRow() + 1
      );

    const eventId =
      Utilities.getUuid();

    googleEvent =
      createGoogleCalendarEvent_(
        event,
        creator
      );

    const googleEventId =
      String(
        googleEvent.getId() || ''
      ).trim();

    writeCalendarRow_(
      sheet,
      row,
      columns,
      {
        eventId:
          eventId,
        createdById:
          employee.id,
        googleEventId:
          googleEventId,
        date:
          event.date,
        time:
          event.time,
        endTime:
          event.endTime,
        location:
          event.location,
        topic:
          event.topic,
        description:
          event.description,
        relate:
          event.relate,
        department:
          event.department,
        status:
          event.status,
        createBy:
          creator
      }
    );

    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      message:
        'Calendar event added and synced to Google Calendar',
      data: {
        'Event ID':
          eventId,
        'Created by ID':
          employee.id,
        'Google Event ID':
          googleEventId,
        Date:
          event.date,
        Time:
          event.time,
        'End Time':
          event.endTime,
        Location:
          event.location,
        Topic:
          event.topic,
        Description:
          event.description,
        Relate:
          event.relate,
        Department:
          event.department,
        Status:
          event.status,
        'Create by':
          creator
      }
    });
  }
  catch (err) {
    if (googleEvent) {
      try {
        googleEvent.deleteEvent();
      }
      catch (rollbackErr) {
        Logger.log(
          'Google Calendar rollback failed: ' +
          rollbackErr
        );
      }
    }

    throw err;
  }
  finally {
    lock.releaseLock();
  }
}

function updateCalendar(
  data,
  session
) {
  requireCalendarViewer_(
    session
  );

  const eventId =
    String(
      data.eventId ||
      data['Event ID'] ||
      ''
    ).trim();

  if (!eventId) {
    throw new Error(
      'Event ID is required'
    );
  }

  const employee =
    getEmployeeRecordById_(
      session.id
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  const event =
    validateCalendarPayload_(
      data
    );

  const sheet =
    getSheet(
      CALENDAR_SHEET
    );

  const columns =
    ensureCalendarColumns_();

  const lock =
    LockService
      .getScriptLock();

  lock.waitLock(10000);

  try {
    const row =
      findCalendarRowByEventId_(
        sheet,
        columns,
        eventId
      );

    if (!row) {
      throw new Error(
        'Calendar event not found'
      );
    }

    const ownerId =
      ensureCalendarOwnerForRow_(
        sheet,
        row,
        columns
      );

    /*
      Shared internal Calendar:
      Master / Executive / Manager / Employee can edit any event.
      Outsource is already blocked by requireCalendarViewer_().
      Keep original owner/Create by for audit history.
    */

    const originalCreator =
      String(
        sheet
          .getRange(
            row,
            columns.createBy
          )
          .getDisplayValue() || ''
      ).trim();

    const creatorLabel =
      originalCreator ||
      calendarCreatorLabelFromEmployee_(
        employee
      );

    const existingGoogleEventId =
      getCalendarGoogleEventIdForRow_(
        sheet,
        row,
        columns
      );

    const googleEvent =
      updateGoogleCalendarEvent_(
        existingGoogleEventId,
        event,
        creatorLabel
      );

    const googleEventId =
      String(
        googleEvent.getId() ||
        existingGoogleEventId ||
        ''
      ).trim();

    writeCalendarRow_(
      sheet,
      row,
      columns,
      {
        eventId:
          eventId,
        createdById:
          ownerId || '',
        googleEventId:
          googleEventId,
        date:
          event.date,
        time:
          event.time,
        endTime:
          event.endTime,
        location:
          event.location,
        topic:
          event.topic,
        description:
          event.description,
        relate:
          event.relate,
        department:
          event.department,
        status:
          event.status,
        createBy:
          creatorLabel
      }
    );

    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      message:
        'Calendar event updated',
      eventId:
        eventId
    });
  }
  finally {
    lock.releaseLock();
  }
}


function deleteCalendar(
  data,
  session
) {
  requireCalendarViewer_(
    session
  );

  const eventId =
    String(
      data.eventId ||
      data['Event ID'] ||
      ''
    ).trim();

  if (!eventId) {
    throw new Error(
      'Event ID is required'
    );
  }

  const employee =
    getEmployeeRecordById_(
      session.id
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  const sheet =
    getSheet(
      CALENDAR_SHEET
    );

  const columns =
    ensureCalendarColumns_();

  const lock =
    LockService
      .getScriptLock();

  lock.waitLock(10000);

  try {
    const row =
      findCalendarRowByEventId_(
        sheet,
        columns,
        eventId
      );

    if (!row) {
      throw new Error(
        'Calendar event not found'
      );
    }

    const ownerId =
      ensureCalendarOwnerForRow_(
        sheet,
        row,
        columns
      );

    /*
      Shared internal Calendar:
      Master / Executive / Manager / Employee can delete any event.
      Outsource is already blocked by requireCalendarViewer_().
      ownerId is retained only for audit/history.
    */

    const googleEventId =
      getCalendarGoogleEventIdForRow_(
        sheet,
        row,
        columns
      );

    deleteGoogleCalendarEvent_(
      googleEventId
    );

    sheet.deleteRow(
      row
    );

    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      message:
        'Calendar event deleted',
      eventId:
        eventId
    });
  }
  finally {
    lock.releaseLock();
  }
}



/*
  ONE-TIME MIGRATION FOR EXISTING SHEET EVENTS
  Run this manually from Apps Script editor once after authorizing Calendar access.
  It creates Google Calendar events only for rows that do not yet have Google Event ID.
*/
function syncExistingCalendarEventsToGoogleCalendar() {
  /*
    Backward-compatible helper.
    In V20 this uses the matching-aware Two-way Sync engine,
    so re-running it will not blindly duplicate Calendar events.
  */
  return syncGoogleCalendarToSheet_();
}


function ensureCalendarColumns_() {
  const sheet =
    getSheet(
      CALENDAR_SHEET
    );

  const required = [
    {
      key:
        'date',
      header:
        'Date'
    },
    {
      key:
        'time',
      header:
        'Time'
    },
    {
      key:
        'endTime',
      header:
        'End Time'
    },
    {
      key:
        'location',
      header:
        'Location'
    },
    {
      key:
        'topic',
      header:
        'Topic'
    },
    {
      key:
        'description',
      header:
        'Description'
    },
    {
      key:
        'relate',
      header:
        'Relate'
    },
    {
      key:
        'department',
      header:
        'Department'
    },
    {
      key:
        'status',
      header:
        'Status'
    },
    {
      key:
        'createBy',
      header:
        'Create by'
    },
    {
      key:
        'eventId',
      header:
        'Event ID'
    },
    {
      key:
        'createdById',
      header:
        'Created by ID'
    },
    {
      key:
        'googleEventId',
      header:
        'Google Event ID'
    }
  ];

  const readWidth =
    Math.max(
      sheet.getLastColumn(),
      required.length
    );

  let headers =
    sheet
      .getRange(
        1,
        1,
        1,
        readWidth
      )
      .getDisplayValues()[0]
      .map(v =>
        String(
          v || ''
        ).trim()
      );

  const result = {};

  required.forEach(
    item => {
      const wanted =
        item.header
          .toLowerCase();

      let column = 0;

      for (
        let i = 0;
        i < headers.length;
        i++
      ) {
        if (
          headers[i]
            .toLowerCase()
          === wanted
        ) {
          column =
            i + 1;
          break;
        }
      }

      if (!column) {
        column =
          sheet.getLastColumn() + 1;

        sheet
          .getRange(
            1,
            column
          )
          .setValue(
            item.header
          );

        while (
          headers.length <
          column
        ) {
          headers.push('');
        }

        headers[
          column - 1
        ] =
          item.header;
      }

      result[
        item.key
      ] =
        column;
    }
  );

  SpreadsheetApp.flush();

  return result;
}


function writeCalendarRow_(
  sheet,
  row,
  columns,
  event
) {
  sheet
    .getRange(
      row,
      columns.date
    )
    .setNumberFormat('@')
    .setValue(
      event.date
    );

  sheet
    .getRange(
      row,
      columns.time
    )
    .setNumberFormat('@')
    .setValue(
      event.time
    );

  sheet
    .getRange(
      row,
      columns.endTime
    )
    .setNumberFormat('@')
    .setValue(
      event.endTime || ''
    );

  sheet
    .getRange(
      row,
      columns.location
    )
    .setValue(
      safeSheetText_(
        event.location
      )
    );

  sheet
    .getRange(
      row,
      columns.topic
    )
    .setValue(
      safeSheetText_(
        event.topic
      )
    );

  sheet
    .getRange(
      row,
      columns.description
    )
    .setValue(
      safeSheetText_(
        event.description || ''
      )
    );

  sheet
    .getRange(
      row,
      columns.relate
    )
    .setValue(
      safeSheetText_(
        event.relate
      )
    );

  sheet
    .getRange(
      row,
      columns.department
    )
    .setValue(
      normalizeCalendarDepartment_(
        event.department
      )
    );

  sheet
    .getRange(
      row,
      columns.status
    )
    .setValue(
      normalizeCalendarStatus_(
        event.status
      )
    );

  sheet
    .getRange(
      row,
      columns.createBy
    )
    .setValue(
      safeSheetText_(
        event.createBy
      )
    );

  sheet
    .getRange(
      row,
      columns.eventId
    )
    .setNumberFormat('@')
    .setValue(
      event.eventId
    );

  sheet
    .getRange(
      row,
      columns.createdById
    )
    .setNumberFormat('@')
    .setValue(
      event.createdById
    );

  sheet
    .getRange(
      row,
      columns.googleEventId
    )
    .setNumberFormat('@')
    .setValue(
      event.googleEventId || ''
    );
}


function findCalendarRowByEventId_(
  sheet,
  columns,
  eventId
) {
  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return 0;
  }

  const ids =
    sheet
      .getRange(
        2,
        columns.eventId,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {
    if (
      String(
        ids[i][0] || ''
      ).trim()
      === eventId
    ) {
      return i + 2;
    }
  }

  return 0;
}


function ensureCalendarOwnerForRow_(
  sheet,
  row,
  columns
) {
  let ownerId =
    String(
      sheet
        .getRange(
          row,
          columns.createdById
        )
        .getDisplayValue() || ''
    ).trim();

  if (ownerId) {
    return ownerId;
  }

  const createBy =
    String(
      sheet
        .getRange(
          row,
          columns.createBy
        )
        .getDisplayValue() || ''
    ).trim();

  if (!createBy) {
    return '';
  }

  const lookup =
    buildCalendarCreatorLookup_();

  const inferred =
    lookup[
      normalizeCalendarCreatorKey_(
        createBy
      )
    ] || '';

  if (inferred) {
    sheet
      .getRange(
        row,
        columns.createdById
      )
      .setNumberFormat('@')
      .setValue(
        inferred
      );

    SpreadsheetApp.flush();
  }

  return inferred;
}


function buildCalendarCreatorLookup_() {
  const lookup = {};

  const employees =
    getEmployeesData_();

  employees.forEach(
    employee => {
      const label =
        [
          String(
            employee.name || ''
          ).trim(),
          String(
            employee.position || ''
          ).trim()
        ]
          .filter(Boolean)
          .join(' · ');

      const key =
        normalizeCalendarCreatorKey_(
          label
        );

      if (!key) {
        return;
      }

      if (
        Object.prototype
          .hasOwnProperty
          .call(
            lookup,
            key
          )
      ) {
        // Duplicate identical name + position: do not infer ownership.
        lookup[key] = '';
      }
      else {
        lookup[key] =
          String(
            employee.id ||
            employee.employeeId ||
            ''
          ).trim();
      }
    }
  );

  return lookup;
}


function normalizeCalendarCreatorKey_(
  value
) {
  return String(
    value || ''
  )
    .trim()
    .replace(
      /\s+/g,
      ' '
    )
    .toLowerCase();
}


function calendarCreatorLabelFromEmployee_(
  employee
) {
  return [
    String(
      employee.name || ''
    ).trim(),
    String(
      employee.position || ''
    ).trim()
  ]
    .filter(Boolean)
    .join(' · ');
}


function validateCalendarPayload_(
  data
) {
  const date =
    String(
      data.date ||
      data.Date ||
      ''
    ).trim();

  const time =
    String(
      data.time ||
      data.Time ||
      ''
    ).trim();

  let endTime =
    String(
      data.endTime ||
      data['End Time'] ||
      data.EndTime ||
      ''
    ).trim();

  const location =
    sanitizeCalendarText_(
      data.location ||
      data.Location ||
      '',
      'Location',
      300
    );

  const topic =
    sanitizeCalendarText_(
      data.topic ||
      data.Topic ||
      '',
      'Topic',
      500
    );

  const description =
    sanitizeCalendarText_(
      data.description ||
      data.Description ||
      '',
      'Description',
      8000
    );

  const relate =
    sanitizeCalendarText_(
      data.relate ||
      data.Relate ||
      '',
      'Relate',
      500
    );

  const department =
    normalizeCalendarDepartment_(
      data.department ||
      data.Department ||
      ''
    );

  const status =
    normalizeCalendarStatus_(
      data.status ||
      data.Status ||
      ''
    );

  validateCalendarDate_(
    date
  );

  validateCalendarTime_(
    time
  );

  if (!endTime) {
    endTime =
      defaultCalendarEndTime_(
        time
      );
  }

  validateCalendarTime_(
    endTime
  );

  const startMinutes =
    calendarTimeToMinutes_(
      time
    );

  const endMinutes =
    calendarTimeToMinutes_(
      endTime
    );

  if (
    startMinutes === null ||
    endMinutes === null ||
    endMinutes <= startMinutes
  ) {
    throw new Error(
      'End Time must be after Start Time'
    );
  }

  if (!topic) {
    throw new Error(
      'Topic is required'
    );
  }

  if (!relate) {
    throw new Error(
      'Relate is required'
    );
  }

  return {
    date:
      date,
    time:
      time,
    endTime:
      endTime,
    location:
      location,
    topic:
      topic,
    description:
      description,
    relate:
      relate,
    department:
      department,
    status:
      status
  };
}


function isOutsourceLevel_(
  level
) {
  return String(
    level || ''
  )
    .trim()
    .toLowerCase()
    .startsWith(
      'outsource'
    );
}


function requireCalendarViewer_(
  session
) {
  if (
    !session ||
    !session.id
  ) {
    throw new Error(
      'Session required'
    );
  }

  if (
    isOutsourceLevel_(
      session.level
    )
  ) {
    throw new Error(
      'Calendar is not available for Outsource accounts'
    );
  }

  const level =
    String(
      session.level || ''
    )
      .trim()
      .toLowerCase();

  if (
    ![
      'master',
      'executive',
      'manager',
      'employee'
    ].includes(
      level
    )
  ) {
    throw new Error(
      'Calendar permission required'
    );
  }

  return session;
}


function validateCalendarDate_(
  value
) {
  const text =
    String(
      value || ''
    ).trim();

  const match =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );

  if (!match) {
    throw new Error(
      'Date must be YYYY-MM-DD'
    );
  }

  const year =
    Number(
      match[1]
    );

  const month =
    Number(
      match[2]
    );

  const day =
    Number(
      match[3]
    );

  const test =
    new Date(
      year,
      month - 1,
      day
    );

  if (
    test.getFullYear() !==
      year ||
    test.getMonth() !==
      month - 1 ||
    test.getDate() !==
      day
  ) {
    throw new Error(
      'Invalid Date'
    );
  }
}


function validateCalendarTime_(
  value
) {
  const text =
    String(
      value || ''
    ).trim();

  const match =
    text.match(
      /^(\d{2}):(\d{2})$/
    );

  if (!match) {
    throw new Error(
      'Time must be HH:mm'
    );
  }

  const hour =
    Number(
      match[1]
    );

  const minute =
    Number(
      match[2]
    );

  if (
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    throw new Error(
      'Invalid Time'
    );
  }
}


function sanitizeCalendarText_(
  value,
  fieldName,
  maxLength
) {
  const text =
    String(
      value || ''
    ).trim();

  if (
    text.length >
    maxLength
  ) {
    throw new Error(
      fieldName +
      ' is too long'
    );
  }

  return text;
}


function safeSheetText_(
  value
) {
  const text =
    String(
      value || ''
    );

  return /^[=+\-@]/.test(
    text
  )
    ? "'" + text
    : text;
}


/* =========================
   GET CONTACTS
========================= */

function getContacts() {
  const sheet =
    getSheet(CONTACT_SHEET);

  const rows =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (rows.length <= 1) {
    return jsonResponse({
      success: true,
      count: 0,
      data: []
    });
  }

  const headers =
    rows[0];

  const data =
    rows
      .slice(1)
      .filter(row =>
        String(row[0]).trim() !== ''
      )
      .map(row =>
        rowToObject(headers, row)
      );

  return jsonResponse({
    success: true,
    count: data.length,
    data: data
  });
}


/* =========================
   SERVICE / PROMOTION DATA
========================= */

function getServices() {
  const data =
    getServicesData_();

  return jsonResponse({
    success: true,
    count: data.length,
    data: data,
    services: data
  });
}


function getServicesData_() {
  return sheetValuesToObjects(
    getSheet(
      SERVICE_SHEET
    )
  );
}


function getPromotions() {
  const data =
    getPromotionsData_();

  return jsonResponse({
    success: true,
    count: data.length,
    data: data,
    promotions: data
  });
}


function getPromotionsData_() {
  return sheetValuesToObjects(
    getSheet(
      PROMOTION_SHEET
    )
  );
}


/* =========================
   GET ALL DATA
========================= */

function getAllData() {
  const contactSheet =
    getSheet(CONTACT_SHEET);

  const employees =
    getEmployeesData_();

  const doctors =
    getDoctorsData_();

  const branches =
    getBranchesData_();

  const services =
    getServicesData_();

  const promotions =
    getPromotionsData_();

  /* Calendar is read directly from Google Calendar in V31. */
  const calendar =
    [];

  const contacts =
    sheetValuesToObjects(
      contactSheet
    );

  return jsonResponse({
    success: true,

    employees:
      employees,

    doctors:
      doctors,

    branches:
      branches,

    services:
      services,

    promotions:
      promotions,

    calendar:
      calendar,

    contacts:
      contacts,

    counts: {
      employees:
        employees.length,

      doctors:
        doctors.length,

      branches:
        branches.length,

      services:
        services.length,

      promotions:
        promotions.length,

      calendar:
        calendar.length,

      contacts:
        contacts.length,

      personnel:
        employees.length
    },

    timestamp:
      new Date().toISOString()
  });
}


/* =========================
   UPDATE EMPLOYEE LEVEL
========================= */

function updateEmployeeLevel(
  data,
  session
) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  const newLevel =
    String(
      data.level || ''
    ).trim();

  const allowedLevels = ALLOWED_LEVELS;

  if (!employeeId) {
    throw new Error(
      'Employee ID is required'
    );
  }

  if (
    newLevel &&
    !allowedLevels.includes(
      newLevel
    )
  ) {
    throw new Error(
      'Invalid Level'
    );
  }

  const sheet =
    getSheet(EMPLOYEE_SHEET);

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  for (
    let i = 1;
    i < values.length;
    i++
  ) {
    if (
      String(
        values[i][0]
      ).trim()
      === employeeId
    ) {
      const targetEmployee =
        getEmployeeRecordFromRow_(
          sheet,
          i + 1,
          ensureEmployeeSecurityColumns_()
        );

      assertManagerCanManageEmployee_(
        session,
        targetEmployee,
        newLevel
      );

      sheet
        .getRange(
          i + 1,
          2
        )
        .setValue(
          newLevel
        );

      SpreadsheetApp.flush();

      if (
        hasWebsiteAccessLevel_(
          newLevel
        )
      ) {
        const refreshed =
          getEmployeeRecordById_(
            employeeId
          );

        if (
          refreshed &&
          (
            !refreshed.passwordHash ||
            !refreshed.passwordSalt
          )
        ) {
          setPasswordForEmployee_(
            employeeId,
            DEFAULT_PASSWORD,
            true
          );
        }
      }

      return jsonResponse({
        success: true,
        message:
          newLevel
            ? 'Permission updated'
            : 'Website access disabled',
        employeeId:
          employeeId,
        level:
          newLevel
      });
    }
  }

  return jsonResponse({
    success: false,
    message:
      'Employee not found'
  });
}


/* =========================
   UPDATE EMPLOYEE
========================= */

function updateEmployee(
  data,
  session
) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  if (!employeeId) {
    throw new Error(
      'Employee ID is required'
    );
  }

  const sheet =
    getSheet(EMPLOYEE_SHEET);

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(
        values[i][0]
      ).trim()
      === employeeId
    ) {
      const targetEmployee =
        getEmployeeRecordFromRow_(
          sheet,
          i + 1,
          ensureEmployeeSecurityColumns_()
        );

      assertManagerCanManageEmployee_(
        session,
        targetEmployee,
        data.level
      );

      if (
        data.level !== undefined
      ) {
        const allowedLevels = ALLOWED_LEVELS;

        if (
          String(data.level || '').trim() &&
          !allowedLevels.includes(
            String(data.level).trim()
          )
        ) {
          throw new Error(
            'Invalid Level'
          );
        }

        sheet
          .getRange(
            i + 1,
            2
          )
          .setValue(
            String(data.level || '').trim()
          );
      }

      if (
        data.position !== undefined
      ) {
        sheet
          .getRange(
            i + 1,
            3
          )
          .setValue(
            String(data.position)
          );
      }

      if (
        data.name !== undefined
      ) {
        sheet
          .getRange(
            i + 1,
            4
          )
          .setValue(
            String(data.name)
          );
      }

      if (
        data.tel !== undefined
      ) {
        sheet
          .getRange(
            i + 1,
            5
          )
          .setValue(
            String(data.tel)
          );
      }

      if (
        data.email !== undefined
      ) {
        sheet
          .getRange(
            i + 1,
            6
          )
          .setValue(
            String(data.email)
          );
      }

      SpreadsheetApp.flush();

      if (
        data.level !== undefined &&
        hasWebsiteAccessLevel_(
          String(data.level || '').trim()
        )
      ) {
        const refreshed =
          getEmployeeRecordById_(
            employeeId
          );

        if (
          refreshed &&
          (
            !refreshed.passwordHash ||
            !refreshed.passwordSalt
          )
        ) {
          setPasswordForEmployee_(
            employeeId,
            DEFAULT_PASSWORD,
            true
          );
        }
      }

      return jsonResponse({
        success: true,
        message:
          'Employee updated',
        employeeId:
          employeeId
      });
    }
  }

  return jsonResponse({
    success: false,
    message:
      'Employee not found'
  });
}


/* =========================
   DELETE EMPLOYEE
========================= */

function deleteEmployee(
  data,
  session
) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  if (!employeeId) {
    throw new Error(
      'Employee ID is required'
    );
  }

  if (
    employeeId ===
      MASTER_BOOTSTRAP_ID
  ) {
    throw new Error(
      'Primary Master account cannot be deleted'
    );
  }

  if (
    session &&
    String(
      session.id || ''
    ).trim() ===
      employeeId
  ) {
    throw new Error(
      'You cannot delete the account currently in use'
    );
  }

  const sheet =
    getSheet(
      EMPLOYEE_SHEET
    );

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    throw new Error(
      'Employee not found'
    );
  }

  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < ids.length;
    i++
  ) {
    if (
      String(
        ids[i][0] || ''
      ).trim() ===
        employeeId
    ) {
      const employee =
        getEmployeeRecordFromRow_(
          sheet,
          i + 2,
          ensureEmployeeSecurityColumns_()
        );

      assertManagerCanManageEmployee_(
        session,
        employee
      );

      sheet.deleteRow(
        i + 2
      );

      SpreadsheetApp.flush();

      return jsonResponse({
        success: true,
        message:
          'Employee deleted',
        employeeId:
          employeeId,
        name:
          employee.name || ''
      });
    }
  }

  throw new Error(
    'Employee not found'
  );
}


/* =========================
   ADD EMPLOYEE
========================= */

function addEmployee(
  data,
  session
) {
  const employeeId =
    String(
      data.employeeId || ''
    ).trim();

  if (!employeeId) {
    throw new Error(
      'Employee ID is required'
    );
  }

  const sheet =
    getSheet(EMPLOYEE_SHEET);

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  const exists =
    values
      .slice(1)
      .some(row =>
        String(
          row[0]
        ).trim()
        === employeeId
      );

  if (exists) {
    return jsonResponse({
      success: false,
      message:
        'Employee ID already exists'
    });
  }

  const allowedLevels =
    ALLOWED_LEVELS;

  const requestedLevel =
    String(
      data.level || ''
    ).trim();

  if (
    requestedLevel &&
    !allowedLevels.includes(
      requestedLevel
    )
  ) {
    throw new Error(
      'Invalid Level'
    );
  }

  const level =
    requestedLevel;

  assertManagerCanManageEmployee_(
    session,
    null,
    level
  );

  sheet.appendRow([
    employeeId,
    level,
    data.position || '',
    data.name || '',
    data.tel || '',
    data.email || ''
  ]);

  SpreadsheetApp.flush();

  if (
    hasWebsiteAccessLevel_(
      level
    )
  ) {
    setPasswordForEmployee_(
      employeeId,
      DEFAULT_PASSWORD,
      true
    );
  }

  return jsonResponse({
    success: true,
    message:
      level
        ? 'Employee added with website access'
        : 'Employee added as directory record only',
    employeeId:
      employeeId,
    accessEnabled:
      !!level
  });
}


/* =========================
   ADD CONTACT
========================= */

function addContact(
  data,
  session
) {
  const subject =
    String(
      data.subject || ''
    ).trim();

  const detail =
    String(
      data.detail || ''
    ).trim();

  if (!subject) {
    throw new Error(
      'Subject is required'
    );
  }

  if (!detail) {
    throw new Error(
      'Detail is required'
    );
  }

  const employee =
    getEmployeeRecordById_(
      session.id
    );

  if (!employee) {
    throw new Error(
      'Employee not found'
    );
  }

  const contactSheet =
    getSheet(CONTACT_SHEET);

  contactSheet.appendRow([
    employee.id,
    employee.name,
    employee.position,
    subject,
    detail
  ]);

  SpreadsheetApp.flush();

  return jsonResponse({
    success: true,
    message:
      'Contact submitted',
    employeeId:
      employee.id,
    name:
      employee.name,
    position:
      employee.position
  });
}


/* =========================
   HELPERS
========================= */

function getSheet(sheetName) {
  const ss =
    SpreadsheetApp
      .openById(
        SPREADSHEET_ID
      );

  const sheet =
    ss.getSheetByName(
      sheetName
    );

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' +
      sheetName
    );
  }

  return sheet;
}


function isCellImage_(value) {
  try {
    return !!(
      value &&
      value.valueType ==
        SpreadsheetApp.ValueType.IMAGE &&
      typeof value.getContentUrl ===
        'function'
    );
  }
  catch (err) {
    return false;
  }
}


function isHttpUrl_(value) {
  return (
    typeof value === 'string' &&
    /^https?:\/\//i.test(
      value.trim()
    )
  );
}


function sheetValuesToObjects(
  sheet
) {
  const rows =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (
    rows.length <= 1
  ) {
    return [];
  }

  const headers =
    rows[0];

  return rows
    .slice(1)
    .filter(row =>
      String(
        row[0]
      ).trim() !== ''
    )
    .map(row =>
      rowToObject(
        headers,
        row
      )
    );
}


function rowToObject(
  headers,
  row
) {
  const obj = {};

  headers.forEach(
    (header, i) => {
      const key =
        String(
          header || ''
        ).trim();

      if (key) {
        obj[key] =
          row[i] || '';
      }
    }
  );

  return obj;
}


function jsonResponse(data) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService
        .MimeType
        .JSON
    );
}


function errorResponse(err) {
  return jsonResponse({
    success: false,
    message:
      err &&
      err.message
        ? err.message
        : String(err)
  });
}
