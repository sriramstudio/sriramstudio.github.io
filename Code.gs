// ============================================================
// SRIRAM STUDIO — Google Apps Script Backend
// Paste this entire file into your Apps Script editor.
// Deploy as Web App: Execute as Me, Access: Anyone
//
// AUTH: every action except public registration and the login check
// requires the admin PIN as a `pin` parameter, verified server-side.
// Keep this in step with sriramstudio_admin.html — deploy the HTML first.
// ============================================================

// ─── Sheet bootstrap ─────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initHeaders(sheet, name);
  }
  return sheet;
}

function initHeaders(sheet, name) {
  const headerStyle = {
    bg: '#2C1A0E',
    fg: '#FFFFFF',
    bold: true
  };
  if (name === 'Enrollments') {
    const h = ['ID','Enrolled At','Type','Student Name','Date of Birth','Gender',
      'Blood Group','School/College','Guardian Name','Relation','Phone','WhatsApp',
      'Email','Address','Program','Location','Batch','Joining Date',
      'Pracheen Kala Kendra','Workshop Name','Workshop Date','Workshop Fee',
      'Heard From','Notes','Status','Left On','Review'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setFontWeight('bold')
      .setBackground('#2C1A0E').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 90);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(4, 180);
  } else if (name === 'Receipts') {
    const h = ['Receipt No','Issued At','Student Name','Contact',
      'Amount (₹)','Fee Month','Fee Year','Payment Mode','UPI Reference',
      'Fee Type','Date Received','Note','Students'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setFontWeight('bold')
      .setBackground('#2C1A0E').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(3, 180);
  } else if (name === 'Config') {
    sheet.appendRow(['Key', 'Value']);
    sheet.appendRow(['pin', '1234']);
    sheet.appendRow(['receipt_seq', '1']);
    sheet.appendRow(['notify_email', '']);
    sheet.appendRow(['notify_whatsapp', '']);
    sheet.appendRow(['callmebot_key', '']);
    sheet.getRange(1,1,1,2).setFontWeight('bold');
  }
}

// ─── Status column ────────────────────────────────────────────
// 'Status' was added after the sheet was already in use, so initHeaders
// never runs for it. Anything blank counts as active.
const STATUS_LEFT   = 'Left';
const STATUS_ACTIVE = 'Active';
// Words that have been used in the sheet to mean the student has left.
const LEFT_WORDS = ['discontinue', 'discontinued', 'left', 'inactive',
                    'stopped', 'quit', 'dropped', 'not continuing'];

function isLeftWord_(v) {
  const t = (v || '').toString().trim().toLowerCase();
  if (!t) return false;
  for (let i = 0; i < LEFT_WORDS.length; i++) {
    if (t.indexOf(LEFT_WORDS[i]) >= 0) return true;
  }
  return false;
}

// Returns the zero-based index of a column, appending it if the sheet
// predates it. Columns are only ever added at the end, so existing data
// never shifts.
function ensureColumn_(sheetName, header) {
  const sheet = getSheet(sheetName);
  const width = Math.max(1, sheet.getLastColumn());
  const head  = sheet.getRange(1, 1, 1, width).getValues()[0]
                     .map(function (v) { return (v === null ? '' : v.toString().trim()); });
  const at = head.indexOf(header);
  if (at >= 0) return at;
  sheet.getRange(1, width + 1).setValue(header)
       .setFontWeight('bold').setBackground('#2C1A0E').setFontColor('#FFFFFF');
  return width;
}

function ensureStatusColumn_() { return ensureColumn_('Enrollments', 'Status'); }

// A student who has left keeps their row and their receipt history, but is
// hidden from the receipt autocomplete: a rejoiner gets a fresh row, and the
// old one must not be pickable.
function ensureEnrollmentColumns_() {
  ensureColumn_('Enrollments', 'Status');
  ensureColumn_('Enrollments', 'Left On');
  ensureColumn_('Enrollments', 'Review');
}

// Looks for an existing enrolment under the same name.
// Prefers a phone match; otherwise falls back to the most recent by name.
function findExistingStudent_(name, phone) {
  const n = (name || '').toString().trim().toLowerCase();
  if (!n) return null;
  const digits = (phone || '').toString().replace(/\D/g, '');
  const sheet  = getSheet('Enrollments');
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  const head   = data[0].map(function (v) { return (v === null ? '' : v.toString().trim()); });
  const iName  = head.indexOf('Student Name');
  const iPhone = head.indexOf('Phone');
  const iId    = head.indexOf('ID');
  const iStat  = head.indexOf('Status');

  let byName = null;
  for (let i = data.length - 1; i >= 1; i--) {
    const rn = (data[i][iName] === undefined || data[i][iName] === null)
                 ? '' : data[i][iName].toString().trim().toLowerCase();
    if (rn !== n) continue;
    const hit = {
      id:     iId   >= 0 ? (data[i][iId]   || '').toString() : '',
      status: iStat >= 0 ? (data[i][iStat] || '').toString() : ''
    };
    const rd = iPhone >= 0 ? (data[i][iPhone] || '').toString().replace(/\D/g, '') : '';
    if (digits && rd && rd === digits) return hit;   // strongest match
    if (!byName) byName = hit;                       // most recent same name
  }
  return byName;
}

// Run by hand to create the tracking columns now rather than waiting for the
// next enrolment to add them.
function addStatusColumn() {
  ensureEnrollmentColumns_();
  const sheet = getSheet('Enrollments');
  const head  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function (v) { return (v === null ? '' : v.toString().trim()); });
  const msg = ['Status', 'Left On', 'Review'].map(function (h) {
    return h + ' = column ' + (head.indexOf(h) + 1);
  }).join(', ');
  Logger.log(msg);
  return msg;
}

// ─── Main handler ─────────────────────────────────────────────
// ─── Authorisation ────────────────────────────────────────────
// The web app URL is public (it is embedded in register.html), so the
// PIN is the only thing separating a visitor from the student database.
// It is checked here, on the server. The client cannot be trusted.

const PUBLIC_ACTIONS      = ['verifyPin'];              // reachable without a PIN
const PUBLIC_ENROLL_MODES = ['app-admission', 'app-workshop'];  // the public form
const LOCK_MAX            = 10;   // failed attempts before lockout
const LOCK_MINS           = 15;   // lockout duration

function readConfig(key) {
  const values = getSheet('Config').getDataRange().getValues();
  for (const row of values) if (row[0] === key) return row[1].toString();
  return '';
}

function pinOk(pin) {
  if (pin === undefined || pin === null || pin === '') return false;
  const stored = readConfig('pin') || '1234';
  return stored === pin.toString();
}

// A 4-digit PIN is only 10,000 guesses, so throttle failures.
// Script-wide rather than per-IP (Apps Script exposes no client IP),
// which is blunt but makes brute force impractical.
function failCount() {
  return parseInt(CacheService.getScriptCache().get('pin_fails') || '0', 10);
}
function noteFailure() {
  const n = failCount() + 1;
  CacheService.getScriptCache().put('pin_fails', String(n), LOCK_MINS * 60);
  return n;
}
function clearFailures() {
  CacheService.getScriptCache().remove('pin_fails');
}

// Decide whether this request needs a PIN. Fails closed on bad input.
function requiresPin(e) {
  const action = e.parameter.action;
  if (PUBLIC_ACTIONS.indexOf(action) !== -1) return false;
  if (action === 'addEnrollment') {
    try {
      const probe = JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
      return PUBLIC_ENROLL_MODES.indexOf(probe.mode) === -1;
    } catch (err) {
      return true;
    }
  }
  return true;
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;

    // ── Authorisation gate ──
    const needsPin = requiresPin(e);
    if (needsPin || action === 'verifyPin') {
      if (failCount() >= LOCK_MAX) {
        return output({
          error:  'Too many failed PIN attempts. Try again in ' + LOCK_MINS + ' minutes.',
          locked: true
        });
      }
    }
    if (needsPin && !pinOk(e.parameter.pin)) {
      noteFailure();
      return output({ error: 'Unauthorised.', auth: true });
    }

    switch (action) {

      case 'init':
        getSheet('Enrollments');
        getSheet('Receipts');
        getSheet('Config');
        result = { success: true, message: 'Sriram Studio sheets ready.' };
        break;

      case 'addEnrollment': {
        const data = JSON.parse(decodeURIComponent(e.parameter.data));
        result = addEnrollment(data);
        break;
      }

      case 'addReceipt': {
        const data = JSON.parse(decodeURIComponent(e.parameter.data));
        result = addReceipt(data);
        break;
      }

      case 'getEnrollments':
        result = getEnrollments();
        break;

      case 'listStudents':
        result = listStudents();
        break;

      case 'searchStudents':
        result = searchStudents(e.parameter.q || '');
        break;

      case 'getReceipts':
        result = getReceipts();
        break;

      case 'verifyPin': {
        const vp = verifyPin(e.parameter.pin);
        if (vp.valid) clearFailures(); else noteFailure();
        result = vp;
        break;
      }

      case 'setPin':
        result = setPin(e.parameter.current, e.parameter.newpin);
        break;

      case 'deleteEnrollment':
        result = deleteEnrollment(e.parameter.id);
        break;

      case 'updateConfig': {
        const cfgKey = e.parameter.key;
        const cfgVal = e.parameter.value || '';
        result = (cfgKey === 'pin')
          ? { success: false, error: 'Use setPin to change the PIN.' }
          : updateConfig(cfgKey, cfgVal);
        break;
      }

      default:
        result = { error: 'Unknown action: ' + action };
    }

    return output(result);

  } catch (err) {
    return output({ error: err.message, stack: err.stack });
  }
}

function output(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Enrollment ───────────────────────────────────────────────
function addEnrollment(d) {
  const sheet = getSheet('Enrollments');

  // ── Enrollment ID: SR-YYYY-MMDDHHmmssSSS (never duplicates) ──
  const t    = new Date();
  const pad2 = n => String(n).padStart(2, '0');
  const pad3 = n => String(n).padStart(3, '0');
  const id   =
    'SR-' + t.getFullYear() + '-' +
    pad2(t.getMonth() + 1) + pad2(t.getDate()) +
    pad2(t.getHours())     + pad2(t.getMinutes()) +
    pad2(t.getSeconds())   + pad3(t.getMilliseconds());

  const now = Utilities.formatDate(t, Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');

  ensureEnrollmentColumns_();

  // A returning student gets a fresh row by design; the old one stays marked
  // Left. Flag it either way so an accidental duplicate cannot slip past.
  const prior = findExistingStudent_(d.studentName, d.phone);
  const review = !prior ? ''
    : (isLeftWord_(prior.status)
        ? 'Rejoining? Earlier record ' + prior.id + ' is marked Left'
        : 'Possible duplicate of active record ' + prior.id);

  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                    .map(function (v) { return (v === null ? '' : v.toString().trim()); });
  const put  = {
    'ID': id,
    'Enrolled At': now,
    'Type':
      d.mode === 'admission'     ? 'New Admission' :
      d.mode === 'workshop'      ? 'Workshop' :
      d.mode === 'legacy'        ? 'Existing Student' :
      d.mode === 'app-admission' ? 'Application \u2013 Admission' :
      d.mode === 'app-workshop'  ? 'Application \u2013 Workshop' : 'Existing Student',
    'Student Name': d.studentName || '',
    'Date of Birth': d.dob || '',
    'Gender': d.gender || '',
    'Blood Group': d.bloodGroup || '',
    'School/College': d.school || '',
    'Guardian Name': d.guardianName || '',
    'Relation': d.relation || '',
    'Phone': d.phone || '',
    'WhatsApp': d.whatsapp || '',
    'Email': d.email || '',
    'Address': d.address || '',
    'Program': d.program || '',
    'Location': d.location || '',
    'Batch': d.batch || '',
    // Starting month and year are chosen from dropdowns on every mode now.
    // The older single-date fields are still honoured for anything that has
    // not been updated to send them.
    // The admin panel supplies month and year on every mode. The public form
    // deliberately does not ask, so an application falls back to the month it
    // was submitted in — better than blank, and Anjali can correct it.
    // Existing Students stay blank unless a month is given: their real joining
    // date predates this system and guessing it would be worse than nothing.
    'Joining Date': (d.startMonth && d.startYear) ? (d.startMonth + ' ' + d.startYear)
                    : d.mode === 'legacy' ? (d.approxJoining || '')
                    : (d.startDate || d.workshopDate ||
                       Utilities.formatDate(t, Session.getScriptTimeZone(), 'MMMM yyyy')),
    'Pracheen Kala Kendra': d.pracheen || '',
    'Workshop Name': d.workshopName || '',
    'Workshop Date': d.workshopDate || '',
    'Workshop Fee': d.workshopFee ? '\u20b9' + d.workshopFee : '',
    'Heard From': d.hearFrom || '',
    'Notes': d.notes || '',
    'Status': STATUS_ACTIVE,
    'Left On': '',
    'Review': review
  };
  // Headers get renamed in the sheet — 'Joining Date' is really
  // 'Joining Date/Approx Joining Month'. Match on prefix so a rename does not
  // silently drop the value.
  const row = head.map(function (h) {
    if (put.hasOwnProperty(h)) return put[h];
    const alias = Object.keys(put).filter(function (k) { return h.indexOf(k) === 0; })[0];
    return alias ? put[alias] : '';
  });
  sheet.appendRow(row);

  // Send notifications for self-registered applications only
  if (d.mode === 'app-admission' || d.mode === 'app-workshop') {
    try { sendNotifications(d, id, now); } catch(e) { Logger.log('Notification error: ' + e.message); }
  }

  return { success: true, id, enrolledAt: now, review: review };
}

function getEnrollments() {
  const sheet = getSheet('Enrollments');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { records: [] };
  const headers = data[0];
  const records = data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i].toString() : ''; });
      return obj;
    })
    .reverse();
  return { records };
}

// The whole active roster in one slim call. The autocomplete filters this
// locally: a request per keystroke against Apps Script is slow, and replies
// arriving out of order made the suggestions flicker between queries.
function listStudents() {
  const sheet = getSheet('Enrollments');
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { students: [] };
  const headers = data[0];

  const col = function (name) {
    const idx = headers.indexOf(name);
    return function (row) {
      return (idx >= 0 && row[idx] !== undefined && row[idx] !== null) ? row[idx].toString().trim() : '';
    };
  };
  const nameAt = headers.indexOf('Student Name');
  if (nameAt < 0) return { students: [] };

  const seen = {}, students = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const row  = data[i];
    const name = (row[nameAt] === undefined || row[nameAt] === null) ? '' : row[nameAt].toString().trim();
    if (!name) continue;
    if (isLeftWord_(col('Status')(row))) continue;      // left students are not billable

    const phone = col('Phone')(row) || col('WhatsApp')(row);
    const key   = name.toLowerCase() + '|' + phone.replace(/\D/g, '');
    if (seen[key]) continue;
    seen[key] = true;
    students.push({
      studentName: name,
      phone:    phone,
      program:  col('Program')(row),
      location: col('Location')(row),
      status:   col('Status')(row)
    });
  }
  return { students: students };
}

function searchStudents(q) {
  if (!q || q.length < 2) return { results: [] };
  const sheet = getSheet('Enrollments');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { results: [] };
  const headers = data[0];

  // Use safe column lookup — returns '' if column was deleted
  const col = (name) => {
    const idx = headers.indexOf(name);
    return (row) => idx >= 0 && row[idx] !== undefined ? row[idx].toString() : '';
  };

  const lower = q.toLowerCase();
  const seen = new Set();
  const results = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const nameIdx = headers.indexOf('Student Name');
    if (nameIdx < 0) break;
    const name = row[nameIdx] !== undefined ? row[nameIdx].toString() : '';
    if (!name) continue;
    // Someone who has left keeps their record and receipt history, but must
    // not be pickable for a new receipt — a rejoiner has a fresh row.
    if (isLeftWord_(col('Status')(row))) continue;
    const phone = col('Phone')(row) || col('WhatsApp')(row);
    // Key on name + phone, not name alone: two different students can share
    // a name, and collapsing them would silently hide one of them.
    const key = name.toLowerCase() + '|' + phone.replace(/\D/g, '');
    if (name.toLowerCase().includes(lower) && !seen.has(key)) {
      seen.add(key);
      results.push({
        studentName: name,
        phone:    phone,
        program:  col('Program')(row),
        location: col('Location')(row),
        status:   col('Status')(row)
      });
      if (results.length >= 6) break;
    }
  }
  return { results };
}

function deleteEnrollment(id) {
  const sheet = getSheet('Enrollments');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Record not found' };
}

// ─── Receipts ─────────────────────────────────────────────────
// One receipt can cover several students — siblings are usually paid for
// together with a single clubbed amount. 'Student Name' keeps the combined
// display string so older receipts and the Records tab are unaffected;
// 'Students' holds the names separately so the data stays queryable.
function ensureReceiptStudentsColumn_() {
  const sheet = getSheet('Receipts');
  const width = Math.max(1, sheet.getLastColumn());
  const head  = sheet.getRange(1, 1, 1, width).getValues()[0]
                     .map(function (v) { return (v === null ? '' : v.toString().trim()); });
  const at = head.indexOf('Students');
  if (at >= 0) return at;
  sheet.getRange(1, width + 1).setValue('Students')
       .setFontWeight('bold').setBackground('#2C1A0E').setFontColor('#FFFFFF');
  return width;
}

function addReceipt(d) {
  ensureReceiptStudentsColumn_();
  const sheet   = getSheet('Receipts');
  const config  = getSheet('Config');
  const cfgData = config.getDataRange().getValues();

  let seq = 1, seqRow = -1;
  for (let i = 0; i < cfgData.length; i++) {
    if (cfgData[i][0] === 'receipt_seq') {
      seq    = parseInt(cfgData[i][1]) || 1;
      seqRow = i + 1;
      break;
    }
  }
  if (seqRow > 0) config.getRange(seqRow, 2).setValue(seq + 1);

  const receiptNo = 'SS-' + new Date().getFullYear() + '-' + String(seq).padStart(4, '0');
  const issuedAt  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');

  sheet.appendRow([
    receiptNo, issuedAt,
    d.studentName   || '',
    d.guardianPhone || '',
    d.amount        || '',
    d.month         || '',
    d.year          || '',
    d.payMode       || '',
    d.upiRef        || '',
    d.feeType       || 'Monthly Fee',
    d.dateReceived  || '',
    d.note          || '',
    (d.students && d.students.length) ? d.students.join(' | ') : (d.studentName || '')
  ]);

  return { success: true, receiptNo, issuedAt };
}

function getReceipts() {
  const sheet = getSheet('Receipts');
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { records: [] };
  const headers = data[0];
  const records = data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i].toString() : ''; });
      return obj;
    })
    .reverse();
  return { records };
}

// ─── Notifications ────────────────────────────────────
function sendNotifications(d, id, enrolledAt) {
  const config  = getSheet('Config');
  const values  = config.getDataRange().getValues();
  let notifyEmail = '', notifyWa = '', callbotKey = '';
  for (const row of values) {
    if (row[0] === 'notify_email')    notifyEmail = row[1].toString().trim();
    if (row[0] === 'notify_whatsapp') notifyWa    = row[1].toString().trim();
    if (row[0] === 'callmebot_key')   callbotKey  = row[1].toString().trim();
  }

  const type     = d.mode === 'app-admission' ? 'New Admission Application' : 'Workshop Application';
  const name     = d.studentName  || '—';
  const guardian = d.guardianName || '—';
  const phone    = d.phone        || '—';
  const centre   = (d.location    || '—').split('–')[0].trim();
  const workshop = d.workshopName ? '\nWorkshop: ' + d.workshopName : '';

  // ── Email notification ──────────────────────────────
  if (notifyEmail) {
    try {
      const subject = '🪷 New Registration — ' + name + ' | Sriram Studio';
      const body =
        'A new application has been submitted via the Sriram Studio registration form.\n\n' +
        'Reference ID : ' + id       + '\n' +
        'Type         : ' + type     + '\n' +
        'Student      : ' + name     + '\n' +
        'Guardian     : ' + guardian + '\n' +
        'Contact      : ' + phone    + '\n' +
        'Centre       : ' + centre   +
        workshop +
        '\nSubmitted    : ' + enrolledAt + '\n\n' +
        'Open your Sriram Studio admin panel to view and manage this application.\n\n' +
        '—\nSriram Studio Admin System\nsriramstudio.github.io';
      MailApp.sendEmail(notifyEmail, subject, body);
    } catch(e) { Logger.log('Email error: ' + e.message); }
  }

  // ── WhatsApp via CallMeBot ─────────────────────────
  if (notifyWa && callbotKey) {
    try {
      const msg =
        '🪷 New Registration%0A' +
        'Student: ' + encodeURIComponent(name) + '%0A' +
        'Type: ' + encodeURIComponent(type) + '%0A' +
        'Contact: ' + encodeURIComponent(phone) +
        (d.workshopName ? '%0AWorkshop: ' + encodeURIComponent(d.workshopName) : '') + '%0A' +
        'Centre: ' + encodeURIComponent(centre) + '%0A' +
        'Ref: ' + id;
      const url = 'https://api.callmebot.com/whatsapp.php?phone=' + notifyWa +
                  '&text=' + msg + '&apikey=' + callbotKey;
      UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch(e) { Logger.log('WhatsApp error: ' + e.message); }
  }
}

// ─── Update Config ─────────────────────────────────────
function updateConfig(key, value) {
  const sheet  = getSheet('Config');
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return { success: true };
    }
  }
  sheet.appendRow([key, value]);
  return { success: true };
}

// ─── PIN ──────────────────────────────────────────────────────
function verifyPin(pin) {
  const sheet  = getSheet('Config');
  const values = sheet.getDataRange().getValues();
  for (const row of values) {
    if (row[0] === 'pin') return { valid: row[1].toString() === pin.toString() };
  }
  return { valid: pin === '1234' }; // fallback
}

function setPin(current, newPin) {
  const check = verifyPin(current);
  if (!check.valid) return { success: false, error: 'Current PIN is incorrect.' };
  if (!newPin || newPin.length !== 4 || !/^\d+$/.test(newPin))
    return { success: false, error: 'New PIN must be exactly 4 digits.' };
  const sheet  = getSheet('Config');
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === 'pin') {
      sheet.getRange(i + 1, 2).setValue(newPin);
      return { success: true };
    }
  }
  return { success: false, error: 'Config sheet error.' };
}


// ─── Centres ──────────────────────────────────────────────────
// The legacy roster was imported in Aug 2026 and its source tab removed;
// Enrollments is the single source of truth for students now. These remain
// because the column audit and the alignment repair both need to recognise a
// real branch name.

// The Center column is also used for status words ("discontinue"), so a
// value only becomes a Location if it actually names one of the branches.
const KNOWN_CENTRES = ['Bhawanipur', 'Wood Street', 'Kankurgachi', 'Salt Lake'];

function matchCentre_(v) {
  const t = (v || '').toString().trim().toLowerCase();
  if (!t) return null;
  for (let i = 0; i < KNOWN_CENTRES.length; i++) {
    const c = KNOWN_CENTRES[i].toLowerCase();
    if (t.indexOf(c) >= 0 || c.indexOf(t) >= 0) return KNOWN_CENTRES[i];
  }
  return null;
}

// Appends new students and fills blanks on ones already there.
// Safe to re-run as the roster gains phone numbers and centres.
// ─── Column audit (read-only, editor only) ────────────────────
// Reports where values actually sit versus what the headers claim, so a
// misalignment can be diagnosed before anything is moved. Writes nothing.

function colLetter_(n) {
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function auditEnrollments() {
  const sheet = getSheet('Enrollments');
  const data  = sheet.getDataRange().getValues();
  const norm  = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  if (data.length <= 1) { Logger.log('No data rows.'); return 'No data rows.'; }

  const head = data[0].map(norm);
  let out = 'ENROLLMENTS COLUMN AUDIT\n========================\n';
  out += 'Data rows : ' + (data.length - 1) + '\n';
  out += 'Columns   : ' + head.length + '\n\n';

  out += 'PER COLUMN  [count filled]  sample values\n';
  out += '-----------------------------------------\n';
  for (let c = 0; c < head.length; c++) {
    let filled = 0;
    const samples = [];
    for (let r = 1; r < data.length; r++) {
      const v = norm(data[r][c]);
      if (!v) continue;
      filled++;
      if (samples.length < 2 && samples.indexOf(v) < 0) {
        samples.push(v.length > 28 ? v.substring(0, 28) + '~' : v);
      }
    }
    out += colLetter_(c + 1) + ' ' + (head[c] || '(no header)') + '  [' + filled + ']';
    if (samples.length) out += '  e.g. ' + samples.join('  /  ');
    out += '\n';
  }

  // Values that clearly sit in the wrong place.
  out += '\nMISPLACED VALUES\n----------------\n';
  const iLoc = head.indexOf('Location');
  let flagged = 0;
  for (let c = 0; c < head.length; c++) {
    if (c === iLoc) continue;
    let hits = 0;
    let example = '';
    for (let r = 1; r < data.length; r++) {
      const v = norm(data[r][c]);
      if (v && matchCentre_(v)) { hits++; if (!example) example = v; }
    }
    if (hits) {
      flagged++;
      out += 'Centre name in ' + colLetter_(c + 1) + ' "' + (head[c] || '?') +
             '" on ' + hits + ' row(s), e.g. "' + example + '"\n';
    }
  }
  if (!flagged) out += 'No centre names found outside Location.\n';

  Logger.log(out);
  return out;
}

// Dumps single rows header-by-header so a shift is visible directly.
// Pass nothing for a spread of rows, or a number for one specific row.
function auditEnrollmentRow(rowNumber) {
  const sheet = getSheet('Enrollments');
  const data  = sheet.getDataRange().getValues();
  const norm  = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const head  = data[0].map(norm);

  let rows;
  if (rowNumber) rows = [rowNumber];
  else {
    rows = [2, 3, 4];
    const mid = Math.floor(data.length / 2);
    if (mid > 4) rows.push(mid);
    if (data.length - 1 > 5) rows.push(data.length - 1, data.length);
  }

  let out = 'ROW DUMP (non-empty cells only)\n===============================\n';
  rows.forEach(function (rn) {
    if (rn < 2 || rn > data.length) return;
    out += '\n--- sheet row ' + rn + ' ---\n';
    const row = data[rn - 1];
    for (let c = 0; c < head.length; c++) {
      const v = norm(row[c]);
      if (!v) continue;
      out += '  ' + colLetter_(c + 1) + ' ' + (head[c] || '(no header)') + ' = ' +
             (v.length > 40 ? v.substring(0, 40) + '~' : v) + '\n';
    }
  });
  Logger.log(out);
  return out;
}


// ─── Column alignment repair (editor only) ────────────────────
// The Program, Batch and Pracheen Kala Kendra columns were deleted from the
// sheet, but the old addEnrollment wrote its 24 values BY POSITION. Rows
// written after that deletion overflowed: values 15-24 landed to the right of
// their headers and spilled into three unheadered columns past Notes.
//
// For an affected row the current columns hold:
//   O=Program(blank)  P=Location  Q=Batch(blank)  R=Joining Date
//   S=Pracheen(blank) T=Workshop Name  U=Workshop Date  V=Workshop Fee
//   W=Heard From      X=Notes
//
// previewColumnRepair() writes nothing. repairColumnAlignment() snapshots the
// whole tab to a backup sheet before touching a cell.

// Zero-based: where each of columns O..U should take its value from.
const REPAIR_PULL = [15, 17, 19, 20, 21, 22, 23];   // -> O,P,Q,R,S,T,U
const REPAIR_FIRST = 14;                            // column O
const REPAIR_CLEAR = [21, 22, 23];                  // V,W,X are not real columns

function repairHeadersOk_(head) {
  const want = ['Location', 'Joining Date', 'Workshop Name', 'Workshop Date',
                'Workshop Fee', 'Heard From', 'Notes'];
  for (let i = 0; i < want.length; i++) {
    const h = (head[REPAIR_FIRST + i] || '').toString();
    if (h.indexOf(want[i]) !== 0) {
      return 'Column ' + colLetter_(REPAIR_FIRST + i + 1) + ' should start with "' +
             want[i] + '" but reads "' + h + '"';
    }
  }
  return null;
}

function isShiftedRow_(row) {
  const norm = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  // Nothing should ever sit past Notes in those unheadered columns.
  for (let i = 0; i < REPAIR_CLEAR.length; i++) {
    if (norm(row[REPAIR_CLEAR[i]])) return true;
  }
  // Or a centre name sitting in the Joining Date column.
  if (matchCentre_(norm(row[15]))) return true;
  return false;
}

function buildColumnRepair_() {
  const sheet = getSheet('Enrollments');
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { error: 'No data rows.' };

  const norm = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const head = data[0].map(norm);

  const bad = repairHeadersOk_(head);
  if (bad) return { error: 'Refusing to touch anything - headers are not what the repair expects.\n' + bad };

  const fixes = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!isShiftedRow_(row)) continue;
    const before = [], after = [];
    for (let i = 0; i < 10; i++) {           // O..X
      before.push(norm(row[REPAIR_FIRST + i]));
    }
    for (let i = 0; i < REPAIR_PULL.length; i++) after.push(norm(row[REPAIR_PULL[i]]));
    while (after.length < 10) after.push('');
    fixes.push({ sheetRow: r + 1, name: norm(row[3]), id: norm(row[0]), before: before, after: after });
  }
  return { head: head, fixes: fixes, total: data.length - 1 };
}

function previewColumnRepair() {
  const r = buildColumnRepair_();
  if (r.error) { Logger.log(r.error); return r.error; }

  const labels = [];
  for (let i = 0; i < 10; i++) labels.push(colLetter_(REPAIR_FIRST + i + 1));

  let out = 'COLUMN REPAIR - PREVIEW\n=======================\n';
  out += 'Data rows        : ' + r.total + '\n';
  out += 'Rows to realign  : ' + r.fixes.length + '\n';
  out += 'Rows left as-is  : ' + (r.total - r.fixes.length) + '\n\n';

  if (!r.fixes.length) { out += 'Nothing to do.\n'; Logger.log(out); return out; }

  // Prove nothing is lost: every non-empty value before must appear after.
  let lost = 0;
  r.fixes.forEach(function (f) {
    f.before.forEach(function (v) {
      if (v && f.after.indexOf(v) < 0) lost++;
    });
  });
  out += 'Non-empty values that would be LOST: ' + lost + '\n';
  out += (lost === 0 ? 'Every value is preserved, only moved.\n\n'
                     : '*** STOP - values would be dropped. Do not run the repair. ***\n\n');

  out += 'Columns shown: ' + labels.join(' ') + '\n\n';
  r.fixes.slice(0, 6).forEach(function (f) {
    out += 'row ' + f.sheetRow + '  ' + f.name + '\n';
    for (let i = 0; i < 10; i++) {
      const b = f.before[i], a = f.after[i];
      if (!b && !a) continue;
      const mark = (b === a) ? '   ' : ' ->';
      out += '   ' + labels[i] + ' "' + (b.length > 30 ? b.substring(0, 30) + '~' : b) + '"' +
             mark + ' "' + (a.length > 30 ? a.substring(0, 30) + '~' : a) + '"\n';
    }
    out += '\n';
  });
  if (r.fixes.length > 6) out += '... and ' + (r.fixes.length - 6) + ' more rows, same transformation.\n';

  out += '\nNothing was changed. Report only.\n';
  Logger.log(out);
  return out;
}

function repairColumnAlignment() {
  const r = buildColumnRepair_();
  if (r.error) { Logger.log(r.error); return r.error; }
  if (!r.fixes.length) { Logger.log('Nothing to repair.'); return 'Nothing to repair.'; }

  // Refuse if the move would drop anything.
  let lost = 0;
  r.fixes.forEach(function (f) {
    f.before.forEach(function (v) { if (v && f.after.indexOf(v) < 0) lost++; });
  });
  if (lost) {
    const msg = 'Refused: ' + lost + ' value(s) would be lost. Run previewColumnRepair().';
    Logger.log(msg);
    return msg;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet('Enrollments');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  sheet.copyTo(ss).setName('Enrollments backup ' + stamp);

  r.fixes.forEach(function (f) {
    sheet.getRange(f.sheetRow, REPAIR_FIRST + 1, 1, 10).setValues([f.after]);
  });

  const msg = 'Realigned ' + r.fixes.length + ' row(s). Backup saved as "Enrollments backup ' + stamp + '".';
  Logger.log(msg);
  return msg;
}


// ─── Drop the leftover unheadered columns (editor only) ───────
// After repairColumnAlignment, V/W/X hold nothing and have no headers. They
// only existed because the old positional write overflowed past Notes.
//
// Refuses if they are not genuinely empty — running this before the repair
// would destroy the Heard From and Notes values still sitting in W and X.

function findOverflowColumns_() {
  const sheet = getSheet('Enrollments');
  const data  = sheet.getDataRange().getValues();
  const norm  = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const head  = data[0].map(norm);

  const targets = [];
  for (let c = 0; c < head.length; c++) {
    if (head[c] !== '') continue;              // only unheadered columns
    let filled = 0, example = '';
    for (let r = 1; r < data.length; r++) {
      const v = norm(data[r][c]);
      if (v) { filled++; if (!example) example = v; }
    }
    targets.push({ index: c, letter: colLetter_(c + 1), filled: filled, example: example });
  }
  return { targets: targets, head: head };
}

function dropEmptyOverflowColumns() {
  const info = findOverflowColumns_();
  if (!info.targets.length) {
    const msg = 'No unheadered columns found. Nothing to drop.';
    Logger.log(msg); return msg;
  }

  const occupied = info.targets.filter(function (t) { return t.filled > 0; });
  if (occupied.length) {
    let msg = 'REFUSED - these columns still hold data:\n';
    occupied.forEach(function (t) {
      msg += '  ' + t.letter + ' : ' + t.filled + ' value(s), e.g. "' + t.example + '"\n';
    });
    msg += '\nRun repairColumnAlignment() first - that data is the Heard From and\n';
    msg += 'Notes values for the misaligned rows. Deleting now would lose them.';
    Logger.log(msg);
    return msg;
  }

  // Delete right to left so earlier indices stay valid.
  const sheet = getSheet('Enrollments');
  const letters = info.targets.map(function (t) { return t.letter; });
  info.targets.slice().sort(function (a, b) { return b.index - a.index; })
      .forEach(function (t) { sheet.deleteColumn(t.index + 1); });

  const msg = 'Dropped ' + info.targets.length + ' empty unheadered column(s): ' +
              letters.join(', ') + '. Columns to their right have shifted left; ' +
              'everything looks columns up by name, so nothing breaks.';
  Logger.log(msg);
  return msg;
}


// ─── Backfill the joining month (editor only) ─────────────────
// Enrolments added before the starting-month field existed have a blank
// joining column. Their Enrolled At date is the best available answer, and it
// is what a new enrolment would record today anyway.
//
// Only ever fills blanks. Skips the imported legacy roster: those students
// joined long before this system, so the import date says nothing useful.

function monthYearOf_(v) {
  if (v === null || v === undefined || v === '') return '';
  const d = (Object.prototype.toString.call(v) === '[object Date]') ? v : new Date(v.toString());
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
}

function buildJoiningBackfill_() {
  const sheet = getSheet('Enrollments');
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { error: 'No data rows.' };

  const norm = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const head = data[0].map(norm);

  let iJoin = -1;
  for (let c = 0; c < head.length; c++) {
    if (head[c].indexOf('Joining Date') === 0) { iJoin = c; break; }
  }
  if (iJoin < 0) return { error: 'No joining column found. Headers: ' + head.join(' | ') };

  const iId   = head.indexOf('ID');
  const iWhen = head.indexOf('Enrolled At');
  const iType = head.indexOf('Type');
  const iName = head.indexOf('Student Name');

  const fills = [], skippedLegacy = [], alreadySet = [], unreadable = [];
  for (let r = 1; r < data.length; r++) {
    const row  = data[r];
    const id   = norm(row[iId]);
    const name = norm(row[iName]);

    if (id.indexOf('SR-LEGACY-') === 0) { skippedLegacy.push(name); continue; }
    if (norm(row[iJoin])) { alreadySet.push(name); continue; }

    const my = monthYearOf_(row[iWhen]);
    if (!my) { unreadable.push({ row: r + 1, name: name, raw: norm(row[iWhen]) }); continue; }
    fills.push({ sheetRow: r + 1, name: name, type: norm(row[iType]), value: my });
  }

  return {
    iJoin: iJoin, joinHeader: head[iJoin],
    fills: fills, skippedLegacy: skippedLegacy,
    alreadySet: alreadySet, unreadable: unreadable, total: data.length - 1
  };
}

function previewJoiningBackfill() {
  const r = buildJoiningBackfill_();
  if (r.error) { Logger.log(r.error); return r.error; }

  let out = 'JOINING MONTH BACKFILL - PREVIEW\n================================\n';
  out += 'Target column     : ' + colLetter_(r.iJoin + 1) + ' "' + r.joinHeader + '"\n';
  out += 'Data rows         : ' + r.total + '\n\n';
  out += 'Will be filled in : ' + r.fills.length + '\n';
  out += 'Legacy, skipped   : ' + r.skippedLegacy.length + '\n';
  out += 'Already has a value, left alone : ' + r.alreadySet.length + '\n';
  out += 'Enrolled At unreadable          : ' + r.unreadable.length + '\n\n';

  if (r.unreadable.length) {
    out += 'COULD NOT READ THE ENROLMENT DATE:\n';
    r.unreadable.forEach(function (x) {
      out += '  row ' + x.row + '  ' + x.name + '  raw: "' + x.raw + '"\n';
    });
    out += '\n';
  }

  if (r.fills.length) {
    out += 'TO BE FILLED:\n';
    r.fills.slice(0, 30).forEach(function (f) {
      out += '  row ' + f.sheetRow + '  ' + f.name + '  [' + f.type + ']  ->  ' + f.value + '\n';
    });
    if (r.fills.length > 30) out += '  ... ' + (r.fills.length - 30) + ' more ...\n';
  } else {
    out += 'Nothing to fill.\n';
  }

  out += '\nNothing was changed. Report only.\n';
  Logger.log(out);
  return out;
}

function backfillJoiningMonth() {
  const r = buildJoiningBackfill_();
  if (r.error) { Logger.log(r.error); return r.error; }
  if (!r.fills.length) { Logger.log('Nothing to fill.'); return 'Nothing to fill.'; }

  const sheet = getSheet('Enrollments');
  r.fills.forEach(function (f) {
    sheet.getRange(f.sheetRow, r.iJoin + 1).setValue(f.value);
  });

  const msg = 'Filled ' + r.fills.length + ' joining month(s). ' +
              r.skippedLegacy.length + ' legacy row(s) left blank, ' +
              r.alreadySet.length + ' already had a value.';
  Logger.log(msg);
  return msg;
}


// ─── Analytics (read-only, editor only) ───────────────────────
// Runs inside the sheet and reports aggregates only, so no student's personal
// details need to leave it. Writes nothing.

// "ANSHIKA  SHOME" and "Anshika Shome" are the same child. Lowercasing alone
// misses that — the double space has to go too.
// A clubbed receipt often shares the surname: "Riya & Diya Sen" names both
// children, but only "diya sen" appears as a contiguous run. Treat every word
// of the student's name appearing in the receipt as a match.
function blobNamesStudent_(blob, key) {
  if (!blob || !key) return false;
  if (blob.indexOf(key) >= 0) return true;
  const parts = key.split(' ').filter(function (p) { return p.length > 1; });
  if (parts.length < 2) return false;                 // one word is too loose
  for (let i = 0; i < parts.length; i++) {
    if (blob.indexOf(parts[i]) < 0) return false;
  }
  return true;
}

function normName_(v) {
  return (v === null || v === undefined) ? ''
    : v.toString().trim().toLowerCase().replace(/\s+/g, ' ');
}

const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

function monthIndex_(name) {
  const n = (name || '').toString().trim().toLowerCase();
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    if (MONTH_NAMES[i].toLowerCase() === n) return i;
  }
  return -1;
}

function headerIndex_(head, prefix) {
  for (let i = 0; i < head.length; i++) {
    if (head[i].indexOf(prefix) === 0) return i;
  }
  return -1;
}

function pad_(str, n) {
  let out = (str === null || str === undefined) ? '' : str.toString();
  while (out.length < n) out += ' ';
  return out;
}

function money_(n) {
  return 'Rs. ' + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function tally_(map, key) {
  const k = key || '(blank)';
  map[k] = (map[k] || 0) + 1;
}

function renderTally_(map, total, limit) {
  const keys = Object.keys(map).sort(function (a, b) { return map[b] - map[a]; });
  let out = '';
  keys.slice(0, limit || 20).forEach(function (k) {
    const pct = total ? Math.round(map[k] * 1000 / total) / 10 : 0;
    out += '  ' + pad_(k, 34) + pad_(map[k], 6) + pct + '%\n';
  });
  return out;
}

function analyticsReport() {
  const norm = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const eData = getSheet('Enrollments').getDataRange().getValues();
  const rData = getSheet('Receipts').getDataRange().getValues();
  const eHead = eData[0].map(norm);
  const rHead = rData[0].map(norm);

  const eName = eHead.indexOf('Student Name');
  const eType = eHead.indexOf('Type');
  const eLoc  = eHead.indexOf('Location');
  const eStat = eHead.indexOf('Status');
  const ePh   = eHead.indexOf('Phone');
  const eDob  = eHead.indexOf('Date of Birth');
  const eJoin = headerIndex_(eHead, 'Joining Date');
  const eHeard = eHead.indexOf('Heard From');

  let out = 'SRIRAM STUDIO - ANALYTICS\n=========================\n';
  out += 'Generated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy') + '\n\n';

  // ── Roster ──
  const byType = {}, byCentre = {}, byHeard = {}, byJoin = {};
  let active = 0, left = 0, noPhone = 0, noCentre = 0, noDob = 0, noJoin = 0;
  const ages = { 'under 6': 0, '6-9': 0, '10-13': 0, '14-17': 0, '18+': 0 };
  const activeNames = {};
  const today = new Date();

  for (let i = 1; i < eData.length; i++) {
    const row = eData[i];
    const name = norm(row[eName]);
    if (!name) continue;
    const isLeft = isLeftWord_(norm(row[eStat]));
    if (isLeft) { left++; } else { active++; activeNames[normName_(name)] = name; }

    tally_(byType, norm(row[eType]));
    const centre = norm(row[eLoc]).split('\u2013')[0].trim();
    tally_(byCentre, centre || '(not recorded)');
    if (eHeard >= 0 && norm(row[eHeard])) tally_(byHeard, norm(row[eHeard]));
    if (eJoin >= 0) { const j = norm(row[eJoin]); if (j) tally_(byJoin, j); else noJoin++; }

    if (!norm(row[ePh])) noPhone++;
    if (!centre) noCentre++;
    const dobRaw = row[eDob];
    if (!norm(dobRaw)) { noDob++; }
    else {
      const d = (Object.prototype.toString.call(dobRaw) === '[object Date]') ? dobRaw : new Date(norm(dobRaw));
      if (!isNaN(d.getTime())) {
        const age = Math.floor((today - d) / (365.25 * 24 * 3600 * 1000));
        if (age < 6) ages['under 6']++;
        else if (age < 10) ages['6-9']++;
        else if (age < 14) ages['10-13']++;
        else if (age < 18) ages['14-17']++;
        else ages['18+']++;
      }
    }
  }
  const roster = active + left;

  out += 'ROSTER\n------\n';
  out += '  Total on record                 ' + roster + '\n';
  out += '  Active                          ' + active + '\n';
  out += '  Left                            ' + left + '\n\n';

  out += 'BY TYPE\n' + renderTally_(byType, roster) + '\n';
  out += 'BY CENTRE\n' + renderTally_(byCentre, roster) + '\n';

  out += 'DATA GAPS (of ' + roster + ')\n';
  out += '  No phone number                 ' + noPhone + '\n';
  out += '  No centre recorded              ' + noCentre + '\n';
  out += '  No date of birth                ' + noDob + '\n';
  out += '  No joining month                ' + noJoin + '\n\n';

  const withDob = roster - noDob;
  if (withDob > 0) {
    out += 'AGE (of the ' + withDob + ' with a date of birth)\n';
    Object.keys(ages).forEach(function (k) {
      if (ages[k]) out += '  ' + pad_(k, 34) + ages[k] + '\n';
    });
    out += '\n';
  }
  if (Object.keys(byHeard).length) {
    out += 'HOW THEY HEARD OF THE STUDIO\n' + renderTally_(byHeard, 0, 10) + '\n';
  }

  // ── Receipts ──
  const rNo   = rHead.indexOf('Receipt No');
  const rName = rHead.indexOf('Student Name');
  const rStu  = rHead.indexOf('Students');
  const rAmt  = headerIndex_(rHead, 'Amount');
  const rMon  = rHead.indexOf('Fee Month');
  const rYr   = rHead.indexOf('Fee Year');
  const rType = rHead.indexOf('Fee Type');
  const rMode = rHead.indexOf('Payment Mode');

  if (rData.length <= 1) {
    out += 'RECEIPTS\n--------\n  None issued yet.\n';
    Logger.log(out);
    return out;
  }

  const byPeriod = {}, byFeeType = {}, byMode = {};
  const perStudent = {}, blobs = [];
  let totalAmt = 0, count = 0, clubbed = 0;
  let latestPeriod = -1;

  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    if (!norm(row[rNo])) continue;
    count++;
    const amt = parseFloat(norm(row[rAmt]).replace(/[^0-9.]/g, '')) || 0;
    totalAmt += amt;

    const mi = monthIndex_(norm(row[rMon]));
    const yr = parseInt(norm(row[rYr]), 10);
    let period = -1;
    if (mi >= 0 && yr) {
      period = yr * 12 + mi;
      const label = MONTH_NAMES[mi].substring(0, 3) + ' ' + yr;
      if (!byPeriod[label]) byPeriod[label] = { n: 0, amt: 0, key: period };
      byPeriod[label].n++;
      byPeriod[label].amt += amt;
      if (period > latestPeriod) latestPeriod = period;
    }

    if (rType >= 0) tally_(byFeeType, norm(row[rType]) || '(blank)');
    if (rMode >= 0) tally_(byMode, norm(row[rMode]) || '(blank)');

    const raw = rStu >= 0 ? norm(row[rStu]) : '';
    const names = raw ? raw.split('|') : [norm(row[rName])];
    if (names.length > 1) clubbed++;
    names.forEach(function (nm) {
      const key = normName_(nm);
      if (!key) return;
      if (!perStudent[key]) perStudent[key] = { name: nm.trim(), n: 0, last: -1 };
      perStudent[key].n++;
      if (period > perStudent[key].last) perStudent[key].last = period;
    });

    // Siblings were merged into one receipt long before the Students column
    // existed, so "Riya & Diya Sen" matches neither child by name. Keep the
    // whole string to search through afterwards.
    const blob = normName_(rStu >= 0 && raw ? raw.replace(/\|/g, ' ') : norm(row[rName]));
    if (blob) blobs.push({ text: blob, period: period });
  }

  out += 'RECEIPTS\n--------\n';
  out += '  Receipts issued                 ' + count + '\n';
  out += '  Total collected                 ' + money_(totalAmt) + '\n';
  out += '  Average receipt                 ' + money_(count ? totalAmt / count : 0) + '\n';
  out += '  Covering more than one student  ' + clubbed + '\n\n';

  const periods = Object.keys(byPeriod).sort(function (a, b) { return byPeriod[a].key - byPeriod[b].key; });
  out += 'BY FEE PERIOD\n';
  periods.slice(-15).forEach(function (p) {
    out += '  ' + pad_(p, 12) + pad_(byPeriod[p].n + ' receipts', 16) + money_(byPeriod[p].amt) + '\n';
  });
  out += '\n';

  if (Object.keys(byFeeType).length) out += 'BY FEE TYPE\n' + renderTally_(byFeeType, count) + '\n';
  if (Object.keys(byMode).length)    out += 'BY PAYMENT MODE\n' + renderTally_(byMode, count) + '\n';

  // ── Who is paying ──
  const billed = [], neverBilled = [], lapsed = [];
  let viaClub = 0;
  Object.keys(activeNames).forEach(function (k) {
    let rec = perStudent[k];

    // No exact match? The student may be named inside a clubbed receipt.
    if (!rec || !rec.n) {
      let last = -1, hits = 0;
      for (let b = 0; b < blobs.length; b++) {
        if (blobNamesStudent_(blobs[b].text, k)) {
          hits++;
          if (blobs[b].period > last) last = blobs[b].period;
        }
      }
      if (hits) { rec = { name: activeNames[k], n: hits, last: last }; viaClub++; }
    }

    if (!rec || !rec.n) { neverBilled.push(activeNames[k]); return; }
    billed.push(rec);
    if (latestPeriod >= 0 && rec.last >= 0 && (latestPeriod - rec.last) >= 2) {
      lapsed.push({ name: rec.name, behind: latestPeriod - rec.last });
    }
  });

  out += 'PAYMENT COVERAGE (active students only)\n';
  out += '  Active students                 ' + active + '\n';
  out += '  Have at least one receipt       ' + billed.length + '\n';
  out += '    of which matched only inside a clubbed receipt  ' + viaClub + '\n';
  out += '  Never had a receipt             ' + neverBilled.length + '\n';
  out += '  Last paid 2+ periods ago        ' + lapsed.length + '\n\n';

  if (lapsed.length) {
    lapsed.sort(function (a, b) { return b.behind - a.behind; });
    out += 'FURTHEST BEHIND (top 15)\n';
    lapsed.slice(0, 15).forEach(function (x) {
      out += '  ' + pad_(x.name, 30) + x.behind + ' periods\n';
    });
    out += '\n';
  }
  if (neverBilled.length) {
    out += 'ACTIVE BUT NEVER INVOICED (first 15 of ' + neverBilled.length + ')\n';
    neverBilled.slice(0, 15).forEach(function (n) { out += '  ' + n + '\n'; });
    out += '\n';
  }

  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Duplicate students (read-only, editor only) ──────────────
// The roster import matched on a lowercased name, so "ANSHIKA  SHOME" and
// "Anshika Shome" were treated as different people and both were kept.
// Reports only — genuinely different students do share names, so nothing is
// merged automatically.

function findDuplicateStudents() {
  const norm  = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const data  = getSheet('Enrollments').getDataRange().getValues();
  const head  = data[0].map(norm);
  const iName = head.indexOf('Student Name');
  const iId   = head.indexOf('ID');
  const iType = head.indexOf('Type');
  const iStat = head.indexOf('Status');
  const iPh   = head.indexOf('Phone');

  const groups = {};
  for (let r = 1; r < data.length; r++) {
    const raw = norm(data[r][iName]);
    if (!raw) continue;
    const key = normName_(raw);
    (groups[key] = groups[key] || []).push({
      row: r + 1, raw: raw,
      id: norm(data[r][iId]), type: norm(data[r][iType]),
      status: norm(data[r][iStat]) || 'Active',
      phone: norm(data[r][iPh]).replace(/\D/g, '')
    });
  }

  const likely = [], sameNameDifferentPeople = [];
  Object.keys(groups).forEach(function (k) {
    const g = groups[k];
    if (g.length < 2) return;
    // Spelled differently but the same name, or the same contact: one person.
    const spellings = {};
    g.forEach(function (x) { spellings[x.raw] = true; });
    const phones = g.map(function (x) { return x.phone; });
    const distinctPhones = phones.filter(function (p, i) { return p && phones.indexOf(p) === i; });
    if (Object.keys(spellings).length > 1 || distinctPhones.length <= 1) likely.push(g);
    else sameNameDifferentPeople.push(g);
  });

  let out = 'DUPLICATE STUDENTS\n==================\n';
  out += 'Rows scanned : ' + (data.length - 1) + '\n';
  out += 'Likely duplicates      : ' + likely.length + ' name(s)\n';
  out += 'Same name, different people : ' + sameNameDifferentPeople.length + ' name(s)\n\n';

  const show = function (g) {
    g.forEach(function (x) {
      out += '   row ' + pad_(x.row, 6) + pad_('"' + x.raw + '"', 28) +
             pad_(x.type, 24) + pad_(x.status, 9) + (x.phone || '(no phone)') + '\n';
    });
    out += '\n';
  };

  if (likely.length) {
    out += 'LIKELY THE SAME CHILD - different spelling, or one contact number:\n\n';
    likely.slice(0, 25).forEach(show);
    if (likely.length > 25) out += '... ' + (likely.length - 25) + ' more.\n\n';
  }
  if (sameNameDifferentPeople.length) {
    out += 'SAME NAME BUT DIFFERENT PHONES - probably genuinely two students:\n\n';
    sameNameDifferentPeople.slice(0, 15).forEach(show);
  }

  out += 'Report only. Nothing was changed. Merge by hand: keep the row with the\n';
  out += 'fuller record, move anything useful across, then delete the other.\n';
  Logger.log(out);
  return out;
}


// ─── Unbilled review (read-only, editor only) ─────────────────
// Some students look unbilled only because the roster spells them differently
// from the receipt history — "Tashvi Kocahr" against "Tashvi Kochar". This
// separates a misspelling from someone who genuinely has never paid, so the
// remainder is a collections list rather than a list of data errors.

function levenshtein_(a, b) {
  a = a.toLowerCase(); b = b.toLowerCase();
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let prev = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1,
                        prev[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}

function reviewUnbilledStudents() {
  const norm  = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const eData = getSheet('Enrollments').getDataRange().getValues();
  const rData = getSheet('Receipts').getDataRange().getValues();
  const eHead = eData[0].map(norm);
  const rHead = rData[0].map(norm);

  const eName = eHead.indexOf('Student Name');
  const eStat = eHead.indexOf('Status');
  const ePh   = eHead.indexOf('Phone');
  const rName = rHead.indexOf('Student Name');
  const rStu  = rHead.indexOf('Students');

  // Every name the receipts know about, and the blobs to search within.
  const paid = {}, blobs = [], receiptNames = {};
  for (let i = 1; i < rData.length; i++) {
    const rawStu = rStu >= 0 ? norm(rData[i][rStu]) : '';
    const whole  = rawStu ? rawStu.replace(/\|/g, ' ') : norm(rData[i][rName]);
    const blob   = normName_(whole);
    if (!blob) continue;
    blobs.push(blob);
    (rawStu ? rawStu.split('|') : [norm(rData[i][rName])]).forEach(function (nm) {
      const k = normName_(nm);
      if (!k) return;
      paid[k] = true;
      receiptNames[k] = (receiptNames[k] || 0) + 1;
    });
  }
  const receiptKeys = Object.keys(receiptNames);

  // Every active student, so a "correction" pointing at another child on the
  // roster can be rejected rather than offered.
  const rosterNames = {};
  for (let r = 1; r < eData.length; r++) {
    const nm = norm(eData[r][eName]);
    if (nm && !isLeftWord_(norm(eData[r][eStat]))) rosterNames[normName_(nm)] = nm;
  }

  // Active students the receipts cannot account for.
  const unbilled = [];
  for (let r = 1; r < eData.length; r++) {
    const raw = norm(eData[r][eName]);
    if (!raw) continue;
    if (isLeftWord_(norm(eData[r][eStat]))) continue;
    const key = normName_(raw);
    if (paid[key]) continue;
    let inBlob = false;
    for (let b = 0; b < blobs.length; b++) {
      if (blobNamesStudent_(blobs[b], key)) { inBlob = true; break; }
    }
    if (inBlob) continue;
    unbilled.push({ row: r + 1, name: raw, key: key, phone: norm(eData[r][ePh]) });
  }

  // Nearest receipt name for each.
  const likelyTypo = [], ambiguous = [], genuine = [];
  unbilled.forEach(function (u) {
    // Nearest receipt name that is NOT another student on the roster.
    let best = null, bestD = 99, nearAny = null, nearAnyD = 99;
    for (let i = 0; i < receiptKeys.length; i++) {
      const cand = receiptKeys[i];
      const d = levenshtein_(u.key, cand);
      if (d < nearAnyD) { nearAnyD = d; nearAny = cand; }
      if (rosterNames[cand]) continue;      // that receipt belongs to a real student
      if (d < bestD) { bestD = d; best = cand; }
    }
    // Tighter than before: on this roster, given names routinely differ by one
    // or two characters while belonging to different children.
    const tolerance = Math.min(2, Math.max(1, Math.round(u.key.length * 0.15)));

    if (best && bestD <= tolerance) {
      likelyTypo.push({ name: u.name, row: u.row, near: best, d: bestD, n: receiptNames[best] });
    } else if (nearAny && rosterNames[nearAny] && nearAnyD <= 2) {
      ambiguous.push({ name: u.name, row: u.row, near: rosterNames[nearAny], d: nearAnyD });
    } else {
      genuine.push({ name: u.name, row: u.row, phone: u.phone, near: best || nearAny, d: bestD === 99 ? nearAnyD : bestD });
    }
  });

  let out = 'UNBILLED STUDENTS - REVIEW\n==========================\n';
  out += 'Active students the receipts cannot account for : ' + unbilled.length + '\n';
  out += '  Probably a spelling difference               : ' + likelyTypo.length + '\n';
  out += '  Too close to another student to call         : ' + ambiguous.length + '\n';
  out += '  No close match - genuinely never invoiced    : ' + genuine.length + '\n\n';

  if (likelyTypo.length) {
    likelyTypo.sort(function (a, b) { return a.d - b.d; });
    out += 'PROBABLY THE SAME CHILD, SPELLED DIFFERENTLY\n';
    out += 'Fix the spelling in Enrollments and they stop showing as unpaid.\n\n';
    likelyTypo.forEach(function (x) {
      out += '  row ' + pad_(x.row, 6) + pad_('"' + x.name + '"', 28) +
             ' receipts say "' + x.near + '"  (' + x.d + ' char, ' + x.n + ' receipt(s))\n';
    });
    out += '\n';
  }

  if (ambiguous.length) {
    out += 'CANNOT TELL - the nearest receipt name is another student on your\n';
    out += 'roster, so this is either a typo or simply a similar name. Check by\n';
    out += 'hand; do not merge on my say-so.\n\n';
    ambiguous.forEach(function (x) {
      out += '  row ' + pad_(x.row, 6) + pad_('"' + x.name + '"', 28) +
             ' similar to enrolled student "' + x.near + '" (' + x.d + ' char)\n';
    });
    out += '\n';
  }

  if (genuine.length) {
    out += 'NO MATCHING RECEIPT AT ALL - the actual collections list\n\n';
    genuine.forEach(function (x) {
      out += '  row ' + pad_(x.row, 6) + pad_('"' + x.name + '"', 28) +
             pad_(x.phone || '(no phone)', 14) + 'nearest: "' + x.near + '" (' + x.d + ' off)\n';
    });
    out += '\n';
  }

  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Duplicate receipts (read-only, editor only) ──────────────
// A receipt reissued after a correction, or submitted twice, inflates revenue.
// Groups receipts covering the same student(s) for the same fee period, then
// sorts them by what the evidence actually suggests happened.
//
// Reports only. Issued receipts are historical records of money that changed
// hands, and none of this is certain enough to act on unchecked.

function shortDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  const d = (Object.prototype.toString.call(v) === '[object Date]') ? v : new Date(v.toString());
  if (isNaN(d.getTime())) return v.toString().substring(0, 12);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM');
}

// A note saying "fee for August" on a receipt filed under July means the
// period was mis-keyed — both payments are real.
function noteMonth_(note) {
  const t = (note || '').toString().toLowerCase();
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    if (t.indexOf(MONTH_NAMES[i].toLowerCase()) >= 0) return MONTH_NAMES[i];
  }
  return '';
}

function findDuplicateReceipts() {
  const norm = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const data = getSheet('Receipts').getDataRange().getValues();
  if (data.length <= 1) { Logger.log('No receipts.'); return 'No receipts.'; }

  const head = data[0].map(norm);
  const iNo   = head.indexOf('Receipt No');
  const iWhen = head.indexOf('Issued At');
  const iName = head.indexOf('Student Name');
  const iStu  = head.indexOf('Students');
  const iAmt  = headerIndex_(head, 'Amount');
  const iMon  = head.indexOf('Fee Month');
  const iYr   = head.indexOf('Fee Year');
  const iType = head.indexOf('Fee Type');
  const iMode = head.indexOf('Payment Mode');
  const iNote = head.indexOf('Note');

  const groups = {};
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const no  = norm(row[iNo]);
    if (!no) continue;
    const rawStu = iStu >= 0 ? norm(row[iStu]) : '';
    const who = (rawStu ? rawStu.split('|') : [norm(row[iName])])
                  .map(normName_).filter(Boolean).sort().join(' + ');
    if (!who) continue;
    const month = norm(row[iMon]);
    const key = who + ' || ' + (month + ' ' + norm(row[iYr])).trim();
    (groups[key] = groups[key] || []).push({
      row: r + 1, no: no, when: shortDate_(row[iWhen]), who: who,
      period: (month + ' ' + norm(row[iYr])).trim() || '(no period)',
      month: month,
      amt: parseFloat(norm(row[iAmt]).replace(/[^0-9.]/g, '')) || 0,
      type: norm(row[iType]) || '(blank)',
      mode: norm(row[iMode]),
      note: norm(row[iNote]),
      seq: parseInt((no.match(/(\d+)\s*$/) || [0, 0])[1], 10)
    });
  }

  const repeated = [], modeFix = [], spread = [], revised = [], misKeyed = [], otherType = [];
  let overRepeat = 0, overMode = 0, overSpread = 0, overRevised = 0;

  Object.keys(groups).forEach(function (k) {
    const g = groups[k];
    if (g.length < 2) return;

    const byType = {};
    g.forEach(function (x) { (byType[x.type] = byType[x.type] || []).push(x); });

    let flagged = false;
    Object.keys(byType).forEach(function (t) {
      const rows = byType[t];
      if (rows.length < 2) return;
      flagged = true;

      // Notes naming different months: the period was mis-keyed, not doubled.
      const noteMonths = {};
      rows.forEach(function (x) { const m = noteMonth_(x.note); if (m) noteMonths[m] = true; });
      if (Object.keys(noteMonths).length > 1) { misKeyed.push(rows); return; }

      const amounts = rows.map(function (x) { return x.amt; });
      const allSame = amounts.every(function (a) { return a === amounts[0]; });
      if (!allSame) {
        revised.push(rows);
        overRevised += amounts.reduce(function (a, b) { return a + b; }, 0) - Math.max.apply(null, amounts);
        return;
      }

      const extra = amounts[0] * (rows.length - 1);
      const days  = {}, modes = {};
      rows.forEach(function (x) { days[x.when] = true; modes[x.mode] = true; });
      const sameDay = Object.keys(days).length === 1;
      const seqs = rows.map(function (x) { return x.seq; }).sort(function (a, b) { return a - b; });
      const consecutive = seqs.every(function (v, i) { return i === 0 || v === seqs[i - 1] + 1; });

      if (sameDay && consecutive && Object.keys(modes).length === 1) {
        repeated.push(rows); overRepeat += extra;
      } else if (sameDay && Object.keys(modes).length > 1) {
        modeFix.push(rows); overMode += extra;
      } else {
        spread.push(rows); overSpread += extra;
      }
    });
    if (!flagged && Object.keys(byType).length > 1) otherType.push(g);
  });

  const line = function (x) {
    return '     row ' + pad_(x.row, 6) + pad_(x.no, 15) + pad_(x.when, 8) +
           pad_(money_(x.amt), 12) + pad_(x.mode, 6) +
           (x.note ? x.note.substring(0, 40) : '') + '\n';
  };
  const show = function (rows, out) {
    let t = '  ' + rows[0].who + '   [' + rows[0].period + ']\n';
    rows.forEach(function (x) { t += line(x); });
    return t + '\n';
  };

  let out = 'POSSIBLE DUPLICATE RECEIPTS\n===========================\n';
  out += 'Receipts scanned : ' + (data.length - 1) + '\n\n';
  out += pad_('Submitted more than once', 38) + pad_(repeated.length + ' grp', 9) + money_(overRepeat) + '\n';
  out += pad_('Reissued with a different pay mode', 38) + pad_(modeFix.length + ' grp', 9) + money_(overMode) + '\n';
  out += pad_('Same amount, different days', 38) + pad_(spread.length + ' grp', 9) + money_(overSpread) + '\n';
  out += pad_('Differing amounts (a reissue?)', 38) + pad_(revised.length + ' grp', 9) + money_(overRevised) + '\n';
  out += pad_('Period mis-keyed - NOT duplicates', 38) + pad_(misKeyed.length + ' grp', 9) + 'Rs. 0\n';
  out += pad_('Different fee type - usually fine', 38) + pad_(otherType.length + ' grp', 9) + 'Rs. 0\n';
  out += '\nMost likely over-counted: ' + money_(overRepeat + overMode) + '\n';
  out += 'Needs checking on top of that: ' + money_(overSpread + overRevised) + '\n\n';

  if (repeated.length) {
    out += 'SUBMITTED MORE THAN ONCE\n';
    out += 'Same day, same payment mode, consecutive receipt numbers. Almost\n';
    out += 'certainly the Generate button firing repeatedly. Keep the first.\n\n';
    repeated.forEach(function (g) { out += show(g); });
  }
  if (modeFix.length) {
    out += 'REISSUED WITH A DIFFERENT PAYMENT MODE\n';
    out += 'Same day and amount, one Cash and one UPI. Looks like a correction\n';
    out += 'made by issuing a second receipt. Keep the one with the right mode.\n\n';
    modeFix.forEach(function (g) { out += show(g); });
  }
  if (spread.length) {
    out += 'SAME AMOUNT, DIFFERENT DAYS\n';
    out += 'Could be a genuine second payment. Check before deleting.\n\n';
    spread.forEach(function (g) { out += show(g); });
  }
  if (revised.length) {
    out += 'DIFFERING AMOUNTS FOR THE SAME PERIOD\n';
    out += 'A corrected reissue, or a part payment and then the balance.\n\n';
    revised.forEach(function (g) { out += show(g); });
  }
  if (misKeyed.length) {
    out += 'PERIOD MIS-KEYED - BOTH PAYMENTS ARE REAL\n';
    out += 'The notes name different months, so the Fee Month is wrong on one.\n';
    out += 'Fix the month; do not delete either.\n\n';
    misKeyed.forEach(function (g) { out += show(g); });
  }
  if (otherType.length) {
    out += 'SAME PERIOD, DIFFERENT FEE TYPE (' + otherType.length + ' group(s), normally legitimate)\n';
    out += 'Registration or costume fees alongside the monthly fee.\n\n';
  }

  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}
