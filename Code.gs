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

// Apps Script responses are slow enough that a receipt can save while the
// panel shows a connection error, and the natural reaction is to press
// Generate again. Six consecutive receipt numbers for one student at one
// timestamp is what that looks like in the data. Remembering each receipt
// briefly makes a retry return the original instead of minting another.
const RECEIPT_DUP_WINDOW_SEC = 15 * 60;

function receiptFingerprint_(d) {
  const who = (d.students && d.students.length)
    ? d.students.map(normName_).sort().join('+')
    : normName_(d.studentName);
  const fp = 'rcpt|' + who + '|' + (d.month || '') + (d.year || '') +
             '|' + (d.amount || '') + '|' + (d.feeType || '');
  return fp.substring(0, 240);
}

function addReceipt(d) {
  ensureReceiptStudentsColumn_();

  // Same student(s), period, amount and fee type within the window: this is a
  // retry, not a second payment. A genuine second payment differs in amount or
  // falls outside it, and either way the duplicate report still catches it.
  const cache = CacheService.getScriptCache();
  const fp    = receiptFingerprint_(d);
  const seen  = cache.get(fp);
  if (seen) {
    const parts = seen.split('||');
    return { success: true, receiptNo: parts[0], issuedAt: parts[1], duplicate: true };
  }
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

  cache.put(fp, receiptNo + '||' + issuedAt, RECEIPT_DUP_WINDOW_SEC);
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
  const today = new Date();

  for (let i = 1; i < eData.length; i++) {
    const row = eData[i];
    const name = norm(row[eName]);
    if (!name) continue;
    const isLeft = isLeftWord_(norm(row[eStat]));
    if (isLeft) { left++; } else { active++; }

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
  let totalAmt = 0, count = 0, clubbed = 0;

  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    if (!norm(row[rNo])) continue;
    count++;
    const amt = parseFloat(norm(row[rAmt]).replace(/[^0-9.]/g, '')) || 0;
    totalAmt += amt;

    const mi = monthIndex_(norm(row[rMon]));
    const yr = parseInt(norm(row[rYr]), 10);
    if (mi >= 0 && yr) {
      const label = MONTH_NAMES[mi].substring(0, 3) + ' ' + yr;
      if (!byPeriod[label]) byPeriod[label] = { n: 0, amt: 0, key: yr * 12 + mi };
      byPeriod[label].n++;
      byPeriod[label].amt += amt;
    }

    if (rType >= 0) tally_(byFeeType, norm(row[rType]) || '(blank)');
    if (rMode >= 0) tally_(byMode, norm(row[rMode]) || '(blank)');

    // Siblings clubbed onto one receipt: 'Riya Sen | Diya Sen', or on older
    // receipts 'Riya & Diya Sen' in a single cell. splitReceiptNames_ handles
    // both; here only the count of clubbed receipts is wanted.
    if (splitReceiptNames_(rStu >= 0 ? norm(row[rStu]) : '', norm(row[rName])).length > 1) clubbed++;
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
  // Delegated to buildFeeCoverage_ so this report and feeCoverageByMonth can
  // never disagree about who has paid: one resolver, one set of numbers.
  const cov = buildFeeCoverage_();
  const billed = [], neverBilled = [], lapsed = [];
  cov.students.forEach(function (s) {
    if (!s.billable || s.left) return;
    if (s.receipts) billed.push(s); else { neverBilled.push(s.name); return; }
    if (s.lastPaid >= 0 && (cov.now - s.lastPaid) >= 2) {
      lapsed.push({ name: s.name, behind: cov.now - s.lastPaid });
    }
  });
  const closedMonths = cov.months.filter(function (m) {
    return m.period < cov.now && m.expected.length;
  });
  const lastClosed = closedMonths.length ? closedMonths[closedMonths.length - 1] : null;

  out += 'PAYMENT COVERAGE (active monthly-fee students)\n';
  out += '  Active students                 ' + cov.counts.active + '\n';
  out += '  Have at least one receipt       ' + billed.length + '\n';
  out += '  Never had a receipt             ' + neverBilled.length + '\n';
  out += '  Last paid 2+ months ago         ' + lapsed.length + '\n';
  if (lastClosed) {
    out += '  ' + pad_(periodLong_(lastClosed.period) + ' fully paid', 32) +
           (lastClosed.complete ? 'yes'
            : 'no - ' + lastClosed.unpaid.length + ' of ' +
              lastClosed.expected.length + ' outstanding') + '\n';
  }
  out += '\n';

  if (lapsed.length) {
    lapsed.sort(function (a, b) { return b.behind - a.behind; });
    out += 'FURTHEST BEHIND (top 15)\n';
    lapsed.slice(0, 15).forEach(function (x) {
      out += '  ' + pad_(x.name, 30) + x.behind + ' months\n';
    });
    out += '\n';
  }
  if (neverBilled.length) {
    out += 'ACTIVE BUT NEVER INVOICED (first 15 of ' + neverBilled.length + ')\n';
    neverBilled.slice(0, 15).forEach(function (n) { out += '  ' + n + '\n'; });
    out += '\n';
  }

  out += 'Month-by-month coverage: feeCoverageByMonth().\n';
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
  const iCon  = head.indexOf('Contact');

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
      contact: iCon >= 0 ? norm(row[iCon]).replace(/\D/g, '') : '',
      seq: parseInt((no.match(/(\d+)\s*$/) || [0, 0])[1], 10)
    });
  }

  const repeated = [], modeFix = [], spread = [], revised = [], misKeyed = [],
        otherType = [], namesake = [];
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

      // Two children can share a name — there are two Krisha Agarwals on the
      // roster. Different contact numbers mean different families, so these
      // are separate fees, not one fee receipted twice.
      const contacts = {};
      rows.forEach(function (x) { if (x.contact) contacts[x.contact] = true; });
      if (Object.keys(contacts).length > 1) { namesake.push(rows); return; }

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
           pad_(money_(x.amt), 12) + pad_(x.mode, 6) + pad_(x.contact || '-', 12) +
           (x.note ? x.note.substring(0, 32) : '') + '\n';
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
  out += pad_('Same name, different family - NOT dup', 38) + pad_(namesake.length + ' grp', 9) + 'Rs. 0\n';
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
  if (namesake.length) {
    out += 'SAME NAME, DIFFERENT CONTACT NUMBERS - TWO DIFFERENT CHILDREN\n';
    out += 'Separate families paying separate fees. Do not delete either.\n\n';
    namesake.forEach(function (g) { out += show(g); });
  }
  if (otherType.length) {
    out += 'SAME PERIOD, DIFFERENT FEE TYPE (' + otherType.length + ' group(s), normally legitimate)\n';
    out += 'Registration or costume fees alongside the monthly fee.\n\n';
  }

  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Fee coverage by month (read-only, editor only) ───────────
// "Has everybody paid for August?" cannot be answered by counting receipts:
// one receipt often covers two or three siblings, and the same child is
// spelled differently in the roster and on the receipt. This section resolves
// every name printed on every receipt back to a row in Enrollments — the
// golden source — and then asks, month by month, which students the receipts
// actually account for.
//
// Nothing here writes. Run previewReceiptNameMatching() first and read how
// each name was resolved before trusting the coverage numbers.

// Two spellings, one child. Whichever spelling is in Enrollments wins; the
// others are treated as that child on receipts. If neither spelling — or both,
// as separate rows — is on the roster, the report says so instead of guessing.
const NAME_EQUIVALENTS = [
  ['Vani Maskara', 'Vaani Maskara'],
  // Three characters apart, which is past what the matcher will risk on its
  // own, but the surname and the family are unmistakable.
  ['Krishanya Kanoria', 'Krishnaya Kanodia'],
  // A phone keyboard turned this one into a possessive. ’ is the curly
  // apostrophe the receipt actually carries; the straight one is here in case
  // it is ever typed the ordinary way.
  ['Jisha Desai', 'Josh’s Desai', "Josh's Desai"]
];

// Two children, similar names. Both are on the roster and must never be
// collapsed into one another by the spelling-tolerant matcher.
const DISTINCT_STUDENTS = [
  ['Jia Bhimani', 'Jeena Bhimani'],
  // The same letters with the h and n swapped, and two different children.
  // A receipt reading 'Janvi Jain' is one character from each and cannot be
  // credited to either without a contact number to settle it.
  ['Jahnvi Jain', 'Janhvi Jain']
];

// How many months the headline table shows.
const COVERAGE_MONTHS = 24;

function periodNow_() {
  const d = new Date();
  return d.getFullYear() * 12 + d.getMonth();
}

function periodLabel_(p) {
  if (p === null || p === undefined || p < 0) return '(none)';
  return MONTH_NAMES[p % 12].substring(0, 3) + ' ' + Math.floor(p / 12);
}

function periodLong_(p) {
  if (p === null || p === undefined || p < 0) return '(none)';
  return MONTH_NAMES[p % 12] + ' ' + Math.floor(p / 12);
}

// 'March', 'Mar' and 'Sept' all mean the same month.
function monthIndexLoose_(name) {
  const n = (name || '').toString().trim().toLowerCase();
  if (!n) return -1;
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    const full = MONTH_NAMES[i].toLowerCase();
    if (full === n) return i;
    if (n.length >= 3 && full.indexOf(n) === 0) return i;
  }
  return -1;
}

// The sheet holds months as free text ('March 2027'), as real dates, and as
// whatever a form once sent. Reduce any of them to a year*12+month number.
function parsePeriod_(v) {
  if (v === null || v === undefined || v === '') return -1;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? -1 : v.getFullYear() * 12 + v.getMonth();
  }
  const s = v.toString().trim();
  if (!s) return -1;

  const named = s.match(/([A-Za-z]{3,})[^A-Za-z0-9]*(\d{4})/);
  if (named) {
    const mi = monthIndexLoose_(named[1]);
    if (mi >= 0) return parseInt(named[2], 10) * 12 + mi;
  }
  const iso = s.match(/^(\d{4})[-\/.](\d{1,2})/);
  if (iso) {
    const mo = parseInt(iso[2], 10);
    if (mo >= 1 && mo <= 12) return parseInt(iso[1], 10) * 12 + mo - 1;
  }
  const my = s.match(/^(\d{1,2})[-\/.](\d{4})$/);
  if (my) {
    const mo = parseInt(my[1], 10);
    if (mo >= 1 && mo <= 12) return parseInt(my[2], 10) * 12 + mo - 1;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear() * 12 + d.getMonth();
  return -1;
}

// One receipt, several children. The Students column separates them with '|',
// but receipts issued before that column existed printed 'Riya & Diya Sen'
// into a single cell — and only the last child carries the surname.
function splitReceiptNames_(rawStudents, rawName) {
  const src = (rawStudents !== null && rawStudents !== undefined &&
               rawStudents.toString().trim())
    ? rawStudents.toString()
    : (rawName === null || rawName === undefined ? '' : rawName.toString());
  if (!src.trim()) return [];

  const parts = src.split(/\s*(?:\||&|\+|\/|,|\band\b)\s*/i)
    .map(function (p) { return p.replace(/\s+/g, ' ').trim(); })
    .filter(function (p) { return p.length > 0; });
  if (!parts.length) return [];

  const tail = parts[parts.length - 1].split(' ');
  const surname = tail.length > 1 ? tail[tail.length - 1] : '';
  return parts.map(function (p) {
    return (surname && p.indexOf(' ') < 0) ? (p + ' ' + surname) : p;
  });
}

function digitsOf_(v) {
  return (v === null || v === undefined) ? '' : v.toString().replace(/\D/g, '').slice(-10);
}

// Spelling tolerance. Given names on this roster routinely differ by one or
// two characters while belonging to different children, so this stays tight —
// anything further apart is reported for a human to decide, not merged.
function nameTolerance_(key) {
  return key.replace(/\s/g, '').length <= 8 ? 1 : 2;
}

// Families often pay two months at once and Anjali writes it in the note —
// "July and August fees", "2 months Aug Sept". The Fee Month column can only
// hold one of them, so the other month would read as unpaid.
//
// Naming a month is not the same as paying for it. "Fee for June & late fee
// for May" settles June and adds a penalty for May — May's own tuition may
// still be outstanding. Same for "100 advanced for August": an advance is not
// August's fee. "Balance of July" is different, that money is July's fee.
const NOTE_NOT_A_FEE = /(late\s*fees?|latefees?|penalty|fine|advance)/;

// Every month mentioned in the note, each marked with whether it was mentioned
// as a fee or only as a penalty or an advance. The note is cut into clauses
// first, so a qualifier in one clause does not taint the months in another.
function noteMonthMentions_(note) {
  const t = (note || '').toString().toLowerCase();
  if (!t) return [];
  const out = [];
  t.split(/[,&+;\/]|\band\b/).forEach(function (clause) {
    const qualified = NOTE_NOT_A_FEE.test(clause);
    clause.split(/[^a-z]+/).forEach(function (word) {
      if (word.length < 3) return;
      const mi = monthIndexLoose_(word);
      if (mi >= 0) out.push({ mi: mi, qualified: qualified });
    });
  });
  return out;
}

// Months the note offers as actually paid for, and months it mentions only as
// a late fee or an advance — which are reported, never credited.
function noteMonthsAll_(note) {
  const seen = {}, out = [];
  noteMonthMentions_(note).forEach(function (m) {
    if (m.qualified || seen[m.mi]) return;
    seen[m.mi] = true;
    out.push(m.mi);
  });
  return out;
}

function noteMonthsQualifiedOnly_(note) {
  const paid = {}, out = [], seen = {};
  const all = noteMonthMentions_(note);
  all.forEach(function (m) { if (!m.qualified) paid[m.mi] = true; });
  all.forEach(function (m) {
    if (!m.qualified || paid[m.mi] || seen[m.mi]) return;
    seen[m.mi] = true;
    out.push(m.mi);
  });
  return out;
}

// "2 months", "two months fee", "both months" — the note says there is more
// than one month even when it does not say which.
function noteSaysSeveral_(note) {
  const t = (note || '').toString().toLowerCase();
  return /\b(2|3|two|three|double|both)\s*(-|\s)?\s*(month|mth|mnth)/.test(t) ||
         /\bmonths\b/.test(t);
}

// A month named in a note carries no year. Pick the year that puts it nearest
// the month the receipt is filed under, which is what makes a December receipt
// noting "December and January" land in the following January.
function noteYearFor_(mi, primaryPeriod, explicitYear) {
  if (explicitYear) return explicitYear * 12 + mi;
  const base = Math.floor(primaryPeriod / 12);
  let best = -1, bestGap = 999;
  [base - 1, base, base + 1].forEach(function (y) {
    if (y <= 0) return;
    const gap = Math.abs((y * 12 + mi) - primaryPeriod);
    if (gap < bestGap) { bestGap = gap; best = y * 12 + mi; }
  });
  return best;
}

// Which months one receipt actually covers. The Fee Month is always one of
// them. Extra months come from the note, and only when it names two or more —
// a note naming a single different month is a mis-keyed Fee Month, which
// findDuplicateReceipts already reports, and crediting both would be wrong.
function receiptPeriods_(primaryPeriod, note) {
  const named = noteMonthsAll_(note);
  const result = { periods: [], multi: false, vague: false, ignored: [] };
  if (primaryPeriod >= 0) result.periods.push(primaryPeriod);

  const explicit = (note || '').toString().match(/\b(20\d{2})\b/);
  const year = explicit ? parseInt(explicit[1], 10) : 0;
  const near = function (mi) {
    const p = noteYearFor_(mi, primaryPeriod, year);
    // Six months either side. Further off is prose, not a fee period.
    return (p < 0 || Math.abs(p - primaryPeriod) > 6) ? -1 : p;
  };

  if (primaryPeriod >= 0) {
    noteMonthsQualifiedOnly_(note).forEach(function (mi) {
      const p = near(mi);
      if (p >= 0 && p !== primaryPeriod) result.ignored.push(p);
    });
  }

  if (named.length >= 2 && primaryPeriod >= 0) {
    named.forEach(function (mi) {
      const p = near(mi);
      if (p >= 0 && result.periods.indexOf(p) < 0) result.periods.push(p);
    });
    result.multi = result.periods.length > 1;
  } else if (named.length < 2 && noteSaysSeveral_(note) && primaryPeriod >= 0) {
    // "2 months fee" without saying which two. Guessing the second month would
    // be inventing a payment, so it is reported instead.
    result.vague = true;
  }
  result.periods.sort(function (a, b) { return a - b; });
  return result;
}


// ─── The model every report below reads from ──────────────────

function buildFeeCoverage_() {
  const norm = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const eData = getSheet('Enrollments').getDataRange().getValues();
  const rData = getSheet('Receipts').getDataRange().getValues();
  const eHead = eData[0].map(norm);

  const eName = eHead.indexOf('Student Name');
  const eId   = eHead.indexOf('ID');
  const eType = eHead.indexOf('Type');
  const eStat = eHead.indexOf('Status');
  const ePh   = eHead.indexOf('Phone');
  const eLoc  = eHead.indexOf('Location');
  const eWhen = eHead.indexOf('Enrolled At');
  const eJoin = headerIndex_(eHead, 'Joining Date');
  const eLeft = headerIndex_(eHead, 'Left On');

  const warnings = [];

  // ── The roster: the golden source ──
  const students = [], byKey = {};
  for (let r = 1; r < eData.length; r++) {
    const raw = norm(eData[r][eName]);
    if (!raw) continue;
    const s = {
      row: r + 1,
      id: eId >= 0 ? norm(eData[r][eId]) : '',
      name: raw,
      key: normName_(raw),
      type: eType >= 0 ? norm(eData[r][eType]) : '',
      left: isLeftWord_(eStat >= 0 ? norm(eData[r][eStat]) : ''),
      phone: ePh >= 0 ? digitsOf_(eData[r][ePh]) : '',
      centre: eLoc >= 0 ? norm(eData[r][eLoc]).split('–')[0].trim() : '',
      joinRaw: eJoin >= 0 ? norm(eData[r][eJoin]) : '',
      leftRaw: eLeft >= 0 ? norm(eData[r][eLeft]) : '',
      enrolled: eWhen >= 0 ? parsePeriod_(eData[r][eWhen]) : -1,
      paid: {}, receipts: 0, firstPaid: -1, lastPaid: -1,
      start: -1, end: -1, billable: false
    };
    students.push(s);
    (byKey[s.key] = byKey[s.key] || []).push(s);
  }

  // A provisional enrolled-from/enrolled-to window, needed while the receipts
  // are still being read: when two children share a name and the receipt has
  // no contact number, the month it covers is the only thing that can tell
  // them apart. Pass 2 refines these once the receipts are known.
  // A joining date more than twenty years old is a typo, not a fact.
  const nowP = periodNow_();
  students.forEach(function (s) {
    let start = parsePeriod_(s.joinRaw);
    if (start < 0 || start < nowP - 240) start = s.enrolled;
    if (start < 0 || start < nowP - 240) start = -1;
    s.start = start;
    s.end = s.left ? parsePeriod_(s.leftRaw) : nowP;
    if (s.end < 0) s.end = nowP;
  });

  // ── Configured equivalent spellings ──
  const aliasMap = {}, config = { equivalents: [], distinct: [] };

  NAME_EQUIVALENTS.forEach(function (group) {
    const present = group.filter(function (n) { return byKey[normName_(n)]; });
    const entry = { group: group, present: present, status: '' };
    if (present.length === 1) {
      const target = normName_(present[0]);
      group.forEach(function (n) {
        const k = normName_(n);
        if (k !== target) aliasMap[k] = target;
      });
      entry.status = 'receipts credited to "' + byKey[target][0].name +
                     '" (row ' + byKey[target][0].row + ')';
    } else if (present.length === 0) {
      entry.status = 'NOT ON THE ROSTER - no spelling of this name is in Enrollments';
      warnings.push('Equivalent group [' + group.join(' = ') + '] matches nobody in Enrollments.');
    } else {
      entry.status = 'BOTH SPELLINGS ARE SEPARATE ROWS - merge them in Enrollments first';
      warnings.push('Equivalent group [' + group.join(' = ') + '] exists twice on the roster, so ' +
                    'receipts are being credited to two different rows.');
    }
    config.equivalents.push(entry);
  });

  // Each pinned name is bound to its own roster row so the spelling-tolerant
  // matcher can never hand one sibling's receipt to the other.
  const pinnedTo = {}, pinGroup = {};
  DISTINCT_STUDENTS.forEach(function (group, gi) {
    const entry = { group: group, bound: [], status: '' };
    group.forEach(function (n) {
      const k = normName_(n);
      const hit = byKey[k];
      if (hit && hit.length === 1) {
        pinnedTo[k] = hit[0].key;
        pinGroup[hit[0].key] = 'pin' + gi;
        entry.bound.push('"' + n + '" -> row ' + hit[0].row + '  "' + hit[0].name + '"');
      } else if (hit) {
        entry.bound.push('"' + n + '" -> AMBIGUOUS, ' + hit.length + ' rows share this name');
        warnings.push('"' + n + '" appears ' + hit.length + ' times in Enrollments.');
      } else {
        entry.bound.push('"' + n + '" -> NOT FOUND in Enrollments');
        warnings.push('"' + n + '" is configured as a distinct student but is not on the roster. ' +
                      'Check how Enrollments spells it and update DISTINCT_STUDENTS.');
      }
    });
    entry.status = entry.bound.join(' | ');
    config.distinct.push(entry);
  });

  const rosterKeys = Object.keys(byKey);

  // Receipts are often written with the first name only — 'Hiral', 'Priana'.
  // That is enough when exactly one child on the roster answers to it.
  const byFirst = {};
  students.forEach(function (s) {
    const w = s.key.split(' ')[0];
    if (w) (byFirst[w] = byFirst[w] || []).push(s);
  });
  const onlyChildNamed_ = function (word) {
    const hits = byFirst[word];
    if (!hits) return null;
    if (hits.length === 1) return hits[0];
    const here = hits.filter(function (s) { return !s.left; });
    return here.length === 1 ? here[0] : null;
  };

  // ── Resolve one printed name to one roster row ──
  const cache = {};

  // Exact, then the alias lists, then spelling tolerance. No decomposition:
  // the sibling reader below calls this on each piece it tries.
  const resolveSimple = function (fragment, contact, period) {
    const key0 = normName_(fragment);
    if (!key0) return { how: 'blank', student: null };

    const key = pinnedTo[key0] || aliasMap[key0] || key0;
    const how = (pinnedTo[key0] && pinnedTo[key0] !== key0) ? 'pinned'
              : (aliasMap[key0] ? 'alias' : 'exact');

    const hits = byKey[key];
    if (hits && hits.length === 1) {
      return { how: how, student: hits[0], from: fragment };
    }
    if (hits && hits.length > 1) {
      // Two children genuinely share this name. The contact number on the
      // receipt tells them apart; failing that, only one of them may be
      // enrolled in the month the receipt covers.
      return tellApart_(hits, contact, period, fragment, true, {
        near: hits[0].name, d: 0,
        why: hits.length + ' rows in Enrollments carry this exact name (rows ' +
             hits.map(function (s) { return s.row; }).join(', ') + ')'
      });
    }

    const fuzzy = fuzzyResolve_(key, byKey, rosterKeys, pinGroup, fragment);
    // A near-miss sitting the same distance from two children — 'Janvi Jain'
    // between Jahnvi and Janhvi. A contact number is a hard identifier and
    // settles it. The month is not: the name matches neither child exactly, so
    // picking whichever happens to be enrolled would be a guess.
    if (fuzzy.how === 'ambiguous' && fuzzy.candidates && fuzzy.candidates.length > 1) {
      return tellApart_(fuzzy.candidates, contact, period, fragment, false,
                        { near: fuzzy.near, d: fuzzy.d, why: fuzzy.why });
    }
    return fuzzy;
  };

  // Several children answer to one printed name. Settle it on evidence, or
  // report it — never pick one for the sake of a tidier number.
  function tellApart_(hits, contact, period, fragment, allowMonth, ctx) {
    const byPhone = contact
      ? hits.filter(function (s) { return s.phone && s.phone === contact; }) : [];
    if (byPhone.length === 1) {
      return { how: 'by phone', student: byPhone[0], from: fragment };
    }
    if (allowMonth && period >= 0) {
      const live = hits.filter(function (s) {
        return s.start >= 0 && period >= s.start && period <= s.end;
      });
      if (live.length === 1) return { how: 'by month', student: live[0], from: fragment };
    }
    return { how: 'ambiguous', student: null, from: fragment,
             near: ctx.near, d: ctx.d, why: ctx.why, candidates: hits };
  }

  // Several children written into one name with no separator at all:
  // 'Anshika Anvika Shome' is two sisters, 'Jeena Jia' is two more with the
  // surname left off entirely. Only accepted when it yields two or more
  // different children, which a single mangled name never does.
  const readSiblings_ = function (fragment, contact, period) {
    const words = normName_(fragment).split(' ').filter(Boolean);
    if (words.length < 2) return null;

    const tryAll = function (names) {
      const out = [], used = {};
      for (let i = 0; i < names.length; i++) {
        const r = resolveSimple(names[i], contact, period);
        if (!r.student || used[r.student.row]) return null;   // all or nothing
        used[r.student.row] = true;
        out.push({ how: 'split', student: r.student, from: names[i], via: fragment });
      }
      return out.length >= 2 ? out : null;
    };

    // The last word is the surname they share.
    const surname = words[words.length - 1];
    let out = tryAll(words.slice(0, -1).map(function (w) { return w + ' ' + surname; }));
    if (out) return out;

    // No surname anywhere: name one child from their first name, then lend
    // that child's surname to the others.
    for (let i = 0; i < words.length; i++) {
      const anchor = onlyChildNamed_(words[i]);
      if (!anchor) continue;
      const sur = anchor.key.split(' ').slice(-1)[0];
      out = tryAll(words.map(function (w, j) { return j === i ? anchor.key : w + ' ' + sur; }));
      if (out) return out;
    }
    return null;
  };

  // The full ladder for one printed name. Returns an array, because one
  // mangled name can turn out to be two children.
  const resolve = function (fragment, contact, period) {
    const key0 = normName_(fragment);
    if (!key0) return [{ how: 'blank', student: null, from: fragment }];
    const cacheable = !contact;
    if (cacheable && cache[key0]) return cache[key0];

    let out;
    const first = resolveSimple(fragment, contact, period);
    if (first.student) {
      out = [first];
    } else {
      const siblings = readSiblings_(fragment, contact, period);
      if (siblings) {
        out = siblings;
      } else if (key0.indexOf(' ') < 0) {
        const only = onlyChildNamed_(key0);
        if (only) {
          out = [{ how: 'first name', student: only, from: fragment }];
        } else {
          // Say so when a bare first name fails because it is not unique —
          // that is a different problem from a name nobody recognises.
          const share = byFirst[key0];
          if (share && share.length > 1) {
            first.why = share.length + ' children answer to the first name "' + key0 +
                        '" (' + share.map(function (s) {
                          return s.name + ', row ' + s.row + (s.left ? ', Left' : '');
                        }).join('; ') + ') - the receipt needs a surname';
          }
          out = [first];
        }
      } else {
        out = [first];
      }
    }
    if (cacheable) cache[key0] = out;
    return out;
  };

  // ── Pass 1: attribute every receipt ──
  const rHead = rData.length ? rData[0].map(norm) : [];
  const rNo   = rHead.indexOf('Receipt No');
  const rName = rHead.indexOf('Student Name');
  const rStu  = rHead.indexOf('Students');
  const rAmt  = headerIndex_(rHead, 'Amount');
  const rMon  = rHead.indexOf('Fee Month');
  const rYr   = rHead.indexOf('Fee Year');
  const rType = rHead.indexOf('Fee Type');
  const rNote = rHead.indexOf('Note');
  const rCon  = rHead.indexOf('Contact');
  const rRecd = rHead.indexOf('Date Received');
  const rIss  = rHead.indexOf('Issued At');

  const receipts = [], unresolved = [], ambiguous = [], resolvedNames = {};
  const periodStats = {};
  let monthlyCount = 0, otherTypeCount = 0, noPeriod = 0, periodGuessed = 0;
  let multiMonth = 0, vagueMonth = 0;

  for (let i = 1; i < rData.length; i++) {
    const row = rData[i];
    if (rNo >= 0 && !norm(row[rNo])) continue;

    // Registration, workshop and one-off fees are money, but they are not a
    // month's tuition and must not make a month look covered.
    const feeType = rType >= 0 ? norm(row[rType]) : '';
    if (feeType && normName_(feeType).indexOf('monthly') !== 0) { otherTypeCount++; continue; }
    monthlyCount++;

    // Fee Month/Fee Year is the truth. Where it is blank a note saying "fee
    // for August" is the next best thing, and the date received the last
    // resort — both are counted and flagged rather than trusted silently.
    let period = -1, periodSource = 'fee month';
    const mi = monthIndexLoose_(rMon >= 0 ? norm(row[rMon]) : '');
    const yr = parseInt(rYr >= 0 ? norm(row[rYr]) : '', 10);
    if (mi >= 0 && yr) {
      period = yr * 12 + mi;
    } else {
      const fallback = parsePeriod_(rRecd >= 0 && norm(row[rRecd]) ? row[rRecd]
                                    : (rIss >= 0 ? row[rIss] : ''));
      const fallbackYear = yr || (fallback >= 0 ? Math.floor(fallback / 12) : 0);
      const noted = monthIndexLoose_(noteMonth_(rNote >= 0 ? norm(row[rNote]) : ''));
      if (mi >= 0 && fallbackYear) {
        period = fallbackYear * 12 + mi; periodSource = 'year from date'; periodGuessed++;
      } else if (noted >= 0 && fallbackYear) {
        period = fallbackYear * 12 + noted; periodSource = 'note'; periodGuessed++;
      } else if (fallback >= 0) {
        period = fallback; periodSource = 'date received'; periodGuessed++;
      } else {
        noPeriod++; periodSource = 'none';
      }
    }

    // One receipt can settle two months at once. The note is the only place
    // that is recorded.
    const noteText = rNote >= 0 ? norm(row[rNote]) : '';
    const span = receiptPeriods_(period, noteText);
    if (span.multi) multiMonth++;
    if (span.vague) vagueMonth++;

    const contact = rCon >= 0 ? digitsOf_(row[rCon]) : '';
    const amount  = parseFloat(norm(rAmt >= 0 ? row[rAmt] : '').replace(/[^0-9.]/g, '')) || 0;
    const printed = rStu >= 0 ? norm(row[rStu]) : '';
    const fragments = splitReceiptNames_(printed, rName >= 0 ? norm(row[rName]) : '');

    const rec = {
      row: i + 1,
      no: rNo >= 0 ? norm(row[rNo]) : '(no number)',
      printed: printed || (rName >= 0 ? norm(row[rName]) : ''),
      period: period, periodSource: periodSource,
      periods: span.periods, multi: span.multi, vague: span.vague,
      ignored: span.ignored, note: noteText,
      amount: amount, contact: contact,
      matched: [], unmatched: []
    };

    const seen = {}, resolutions = [];
    fragments.forEach(function (f) {
      const label = normName_(f);
      if (!label || seen[label]) return;      // "Riya Sen | Riya Sen" is one child
      seen[label] = true;
      resolve(f, contact, period).forEach(function (res) {
        resolutions.push({ label: label, res: res });
      });
    });

    // Siblings have similar names — 'Riya Sen' is one character from 'Diya
    // Sen'. A derived match must never be credited to a child another name on
    // the same receipt already matched outright; that is a missing roster row,
    // not a typo. The firmest evidence claims its student first.
    const RANK = { exact: 0, alias: 0, pinned: 0, 'by phone': 0, 'by month': 0,
                   'first name': 1, split: 2, spelling: 3 };
    const rankOf = function (x) {
      const r = RANK[x.res.how];
      return r === undefined ? 9 : r;
    };
    resolutions.sort(function (a, b) { return rankOf(a) - rankOf(b); });

    const taken = {};
    resolutions.forEach(function (x) {
      let res = x.res;
      if (res.student && taken[res.student.row] && rankOf(x) > 0) {
        res = { how: 'unmatched', student: null, from: res.from, near: res.student.name,
                d: res.d, clash: true };
      } else if (res.student) {
        taken[res.student.row] = true;
      }

      const tally = resolvedNames[x.label] || { n: 0, parts: [] };
      tally.n++;
      tally.how = res.how;
      tally.to = res.student ? res.student.name : '';
      tally.row = res.student ? res.student.row : '';
      tally.near = res.near || '';
      tally.d = res.d;
      tally.why = res.why || tally.why;
      if (res.clash) tally.clash = true;
      // One printed name that turned out to be several children.
      if (res.how === 'split' && res.student &&
          tally.parts.indexOf(res.student.name) < 0) tally.parts.push(res.student.name);
      resolvedNames[x.label] = tally;

      if (res.student) {
        rec.matched.push(res);
      } else {
        rec.unmatched.push({ text: res.from, how: res.how, near: res.near, d: res.d,
                             clash: res.clash });
        const entry = { row: rec.row, no: rec.no, text: res.from, printed: rec.printed,
                        period: period, how: res.how, near: res.near, d: res.d,
                        why: res.why, clash: res.clash };
        if (res.how === 'ambiguous') ambiguous.push(entry); else unresolved.push(entry);
      }
    });

    receipts.push(rec);

    if (span.periods.length) {
      // The receipt itself, and the money on it, belong to the month it is
      // filed under — that is when the cash came in. Coverage is different:
      // a receipt settling two months covers the student in both.
      if (period >= 0) {
        const st = periodStats[period] = periodStats[period] ||
          { receipts: 0, amount: 0, unmatched: 0, ambiguous: 0, spanning: 0 };
        st.receipts++;
        st.amount += amount;
        if (span.multi) st.spanning++;
        rec.unmatched.forEach(function (u) {
          if (u.how === 'ambiguous') st.ambiguous++; else st.unmatched++;
        });
      }
      rec.matched.forEach(function (m) {
        const s = m.student;
        s.receipts++;
        span.periods.forEach(function (p) {
          s.paid[p] = (s.paid[p] || 0) + 1;
          if (s.firstPaid < 0 || p < s.firstPaid) s.firstPaid = p;
          if (p > s.lastPaid) s.lastPaid = p;
        });
      });
    } else {
      rec.matched.forEach(function (m) { m.student.receipts++; });
    }
  }

  // ── Pass 2: who is expected to pay, and for which months ──
  // A workshop attendee or an application nobody acted on is not a monthly
  // student — unless the receipts say otherwise, in which case they win.
  const now = periodNow_();
  let firstPeriod = now, lastPeriod = now;
  const gaps = { noJoining: [], noLeftOn: [] };

  students.forEach(function (s) {
    const t = normName_(s.type);
    const casual = t.indexOf('workshop') >= 0 || t.indexOf('application') === 0;
    s.billable = !casual || s.receipts > 0;

    const joinP = parsePeriod_(s.joinRaw);
    let start = (joinP >= 0 && joinP >= now - 240) ? joinP : -1;
    if (start < 0 && s.enrolled >= now - 240) start = s.enrolled;
    if (s.firstPaid >= 0 && (start < 0 || s.firstPaid < start)) start = s.firstPaid;
    s.startKnown = joinP >= 0 && joinP >= now - 240;
    s.start = start;
    if (s.billable && !s.startKnown) gaps.noJoining.push(s);

    let end = now;
    if (s.left) {
      const leftP = parsePeriod_(s.leftRaw);
      s.endKnown = leftP >= 0;
      end = leftP >= 0 ? leftP : (s.lastPaid >= 0 ? s.lastPaid : start);
      if (!s.endKnown && s.billable) gaps.noLeftOn.push(s);
    } else {
      s.endKnown = true;
    }
    if (s.lastPaid > end) end = s.lastPaid;
    s.end = end;

    if (s.billable && s.start >= 0 && s.start < firstPeriod) firstPeriod = s.start;
    if (s.billable && s.end > lastPeriod) lastPeriod = s.end;
  });
  Object.keys(periodStats).forEach(function (p) {
    const n = parseInt(p, 10);
    if (n < firstPeriod) firstPeriod = n;
    if (n > lastPeriod) lastPeriod = n;
  });
  // One receipt filed under the year 1900 must not stretch the table by a
  // thousand rows.
  if (firstPeriod < now - 240) firstPeriod = now - 240;
  if (lastPeriod > now + 12) lastPeriod = now + 12;

  // ── Pass 3: month by month ──
  const months = [];
  for (let p = firstPeriod; p <= lastPeriod; p++) {
    const expected = [], paid = [], unpaid = [], outside = [];
    students.forEach(function (s) {
      if (!s.billable || s.start < 0) return;
      const due = p >= s.start && p <= s.end;
      const has = !!s.paid[p];
      if (due) { expected.push(s); (has ? paid : unpaid).push(s); }
      else if (has) outside.push(s);
    });
    const st = periodStats[p] || { receipts: 0, amount: 0, unmatched: 0, ambiguous: 0, spanning: 0 };
    months.push({
      period: p, expected: expected, paid: paid, unpaid: unpaid, outside: outside,
      receipts: st.receipts, amount: st.amount,
      unmatchedNames: st.unmatched, ambiguousNames: st.ambiguous,
      complete: expected.length > 0 && unpaid.length === 0 &&
                st.unmatched === 0 && st.ambiguous === 0
    });
  }

  return {
    students: students, byKey: byKey, months: months, receipts: receipts,
    unresolved: unresolved, ambiguousReceipts: ambiguous, resolvedNames: resolvedNames,
    config: config, warnings: warnings, gaps: gaps,
    firstPeriod: firstPeriod, lastPeriod: lastPeriod, now: now,
    counts: {
      monthly: monthlyCount, otherType: otherTypeCount,
      noPeriod: noPeriod, periodGuessed: periodGuessed,
      multiMonth: multiMonth, vagueMonth: vagueMonth,
      roster: students.length,
      billable: students.filter(function (s) { return s.billable; }).length,
      active: students.filter(function (s) { return s.billable && !s.left; }).length
    }
  };
}

// A receipt name with no exact roster match. Accept a near-miss only when it
// is unambiguous, and never across a configured distinct pair.
function fuzzyResolve_(key, byKey, rosterKeys, pinGroup, fragment) {
  let best = null, bestD = 99, second = 99;
  for (let i = 0; i < rosterKeys.length; i++) {
    const d = levenshtein_(key, rosterKeys[i]);
    if (d < bestD) { second = bestD; bestD = d; best = rosterKeys[i]; }
    else if (d < second) { second = d; }
  }
  if (!best) return { how: 'unmatched', student: null, from: fragment };

  const near = byKey[best][0].name;
  if (bestD > nameTolerance_(key)) {
    return { how: 'unmatched', student: null, from: fragment, near: near, d: bestD };
  }
  // Equally close to two different children, or to a name two children share.
  if (second <= bestD) {
    // Name them: which two roster rows are tied is the whole decision.
    const tied = [], candidates = [];
    for (let i = 0; i < rosterKeys.length; i++) {
      if (levenshtein_(key, rosterKeys[i]) === bestD) {
        byKey[rosterKeys[i]].forEach(function (s) {
          candidates.push(s);
          tied.push('"' + s.name + '" (row ' + s.row + ')');
        });
      }
    }
    return { how: 'ambiguous', student: null, from: fragment, near: near, d: bestD,
             candidates: candidates,
             why: 'equally close (' + bestD + ' char) to ' + tied.join(' and ') +
                  ' - either would be a guess' };
  }
  if (byKey[best].length > 1) {
    return { how: 'ambiguous', student: null, from: fragment, near: near, d: bestD,
             why: byKey[best].length + ' rows in Enrollments carry the name "' + near +
                  '" (rows ' + byKey[best].map(function (s) { return s.row; }).join(', ') + ')' };
  }
  // A near-miss onto one of a pinned pair needs clear daylight from the other.
  if (pinGroup[best]) {
    for (let i = 0; i < rosterKeys.length; i++) {
      if (rosterKeys[i] === best || pinGroup[rosterKeys[i]] !== pinGroup[best]) continue;
      if (levenshtein_(key, rosterKeys[i]) <= bestD + 1) {
        return { how: 'ambiguous', student: null, from: fragment, near: near, d: bestD };
      }
    }
  }
  return { how: 'spelling', student: byKey[best][0], from: fragment, d: bestD };
}


// ─── Report 1: the headline table ─────────────────────────────

function feeCoverageByMonth() {
  const cov = buildFeeCoverage_();
  let out = 'FEE COVERAGE BY MONTH\n=====================\n';
  out += 'Generated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy') + '\n';
  out += 'Enrollments is the golden source. A receipt covering siblings counts\n';
  out += 'once for each child named on it.\n\n';

  out += 'ROSTER\n';
  out += '  Rows in Enrollments             ' + cov.counts.roster + '\n';
  out += '  Monthly-fee students            ' + cov.counts.billable +
         '   (workshops and untouched applications excluded)\n';
  out += '  ...of them active today         ' + cov.counts.active + '\n\n';

  out += 'RECEIPTS\n';
  out += '  Monthly-fee receipts read       ' + cov.counts.monthly + '\n';
  out += '  Other fee types skipped         ' + cov.counts.otherType + '\n';
  out += '  Month inferred, not stated      ' + cov.counts.periodGuessed + '\n';
  out += '  Covering more than one month    ' + cov.counts.multiMonth +
         '   (from the note)\n';
  if (cov.counts.vagueMonth) {
    out += '  Says several months, not which  ' + cov.counts.vagueMonth +
           '   <- see previewMultiMonthReceipts()\n';
  }
  out += '  No month at all - not counted   ' + cov.counts.noPeriod + '\n';
  out += '  Names no roster row matches     ' + cov.unresolved.length + '\n';
  out += '  Names too close to call         ' + cov.ambiguousReceipts.length + '\n\n';

  out += pad_('MONTH', 10) + pad_('DUE', 6) + pad_('PAID', 6) + pad_('UNPAID', 8) +
         pad_('COVER', 8) + pad_('RCPTS', 7) + pad_('COLLECTED', 14) + 'FLAGS\n';
  out += '-------------------------------------------------------------------------\n';

  const shown = cov.months.slice(-COVERAGE_MONTHS);
  shown.forEach(function (m) {
    const pct = m.expected.length ? Math.round(m.paid.length * 1000 / m.expected.length) / 10 : 0;
    let flags = '';
    if (m.complete) flags = 'ALL PAID';
    if (m.unmatchedNames) flags += (flags ? ' ' : '') + m.unmatchedNames + ' unknown name(s)';
    if (m.ambiguousNames) flags += (flags ? ' ' : '') + m.ambiguousNames + ' unclear name(s)';
    if (m.period === cov.now) flags += (flags ? ' ' : '') + '(current month)';
    out += pad_(periodLabel_(m.period), 10) + pad_(m.expected.length, 6) +
           pad_(m.paid.length, 6) + pad_(m.unpaid.length, 8) +
           pad_(pct + '%', 8) + pad_(m.receipts, 7) + pad_(money_(m.amount), 14) + flags + '\n';
  });
  out += '\n';

  // The current month is usually still being collected, so the month before it
  // is the first one that ought to be complete.
  const closed = cov.months.filter(function (m) { return m.period < cov.now && m.expected.length; });
  const last = closed.length ? closed[closed.length - 1] : null;
  if (last) {
    out += 'LATEST CLOSED MONTH - ' + periodLong_(last.period) + '\n';
    if (last.complete) {
      out += '  Every student due that month has a receipt. Nothing outstanding.\n\n';
    } else {
      out += '  ' + last.unpaid.length + ' of ' + last.expected.length + ' students have no receipt';
      if (last.unmatchedNames || last.ambiguousNames) {
        out += ', and ' + (last.unmatchedNames + last.ambiguousNames) +
               ' name(s) on receipts could not be tied to the roster';
      }
      out += '.\n';
      out += '  Run feeCoverageForMonth(\'' + periodLong_(last.period) + '\') for the names.\n\n';
    }
  }

  const clean = shown.filter(function (m) { return m.complete; }).length;
  out += 'Fully covered months in this window : ' + clean + ' of ' + shown.length + '\n\n';

  if (cov.warnings.length) {
    out += 'CONFIGURATION WARNINGS\n';
    cov.warnings.forEach(function (w) { out += '  ! ' + w + '\n'; });
    out += '\n';
  }

  out += 'Report only. Nothing was changed.\n';
  out += 'Next: previewReceiptNameMatching() to check how names were resolved,\n';
  out += 'feeCoverageForMonth(\'August 2026\') for one month in full, and\n';
  out += 'feeGapsByStudent() for who is behind.\n';
  Logger.log(out);
  return out;
}


// ─── Report 2: one month in full ──────────────────────────────

function feeCoverageForMonth(monthText) {
  const cov = buildFeeCoverage_();
  let target = parsePeriod_(monthText);
  if (target < 0) {
    // No argument given: the latest month any receipt is booked against.
    for (let i = cov.months.length - 1; i >= 0; i--) {
      if (cov.months[i].receipts) { target = cov.months[i].period; break; }
    }
    if (target < 0) target = cov.now;
  }
  const m = cov.months.filter(function (x) { return x.period === target; })[0];

  let out = 'FEE COVERAGE - ' + periodLong_(target) + '\n';
  out += '=====================================\n';
  if (!m) {
    out += 'That month is outside the range the data covers (' +
           periodLong_(cov.firstPeriod) + ' to ' + periodLong_(cov.lastPeriod) + ').\n';
    Logger.log(out);
    return out;
  }

  const pct = m.expected.length ? Math.round(m.paid.length * 1000 / m.expected.length) / 10 : 0;
  out += '  Students due to pay              ' + m.expected.length + '\n';
  out += '  Covered by a receipt             ' + m.paid.length + '  (' + pct + '%)\n';
  out += '  No receipt                       ' + m.unpaid.length + '\n';
  out += '  Receipts booked to this month    ' + m.receipts + '\n';
  out += '  Collected                        ' + money_(m.amount) + '\n';
  if (m.outside.length) {
    out += '  Paid but not due this month      ' + m.outside.length +
           '  (joined later or already left)\n';
  }
  out += '\n';

  if (m.complete) {
    out += 'ALL ' + m.expected.length + ' STUDENTS DUE IN ' +
           periodLong_(target).toUpperCase() + ' ARE PAID FOR.\n\n';
  }

  if (m.unpaid.length) {
    out += 'NO RECEIPT FOR ' + periodLong_(target).toUpperCase() + ' (' + m.unpaid.length + ')\n';
    out += '  ' + pad_('row', 6) + pad_('student', 28) + pad_('centre', 14) +
           pad_('phone', 13) + pad_('due since', 11) + 'last paid\n';
    m.unpaid.slice().sort(function (a, b) { return b.lastPaid - a.lastPaid; }).forEach(function (s) {
      out += '  ' + pad_(s.row, 6) + pad_(s.name, 28) + pad_(s.centre || '-', 14) +
             pad_(s.phone || '-', 13) + pad_(periodLabel_(s.start) + (s.startKnown ? '' : '?'), 11) +
             (s.lastPaid >= 0 ? periodLabel_(s.lastPaid) : 'never') + '\n';
    });
    out += '  A "?" after the joining month means Enrollments has no joining\n';
    out += '  month for that student and it was inferred.\n\n';
  }

  const flagged = cov.unresolved.concat(cov.ambiguousReceipts)
    .filter(function (u) { return u.period === target; });
  if (flagged.length) {
    out += 'RECEIPT NAMES THIS MONTH THAT NO ROSTER ROW ACCOUNTS FOR (' + flagged.length + ')\n';
    out += 'Until these are resolved the unpaid list above may be overstated.\n';
    flagged.forEach(function (u) {
      out += '  row ' + pad_(u.row, 6) + pad_(u.no, 14) + pad_('"' + u.text + '"', 26) +
             (u.near ? 'nearest roster name "' + u.near + '" (' + u.d + ' char off)'
                     : 'no close roster name') +
             (u.how === 'ambiguous' ? '  - too close to call' : '') +
             (u.clash ? '  - that row is already taken by a sibling on this receipt' : '') + '\n';
      if (u.printed && normName_(u.printed) !== normName_(u.text)) {
        out += '        receipt reads: ' + u.printed + '\n';
      }
    });
    out += '\n';
  }

  if (m.outside.length) {
    out += 'PAID BUT NOT DUE THIS MONTH\n';
    out += 'Either the payment is early or late, or their joining/leaving month\n';
    out += 'in Enrollments is wrong.\n';
    m.outside.forEach(function (s) {
      out += '  row ' + pad_(s.row, 6) + pad_(s.name, 28) +
             'due ' + pad_(periodLabel_(s.start) + ' - ' + periodLabel_(s.end), 20) +
             (s.left ? 'marked Left' : 'active') + '\n';
    });
    out += '\n';
  }

  out += 'PAID (' + m.paid.length + ')\n';
  m.paid.forEach(function (s) {
    out += '  ' + pad_(s.name, 30) +
           (s.paid[target] > 1 ? '(' + s.paid[target] + ' receipts - see findDuplicateReceipts)' : '') + '\n';
  });
  out += '\nReport only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Report 3: how every receipt name was resolved ────────────
// Run this before believing either of the reports above it.

function previewReceiptNameMatching() {
  const cov = buildFeeCoverage_();
  let out = 'RECEIPT NAME MATCHING - PREVIEW\n===============================\n';
  out += 'Every distinct name printed on a monthly-fee receipt, and the row in\n';
  out += 'Enrollments it was credited to. Writes nothing.\n\n';

  const names = Object.keys(cov.resolvedNames);
  const buckets = {};
  names.forEach(function (k) {
    const r = cov.resolvedNames[k];
    (buckets[r.how] = buckets[r.how] || []).push({ from: k, r: r });
  });
  const size = function (b) { return (buckets[b] || []).length; };
  const receiptsFor = function (list) {
    return list.reduce(function (n, x) { return n + x.r.n; }, 0);
  };

  out += 'Distinct names on receipts        ' + names.length + '\n';
  out += '  Matched the roster exactly      ' + size('exact') + '\n';
  out += '  Matched through the alias list  ' + (size('alias') + size('pinned')) + '\n';
  out += '  Matched on spelling tolerance   ' + size('spelling') + '\n';
  out += '  Read as several children        ' + size('split') + '\n';
  out += '  Matched on first name alone     ' + size('first name') + '\n';
  out += '  Split by phone or by month      ' + (size('by phone') + size('by month')) + '\n';
  out += '  TOO CLOSE TO CALL               ' + size('ambiguous') + '\n';
  out += '  NO ROSTER ROW AT ALL            ' + size('unmatched') + '\n\n';

  out += 'CONFIGURED EQUIVALENT SPELLINGS - one child, more than one spelling\n';
  if (!cov.config.equivalents.length) out += '  (none configured)\n';
  cov.config.equivalents.forEach(function (e) {
    out += '  ' + pad_(e.group.join(' = '), 34) + '  ' + e.status + '\n';
  });
  out += '\nCONFIGURED DISTINCT STUDENTS - similar names, different children\n';
  if (!cov.config.distinct.length) out += '  (none configured)\n';
  cov.config.distinct.forEach(function (e) {
    out += '  ' + e.group.join('  vs  ') + '\n';
    e.bound.forEach(function (b) { out += '      ' + b + '\n'; });
  });
  out += '\n';

  if (size('spelling')) {
    out += 'CREDITED DESPITE A SPELLING DIFFERENCE - check these\n';
    buckets.spelling.sort(function (a, b) { return a.r.d - b.r.d; }).forEach(function (x) {
      out += '  ' + pad_('"' + x.from + '"', 28) + '-> ' + pad_(x.r.to, 26) +
             '(' + x.r.d + ' char, ' + x.r.n + ' receipt(s), row ' + x.r.row + ')\n';
    });
    out += '\n';
  }
  if (size('alias') || size('pinned')) {
    out += 'CREDITED THROUGH THE ALIAS LIST\n';
    (buckets.alias || []).concat(buckets.pinned || []).forEach(function (x) {
      out += '  ' + pad_('"' + x.from + '"', 28) + '-> ' + pad_(x.r.to, 26) +
             '(' + x.r.n + ' receipt(s), row ' + x.r.row + ')\n';
    });
    out += '\n';
  }
  if (size('split')) {
    out += 'READ AS SEVERAL CHILDREN IN ONE NAME - check these\n';
    out += 'No separator was typed, so the name was tried as siblings sharing a\n';
    out += 'surname. Only accepted where every piece matched a different child.\n';
    const shownSplit = {};
    buckets.split.forEach(function (x) {
      if (shownSplit[x.from]) return;
      shownSplit[x.from] = true;
      out += '  ' + pad_('"' + x.from + '"', 28) + '-> ' + x.r.parts.join('  +  ') + '\n';
    });
    out += '\n';
  }
  if (size('first name')) {
    out += 'MATCHED ON FIRST NAME ALONE - check these\n';
    out += 'The receipt gave no surname, and exactly one child answers to it.\n';
    buckets['first name'].forEach(function (x) {
      out += '  ' + pad_('"' + x.from + '"', 28) + '-> ' + pad_(x.r.to, 26) +
             '(' + x.r.n + ' receipt(s), row ' + x.r.row + ')\n';
    });
    out += '\n';
  }
  if (size('by phone') || size('by month')) {
    out += 'TWO CHILDREN SHARE THIS NAME - told apart by contact number or month\n';
    (buckets['by phone'] || []).concat(buckets['by month'] || []).forEach(function (x) {
      out += '  ' + pad_('"' + x.from + '"', 28) + '-> ' + pad_(x.r.to, 26) +
             '(' + x.r.how + ', row ' + x.r.row + ')\n';
    });
    out += '\n';
  }
  if (size('ambiguous')) {
    out += 'TOO CLOSE TO CALL - credited to nobody, so a month may look short\n';
    out += 'Fix the spelling in Enrollments, or add the pair to NAME_EQUIVALENTS\n';
    out += 'or DISTINCT_STUDENTS at the top of this section.\n';
    buckets.ambiguous.forEach(function (x) {
      out += '  ' + pad_('"' + x.from + '"', 28) + 'nearest "' + (x.r.near || '?') + '"' +
             (x.r.d !== undefined ? ' (' + x.r.d + ' char)' : '') +
             '  ' + x.r.n + ' receipt(s)\n';
      if (x.r.why) out += '        why: ' + x.r.why + '\n';
    });
    out += '\n';
  }
  if (size('unmatched')) {
    out += 'NO ROSTER ROW AT ALL - ' + receiptsFor(buckets.unmatched) +
           ' receipt line(s) for someone not in Enrollments\n';
    out += 'Either the child is missing from the roster, or the receipt names a\n';
    out += 'guardian, or the spelling is too far off to match safely.\n';
    buckets.unmatched.forEach(function (x) {
      out += '  ' + pad_('"' + x.from + '"', 28) +
             (x.r.near ? 'nearest "' + x.r.near + '" (' + x.r.d + ' char off)' : 'nothing close') +
             '  ' + x.r.n + ' receipt(s)' +
             (x.r.clash ? '  - a sibling on the same receipt already matched that row' : '') + '\n';
      if (x.r.why) out += '        why: ' + x.r.why + '\n';
    });
    out += '\n';
  }

  if (cov.counts.noPeriod) {
    out += cov.counts.noPeriod + ' monthly receipt(s) name no fee month at all and are\n';
    out += 'counted towards nobody. Fill in Fee Month/Fee Year on those rows.\n\n';
  }
  if (cov.warnings.length) {
    out += 'CONFIGURATION WARNINGS\n';
    cov.warnings.forEach(function (w) { out += '  ! ' + w + '\n'; });
    out += '\n';
  }

  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Report 3b: receipts that settle more than one month ──────
// Anjali records these in the note, because the Fee Month column only holds
// one month. Every one of them is listed here so the reading can be checked.

function previewMultiMonthReceipts() {
  const cov = buildFeeCoverage_();
  const multi = cov.receipts.filter(function (r) { return r.multi; });
  const vague = cov.receipts.filter(function (r) { return r.vague; });

  let out = 'RECEIPTS COVERING MORE THAN ONE MONTH\n=====================================\n';
  out += 'The Fee Month column holds one month, so a receipt settling two is\n';
  out += 'recognised from its note. Both months then count as paid for every\n';
  out += 'child on the receipt.\n\n';
  out += '  Read as covering several months  ' + multi.length + '\n';
  out += '  Note says several, but not which ' + vague.length + '\n\n';

  if (multi.length) {
    out += 'READ AS SEVERAL MONTHS - check these\n';
    out += '  ' + pad_('row', 6) + pad_('receipt', 14) + pad_('filed under', 16) +
           pad_('counted for', 34) + 'note\n';
    multi.forEach(function (r) {
      out += '  ' + pad_(r.row, 6) + pad_(r.no, 14) + pad_(periodLong_(r.period), 16) +
             pad_(r.periods.map(periodLong_).join(' + '), 34) +
             (r.note || '').substring(0, 40) + '\n';
    });
    out += '\n';
  }

  if (vague.length) {
    out += 'THE NOTE SAYS SEVERAL MONTHS BUT DOES NOT NAME THEM\n';
    out += 'Only the Fee Month is counted for these. Guessing the other month\n';
    out += 'would be inventing a payment. Name the months in the note and they\n';
    out += 'will be picked up on the next run.\n';
    out += '  ' + pad_('row', 6) + pad_('receipt', 14) + pad_('filed under', 16) +
           pad_('amount', 12) + 'note\n';
    vague.forEach(function (r) {
      out += '  ' + pad_(r.row, 6) + pad_(r.no, 14) + pad_(periodLong_(r.period), 16) +
             pad_(money_(r.amount), 12) + (r.note || '').substring(0, 40) + '\n';
    });
    out += '\n';
  }

  const held = cov.receipts.filter(function (r) { return r.ignored && r.ignored.length; });
  if (held.length) {
    out += 'MONTHS MENTIONED BUT NOT COUNTED (' + held.length + ')\n';
    out += 'A late fee, a fine or an advance names a month without settling that\n';
    out += "month's tuition. \"Fee for June & late fee for May\" pays June and adds\n";
    out += 'a penalty for May — May itself may still be outstanding, and is left\n';
    out += 'alone here. If one of these really did settle the month, say so in the\n';
    out += 'note ("Fee for May & June") and it will be counted.\n';
    out += '  ' + pad_('row', 6) + pad_('receipt', 14) + pad_('counted for', 30) +
           pad_('not counted', 16) + 'note\n';
    held.forEach(function (r) {
      out += '  ' + pad_(r.row, 6) + pad_(r.no, 14) +
             pad_(r.periods.map(periodLong_).join(' + '), 30) +
             pad_(r.ignored.map(periodLong_).join(', '), 16) +
             (r.note || '').substring(0, 40) + '\n';
    });
    out += '\n';
  }

  out += 'Money is counted in the month the receipt is filed under, which is\n';
  out += 'when it was received. Only coverage spans both months.\n\n';
  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Report 4: who is behind, and by which months ─────────────

function feeGapsByStudent() {
  const cov = buildFeeCoverage_();
  const rows = [];
  cov.students.forEach(function (s) {
    if (!s.billable || s.start < 0) return;
    const missing = [];
    for (let p = s.start; p <= s.end; p++) if (!s.paid[p]) missing.push(p);
    if (missing.length) rows.push({ s: s, missing: missing });
  });
  rows.sort(function (a, b) { return b.missing.length - a.missing.length; });

  const neverPaid = rows.filter(function (r) { return r.s.lastPaid < 0; });
  const partial   = rows.filter(function (r) { return r.s.lastPaid >= 0; });

  let out = 'FEE GAPS BY STUDENT\n===================\n';
  out += 'Months between a student\'s joining month and today - or the month they\n';
  out += 'left - with no receipt naming them.\n\n';
  out += '  Monthly-fee students            ' + cov.counts.billable + '\n';
  out += '  Paid for every month due        ' + (cov.counts.billable - rows.length) + '\n';
  out += '  Missing at least one month      ' + rows.length + '\n';
  out += '  Never had a monthly receipt     ' + neverPaid.length + '\n\n';

  if (neverPaid.length) {
    out += 'NEVER INVOICED (' + neverPaid.length + ')\n';
    out += '  ' + pad_('row', 6) + pad_('student', 28) + pad_('centre', 14) +
           pad_('phone', 13) + 'due since\n';
    neverPaid.forEach(function (r) {
      out += '  ' + pad_(r.s.row, 6) + pad_(r.s.name, 28) + pad_(r.s.centre || '-', 14) +
             pad_(r.s.phone || '-', 13) + pad_(periodLabel_(r.s.start), 10) +
             '(' + r.missing.length + ' months)\n';
    });
    out += '\n';
  }

  if (partial.length) {
    out += 'MISSING SOME MONTHS (' + partial.length + ')\n';
    out += '  ' + pad_('row', 6) + pad_('student', 28) + pad_('missing', 9) + 'months\n';
    partial.slice(0, 60).forEach(function (r) {
      const list = r.missing.slice(-8).map(periodLabel_).join(', ');
      out += '  ' + pad_(r.s.row, 6) + pad_(r.s.name, 28) + pad_(r.missing.length, 9) +
             (r.missing.length > 8 ? '...' + list : list) + '\n';
    });
    if (partial.length > 60) out += '  ... ' + (partial.length - 60) + ' more.\n';
    out += '\n';
  }

  if (cov.gaps.noJoining.length) {
    out += 'NO JOINING MONTH ON RECORD (' + cov.gaps.noJoining.length + ')\n';
    out += 'Their first expected month was inferred, so their gap count is a\n';
    out += 'guess too. previewJoiningBackfill() then backfillJoiningMonth().\n';
    cov.gaps.noJoining.slice(0, 20).forEach(function (s) {
      out += '  row ' + pad_(s.row, 6) + s.name + '\n';
    });
    if (cov.gaps.noJoining.length > 20) out += '  ... ' + (cov.gaps.noJoining.length - 20) + ' more.\n';
    out += '\n';
  }
  if (cov.gaps.noLeftOn.length) {
    out += 'MARKED LEFT WITH NO LAST FEE MONTH (' + cov.gaps.noLeftOn.length + ')\n';
    out += 'Their last paid month is standing in for it. Fill in Left On.\n';
    cov.gaps.noLeftOn.slice(0, 20).forEach(function (s) {
      out += '  row ' + pad_(s.row, 6) + s.name + '\n';
    });
    out += '\n';
  }

  out += 'Report only. Nothing was changed.\n';
  Logger.log(out);
  return out;
}


// ─── Analytics tabs in the sheet itself ───────────────────────
// The reports above print to the execution log, which only helps whoever is
// sitting in the Apps Script editor. This section lays the same numbers out as
// ordinary tabs Anjali can open, filter and read.
//
// The Dashboard is driven by spreadsheet formulas rather than by script, so
// changing the month in the picker updates it instantly — no refresh, no
// waiting for Apps Script. Only a change to the underlying Enrollments or
// Receipts data needs a refresh from the menu.
//
// These functions WRITE, which nothing else in the analytics section does.
// They write only to their own tabs: Enrollments, Receipts and Config are
// opened read-only and never touched. Every generated tab is stamped with a
// marker, and a refresh refuses to overwrite a tab that lacks it.

const ANALYTICS_MARK = 'SS-ANALYTICS';
const TAB_DASH   = 'Analytics Dashboard';
const TAB_COVER  = 'Analytics Coverage';
const TAB_STUDES = 'Analytics Students';
const TAB_NAMES  = 'Analytics Names';
const TAB_GAPS   = 'Analytics Gaps';
const ANALYTICS_TABS = [TAB_DASH, TAB_COVER, TAB_STUDES, TAB_NAMES, TAB_GAPS];

// Never let a refresh write over one of the real data tabs, whatever a tab
// name is changed to in future.
const PROTECTED_TABS = ['Enrollments', 'Receipts', 'Config'];

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Analytics')
      .addItem('Refresh dashboard', 'refreshAnalytics')
      .addItem('Preview refresh (writes nothing)', 'previewAnalyticsRefresh')
      .addSeparator()
      .addItem('Name matching report', 'showNameMatching')
      .addItem('Two-month receipts report', 'showMultiMonth')
      .addItem('Fee gaps report', 'showFeeGaps')
      .addToUi();
  } catch (e) {
    // No UI (a trigger, or the editor). The menu simply is not built.
  }
}

// Show a report in the sheet when there is a screen to show it on, and always
// write it to the execution log for the editor.
function report_(text) {
  Logger.log(text);
  try {
    const html = HtmlService
      .createHtmlOutput('<pre style="font:12px/1.45 monospace;white-space:pre-wrap">' +
                        text.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</pre>')
      .setWidth(900).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Sriram Studio analytics');
  } catch (e) { /* editor, no UI */ }
  return text;
}

function showNameMatching() { return report_(previewReceiptNameMatching()); }
function showMultiMonth()   { return report_(previewMultiMonthReceipts()); }
function showFeeGaps()      { return report_(feeGapsByStudent()); }


// ─── What the tabs contain ────────────────────────────────────
// Kept apart from the writing so the shape of every tab can be checked
// without a spreadsheet.

function analyticsTabModel_(cov) {
  const months = cov.months;
  const monthNames = months.map(function (m) { return periodLong_(m.period); });

  // ── Students: one row per monthly-fee student, one column per month ──
  const stuHead = ['Row', 'ID', 'Student', 'Centre', 'Phone', 'Status',
                   'Joining', 'Left On', 'Due From', 'Due To',
                   'Months Due', 'Months Paid', 'Months Missing', 'Last Paid']
                  .concat(monthNames);
  const stuRows = [];
  cov.students.forEach(function (s) {
    if (!s.billable || s.start < 0) return;
    let due = 0, paid = 0;
    const cells = months.map(function (m) {
      const isDue = m.period >= s.start && m.period <= s.end;
      const has = !!s.paid[m.period];
      if (isDue) { due++; if (has) { paid++; return 'PAID'; } return 'UNPAID'; }
      return has ? 'EXTRA' : '';
    });
    stuRows.push([s.row, s.id, s.name, s.centre || '', s.phone || '',
                  s.left ? 'Left' : 'Active', s.joinRaw || '', s.leftRaw || '',
                  periodLong_(s.start), periodLong_(s.end),
                  due, paid, due - paid,
                  s.lastPaid >= 0 ? periodLong_(s.lastPaid) : 'never'].concat(cells));
  });
  stuRows.sort(function (a, b) { return b[12] - a[12] || a[2].localeCompare(b[2]); });

  // ── Coverage: one row per month ──
  const covHead = ['Month', 'Due', 'Paid', 'Unpaid', 'Coverage', 'Receipts',
                   'Collected', 'Notes'];
  const covRows = months.map(function (m) {
    const notes = [];
    if (m.complete) notes.push('all paid');
    if (m.period === cov.now) notes.push('current month, still collecting');
    if (m.unmatchedNames) notes.push(m.unmatchedNames + ' receipt name(s) not on the roster');
    if (m.ambiguousNames) notes.push(m.ambiguousNames + ' receipt name(s) too close to call');
    return [periodLong_(m.period), m.expected.length, m.paid.length, m.unpaid.length,
            m.expected.length ? m.paid.length / m.expected.length : 0,
            m.receipts, m.amount, notes.join('; ')];
  });

  // ── Names: how every printed receipt name was resolved ──
  const nameHead = ['Name on receipt', 'How it was matched', 'Credited to',
                    'Enrollments row', 'Receipt lines', 'Detail'];
  const HOW_TEXT = {
    exact: 'exact match',
    alias: 'alias list',
    pinned: 'alias list (pinned)',
    spelling: 'spelling tolerance',
    split: 'read as several children',
    'first name': 'first name only',
    'by phone': 'contact number',
    'by month': 'only one enrolled that month',
    ambiguous: 'TOO CLOSE TO CALL - credited to nobody',
    unmatched: 'NO ROSTER ROW - credited to nobody'
  };
  const nameRows = Object.keys(cov.resolvedNames).map(function (k) {
    const r = cov.resolvedNames[k];
    let detail = r.why || '';
    if (!detail && r.how === 'spelling') detail = r.d + ' character(s) different';
    if (!detail && r.how === 'split' && r.parts) detail = r.parts.join(' + ');
    if (!detail && r.clash) detail = 'a sibling on the same receipt already matched that row';
    return [k, HOW_TEXT[r.how] || r.how, r.to || '', r.row || '', r.n, detail];
  });
  // The ones needing a decision first, then the ones worth a glance, then the
  // exact matches nobody needs to read.
  nameRows.sort(function (a, b) {
    const rank = function (row) {
      if (row[1].indexOf('TOO CLOSE') === 0 || row[1].indexOf('NO ROSTER') === 0) return 0;
      if (row[1] === 'spelling tolerance' || row[1] === 'read as several children' ||
          row[1] === 'first name only') return 1;
      return 2;
    };
    return rank(a) - rank(b) || a[0].localeCompare(b[0]);
  });

  // ── Gaps: students missing at least one month ──
  const gapHead = ['Row', 'Student', 'Centre', 'Phone', 'Status', 'Due Since',
                   'Months Missing', 'Last Paid', 'Which months'];
  const gapRows = [];
  cov.students.forEach(function (s) {
    if (!s.billable || s.start < 0) return;
    const missing = [];
    for (let p = s.start; p <= s.end; p++) if (!s.paid[p]) missing.push(periodLong_(p));
    if (!missing.length) return;
    gapRows.push([s.row, s.name, s.centre || '', s.phone || '',
                  s.left ? 'Left' : 'Active', periodLong_(s.start), missing.length,
                  s.lastPaid >= 0 ? periodLong_(s.lastPaid) : 'never',
                  missing.join(', ')]);
  });
  gapRows.sort(function (a, b) { return b[6] - a[6] || a[1].localeCompare(b[1]); });

  // The month the dashboard opens on: the last one that ought to be finished.
  let defaultMonth = '';
  for (let i = months.length - 1; i >= 0; i--) {
    if (months[i].period < cov.now && months[i].expected.length) {
      defaultMonth = periodLong_(months[i].period);
      break;
    }
  }
  if (!defaultMonth && months.length) defaultMonth = periodLong_(months[months.length - 1].period);

  return {
    students: { head: stuHead, rows: stuRows },
    coverage: { head: covHead, rows: covRows },
    names:    { head: nameHead, rows: nameRows },
    gaps:     { head: gapHead, rows: gapRows },
    monthNames: monthNames,
    defaultMonth: defaultMonth,
    counts: cov.counts,
    firstMonthCol: stuHead.length - monthNames.length + 1   // 1-based
  };
}


// ─── Writing the tabs ─────────────────────────────────────────

function analyticsSheet_(ss, name) {
  if (PROTECTED_TABS.indexOf(name) >= 0) {
    throw new Error('Refusing to write to "' + name + '" — that is a data tab.');
  }
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else if (sh.getLastRow() > 0 &&
             sh.getRange(1, 1).getNote().indexOf(ANALYTICS_MARK) !== 0) {
    throw new Error('A tab called "' + name + '" already exists and was not created by ' +
                    'this report. Rename or delete it, then refresh again. ' +
                    'Nothing has been changed.');
  }
  sh.clear();
  sh.clearConditionalFormatRules();
  const filter = sh.getFilter();
  if (filter) filter.remove();
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.getRange(1, 1).setNote(ANALYTICS_MARK + ' — generated tab. Safe to delete; the ' +
                            'Analytics menu rebuilds it. Do not type here, edits are lost ' +
                            'on the next refresh.');
  return sh;
}

function writeGrid_(sh, head, rows, startRow) {
  const r0 = startRow || 1;
  sh.getRange(r0, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#2C1A0E').setFontColor('#FFFFFF');
  if (rows.length) {
    sh.getRange(r0 + 1, 1, rows.length, head.length).setValues(rows);
  }
  sh.setFrozenRows(r0);
  return r0 + rows.length;
}

function refreshAnalytics() {
  const started = new Date();
  const cov = buildFeeCoverage_();
  const model = analyticsTabModel_(cov);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nStu = model.students.rows.length;
  const nCov = model.coverage.rows.length;
  const firstCol = model.firstMonthCol;
  const lastCol  = model.students.head.length;

  // ── Students matrix ──
  const stu = analyticsSheet_(ss, TAB_STUDES);
  writeGrid_(stu, model.students.head, model.students.rows);
  stu.setFrozenColumns(3);
  if (nStu) {
    const band = stu.getRange(2, firstCol, nStu, Math.max(1, lastCol - firstCol + 1));
    stu.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('UNPAID')
        .setBackground('#F8D7DA').setFontColor('#842029').setRanges([band]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PAID')
        .setBackground('#D1E7DD').setFontColor('#0F5132').setRanges([band]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('EXTRA')
        .setBackground('#FFF3CD').setFontColor('#664D03').setRanges([band]).build()
    ]);
    stu.getRange(1, 1, nStu + 1, lastCol).createFilter();
  }
  stu.autoResizeColumns(1, Math.min(14, lastCol));

  // ── Coverage ──
  const cvr = analyticsSheet_(ss, TAB_COVER);
  writeGrid_(cvr, model.coverage.head, model.coverage.rows);
  if (nCov) {
    cvr.getRange(2, 5, nCov, 1).setNumberFormat('0.0%');
    cvr.getRange(2, 7, nCov, 1).setNumberFormat('"Rs. "#,##0');
    try {
      cvr.insertChart(cvr.newChart().asColumnChart()
        .addRange(cvr.getRange(1, 1, nCov + 1, 1))
        .addRange(cvr.getRange(1, 5, nCov + 1, 1))
        .setPosition(2, 10, 0, 0)
        .setOption('title', 'Fee coverage by month')
        .setOption('legend', { position: 'none' })
        .setOption('height', 320).setOption('width', 520)
        .build());
    } catch (e) {
      Logger.log('Chart skipped: ' + e.message);
    }
  }
  cvr.autoResizeColumns(1, model.coverage.head.length);

  // ── Names and gaps ──
  const nms = analyticsSheet_(ss, TAB_NAMES);
  writeGrid_(nms, model.names.head, model.names.rows);
  if (model.names.rows.length) {
    nms.getRange(1, 1, model.names.rows.length + 1, model.names.head.length).createFilter();
  }
  nms.autoResizeColumns(1, model.names.head.length);

  const gps = analyticsSheet_(ss, TAB_GAPS);
  writeGrid_(gps, model.gaps.head, model.gaps.rows);
  if (model.gaps.rows.length) {
    gps.getRange(1, 1, model.gaps.rows.length + 1, model.gaps.head.length).createFilter();
  }
  gps.autoResizeColumns(1, model.gaps.head.length);

  // ── Dashboard, driven by formulas so the picker responds instantly ──
  const dash = analyticsSheet_(ss, TAB_DASH);
  writeDashboard_(dash, model, nStu, nCov, firstCol, lastCol, started);

  ss.setActiveSheet(dash);
  const msg = 'Analytics refreshed ' +
    Utilities.formatDate(started, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') + '\n' +
    '  ' + TAB_DASH  + '  - pick a month, see who has not paid\n' +
    '  ' + TAB_COVER + '  - ' + nCov + ' months\n' +
    '  ' + TAB_STUDES+ '  - ' + nStu + ' students x ' + model.monthNames.length + ' months\n' +
    '  ' + TAB_NAMES + '  - ' + model.names.rows.length + ' names off the receipts\n' +
    '  ' + TAB_GAPS  + '  - ' + model.gaps.rows.length + ' students missing a month\n' +
    'Enrollments, Receipts and Config were not touched.\n';
  Logger.log(msg);
  return msg;
}

function writeDashboard_(sh, model, nStu, nCov, firstCol, lastCol, when) {
  const S = "'" + TAB_STUDES + "'";
  const C = "'" + TAB_COVER + "'";
  const lastStuRow = nStu + 1;
  // Every month cell for every student, as one block the formulas index into.
  const band = S + '!$' + colLetter_(firstCol) + '$2:$' + colLetter_(lastCol) + '$' + lastStuRow;
  const hdr  = S + '!$' + colLetter_(firstCol) + '$1:$' + colLetter_(lastCol) + '$1';
  // The unpaid list below spills down columns A to E, so the helper cell that
  // holds the month's column number has to live outside that path.
  const col  = 'INDEX(' + band + ',0,$H$1)';

  sh.getRange('A1').setValue('SRIRAM STUDIO — FEE DASHBOARD')
    .setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue(
    'Refreshed ' + Utilities.formatDate(when, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') +
    '.  Enrollments is the golden source. A receipt covering siblings counts for each child on it.');
  sh.getRange('A3').setValue(
    'Change the month below and this page updates at once. After editing Enrollments or ' +
    'Receipts, use  Analytics ▸ Refresh dashboard.');

  sh.getRange('A5').setValue('Month').setFontWeight('bold');
  sh.getRange('B5').setValue(model.defaultMonth);
  if (nCov) {
    sh.getRange('B5').setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(sh.getParent().getSheetByName(TAB_COVER).getRange(2, 1, nCov, 1), true)
        .setAllowInvalid(false).build());
  }
  sh.getRange('B5').setBackground('#FFF3CD').setFontWeight('bold');

  // Which column of the matrix that month is. Parked clear of the unpaid list.
  sh.getRange('G1').setValue('(internal — do not delete)').setFontColor('#BBBBBB');
  sh.getRange('H1').setFormula('=MATCH($B$5,' + hdr + ',0)').setFontColor('#BBBBBB');

  const kpis = [
    ['Students due to pay',   '=IFERROR(COUNTIF(' + col + ',"PAID")+COUNTIF(' + col + ',"UNPAID"),0)'],
    ['Covered by a receipt',  '=IFERROR(COUNTIF(' + col + ',"PAID"),0)'],
    ['No receipt',            '=IFERROR(COUNTIF(' + col + ',"UNPAID"),0)'],
    ['Coverage',              '=IFERROR($B$8/$B$7,0)'],
    ['Receipts booked',       '=IFERROR(VLOOKUP($B$5,' + C + '!$A$2:$H$' + (nCov + 1) + ',6,FALSE),0)'],
    ['Collected',             '=IFERROR(VLOOKUP($B$5,' + C + '!$A$2:$H$' + (nCov + 1) + ',7,FALSE),0)'],
    ['Notes',                 '=IFERROR(VLOOKUP($B$5,' + C + '!$A$2:$H$' + (nCov + 1) + ',8,FALSE),"")']
  ];
  kpis.forEach(function (k, i) {
    sh.getRange(7 + i, 1).setValue(k[0]).setFontWeight('bold');
    sh.getRange(7 + i, 2).setFormula(k[1]);
  });
  sh.getRange('B10').setNumberFormat('0.0%');
  sh.getRange('B12').setNumberFormat('"Rs. "#,##0');

  sh.getRange('A15').setValue('WHO HAS NOT PAID FOR THIS MONTH').setFontWeight('bold');
  sh.getRange('A16').setValue(
    'Straight from the roster. Chase these, or check the Analytics Names tab if a receipt ' +
    'was issued under a spelling nobody recognises.');
  sh.getRange(17, 1, 1, 5).setValues([['Row', 'Student', 'Centre', 'Phone', 'Last paid']])
    .setFontWeight('bold').setBackground('#2C1A0E').setFontColor('#FFFFFF');
  if (nStu) {
    sh.getRange('A18').setFormula(
      '=IFERROR(FILTER({' + S + '!$A$2:$A$' + lastStuRow + ',' +
                          S + '!$C$2:$E$' + lastStuRow + ',' +
                          S + '!$N$2:$N$' + lastStuRow + '},' +
      col + '="UNPAID"),"Everybody due this month has paid.")');
  }

  sh.getRange('D5').setValue('Roster').setFontWeight('bold');
  sh.getRange('E5').setValue(model.counts.roster + ' rows, ' + model.counts.billable +
                             ' monthly-fee students, ' + model.counts.active + ' active today');
  sh.getRange('D6').setValue('Receipts').setFontWeight('bold');
  sh.getRange('E6').setValue(model.counts.monthly + ' monthly receipts read, ' +
                             model.counts.otherType + ' other fee types skipped, ' +
                             model.counts.noPeriod + ' with no fee month');

  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 210);
  sh.setColumnWidth(3, 150);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 420);
  sh.setFrozenRows(4);
}


// ─── Preview, per the house rule: look before writing ─────────

function previewAnalyticsRefresh() {
  const cov = buildFeeCoverage_();
  const model = analyticsTabModel_(cov);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let out = 'ANALYTICS REFRESH - PREVIEW\n===========================\n';
  out += 'Nothing has been written. This is what a refresh would do.\n\n';

  out += 'TABS\n';
  const sizes = {};
  sizes[TAB_DASH]   = 'the dashboard, month picker and unpaid list';
  sizes[TAB_COVER]  = model.coverage.rows.length + ' months';
  sizes[TAB_STUDES] = model.students.rows.length + ' students x ' +
                      model.monthNames.length + ' months';
  sizes[TAB_NAMES]  = model.names.rows.length + ' names printed on receipts';
  sizes[TAB_GAPS]   = model.gaps.rows.length + ' students missing at least one month';

  let blocked = 0;
  ANALYTICS_TABS.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    let state;
    if (!sh) {
      state = 'will be created';
    } else if (sh.getLastRow() > 0 &&
               sh.getRange(1, 1).getNote().indexOf(ANALYTICS_MARK) !== 0) {
      state = 'BLOCKED - a tab of this name exists and was not generated here';
      blocked++;
    } else {
      state = 'will be rebuilt (' + sh.getLastRow() + ' rows replaced)';
    }
    out += '  ' + pad_(name, 22) + pad_(state, 46) + sizes[name] + '\n';
  });

  out += '\nUNTOUCHED\n';
  PROTECTED_TABS.forEach(function (n) {
    const sh = ss.getSheetByName(n);
    out += '  ' + pad_(n, 22) + (sh ? sh.getLastRow() + ' rows - read only, never written' :
                                      'not present') + '\n';
  });

  out += '\nDASHBOARD WOULD OPEN ON : ' + (model.defaultMonth || '(no months yet)') + '\n';
  if (blocked) {
    out += '\n' + blocked + ' tab(s) are blocked. Rename or delete them first; a refresh\n';
    out += 'would stop at the first one and change nothing.\n';
  }
  out += '\nReport only. Nothing was changed.\n';
  return report_(out);
}
