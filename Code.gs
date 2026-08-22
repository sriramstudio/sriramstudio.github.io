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
    'Joining Date': d.mode === 'legacy' ? (d.approxJoining || '') : (d.startDate || d.workshopDate || ''),
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
  const row = head.map(function (h) { return put.hasOwnProperty(h) ? put[h] : ''; });
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


// ─── Legacy roster import (run by hand from the editor) ───────
// The 'Legacy Students' tab is the canonical list of students who joined
// before the registration form existed. Receipts are NOT used as a source:
// they contain misspellings and siblings recorded together in one row.
//
// Imported rows carry Type = "Existing Student" (the label the admin panel's
// Existing Student mode uses) so they never read as new admissions.
//
// Runs as an UPSERT, because the roster is expected to be filled in over
// time: a name already present is not added again, and blank Phone/Location
// cells are filled from the roster when it later gains that detail. Existing
// values are never overwritten.
//
// Editor-only — deliberately not routed through doGet.
// previewLegacyStudents() writes NOTHING. Run it first.

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

function findColumn_(headers, candidates) {
  for (let c = 0; c < candidates.length; c++) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().toLowerCase().indexOf(candidates[c]) >= 0) return i;
    }
  }
  return -1;
}

function buildLegacyRoster_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tab = ss.getSheetByName('Legacy Students');
  if (!tab) {
    const names = ss.getSheets().map(function (x) { return x.getName(); }).join(', ');
    return { error: 'No tab called "Legacy Students". Tabs found: ' + names };
  }

  const lData = tab.getDataRange().getValues();
  if (lData.length <= 1) return { error: 'The Legacy Students tab has no data rows.' };

  const norm   = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  const digits = function (v) { return norm(v).replace(/\D/g, ''); };

  const lHead    = lData[0].map(norm);
  const nameAt   = findColumn_(lHead, ['student name', 'name']);
  const phoneAt  = findColumn_(lHead, ['phone', 'contact', 'mobile', 'whatsapp']);
  const centreAt = findColumn_(lHead, ['center', 'centre', 'location', 'branch']);
  const statusAt = findColumn_(lHead, ['status', 'active', 'current']);
  if (nameAt < 0) return { error: 'No name column found. Headers: ' + lHead.join(' | ') };

  const eData  = getSheet('Enrollments').getDataRange().getValues();
  const eHead  = eData[0].map(norm);
  const eName  = eHead.indexOf('Student Name');
  const ePhone = eHead.indexOf('Phone');
  const eWa    = eHead.indexOf('WhatsApp');
  const eLoc    = eHead.indexOf('Location');
  const eStatus = ensureStatusColumn_();

  // Everyone already in Enrollments, indexed by name.
  const byExisting = {};
  for (let i = 1; i < eData.length; i++) {
    const n = norm(eData[i][eName]);
    if (!n) continue;
    const k = n.toLowerCase();
    (byExisting[k] = byExisting[k] || []).push({
      sheetRow: i + 1,
      phone:    norm(eData[i][ePhone]),
      digits:   digits(eData[i][ePhone]),
      location: norm(eData[i][eLoc]),
      status:   norm(eData[i][eStatus])
    });
  }

  // The roster.
  const rows = [], byName = {};
  for (let i = 1; i < lData.length; i++) {
    const n = norm(lData[i][nameAt]);
    if (!n) continue;
    const rec = {
      row:    i + 1,
      name:   n,
      phone:  phoneAt  >= 0 ? norm(lData[i][phoneAt])  : '',
      centre: centreAt >= 0 ? norm(lData[i][centreAt]) : ''
    };
    rec.digits    = digits(rec.phone);
    rec.centreRaw = rec.centre;
    const mapped  = matchCentre_(rec.centre);
    rec.centre    = mapped || '';                        // only real branches
    const leftover = (!mapped && rec.centreRaw) ? rec.centreRaw : '';
    // A leaving word — wherever it was written — becomes the status.
    const rawStatus = statusAt >= 0 ? norm(lData[i][statusAt]) : '';
    rec.status = (isLeftWord_(rawStatus) || isLeftWord_(leftover))
                   ? STATUS_LEFT
                   : (rawStatus ? STATUS_ACTIVE : '');
    rec.centreNote = isLeftWord_(leftover) ? '' : leftover;
    rows.push(rec);
    const k = n.toLowerCase();
    (byName[k] = byName[k] || []).push(rec);
  }

  const toAdd = [], toUpdate = [], already = [], blocked = [], ambiguous = [];

  rows.forEach(function (rec) {
    const k       = rec.name.toLowerCase();
    const twins   = byName[k];              // same name within the roster
    const matches = byExisting[k] || [];    // same name already enrolled

    if (!matches.length) {
      // Two roster rows with one name and no phone would land as identical
      // rows — refuse rather than create something nobody can tell apart.
      if (twins.length > 1 && !rec.digits) { blocked.push(rec); return; }
      toAdd.push(rec);
      return;
    }

    // Pick which existing row this refers to.
    // An existing row can only stand for one roster row. Without this, two
    // students sharing a name both target the same blank-phone row: the first
    // fills it in, the second overwrites it, and one real student is lost.
    let target = null;
    if (rec.digits) {
      const exact = matches.filter(function (m) { return m.digits === rec.digits; });
      if (exact.length) target = exact[0];
      else {
        const blank = matches.filter(function (m) { return !m.digits && !m.claimed; });
        // A blank-phone row is one we imported earlier; fill it in. Otherwise
        // this is a genuinely different student who happens to share a name.
        if (blank.length) target = blank[0];
        else { toAdd.push(rec); return; }
      }
    } else {
      const free = matches.filter(function (m) { return !m.claimed; });
      if (free.length > 1) { ambiguous.push(rec); return; }
      // The roster says several students share this name but fewer exist in
      // Enrollments, and there is no phone to say which row is which.
      // Treating the extras as "already there" would drop a real student.
      if (!free.length || twins.length > matches.length) { blocked.push(rec); return; }
      target = free[0];
    }
    target.claimed = true;

    // Fill blanks only — never overwrite what is already recorded.
    const sets = [];
    if (rec.phone  && !target.phone)    sets.push({ col: ePhone + 1, val: rec.phone, what: 'phone' });
    if (rec.phone  && !target.phone && eWa >= 0) sets.push({ col: eWa + 1, val: rec.phone, what: 'whatsapp' });
    if (rec.centre && !target.location) sets.push({ col: eLoc + 1, val: rec.centre, what: 'centre' });
    if (rec.status && !target.status)    sets.push({ col: eStatus + 1, val: rec.status, what: 'status' });

    if (sets.length) toUpdate.push({ rec: rec, sheetRow: target.sheetRow, sets: sets });
    else already.push(rec);
  });

  const dupes = Object.keys(byName).filter(function (k) { return byName[k].length > 1; });

  return {
    headers: lHead,
    nameHeader:   lHead[nameAt],
    phoneHeader:  phoneAt  >= 0 ? lHead[phoneAt]  : null,
    centreHeader: centreAt >= 0 ? lHead[centreAt] : null,
    total: rows.length,
    leftCount: rows.filter(function (x) { return x.status === STATUS_LEFT; }).length,
    toAdd: toAdd, toUpdate: toUpdate, already: already,
    blocked: blocked, ambiguous: ambiguous,
    dupes: dupes, byName: byName, eHead: eHead
  };
}

function previewLegacyStudents() {
  const r = buildLegacyRoster_();
  if (r.error) { Logger.log(r.error); return r.error; }

  let out = 'LEGACY ROSTER IMPORT - PREVIEW\n==============================\n\n';
  out += 'Tab headers   : ' + r.headers.join(' | ') + '\n';
  out += 'Name column   : "' + r.nameHeader + '"\n';
  out += 'Phone column  : ' + (r.phoneHeader  ? '"' + r.phoneHeader  + '"' : 'NONE FOUND') + '\n';
  out += 'Centre column : ' + (r.centreHeader ? '"' + r.centreHeader + '"' : 'NONE FOUND') + '\n';
  out += 'Names listed  : ' + r.total + '\n';
  out += 'Marked as left: ' + r.leftCount + '\n\n';

  out += 'New rows to add                  : ' + r.toAdd.length + '\n';
  out += 'Existing rows to fill in         : ' + r.toUpdate.length + '\n';
  out += 'Already complete (nothing to do) : ' + r.already.length + '\n';
  out += 'Blocked - repeated name, no phone: ' + r.blocked.length + '\n';
  out += 'Ambiguous - name enrolled twice  : ' + r.ambiguous.length + '\n\n';

  if (r.blocked.length) {
    out += 'NOT IMPORTED - this name appears more than once in your roster and\n';
    out += 'there is no phone to tell those students apart. Add a phone on these\n';
    out += 'rows, then run again:\n';
    r.blocked.forEach(function (x) { out += '  row ' + x.row + '  ' + x.name + '\n'; });
    out += '\n';
  }
  if (r.ambiguous.length) {
    out += 'SKIPPED - this name is already in Enrollments more than once, and\n';
    out += 'the roster row has no phone to say which one it means:\n';
    r.ambiguous.forEach(function (x) { out += '  row ' + x.row + '  ' + x.name + '\n'; });
    out += '\n';
  }
  if (r.toUpdate.length) {
    out += 'WILL FILL IN (blank cells only, nothing overwritten):\n';
    r.toUpdate.slice(0, 15).forEach(function (u) {
      out += '  ' + u.rec.name + '  ->  ' +
             u.sets.map(function (x) { return x.what + '=' + x.val; }).join(', ') + '\n';
    });
    if (r.toUpdate.length > 15) out += '  ... ' + (r.toUpdate.length - 15) + ' more ...\n';
    out += '\n';
  }

  const oddCentres = r.toAdd.concat(r.already).filter(function (x) { return !!x.centreNote; });
  if (oddCentres.length) {
    out += 'CENTER COLUMN VALUES THAT ARE NOT A BRANCH (' + oddCentres.length + '):\n';
    out += 'These will NOT go into Location. They are kept in Notes instead.\n';
    oddCentres.forEach(function (x) {
      out += '  row ' + x.row + '  ' + x.name + '  ->  "' + x.centreNote + '"\n';
    });
    out += '\n';
  }

  if (r.toAdd.length) {
    const withPhone  = r.toAdd.filter(function (x) { return !!x.digits; }).length;
    const withCentre = r.toAdd.filter(function (x) { return !!x.centre; }).length;
    out += 'Of the new rows: ' + withPhone + ' have a phone, ' + withCentre + ' have a centre.\n';
    out += 'The rest go in with the name only — phone and centre can be added\n';
    out += 'to the roster later and filled in by re-running this import.\n\n';
    out += 'SAMPLE (first 10 and last 5 of ' + r.toAdd.length + '):\n';
    const show = function (x) {
      out += '  ' + x.name + '   ' + (x.phone || '(no phone)') +
             '   ' + (x.centre || '(no centre)') + '\n';
    };
    r.toAdd.slice(0, 10).forEach(show);
    if (r.toAdd.length > 15) out += '  ... ' + (r.toAdd.length - 15) + ' more ...\n';
    if (r.toAdd.length > 10) r.toAdd.slice(-5).forEach(show);
  }

  out += '\nNothing was changed. Report only.\n';
  Logger.log(out);
  return out;
}

// Appends new students and fills blanks on ones already there.
// Safe to re-run as the roster gains phone numbers and centres.
function importLegacyStudents() {
  const r = buildLegacyRoster_();
  if (r.error) { Logger.log(r.error); return r.error; }
  if (!r.toAdd.length && !r.toUpdate.length) { Logger.log('Nothing to do.'); return 'Nothing to do.'; }

  const sheet = getSheet('Enrollments');
  const head  = r.eHead;
  const idx   = function (n) {
    const i = head.indexOf(n);
    if (i < 0) throw new Error('Missing column: ' + n);
    return i;
  };

  // Fill blanks on rows that already exist.
  r.toUpdate.forEach(function (u) {
    u.sets.forEach(function (st) { sheet.getRange(u.sheetRow, st.col).setValue(st.val); });
  });

  // Append the new ones.
  let added = 0;
  if (r.toAdd.length) {
    const now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const rows  = r.toAdd.map(function (rec, i) {
      const row = new Array(head.length).fill('');
      row[idx('ID')]           = 'SR-LEGACY-' + stamp + '-' + String(i + 1).padStart(3, '0');
      row[idx('Enrolled At')]  = now;
      row[idx('Type')]         = 'Existing Student';
      row[idx('Student Name')] = rec.name;
      row[idx('Phone')]        = rec.phone;
      row[idx('WhatsApp')]     = rec.phone;
      row[idx('Location')]     = rec.centre;
      row[idx('Notes')]        = 'Legacy roster - imported from Legacy Students tab' +
                                 (rec.centreNote ? ' | Center column said: ' + rec.centreNote : '');
      row[idx('Status')]       = rec.status || STATUS_ACTIVE;
      return row;
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, head.length).setValues(rows);
    added = rows.length;
  }

  let msg = 'Added ' + added + ' student(s); filled in ' + r.toUpdate.length + ' existing row(s).';
  if (r.blocked.length)   msg += ' ' + r.blocked.length + ' blocked.';
  if (r.ambiguous.length) msg += ' ' + r.ambiguous.length + ' ambiguous.';
  Logger.log(msg);
  return msg;
}
