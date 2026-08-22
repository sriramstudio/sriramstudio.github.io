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
      'Heard From','Notes'];
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
      'Fee Type','Date Received','Note'];
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

  sheet.appendRow([
    id, now,
    d.mode === 'admission'     ? 'New Admission' :
    d.mode === 'workshop'      ? 'Workshop' :
    d.mode === 'legacy'        ? 'Existing Student' :
    d.mode === 'app-admission' ? 'Application – Admission' :
    d.mode === 'app-workshop'  ? 'Application – Workshop' : 'Existing Student',
    d.studentName || '',
    d.dob || '',
    d.gender || '',
    d.bloodGroup || '',
    d.school || '',
    d.guardianName || '',
    d.relation || '',
    d.phone || '',
    d.whatsapp || '',
    d.email || '',
    d.address || '',
    d.program || '',
    d.location || '',
    d.batch || '',
    d.mode === 'legacy' ? (d.approxJoining || '') : (d.startDate || d.workshopDate || ''),
    d.pracheen || '',
    d.workshopName || '',
    d.workshopDate || '',
    d.workshopFee ? '₹' + d.workshopFee : '',
    d.hearFrom || '',
    d.notes || ''
  ]);

  // Send notifications for self-registered applications only
  if (d.mode === 'app-admission' || d.mode === 'app-workshop') {
    try { sendNotifications(d, id, now); } catch(e) { Logger.log('Notification error: ' + e.message); }
  }

  return { success: true, id, enrolledAt: now };
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
        location: col('Location')(row)
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
function addReceipt(d) {
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
    d.note          || ''
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


// ─── Legacy student import (run by hand from the editor) ──────
// Students who joined before the registration form existed appear only in
// the Receipts sheet, so searchStudents — which reads Enrollments — never
// finds them. These rebuild Enrollment rows from the receipt history,
// carrying the contact number across so same-name students stay distinct.
//
// NOT routed through doGet: these are editor-only, never reachable on the web.
// Run previewLegacyImport() first — it writes nothing.

function colIndex_(headers, name) {
  const i = headers.indexOf(name);
  if (i < 0) throw new Error('Missing column in sheet: ' + name);
  return i;
}

function buildLegacyImport_() {
  const rData = getSheet('Receipts').getDataRange().getValues();
  const eData = getSheet('Enrollments').getDataRange().getValues();
  if (rData.length <= 1) return { toAdd: [], text: 'Receipts sheet is empty.' };

  const rHead = rData[0], eHead = eData[0];
  const rName = colIndex_(rHead, 'Student Name');
  const rCont = colIndex_(rHead, 'Contact');
  const eName = colIndex_(eHead, 'Student Name');
  const ePhone = colIndex_(eHead, 'Phone');

  const norm   = v => (v === undefined || v === null) ? '' : v.toString().trim();
  const digits = v => norm(v).replace(/\D/g, '');

  // Who is already enrolled, keyed by name + phone digits.
  const existing = new Set();
  for (let i = 1; i < eData.length; i++) {
    const n = norm(eData[i][eName]);
    if (n) existing.add(n.toLowerCase() + '|' + digits(eData[i][ePhone]));
  }

  // Distinct name+contact pairs across the whole receipt history.
  const seen = {};
  for (let i = 1; i < rData.length; i++) {
    const n = norm(rData[i][rName]);
    if (!n) continue;
    const key = n.toLowerCase() + '|' + digits(rData[i][rCont]);
    if (!seen[key]) seen[key] = { name: n, phone: norm(rData[i][rCont]), digits: digits(rData[i][rCont]), count: 0 };
    seen[key].count++;
  }

  const toAdd = [], skipped = [], noPhone = [], byName = {};
  Object.keys(seen).forEach(k => {
    const rec = seen[k], ln = rec.name.toLowerCase();
    (byName[ln] = byName[ln] || []).push(rec);
    if (existing.has(k)) { skipped.push(rec); return; }
    if (!rec.digits) noPhone.push(rec);
    toAdd.push(rec);
  });
  const collisions = Object.keys(byName).filter(n => byName[n].length > 1).map(n => byName[n]);

  let text = 'LEGACY IMPORT PREVIEW\n---------------------\n';
  text += 'Distinct name+contact pairs in Receipts : ' + Object.keys(seen).length + '\n';
  text += 'Already in Enrollments (will skip)      : ' + skipped.length + '\n';
  text += 'Will be added as "Existing Student"     : ' + toAdd.length + '\n';
  text += 'Of those, with NO phone number          : ' + noPhone.length + '\n\n';

  if (collisions.length) {
    text += 'SAME NAME, DIFFERENT CONTACT — check these are really different people:\n';
    collisions.forEach(g => {
      text += '  ' + g[0].name + '\n';
      g.forEach(r => { text += '     ' + (r.phone || '(no phone)') + '  - ' + r.count + ' receipt(s)\n'; });
    });
    text += '\n';
  } else {
    text += 'No same-name collisions found.\n\n';
  }

  if (noPhone.length) {
    text += 'NO PHONE ON RECEIPT (cannot be told apart if the name repeats):\n';
    noPhone.forEach(r => { text += '  ' + r.name + '\n'; });
    text += '\n';
  }

  text += 'To be added:\n';
  toAdd.forEach(r => { text += '  ' + r.name + '   ' + (r.phone || '(no phone)') + '\n'; });

  return { toAdd: toAdd, skipped: skipped, collisions: collisions, noPhone: noPhone, text: text, eHead: eHead };
}

// Read-only. Writes nothing — run this first and read the log.
function previewLegacyImport() {
  const r = buildLegacyImport_();
  Logger.log(r.text);
  return r.text;
}

// Appends the rows. Safe to re-run: anything already present is skipped.
function importLegacyStudents() {
  const r = buildLegacyImport_();
  if (!r.toAdd || !r.toAdd.length) { Logger.log('Nothing to add.'); return 'Nothing to add.'; }

  const sheet = getSheet('Enrollments');
  const head  = r.eHead;
  const now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

  const rows = r.toAdd.map((rec, i) => {
    const row = new Array(head.length).fill('');
    row[colIndex_(head, 'ID')]           = 'SR-LEGACY-' + stamp + '-' + String(i + 1).padStart(3, '0');
    row[colIndex_(head, 'Enrolled At')]  = now;
    row[colIndex_(head, 'Type')]         = 'Existing Student';
    row[colIndex_(head, 'Student Name')] = rec.name;
    row[colIndex_(head, 'Phone')]        = rec.phone;
    row[colIndex_(head, 'WhatsApp')]     = rec.phone;
    row[colIndex_(head, 'Notes')]        = 'Imported from receipt history';
    return row;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, head.length).setValues(rows);
  Logger.log('Added ' + rows.length + ' legacy students.');
  return 'Added ' + rows.length + ' legacy students.';
}
