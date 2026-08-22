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


// ─── Legacy name reconciliation (run by hand from the editor) ─
// The 'Legacy Students' tab is the canonical list of correct names.
// Receipts predate it and are messy: misspellings, and siblings sometimes
// recorded together in one row. These report how the two line up.
//
// Editor-only — deliberately not routed through doGet.
// previewNameCleanup() writes NOTHING. Read its log before anything else.

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
      cur[j] = Math.min(
        prev[j] + 1,
        cur[j - 1] + 1,
        prev[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1)
      );
    }
    prev = cur;
  }
  return prev[b.length];
}

function findNameColumn_(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toString().toLowerCase();
    if (h.indexOf('name') >= 0) return i;
  }
  return 0;   // fall back to the first column
}

function previewNameCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legacy = ss.getSheetByName('Legacy Students');
  if (!legacy) {
    const names = ss.getSheets().map(function (x) { return x.getName(); }).join(', ');
    const msg = 'No tab called "Legacy Students". Tabs found: ' + names;
    Logger.log(msg);
    return msg;
  }

  const lData = legacy.getDataRange().getValues();
  const rData = getSheet('Receipts').getDataRange().getValues();
  const norm  = function (v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };

  let out = 'LEGACY STUDENTS TAB\n-------------------\n';
  out += 'Headers      : ' + lData[0].map(norm).join(' | ') + '\n';
  out += 'Rows of data : ' + Math.max(0, lData.length - 1) + '\n';

  const lName = findNameColumn_(lData[0].map(norm));
  out += 'Name column  : "' + norm(lData[0][lName]) + '" (column ' + (lName + 1) + ')\n';

  // Canonical names, and any repeats within the list itself.
  const canon = [], countByName = {};
  for (let i = 1; i < lData.length; i++) {
    const n = norm(lData[i][lName]);
    if (!n) continue;
    canon.push(n);
    const k = n.toLowerCase();
    countByName[k] = (countByName[k] || 0) + 1;
  }
  const dupes = Object.keys(countByName).filter(function (k) { return countByName[k] > 1; });
  out += 'Names listed : ' + canon.length + '\n';
  out += 'Repeated names in your list (these need phone numbers): ' + dupes.length + '\n';
  dupes.forEach(function (k) {
    out += '   ' + k + '  x' + countByName[k] + '\n';
  });

  // Distinct receipt names and how often each appears.
  const rHead = rData[0].map(norm);
  const rName = rHead.indexOf('Student Name');
  const rCount = {};
  for (let i = 1; i < rData.length; i++) {
    const n = norm(rData[i][rName]);
    if (!n) continue;
    rCount[n] = (rCount[n] || 0) + 1;
  }
  const rNames = Object.keys(rCount);

  out += '\nRECEIPTS TAB\n------------\n';
  out += 'Receipt rows   : ' + Math.max(0, rData.length - 1) + '\n';
  out += 'Distinct names : ' + rNames.length + '\n';

  // Classify each receipt name against the canonical list.
  const canonLower = canon.map(function (n) { return n.toLowerCase(); });
  const exact = [], close = [], multi = [], none = [];
  const SEPARATORS = /\s(?:&|and|\+)\s|\s*[\/,]\s*/i;

  rNames.forEach(function (rn) {
    const low = rn.toLowerCase();
    if (canonLower.indexOf(low) >= 0) { exact.push(rn); return; }

    // Does this row look like more than one student?
    const hits = canon.filter(function (c) { return low.indexOf(c.toLowerCase()) >= 0; });
    if (hits.length > 1 || SEPARATORS.test(rn)) {
      multi.push({ name: rn, n: rCount[rn], hits: hits });
      return;
    }

    // Nearest canonical name by edit distance.
    let best = null, bestD = 99;
    canon.forEach(function (c) {
      const d = levenshtein_(rn, c);
      if (d < bestD) { bestD = d; best = c; }
    });
    const tolerance = Math.max(2, Math.floor(rn.length * 0.25));
    if (best && bestD <= tolerance) close.push({ from: rn, to: best, d: bestD, n: rCount[rn] });
    else none.push({ name: rn, n: rCount[rn], nearest: best, d: bestD });
  });

  out += '\nMATCHING\n--------\n';
  out += 'Exact match to your list : ' + exact.length + '\n';
  out += 'Likely typo (fixable)    : ' + close.length + '\n';
  out += 'Looks like 2+ students   : ' + multi.length + '\n';
  out += 'No match at all          : ' + none.length + '\n';

  if (close.length) {
    out += '\nLIKELY TYPOS - proposed corrections:\n';
    close.sort(function (a, b) { return a.d - b.d; });
    close.forEach(function (c) {
      out += '  "' + c.from + '"  ->  "' + c.to + '"   (' + c.n + ' receipt(s), ' + c.d + ' char diff)\n';
    });
  }
  if (multi.length) {
    out += '\nLOOKS LIKE MORE THAN ONE STUDENT - needs your decision:\n';
    multi.forEach(function (m) {
      out += '  "' + m.name + '"   (' + m.n + ' receipt(s))';
      if (m.hits.length) out += '   matches: ' + m.hits.join(' + ');
      out += '\n';
    });
  }
  if (none.length) {
    out += '\nNO MATCH - missing from your list, or too different to guess:\n';
    none.forEach(function (x) {
      out += '  "' + x.name + '"   (' + x.n + ' receipt(s))   nearest: "' + x.nearest + '" (' + x.d + ' off)\n';
    });
  }

  out += '\nNothing was changed. This is a report only.\n';
  Logger.log(out);
  return out;
}
