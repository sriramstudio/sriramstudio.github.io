// Runs Code.gs against stubbed Apps Script globals to test the auth gate.
const fs = require('fs'), vm = require('vm');
const SRC = process.argv[2];

function makeSheet(rows) {
  const data = rows.map(r => r.slice());
  const chain = { setFontWeight: () => chain, setBackground: () => chain, setFontColor: () => chain };
  return {
    _data: data,
    getDataRange: () => ({ getValues: () => data.map(r => r.slice()) }),
    appendRow: r => data.push(r.slice()),
    deleteRow: i => data.splice(i - 1, 1),
    deleteColumn: c => data.forEach(row => row.splice(c - 1, 1)),
    getLastRow: () => data.length,
    getLastColumn: () => data.reduce((m, r) => Math.max(m, r.length), 0),
    getRange(r, c, numRows, numCols) {
      const range = {
        setValue: v => { while (data.length < r) data.push([]); data[r - 1][c - 1] = v; return range; },
        setValues: vals => {
          vals.forEach((rowVals, ri) => {
            while (data.length < r + ri) data.push([]);
            rowVals.forEach((v, ci) => { data[r + ri - 1][c + ci - 1] = v; });
          });
          return range;
        },
        getValues: () => {
          const out = [];
          for (let i = 0; i < (numRows || 1); i++) {
            const row = data[r + i - 1] || [];
            const slice = [];
            for (let j = 0; j < (numCols || 1); j++) slice.push(row[c + j - 1] !== undefined ? row[c + j - 1] : '');
            out.push(slice);
          }
          return out;
        },
        setFontWeight: () => range, setBackground: () => range, setFontColor: () => range
      };
      return range;
    },
    setFrozenRows() {}, setColumnWidth() {}
  };
}

function freshWorld(pin) {
  const sheets = {
    Config: makeSheet([['Key', 'Value'], ['pin', pin], ['receipt_seq', '1'],
                       ['notify_email', ''], ['notify_whatsapp', ''], ['callmebot_key', '']]),
    Enrollments: makeSheet([
      ['ID', 'Enrolled At', 'Type', 'Student Name', 'Date of Birth', 'Gender', 'Blood Group',
       'School/College', 'Guardian Name', 'Relation', 'Phone', 'WhatsApp', 'Email', 'Address',
       'Program', 'Location', 'Batch', 'Joining Date', 'Pracheen Kala Kendra', 'Workshop Name',
       'Workshop Date', 'Workshop Fee', 'Heard From', 'Notes', 'Status'],
      ['SR-2026-0101', '01 Jan 2026', 'New Admission', 'Test Child', '2015-01-01', 'F', '', '',
       'Guardian', 'Mother', '9999', '9999', 'a@b.c', '12 Somewhere Rd', '', 'Bhawanipur',
       '', '', '', '', '', '', '', '']
    ]),
    Receipts: makeSheet([['Receipt No', 'Issued At', 'Student Name', 'Contact', 'Amount (₹)',
                          'Fee Month', 'Fee Year', 'Payment Mode', 'UPI Reference', 'Fee Type',
                          'Date Received', 'Note']])
  };
  const cache = {};
  const sandbox = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: n => sheets[n] || null,
        insertSheet: n => (sheets[n] = makeSheet([]))
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: k => (k in cache ? cache[k] : null),
        put: (k, v) => { cache[k] = v; },
        remove: k => { delete cache[k]; }
      })
    },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: t => ({ _t: t, setMimeType() { return this; }, getContent() { return this._t; } })
    },
    // Honour the format string: code depends on the shape it returns.
    Utilities: { formatDate: (d, tz, fmt) => (
      fmt === 'MMMM yyyy'      ? 'January 2026' :
      fmt === 'dd MMM yyyy'    ? '01 Jan 2026' :
      fmt === 'yyyyMMddHHmmss' ? '20260101100000' :
      fmt === 'yyyy-MM-dd HHmm'? '2026-01-01 1000' :
                                 '01 Jan 2026, 10:00 AM') },
    Session: { getScriptTimeZone: () => 'Asia/Kolkata' },
    MailApp: { sendEmail() {} },
    UrlFetchApp: { fetch() {} },
    Logger: { log() {} },
    console
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(SRC, 'utf8'), sandbox, { filename: 'Code.gs' });
  return { sandbox, sheets, cache };
}

const call = (w, params) => JSON.parse(w.sandbox.doGet({ parameter: params }).getContent());
const enc = o => encodeURIComponent(JSON.stringify(o));

let pass = 0, fail = 0;
function check(name, cond, detail) {
  if (cond) { pass++; console.log('  PASS  ' + name); }
  else { fail++; console.log('  FAIL  ' + name + (detail !== undefined ? '   -> ' + JSON.stringify(detail) : '')); }
}

console.log('\n--- reads are protected ---');
let w = freshWorld('1234');
let r = call(w, { action: 'getEnrollments' });
check('getEnrollments with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'getEnrollments', pin: '0000' });
check('getEnrollments with WRONG pin is refused', r.auth === true, r);
r = call(w, { action: 'getEnrollments', pin: '1234' });
check('getEnrollments with correct pin returns data', Array.isArray(r.records) && r.records.length === 1, r);
r = call(w, { action: 'searchStudents', q: 'Test' });
check('searchStudents with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'getReceipts' });
check('getReceipts with NO pin is refused', r.auth === true, r);

console.log('\n--- public registration must still work ---');
w = freshWorld('1234');
r = call(w, { action: 'addEnrollment', data: enc({ mode: 'app-admission', studentName: 'Public Kid' }) });
check('public form (app-admission) works with NO pin', r.success === true, r);
r = call(w, { action: 'addEnrollment', data: enc({ mode: 'app-workshop', studentName: 'Public Kid 2' }) });
check('public form (app-workshop) works with NO pin', r.success === true, r);

console.log('\n--- admin writes must not ---');
w = freshWorld('1234');
r = call(w, { action: 'addEnrollment', data: enc({ mode: 'legacy', studentName: 'Sneaky' }) });
check('admin enrol (legacy) with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'addEnrollment', data: enc({ mode: 'admission', studentName: 'Sneaky2' }) });
check('admin enrol (admission) with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'addEnrollment', data: '%7Bnot json' });
check('malformed data fails CLOSED', r.auth === true, r);
r = call(w, { action: 'deleteEnrollment', id: 'SR-2026-0101' });
check('deleteEnrollment with NO pin is refused', r.auth === true, r);
check('  ...and the record survived', w.sheets.Enrollments._data.length === 2, w.sheets.Enrollments._data.length);
r = call(w, { action: 'addReceipt', data: enc({ studentName: 'x' }) });
check('addReceipt with NO pin is refused', r.auth === true, r);

console.log('\n--- the updateConfig back door ---');
w = freshWorld('1234');
r = call(w, { action: 'updateConfig', key: 'pin', value: '9999' });
check('updateConfig pin with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'updateConfig', key: 'pin', value: '9999', pin: '1234' });
check('updateConfig pin refused EVEN WITH correct pin', r.success === false, r);
r = call(w, { action: 'getEnrollments', pin: '1234' });
check('  ...so the PIN is unchanged', Array.isArray(r.records), r);
r = call(w, { action: 'updateConfig', key: 'notify_email', value: 'new@x.com', pin: '1234' });
check('updateConfig of a normal key still works', r.success === true, r);

console.log('\n--- brute-force throttle ---');
w = freshWorld('1234');
for (let i = 0; i < 10; i++) call(w, { action: 'verifyPin', pin: '0000' });
r = call(w, { action: 'verifyPin', pin: '1234' });
check('locks out after 10 bad PINs, even for the RIGHT pin', r.locked === true, r);
r = call(w, { action: 'getEnrollments', pin: '1234' });
check('  ...and protected reads are locked too', r.locked === true, r);

w = freshWorld('1234');
for (let i = 0; i < 5; i++) call(w, { action: 'verifyPin', pin: '0000' });
r = call(w, { action: 'verifyPin', pin: '1234' });
check('a correct PIN before the limit succeeds', r.valid === true, r);
r = call(w, { action: 'verifyPin', pin: '0000' });
check('  ...and success reset the counter', r.locked === undefined, r);

console.log('\n--- same-name students stay distinct ---');
w = freshWorld('1234');
// A second, different student who happens to share a name.
r = call(w, { action: 'addEnrollment', pin: '1234',
              data: enc({ mode: 'legacy', studentName: 'Test Child', phone: '8888888888' }) });
check('second same-name student added', r.success === true, r);
r = call(w, { action: 'searchStudents', q: 'Test Child', pin: '1234' });
check('BOTH same-name students are returned', r.results && r.results.length === 2, r.results);
check('  ...and their phones differ',
      r.results && r.results.length === 2 && r.results[0].phone !== r.results[1].phone,
      r.results && r.results.map(function (x) { return x.phone; }));

console.log('\n--- editor-only importer is not reachable over the web ---');
w = freshWorld('1234');
r = call(w, { action: 'auditEnrollments' });
check('auditEnrollments with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'auditEnrollments', pin: '1234' });
check('  ...and is not a routed action even WITH the pin',
      typeof r.error === 'string' && r.error.indexOf('Unknown action') === 0, r);

console.log('\n--- Status column ---');
w = freshWorld('1234');
r = call(w, { action: 'addEnrollment', pin: '1234',
              data: enc({ mode: 'legacy', studentName: 'Status Kid', phone: '7777777777' }) });
check('new enrolment is created', r.success === true, r);
var eRows = w.sheets.Enrollments._data;
var statusAt = eRows[0].indexOf('Status');
check('Status is a column on a fresh sheet', statusAt >= 0, eRows[0].length);
check('  ...and a new enrolment is marked Active',
      eRows[eRows.length - 1][statusAt] === 'Active', eRows[eRows.length - 1][statusAt]);
r = call(w, { action: 'searchStudents', q: 'Status Kid', pin: '1234' });
check('search reports the status', r.results && r.results[0] && r.results[0].status === 'Active',
      r.results && r.results[0]);

// The live sheet predates the column, so the migration has to add it.
w = freshWorld('1234');
var enr = w.sheets.Enrollments;
enr._data[0] = enr._data[0].filter(function (h) { return h !== 'Status'; });
var before = enr._data[0].length;
var at1 = w.sandbox.ensureStatusColumn_();
check('migration appends Status when missing', enr._data[0][at1] === 'Status', enr._data[0]);
check('  ...at the end, disturbing no existing column', at1 === before, [at1, before]);
var at2 = w.sandbox.ensureStatusColumn_();
check('  ...and is idempotent', at2 === at1 && enr._data[0].length === before + 1,
      [at1, at2, enr._data[0].length]);

console.log('\n--- one receipt, several students ---');
w = freshWorld('1234');
var recSheet = w.sheets.Receipts;
check('Receipts starts without a Students column',
      recSheet._data[0].indexOf('Students') < 0, recSheet._data[0]);
r = call(w, { action: 'addReceipt', pin: '1234',
              data: enc({ studentName: 'Riya Sen & Diya Sen',
                          students: ['Riya Sen', 'Diya Sen'],
                          amount: '3000', feeType: 'Monthly Fee' }) });
check('receipt is created', r.success === true && !!r.receiptNo, r);
var sAt = recSheet._data[0].indexOf('Students');
check('  ...the Students column is added on demand', sAt >= 0, recSheet._data[0]);
var last = recSheet._data[recSheet._data.length - 1];
check('  ...both names are stored separately', last[sAt] === 'Riya Sen | Diya Sen', last[sAt]);
check('  ...and Student Name keeps the combined display string',
      last[recSheet._data[0].indexOf('Student Name')] === 'Riya Sen & Diya Sen',
      last[recSheet._data[0].indexOf('Student Name')]);
check('  ...with one clubbed amount, not one per child',
      last[recSheet._data[0].indexOf('Amount (\u20b9)')] === '3000',
      last[recSheet._data[0].indexOf('Amount (\u20b9)')]);

// A single-student receipt must still behave exactly as before.
var widthBefore = recSheet._data[0].length;
r = call(w, { action: 'addReceipt', pin: '1234',
              data: enc({ studentName: 'Solo Student', amount: '1500' }) });
last = recSheet._data[recSheet._data.length - 1];
check('a single-student receipt still works', r.success === true, r);
check('  ...and falls back to the one name', last[sAt] === 'Solo Student', last[sAt]);
check('  ...without adding the column twice',
      recSheet._data[0].length === widthBefore, recSheet._data[0].length);

console.log('');
console.log('--- left students and rejoining students ---');
w = freshWorld('1234');
var eh0 = w.sheets.Enrollments._data[0];
w.sheets.Enrollments._data[1][eh0.indexOf('Status')] = 'Left';
r = call(w, { action: 'searchStudents', q: 'Test Child', pin: '1234' });
check('a student who has left is not offered for a receipt',
      !r.results || r.results.length === 0, r.results);

r = call(w, { action: 'addEnrollment', pin: '1234',
              data: enc({ mode: 'admission', studentName: 'Test Child', phone: '9999' }) });
check('the rejoiner gets a fresh row', r.success === true, r);
check('  ...flagged against the earlier Left record',
      (r.review || '').indexOf('Rejoining') === 0, r.review);
r = call(w, { action: 'searchStudents', q: 'Test Child', pin: '1234' });
check('  ...and only the new active row is offered',
      r.results && r.results.length === 1, r.results);

r = call(w, { action: 'addEnrollment', pin: '1234',
              data: enc({ mode: 'admission', studentName: 'Test Child', phone: '9999' }) });
check('a duplicate of an ACTIVE record reads differently',
      (r.review || '').indexOf('Possible duplicate') === 0, r.review);

r = call(w, { action: 'addEnrollment', pin: '1234',
              data: enc({ mode: 'admission', studentName: 'Nobody Else', phone: '5555' }) });
check('a genuinely new student is not flagged', !r.review, r.review);

var eh1 = w.sheets.Enrollments._data[0];
check('Left On column is created', eh1.indexOf('Left On') >= 0, eh1);
check('Review column is created', eh1.indexOf('Review') >= 0, eh1);

console.log('');
console.log('--- column alignment repair ---');
w = freshWorld('1234');
// The live sheet's real shape: Program, Batch and Pracheen were deleted, so
// three unheadered columns sit between Notes and Status.
var CENTRE = 'Bhawanipur - 53A Girish Mukherjee Road';
var liveHead = ['ID','Enrolled At','Type','Student Name','Date of Birth','Gender',
  'Blood Group','School/College','Guardian Name','Relation','Phone','WhatsApp',
  'Email','Address','Location','Joining Date/Approx Joining Month','Workshop Name',
  'Workshop Date','Workshop Fee','Heard From','Notes','','','','Status','Left On','Review'];
function blankRow() { var a = []; for (var i = 0; i < 27; i++) a.push(''); return a; }

// Correctly aligned (written before the columns were deleted).
var good = blankRow();
good[0] = 'SR0534405'; good[3] = 'Anisha Binaykia'; good[13] = 'Lansdowne';
good[14] = CENTRE; good[19] = 'Google Search'; good[20] = 'Wants Odissi';

// Written positionally AFTER the deletion, so values 15-24 overflowed.
var bad = blankRow();
bad[0] = 'SR-2026-0625182635913'; bad[3] = 'ANSHIKA SHOME'; bad[13] = 'Behala';
bad[15] = CENTRE;                  // Location landed in Joining Date
bad[17] = '2026-07-01';            // Joining Date landed in Workshop Date
bad[22] = 'Friend / Family Referral';  // Heard From landed past Notes
bad[23] = 'Please call evenings';      // Notes landed past that

// A legacy row, written by header name and therefore fine.
var legacy = blankRow();
legacy[0] = 'SR-LEGACY-1'; legacy[3] = 'Mahika Sen';
legacy[20] = 'Legacy roster'; legacy[24] = 'Active';

w.sheets.Enrollments = makeSheet([liveHead, good, bad, legacy]);
var plan = w.sandbox.buildColumnRepair_();
check('the repair accepts the real header layout', !plan.error, plan.error);
check('only the shifted row is selected',
      plan.fixes && plan.fixes.length === 1, plan.fixes && plan.fixes.length);
check('  ...and it is the right one',
      plan.fixes && plan.fixes[0].sheetRow === 3, plan.fixes && plan.fixes[0].sheetRow);

var after = plan.fixes[0].after;
check('Location moves back to O', after[0] === CENTRE, after[0]);
check('Joining Date moves back to P', after[1] === '2026-07-01', after[1]);
check('Heard From moves back to T', after[5] === 'Friend / Family Referral', after[5]);
check('Notes moves back to U', after[6] === 'Please call evenings', after[6]);
check('the unheadered columns are cleared',
      after[7] === '' && after[8] === '' && after[9] === '', after.slice(7));

var beforeVals = plan.fixes[0].before.filter(Boolean);
var kept = beforeVals.filter(function (v) { return after.indexOf(v) >= 0; });
check('NOTHING IS LOST - every value survives the move',
      kept.length === beforeVals.length, { before: beforeVals, after: after.filter(Boolean) });

// It must refuse to run against a sheet whose headers are not what it expects.
var wrong = liveHead.slice(); wrong[14] = 'Something Else';
w.sheets.Enrollments = makeSheet([wrong, good, bad]);
var refused = w.sandbox.buildColumnRepair_();
check('it refuses unexpected headers rather than guessing',
      !!refused.error, refused.error);

// A renamed header must not silently drop the value on new enrolments.
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([liveHead]);
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'admission', studentName: 'Alias Test', startDate: '2026-09-01' }) });
var newRow = w.sheets.Enrollments._data[1];
check('a renamed Joining Date header still receives its value',
      newRow[15] === '2026-09-01', newRow[15]);

console.log('');
console.log('--- dropping the leftover unheadered columns ---');
w = freshWorld('1234');
var hd = ['ID','Enrolled At','Type','Student Name','Date of Birth','Gender',
  'Blood Group','School/College','Guardian Name','Relation','Phone','WhatsApp',
  'Email','Address','Location','Joining Date/Approx Joining Month','Workshop Name',
  'Workshop Date','Workshop Fee','Heard From','Notes','','','','Status','Left On','Review'];
function mkRow() { var a = []; for (var i = 0; i < 27; i++) a.push(''); return a; }

// Still holding the misaligned values: must refuse.
var occupied = mkRow();
occupied[3] = 'ANSHIKA SHOME'; occupied[22] = 'Friend / Family Referral';
w.sheets.Enrollments = makeSheet([hd, occupied]);
var res = w.sandbox.dropEmptyOverflowColumns();
check('refuses while the columns still hold data', res.indexOf('REFUSED') === 0, res);
check('  ...and deletes nothing', w.sheets.Enrollments._data[0].length === 27,
      w.sheets.Enrollments._data[0].length);

// After the repair they are empty: safe to drop.
var clean = mkRow();
clean[3] = 'ANSHIKA SHOME'; clean[19] = 'Friend / Family Referral'; clean[24] = 'Active';
w.sheets.Enrollments = makeSheet([hd, clean]);
res = w.sandbox.dropEmptyOverflowColumns();
check('drops them once they are empty', res.indexOf('Dropped 3') === 0, res);
var h2 = w.sheets.Enrollments._data[0];
check('  ...leaving 24 columns', h2.length === 24, h2.length);
check('  ...with Status now next to Notes',
      h2[20] === 'Notes' && h2[21] === 'Status', h2.slice(19));
check('  ...and the row values still under the right headers',
      w.sheets.Enrollments._data[1][19] === 'Friend / Family Referral' &&
      w.sheets.Enrollments._data[1][21] === 'Active',
      w.sheets.Enrollments._data[1].slice(19));

console.log('');
console.log('--- starting month and year ---');
w = freshWorld('1234');
var jhd = ['ID','Enrolled At','Type','Student Name','Date of Birth','Gender',
  'Blood Group','School/College','Guardian Name','Relation','Phone','WhatsApp',
  'Email','Address','Location','Joining Date/Approx Joining Month','Workshop Name',
  'Workshop Date','Workshop Fee','Heard From','Notes','Status','Left On','Review'];
w.sheets.Enrollments = makeSheet([jhd]);
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'admission', studentName: 'Start Kid',
                      startMonth: 'September', startYear: '2026' }) });
var jr = w.sheets.Enrollments._data[1];
check('month and year land in the joining column',
      jr[15] === 'September 2026', jr[15]);
check('  ...and status still defaults to Active', jr[21] === 'Active', jr[21]);

// Anything still sending the old fields must keep working.
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'Old Field Kid',
                      approxJoining: '2024-04' }) });
jr = w.sheets.Enrollments._data[2];
check('the older single-date field is still honoured', jr[15] === '2024-04', jr[15]);

w.sheets.Enrollments = makeSheet([jhd]);
call(w, { data: enc({ mode: 'app-admission', studentName: 'Public Kid' }),
          action: 'addEnrollment' });
jr = w.sheets.Enrollments._data[1];
check('a public application defaults to the month it was submitted',
      /^[A-Z][a-z]+ [0-9]{4}$/.test(jr[15]), jr[15]);
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'No Guess Kid' }) });
jr = w.sheets.Enrollments._data[2];
check('an Existing Student is left blank rather than guessed', jr[15] === '', jr[15]);

console.log('');
console.log('--- joining month backfill ---');
w = freshWorld('1234');
var bh = ['ID','Enrolled At','Type','Student Name','Date of Birth','Gender',
  'Blood Group','School/College','Guardian Name','Relation','Phone','WhatsApp',
  'Email','Address','Location','Joining Date/Approx Joining Month','Workshop Name',
  'Workshop Date','Workshop Fee','Heard From','Notes','Status','Left On','Review'];
function bRow(id, when, type, name, join) {
  var a = []; for (var i = 0; i < 24; i++) a.push('');
  a[0] = id; a[1] = when; a[2] = type; a[3] = name; a[15] = join || '';
  return a;
}
w.sheets.Enrollments = makeSheet([bh,
  bRow('SR-2026-1', new Date(2026, 5, 8), 'Application - Admission', 'Blank Kid', ''),
  bRow('SR-2026-2', new Date(2026, 5, 9), 'New Admission', 'Has One', 'March 2026'),
  bRow('SR-LEGACY-1', new Date(2026, 7, 22), 'Existing Student', 'Legacy Kid', ''),
  bRow('SR-2026-3', 'not a date', 'New Admission', 'Bad Date', '')
]);
var bf = w.sandbox.buildJoiningBackfill_();
check('only the blank non-legacy row is filled',
      bf.fills.length === 1 && bf.fills[0].name === 'Blank Kid',
      bf.fills.map(function (f) { return f.name; }));
check('  ...from its enrolment date', bf.fills[0].value === 'January 2026', bf.fills[0].value);
check('a row that already has a value is left alone',
      bf.alreadySet.length === 1 && bf.alreadySet[0] === 'Has One', bf.alreadySet);
check('the legacy roster is skipped',
      bf.skippedLegacy.length === 1 && bf.skippedLegacy[0] === 'Legacy Kid', bf.skippedLegacy);
check('an unreadable date is reported, not guessed',
      bf.unreadable.length === 1 && bf.unreadable[0].name === 'Bad Date', bf.unreadable);

console.log('');
console.log('--- listStudents (autocomplete source) ---');
w = freshWorld('1234');
r = call(w, { action: 'listStudents' });
check('listStudents needs the PIN', r.auth === true, r);

call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'Krisha Agarwal', phone: '9111111111' }) });
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'Krisha Agarwal', phone: '9222222222' }) });
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'Gone Away', phone: '9333333333' }) });
var ed = w.sheets.Enrollments._data;
ed[ed.length - 1][ed[0].indexOf('Status')] = 'Left';

r = call(w, { action: 'listStudents', pin: '1234' });
var names = (r.students || []).map(function (x) { return x.studentName; });
check('returns the active roster', names.indexOf('Test Child') >= 0, names);
check('a student who has left is excluded', names.indexOf('Gone Away') < 0, names);
var krishas = (r.students || []).filter(function (x) { return x.studentName === 'Krisha Agarwal'; });
check('both same-name students are returned', krishas.length === 2, krishas.length);
check('  ...with different phones',
      krishas.length === 2 && krishas[0].phone !== krishas[1].phone,
      krishas.map(function (x) { return x.phone; }));

call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'Test Child', phone: '9999' }) });
r = call(w, { action: 'listStudents', pin: '1234' });
var dupes = (r.students || []).filter(function (x) { return x.studentName === 'Test Child'; });
check('an identical name and phone is listed once', dupes.length === 1, dupes.length);

console.log('');
console.log('--- duplicate detection and clubbed-receipt coverage ---');
w = freshWorld('1234');
check('normName_ collapses case and stray spaces',
      w.sandbox.normName_('ANSHIKA  SHOME') === w.sandbox.normName_('Anshika Shome'),
      [w.sandbox.normName_('ANSHIKA  SHOME'), w.sandbox.normName_('Anshika Shome')]);

var dh = ['ID','Enrolled At','Type','Student Name','Date of Birth','Gender',
  'Blood Group','School/College','Guardian Name','Relation','Phone','WhatsApp',
  'Email','Address','Location','Joining Date','Workshop Name','Workshop Date',
  'Workshop Fee','Heard From','Notes','Status','Left On','Review'];
function dRow(id, name, phone, status) {
  var a = []; for (var i = 0; i < 24; i++) a.push('');
  a[0] = id; a[2] = 'Existing Student'; a[3] = name; a[10] = phone || '';
  a[21] = status || 'Active'; return a;
}
w.sheets.Enrollments = makeSheet([dh,
  dRow('SR-1', 'ANSHIKA  SHOME', ''),
  dRow('SR-2', 'Anshika Shome', ''),
  dRow('SR-3', 'Krisha Agarwal', '9111111111'),
  dRow('SR-4', 'Krisha Agarwal', '9222222222'),
  dRow('SR-5', 'Diya Sen', '')
]);
var dup = w.sandbox.findDuplicateStudents();
check('the two Anshika spellings are flagged as one child',
      dup.indexOf('Likely duplicates      : 1') >= 0, dup.split('\n')[3]);
check('the two Krishas are kept apart as real students',
      dup.indexOf('Same name, different people : 1') >= 0, dup.split('\n')[4]);

// A sibling named only inside a clubbed receipt must not read as unpaid.
w.sheets.Receipts = makeSheet([
  ['Receipt No','Issued At','Student Name','Contact','Amount (\u20b9)','Fee Month',
   'Fee Year','Payment Mode','UPI Reference','Fee Type','Date Received','Note'],
  ['SS-2026-0001','01 Aug 2026','Riya Sen & Diya Sen','9999','3000','August',
   '2026','Cash','','Monthly Fee','','']
]);
var rep = w.sandbox.analyticsReport();
check('a sibling inside a clubbed receipt counts as paid',
      rep.indexOf('clubbed receipt  1') >= 0,
      rep.split('\n').filter(function (l) { return l.indexOf('clubbed receipt') >= 0; }));
check('  ...and is not listed as never invoiced',
      rep.indexOf('Diya Sen') < rep.indexOf('ACTIVE BUT NEVER INVOICED') ||
      rep.indexOf('ACTIVE BUT NEVER INVOICED') < 0,
      rep.indexOf('ACTIVE BUT NEVER INVOICED'));

console.log('');
console.log('--- unbilled review ---');
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  dRow('SR-1', 'Tashvi Kocahr', ''),
  dRow('SR-2', 'Never Paid', ''),
  dRow('SR-3', 'Diya Sen', ''),
  dRow('SR-4', 'Gone Away', '', 'Left')
]);
w.sheets.Receipts = makeSheet([
  ['Receipt No','Issued At','Student Name','Contact','Amount (₹)','Fee Month',
   'Fee Year','Payment Mode','UPI Reference','Fee Type','Date Received','Note'],
  ['SS-1','01 Aug 2026','Tashvi Kochar','9','2000','August','2026','Cash','','Monthly Fee','',''],
  ['SS-2','01 Aug 2026','Riya Sen & Diya Sen','9','3000','August','2026','Cash','','Monthly Fee','','']
]);
var ub = w.sandbox.reviewUnbilledStudents();
check('a transposed surname is called a spelling difference',
      ub.indexOf('Probably a spelling difference               : 1') >= 0,
      'see the log');
check('  ...and names the receipt spelling', ub.indexOf('tashvi kochar') >= 0, 'no suggestion');
check('a genuine non-payer is separated out',
      ub.indexOf('genuinely never invoiced    : 1') >= 0, 'see the log');
check('a clubbed sibling is not listed at all',
      ub.indexOf('Diya Sen') < 0, 'Diya Sen wrongly listed');
check('a student who has left is excluded',
      ub.indexOf('Gone Away') < 0, 'Gone Away wrongly listed');

console.log('');
console.log('--- suggestions must not merge two real students ---');
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  dRow('SR-1', 'Ananya Jain', ''),
  dRow('SR-2', 'Anaaya Jain', ''),
  dRow('SR-3', 'Tashvi Kocahr', ''),
  dRow('SR-4', 'Riya Sen', ''),
  dRow('SR-5', 'Diya Sen', '')
]);
w.sheets.Receipts = makeSheet([
  ['Receipt No','Issued At','Student Name','Contact','Amount (₹)','Fee Month',
   'Fee Year','Payment Mode','UPI Reference','Fee Type','Date Received','Note'],
  ['SS-1','01 Aug 2026','Anaaya Jain','9','2000','August','2026','Cash','','Monthly Fee','',''],
  ['SS-2','01 Aug 2026','Tashvi Kochar','9','2000','August','2026','Cash','','Monthly Fee','',''],
  ['SS-3','01 Aug 2026','Riya & Diya Sen','9','3000','August','2026','Cash','','Monthly Fee','','']
]);
var rv = w.sandbox.reviewUnbilledStudents();
check('refuses to suggest another enrolled student as a typo',
      rv.indexOf('Probably a spelling difference               : 1') >= 0, 'see log');
check('  ...and the real transposition is still caught',
      rv.indexOf('tashvi kochar') >= 0, 'Tashvi suggestion missing');
check('  ...while Ananya is flagged as too close to call',
      rv.indexOf('Too close to another student to call         : 1') >= 0, 'see log');
check('a shared-surname clubbed receipt covers BOTH siblings',
      rv.indexOf('Riya Sen') < 0 && rv.indexOf('Diya Sen') < 0,
      'a Sen sibling was wrongly listed as unbilled');

console.log('');
console.log('--- duplicate receipts ---');
w = freshWorld('1234');
var rhd = ['Receipt No','Issued At','Student Name','Contact','Amount (₹)',
  'Fee Month','Fee Year','Payment Mode','UPI Reference','Fee Type','Date Received',
  'Note','Students'];
function rRow(no, name, amt, mon, type, students) {
  return [no,'01 Aug 2026',name,'9',amt,mon,'2026','Cash','',type,'','',students||''];
}
w.sheets.Receipts = makeSheet([rhd,
  rRow('SS-1','Aarna Vadera','2000','August','Monthly Fee'),
  rRow('SS-2','Aarna Vadera','2000','August','Monthly Fee'),
  rRow('SS-3','Bela Roy','2000','August','Monthly Fee'),
  rRow('SS-4','Bela Roy','2500','August','Monthly Fee'),
  rRow('SS-5','Cara Das','2000','August','Monthly Fee'),
  rRow('SS-6','Cara Das','500','August','Registration Fee'),
  rRow('SS-7','Riya Sen & Diya Sen','3000','August','Monthly Fee','Riya Sen | Diya Sen'),
  rRow('SS-8','Diya Sen & Riya Sen','3000','August','Monthly Fee','Diya Sen | Riya Sen')
]);
var dr = w.sandbox.findDuplicateReceipts();
check('an identical repeat is flagged',
      dr.indexOf('Submitted more than once') >= 0, 'see log');
check('  ...including a clubbed pair listed in either order',
      dr.indexOf('diya sen + riya sen') >= 0, 'clubbed pair not grouped');
check('a differing amount is called a possible reissue',
      dr.indexOf('DIFFERING AMOUNTS FOR THE SAME PERIOD') >= 0, 'see log');
check('a registration fee alongside monthly is not called a duplicate',
      dr.indexOf('DIFFERENT FEE TYPE') >= 0, 'see log');
check('the revenue at stake is totalled',
      dr.indexOf('Most likely over-counted') >= 0, dr.split(String.fromCharCode(10))[5]);

console.log('');
console.log('--- duplicate receipt classification ---');
w = freshWorld('1234');
function rRow2(no, name, amt, mon, mode, note, when) {
  return [no, when || '01 Aug 2026', name, '9', amt, mon, '2026', mode, '',
          'Monthly Fee', '', note || '', ''];
}
w.sheets.Receipts = makeSheet([rhd,
  rRow2('SS-2026-0100','Repeat Kid','2000','August','UPI','Fee for August'),
  rRow2('SS-2026-0101','Repeat Kid','2000','August','UPI','Fee for August'),
  rRow2('SS-2026-0102','Repeat Kid','2000','August','UPI','Fee for August'),
  rRow2('SS-2026-0200','Mode Kid','1800','August','Cash','Fee for August'),
  rRow2('SS-2026-0201','Mode Kid','1800','August','UPI','Fee for August'),
  rRow2('SS-2026-0300','Miskey Kid','2000','July','Cash','Fee for July','15 Jul 2026'),
  rRow2('SS-2026-0400','Miskey Kid','2000','July','Cash','Fee for August','28 Jul 2026')
]);
var dc = w.sandbox.findDuplicateReceipts();
check('three consecutive same-day receipts read as a repeated submission',
      dc.indexOf('Submitted more than once') >= 0 &&
      dc.indexOf('SUBMITTED MORE THAN ONCE') >= 0, 'see log');
check('a Cash/UPI pair reads as a mode correction, not a repeat',
      dc.indexOf('REISSUED WITH A DIFFERENT PAYMENT MODE') >= 0, 'see log');
check('notes naming different months are NOT called duplicates',
      dc.indexOf('PERIOD MIS-KEYED - BOTH PAYMENTS ARE REAL') >= 0, 'see log');
var over = dc.substring(dc.indexOf('Most likely over-counted'));
check('  ...and are excluded from the money at stake',
      over.indexOf('Rs. 5,800') >= 0, over.split(String.fromCharCode(10))[0]);
check('dates are formatted, not raw Date objects',
      dc.indexOf('GMT+') < 0, 'raw dates still present');

console.log('');
console.log('--- receipt retries are idempotent ---');
w = freshWorld('1234');
var payload = { studentName: 'Divitaa Taparia', students: ['Divitaa Taparia'],
                amount: '1900', month: 'August', year: '2026', feeType: 'Monthly Fee' };
r = call(w, { action: 'addReceipt', pin: '1234', data: enc(payload) });
var firstNo = r.receiptNo;
check('the first receipt is created', r.success === true && !!firstNo, r);
check('  ...and is not flagged as a repeat', !r.duplicate, r);
var rowsAfterFirst = w.sheets.Receipts._data.length;

var again = call(w, { action: 'addReceipt', pin: '1234', data: enc(payload) });
check('a retry does not mint a second number',
      again.receiptNo === firstNo, [firstNo, again.receiptNo]);
check('  ...is marked as a duplicate for the panel', again.duplicate === true, again);
check('  ...and writes no extra row',
      w.sheets.Receipts._data.length === rowsAfterFirst,
      [rowsAfterFirst, w.sheets.Receipts._data.length]);

// A genuinely different payment must still go through.
var other = call(w, { action: 'addReceipt', pin: '1234',
  data: enc({ studentName: 'Divitaa Taparia', students: ['Divitaa Taparia'],
              amount: '500', month: 'August', year: '2026', feeType: 'Monthly Fee' }) });
check('a different amount still creates a receipt',
      other.receiptNo !== firstNo && !other.duplicate, other);

// Sibling order must not defeat the check.
var sib = { studentName: 'A & B', students: ['B Kid', 'A Kid'], amount: '3000',
            month: 'August', year: '2026', feeType: 'Monthly Fee' };
var s1 = call(w, { action: 'addReceipt', pin: '1234', data: enc(sib) });
sib.students = ['A Kid', 'B Kid'];
var s2 = call(w, { action: 'addReceipt', pin: '1234', data: enc(sib) });
check('a clubbed retry in the other order is still caught',
      s2.receiptNo === s1.receiptNo, [s1.receiptNo, s2.receiptNo]);

console.log('');
console.log('--- two children sharing a name ---');
w = freshWorld('1234');
function rRow3(no, name, contact, amt, mon, note) {
  return [no, '02 Jul 2026', name, contact, amt, mon, '2026', 'Cash', '',
          'Monthly Fee', '', note || '', ''];
}
w.sheets.Receipts = makeSheet([rhd,
  rRow3('SS-2026-0175','Krisha Agarwal','9830014153','2000','July','Fee for July'),
  rRow3('SS-2026-0472','Krisha Agarwal','9831439849','2400','July','Fee for July'),
  rRow3('SS-2026-0500','Solo Kid','9800000000','2000','July','Fee for July'),
  rRow3('SS-2026-0501','Solo Kid','9800000000','1800','July','Fee for July')
]);
var nk = w.sandbox.findDuplicateReceipts();
check('two contacts under one name are not called duplicates',
      nk.indexOf('SAME NAME, DIFFERENT CONTACT NUMBERS') >= 0, 'see log');
check('  ...and contribute nothing to the money at stake',
      nk.indexOf('Same name, different family - NOT dup 1 grp') >= 0 ||
      nk.indexOf('different family') >= 0, 'see log');
check('one family with two amounts is still flagged',
      nk.indexOf('DIFFERING AMOUNTS FOR THE SAME PERIOD') >= 0, 'see log');
check('the contact is shown so the split is visible',
      nk.indexOf('9830014153') >= 0 && nk.indexOf('9831439849') >= 0, 'contacts missing');

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);
