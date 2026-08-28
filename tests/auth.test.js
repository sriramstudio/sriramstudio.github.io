// Runs Code.gs against stubbed Apps Script globals to test the auth gate.
const fs = require('fs'), vm = require('vm');
const SRC = process.argv[2];

function makeSheet(rows) {
  const data = rows.map(r => r.slice());
  const notes = {};
  const chain = { setFontWeight: () => chain, setBackground: () => chain, setFontColor: () => chain };
  return {
    _data: data,
    _notes: notes,
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
        setFontWeight: () => range, setBackground: () => range, setFontColor: () => range,
        getNote: () => notes[r + ':' + c] || '',
        setNote: v => { notes[r + ':' + c] = v; return range; }
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
      rep.indexOf('Have at least one receipt       1') >= 0,
      rep.split('\n').filter(function (l) { return l.indexOf('at least one receipt') >= 0; }));
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

console.log('');
console.log('--- fee coverage by month ---');
// Anchored to today rather than a fixed month, so these still mean something
// next year.
var MN = ['January','February','March','April','May','June','July','August',
          'September','October','November','December'];
var nowD  = new Date();
var prevD = new Date(nowD.getFullYear(), nowD.getMonth() - 1, 1);
var backD = new Date(nowD.getFullYear(), nowD.getMonth() - 3, 1);
var curM  = MN[nowD.getMonth()],  curY  = String(nowD.getFullYear());
var prevM = MN[prevD.getMonth()], prevY = String(prevD.getFullYear());
var backM = MN[backD.getMonth()], backY = String(backD.getFullYear());
var prevLabel = prevM + ' ' + prevY;
var backLabel = backM + ' ' + backY;

// dh column positions: 0 ID, 2 Type, 3 Student Name, 10 Phone, 14 Location,
// 15 Joining Date, 21 Status, 22 Left On.
function cRow(id, name, joining, status, leftOn, phone, type) {
  var a = []; for (var i = 0; i < 24; i++) a.push('');
  a[0] = id; a[2] = type || 'Existing Student'; a[3] = name; a[10] = phone || '';
  a[14] = 'Bhawanipur'; a[15] = joining || backLabel;
  a[21] = status || 'Active'; a[22] = leftOn || '';
  return a;
}
function cRcpt(no, printed, students, mon, yr, type, amt, contact) {
  return [no, '01 ' + mon.substring(0, 3) + ' ' + yr, printed, contact || '9000000001',
          amt || '2000', mon, yr, 'Cash', '', type || 'Monthly Fee', '', '', students || ''];
}

w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Riya Sen'),
  cRow('SR-2', 'Diya Sen'),
  cRow('SR-3', 'Vani Maskara'),
  cRow('SR-4', 'Jia Bhimani'),
  cRow('SR-5', 'Jeena Bhimani'),
  cRow('SR-6', 'Tashvi Kochar'),
  cRow('SR-7', 'Late Joiner', curM + ' ' + curY),
  cRow('SR-8', 'Gone Away', backLabel, 'Left', backLabel),
  cRow('SR-9', 'Workshop Only', backLabel, 'Active', '', '', 'Workshop')
]);
w.sheets.Receipts = makeSheet([rhd,
  // Siblings clubbed onto one receipt, the modern way.
  cRcpt('SS-1', 'Riya Sen & Diya Sen', 'Riya Sen | Diya Sen', prevM, prevY),
  // The same pairing the old way: one cell, only the last child has a surname.
  cRcpt('SS-2', 'Jia & Jeena Bhimani', '', prevM, prevY),
  // Spelled differently from the roster.
  cRcpt('SS-3', 'Vaani Maskara', 'Vaani Maskara', prevM, prevY),
  cRcpt('SS-4', 'Tashvi Kocahr', 'Tashvi Kocahr', prevM, prevY),
  // Not a month's tuition, so it must not cover anybody.
  cRcpt('SS-5', 'Late Joiner', 'Late Joiner', prevM, prevY, 'Registration Fee', '1000'),
  // A name that is on no roster row at all.
  cRcpt('SS-6', 'Someone Elsewhere', 'Someone Elsewhere', prevM, prevY)
]);

var mc = w.sandbox.feeCoverageForMonth(prevLabel);
check('both siblings on a clubbed receipt count as paid',
      mc.indexOf('Riya Sen') > mc.indexOf('PAID (') &&
      mc.indexOf('Diya Sen') > mc.indexOf('PAID ('), 'see log');
check('a legacy "Jia & Jeena Bhimani" cell pays for both children',
      mc.indexOf('Jia Bhimani') > mc.indexOf('PAID (') &&
      mc.indexOf('Jeena Bhimani') > mc.indexOf('PAID ('), 'see log');
check('Vaani Maskara on a receipt is Vani Maskara on the roster',
      mc.indexOf('NO RECEIPT FOR') < 0 || mc.indexOf('Vani Maskara') > mc.indexOf('PAID ('),
      'see log');
check('a one-character misspelling still counts as paid',
      mc.indexOf('Tashvi Kochar') > mc.indexOf('PAID ('), 'see log');
check('a registration fee does not cover the month',
      mc.indexOf('Late Joiner') < 0 || mc.indexOf('Late Joiner') > mc.indexOf('NO RECEIPT FOR'),
      'registration fee counted as monthly');
// The six due are Riya, Diya, Vani, Jia, Jeena and Tashvi: the late joiner,
// the student who left and the workshop attendee are all out of scope.
var mcPaid = mc.substring(mc.indexOf('PAID ('));
check('a student who joins later is not due in an earlier month',
      mc.indexOf('Students due to pay              6') >= 0 &&
      mcPaid.indexOf('Late Joiner') < 0,
      mc.split('\n').filter(function (l) { return l.indexOf('due to pay') >= 0; }));
check('a workshop-only enrolment is not billed monthly',
      mc.indexOf('Workshop Only') < 0, 'workshop attendee counted as a monthly student');
check('a student who left before the month is not counted',
      mcPaid.indexOf('Gone Away') < 0 && mc.indexOf('NO RECEIPT FOR') < 0,
      'left student still expected');
check('everyone due that month is covered',
      mc.indexOf('No receipt                       0') >= 0,
      mc.split('\n').filter(function (l) { return l.indexOf('No receipt') >= 0; }));
check('a receipt name matching no roster row is surfaced, not swallowed',
      mc.indexOf('Someone Elsewhere') >= 0, 'unknown name not reported');
check('  ...and holds the month back from being declared complete',
      mc.indexOf('ARE PAID FOR') < 0,
      'a month with an unaccounted-for receipt name was called complete');

var gaps = w.sandbox.feeGapsByStudent();
check('a student with no monthly receipt at all is listed as never invoiced',
      gaps.indexOf('NEVER INVOICED') >= 0 && gaps.indexOf('Late Joiner') >= 0,
      gaps.split('\n').filter(function (l) { return l.indexOf('Never had') >= 0; }));

// Drop the one unaccounted-for name and the month should read as settled.
w.sheets.Receipts._data.pop();
var clean = w.sandbox.feeCoverageForMonth(prevLabel);
check('with every receipt name accounted for, the month reads as fully paid',
      clean.indexOf('ARE PAID FOR') >= 0,
      clean.split('\n').filter(function (l) { return l.indexOf('No receipt') >= 0; }));

// Jia and Jeena are two children, not one spelled two ways.
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-10', 'Jia Bhimani', 'Jia Bhimani', prevM, prevY)
]);
var solo = w.sandbox.feeCoverageForMonth(prevLabel);
check('Jia paying does not mark Jeena paid',
      solo.indexOf('Jeena Bhimani') > solo.indexOf('NO RECEIPT FOR'),
      solo.split('\n').filter(function (l) { return l.indexOf('Jeena') >= 0; }));
check('  ...and Jia herself is still credited',
      solo.indexOf('Jia Bhimani') > solo.indexOf('PAID ('), 'see log');
check('an incomplete month is not reported as fully paid',
      solo.indexOf('ARE PAID FOR') < 0, 'incomplete month called complete');

var pv = w.sandbox.previewReceiptNameMatching();
// Scoped to the pair this fixture actually carries: the configured lists
// describe the real sheet, not the test roster.
check('the preview binds each configured distinct student to its own row',
      pv.indexOf('"Jia Bhimani" -> row') >= 0 && pv.indexOf('"Jeena Bhimani" -> row') >= 0,
      pv.split('\n').filter(function (l) { return l.indexOf('Bhimani"') >= 0; }));
check('the preview shows which spelling the equivalents resolved to',
      pv.indexOf('Vani Maskara') >= 0, 'equivalence not reported');

// A sibling who is missing from the roster must not be credited to the one
// sibling who is on it.
w.sheets.Enrollments = makeSheet([dh, cRow('SR-2', 'Diya Sen')]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-20', 'Riya Sen & Diya Sen', 'Riya Sen | Diya Sen', prevM, prevY)
]);
var clash = w.sandbox.previewReceiptNameMatching();
check('a sibling missing from the roster is not credited to the other one',
      clash.indexOf('NO ROSTER ROW AT ALL') >= 0 && clash.indexOf('"riya sen"') >= 0,
      'the missing sibling was silently absorbed');

var tbl = w.sandbox.feeCoverageByMonth();
check('the headline table reports a month per row',
      tbl.indexOf('FEE COVERAGE BY MONTH') >= 0 &&
      tbl.indexOf(MN[prevD.getMonth()].substring(0, 3) + ' ' + prevY) >= 0, 'see log');
check('  ...and nothing in it writes to the sheet',
      tbl.indexOf('Nothing was changed') >= 0, 'missing the read-only notice');

console.log('');
console.log('--- names run together, and first names alone ---');
// Every fixture here is a real pattern from the studio's receipts.
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Anshika Shome'),
  cRow('SR-2', 'Anvika Shome'),
  cRow('SR-3', 'Nyra Bharuka'),
  cRow('SR-4', 'Nyshita Bharuka'),
  cRow('SR-5', 'Jia Bhimani'),
  cRow('SR-6', 'Jeena Bhimani'),
  cRow('SR-7', 'Hiral Redh'),
  cRow('SR-8', 'Jia Doshi'),
  cRow('SR-9', 'Solo Child')
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-1', 'Anshika Anvika Shome', '', prevM, prevY),
  cRcpt('SS-2', 'Nyra Nyshita Bharuka', '', prevM, prevY),
  cRcpt('SS-3', 'Jeena Jia', '', prevM, prevY),
  cRcpt('SS-4', 'Hiral', '', prevM, prevY)
]);
var sp = w.sandbox.previewReceiptNameMatching();
check('two sisters written as one name are read as two children',
      sp.indexOf('Anshika Shome  +  Anvika Shome') >= 0 ||
      sp.indexOf('Anvika Shome  +  Anshika Shome') >= 0,
      sp.split('\n').filter(function (l) { return l.indexOf('anshika') >= 0; }));
check('  ...and so are two more with a different surname',
      sp.indexOf('Nyra Bharuka') >= 0 && sp.indexOf('Nyshita Bharuka') >= 0, 'see log');
check('"Jeena Jia" with no surname at all resolves to both Bhimanis',
      sp.indexOf('Jeena Bhimani') >= 0 && sp.indexOf('Jia Bhimani') >= 0,
      sp.split('\n').filter(function (l) { return l.indexOf('jeena jia') >= 0; }));
check('  ...and does not grab the other Jia on the roster',
      sp.indexOf('Jia Doshi') < 0, 'the wrong Jia was credited');
check('a first name on its own matches the only child who answers to it',
      sp.indexOf('MATCHED ON FIRST NAME ALONE') >= 0 && sp.indexOf('Hiral Redh') >= 0,
      'see log');

var spm = w.sandbox.feeCoverageForMonth(prevLabel);
// Four receipts, seven children: two Shomes, two Bharukas, two Bhimanis, Hiral.
check('all seven children named across those four receipts count as paid',
      spm.indexOf('Covered by a receipt             7') >= 0,
      spm.split('\n').filter(function (l) { return l.indexOf('Covered by') >= 0; }));
check('  ...and the two nobody paid for are still listed',
      spm.indexOf('Solo Child') > spm.indexOf('NO RECEIPT FOR') &&
      spm.indexOf('Jia Doshi') > spm.indexOf('NO RECEIPT FOR'), 'see log');

// A first name two children answer to is a guess, so it must not be made.
w.sheets.Receipts = makeSheet([rhd, cRcpt('SS-5', 'Jia', '', prevM, prevY)]);
var amb = w.sandbox.previewReceiptNameMatching();
check('a first name two children share is credited to neither',
      amb.indexOf('MATCHED ON FIRST NAME ALONE') < 0, 'a shared first name was guessed');

// Duplicate roster rows: say so, rather than just "too close to call".
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Krisha Agarwal', '', 'Active', '', '9111111111'),
  cRow('SR-2', 'Krisha Agarwal', '', 'Active', '', '9222222222')
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-6', 'Krisha Agarwal', 'Krisha Agarwal', prevM, prevY, 'Monthly Fee', '2000', '9333333333')
]);
var why = w.sandbox.previewReceiptNameMatching();
check('an unsplittable shared name says why it could not be called',
      why.indexOf('rows in Enrollments carry this exact name') >= 0,
      why.split('\n').filter(function (l) { return l.indexOf('why:') >= 0; }));

// The contact number on the receipt is enough to tell them apart.
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-7', 'Krisha Agarwal', 'Krisha Agarwal', prevM, prevY, 'Monthly Fee', '2000', '9222222222')
]);
var byph = w.sandbox.previewReceiptNameMatching();
check('  ...but a matching contact number does tell them apart',
      byph.indexOf('by phone') >= 0, 'contact number ignored');

console.log('');
console.log('--- a near-miss between two real children ---');
// 'Janvi Jain' is one character from both Jahnvi and Janhvi, who are two
// different children. Nothing but a contact number can settle it.
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Jahnvi Jain', '', 'Active', '', '9111111111'),
  cRow('SR-2', 'Janhvi Jain', '', 'Active', '', '9222222222')
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-1', 'Janvi Jain', 'Janvi Jain', prevM, prevY, 'Monthly Fee', '2000', '9000000009')
]);
var tie = w.sandbox.previewReceiptNameMatching();
check('a name one character from two children is credited to neither',
      tie.indexOf('TOO CLOSE TO CALL               1') >= 0,
      tie.split('\n').filter(function (l) { return l.indexOf('TOO CLOSE') >= 0; }));
check('  ...and both candidates are named with their rows',
      tie.indexOf('Jahnvi Jain') >= 0 && tie.indexOf('Janhvi Jain') >= 0 &&
      tie.indexOf('row 2') >= 0 && tie.indexOf('row 3') >= 0,
      tie.split('\n').filter(function (l) { return l.indexOf('why:') >= 0; }));

// The same receipt, now carrying one family's number.
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-2', 'Janvi Jain', 'Janvi Jain', prevM, prevY, 'Monthly Fee', '2000', '9222222222')
]);
var settled = w.sandbox.previewReceiptNameMatching();
check('a contact number settles the tie',
      settled.indexOf('by phone') >= 0 &&
      settled.indexOf('TOO CLOSE TO CALL               0') >= 0,
      settled.split('\n').filter(function (l) { return l.indexOf('janvi') >= 0; }));
check('  ...to the family whose number it is',
      settled.indexOf('-> Janhvi Jain') >= 0, 'credited to the wrong child');

// The month must NOT be used to break a tie the name itself cannot settle.
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Jahnvi Jain', backLabel, 'Left', backLabel, '9111111111'),
  cRow('SR-2', 'Janhvi Jain', '', 'Active', '', '9222222222')
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-3', 'Janvi Jain', 'Janvi Jain', prevM, prevY, 'Monthly Fee', '2000', '')
]);
var noguess = w.sandbox.previewReceiptNameMatching();
check('one child having left does not settle a spelling tie',
      noguess.indexOf('TOO CLOSE TO CALL               1') >= 0,
      'the enrolled child was picked on no real evidence');

console.log('');
console.log('--- the sheet tabs ---');
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Paid Up'),
  cRow('SR-2', 'Behind Child'),
  cRow('SR-3', 'Gone Away', backLabel, 'Left', backLabel),
  cRow('SR-4', 'Workshop Only', backLabel, 'Active', '', '', 'Workshop')
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-1', 'Paid Up', 'Paid Up', backM, backY),
  cRcpt('SS-2', 'Paid Up', 'Paid Up', prevM, prevY),
  cRcpt('SS-3', 'Behind Child', 'Behind Child', backM, backY),
  cRcpt('SS-4', 'Gone Away', 'Gone Away', backM, backY)
]);
var model = w.sandbox.analyticsTabModel_(w.sandbox.buildFeeCoverage_());

check('the students tab carries one row per monthly-fee student',
      model.students.rows.length === 3,
      model.students.rows.map(function (r) { return r[2]; }));
check('  ...and leaves the workshop attendee out',
      model.students.rows.every(function (r) { return r[2] !== 'Workshop Only'; }),
      'workshop attendee included');

var monthCols = model.students.head.slice(model.firstMonthCol - 1);
check('  ...with a column per month, newest last',
      monthCols[monthCols.length - 1] === curM + ' ' + curY,
      [monthCols[0], monthCols[monthCols.length - 1]]);

var behind = model.students.rows.filter(function (r) { return r[2] === 'Behind Child'; })[0];
var prevIdx = model.firstMonthCol - 1 + monthCols.indexOf(prevLabel);
check('a month with no receipt is marked UNPAID',
      behind[prevIdx] === 'UNPAID', behind[prevIdx]);
var paidUp = model.students.rows.filter(function (r) { return r[2] === 'Paid Up'; })[0];
check('  ...and a month with one is marked PAID',
      paidUp[prevIdx] === 'PAID', paidUp[prevIdx]);
var gone = model.students.rows.filter(function (r) { return r[2] === 'Gone Away'; })[0];
check('  ...and a month after the student left is left blank',
      gone[prevIdx] === '', gone[prevIdx]);

check('the students tab is sorted with the worst arrears first',
      model.students.rows[0][12] >= model.students.rows[1][12],
      model.students.rows.map(function (r) { return [r[2], r[12]]; }));

check('the coverage tab holds one row per month',
      model.coverage.rows.length === model.monthNames.length, model.coverage.rows.length);
check('  ...with coverage as a fraction, for percent formatting',
      model.coverage.rows.every(function (r) { return r[4] >= 0 && r[4] <= 1; }),
      model.coverage.rows.map(function (r) { return r[4]; }));

check('the gaps tab lists only students missing a month',
      model.gaps.rows.every(function (r) { return r[6] > 0; }),
      model.gaps.rows.map(function (r) { return [r[1], r[6]]; }));
check('  ...and spells out which months',
      model.gaps.rows.length > 0 && model.gaps.rows[0][8].length > 0, 'no month list');

check('the names tab puts the ones needing a decision at the top',
      model.names.rows.length > 0, 'no names');

check('the dashboard opens on the last month that should be finished',
      model.defaultMonth === prevLabel, [model.defaultMonth, prevLabel]);

console.log('');
console.log('--- the analytics tabs cannot eat the data tabs ---');
var ss = w.sandbox.SpreadsheetApp.getActiveSpreadsheet();
['Enrollments', 'Receipts', 'Config'].forEach(function (name) {
  var threw = false;
  try { w.sandbox.analyticsSheet_(ss, name); } catch (e) { threw = /data tab/.test(e.message); }
  check('refuses to write to ' + name, threw, 'it did not refuse');
});
check('  ...and Enrollments still has every row',
      w.sheets.Enrollments._data.length === 5, w.sheets.Enrollments._data.length);

// A tab of the same name that somebody else made must not be clobbered.
w.sheets['Analytics Coverage'] = makeSheet([['Something', 'Anjali', 'typed']]);
var guarded = false;
try {
  w.sandbox.analyticsSheet_(ss, 'Analytics Coverage');
} catch (e) { guarded = /already exists/.test(e.message); }
check('refuses to overwrite a same-named tab it did not generate', guarded, 'it overwrote it');
check('  ...leaving that tab untouched',
      w.sheets['Analytics Coverage']._data.length === 1,
      w.sheets['Analytics Coverage']._data);

console.log('');
console.log('--- one receipt, two months ---');
w = freshWorld('1234');
var rp = w.sandbox.receiptPeriods_;
var P = function (mName, y) { return y * 12 + MN.indexOf(mName); };

check('a note naming two months covers both',
      rp(P('July', 2026), 'July and August fees').periods.join(',') ===
      [P('July', 2026), P('August', 2026)].join(','),
      rp(P('July', 2026), 'July and August fees').periods.map(function (p) {
        return MN[p % 12] + ' ' + Math.floor(p / 12); }));

check('  ...however it is abbreviated',
      rp(P('July', 2026), '2 months fee - Jul & Aug').periods.length === 2,
      rp(P('July', 2026), '2 months fee - Jul & Aug').periods.length);

check('  ...and is flagged as spanning months',
      rp(P('July', 2026), 'July and August fees').multi === true, 'not flagged');

check('December and January roll into the next year',
      rp(P('December', 2026), 'December and January').periods.join(',') ===
      [P('December', 2026), P('January', 2027)].join(','),
      rp(P('December', 2026), 'December and January').periods.map(function (p) {
        return MN[p % 12] + ' ' + Math.floor(p / 12); }));

check('an explicit year in the note is honoured',
      rp(P('January', 2027), 'fees for January February 2027').periods.join(',') ===
      [P('January', 2027), P('February', 2027)].join(','),
      rp(P('January', 2027), 'fees for January February 2027').periods.length);

// A single different month is a mis-keyed Fee Month, not a second month paid.
check('a note naming one different month does NOT add a month',
      rp(P('July', 2026), 'fee for August').periods.length === 1,
      rp(P('July', 2026), 'fee for August').periods.length);
check('a note naming the same month changes nothing',
      rp(P('August', 2026), 'fee for August').periods.length === 1, 'extra month added');
check('an empty note changes nothing',
      rp(P('August', 2026), '').periods.length === 1, 'extra month added');

// "2 months" without saying which two must not be guessed at.
var vague = rp(P('August', 2026), '2 months fee paid');
check('"2 months fee" with no months named is not guessed',
      vague.periods.length === 1 && vague.vague === true,
      [vague.periods.length, vague.vague]);

check('a month more than six months away is prose, not a fee period',
      rp(P('August', 2026), 'August fee, joined January, sister in December').periods
        .every(function (p) { return Math.abs(p - P('August', 2026)) <= 6; }),
      'a distant month was counted');

console.log('');
console.log('--- two-month receipts in the coverage numbers ---');
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Two Months Kid', backLabel),
  cRow('SR-2', 'One Month Kid', backLabel)
]);
w.sheets.Receipts = makeSheet([rhd,
  // Filed under the earlier month, note says it settles both.
  [ 'SS-1', '01 ' + backM.substring(0, 3) + ' ' + backY, 'Two Months Kid', '9000000001',
    '4000', backM, backY, 'Cash', '', 'Monthly Fee', '',
    backM + ' and ' + prevM + ' fees', 'Two Months Kid' ],
  [ 'SS-2', '01 ' + backM.substring(0, 3) + ' ' + backY, 'One Month Kid', '9000000002',
    '2000', backM, backY, 'Cash', '', 'Monthly Fee', '', '', 'One Month Kid' ]
]);
var mm = w.sandbox.feeCoverageForMonth(prevLabel);
check('the child who paid two months up front counts as paid in the later month',
      mm.indexOf('Two Months Kid') > mm.indexOf('PAID ('),
      mm.split('\n').filter(function (l) { return l.indexOf('Two Months') >= 0; }));
check('  ...while the child who paid one month shows as unpaid',
      mm.indexOf('One Month Kid') > mm.indexOf('NO RECEIPT FOR'),
      mm.split('\n').filter(function (l) { return l.indexOf('One Month') >= 0; }));
check('  ...and no receipt is booked into the later month',
      mm.indexOf('Receipts booked to this month    0') >= 0,
      mm.split('\n').filter(function (l) { return l.indexOf('Receipts booked') >= 0; }));

var back = w.sandbox.feeCoverageForMonth(backLabel);
check('the money stays in the month the receipt was filed under',
      back.indexOf('Rs. 6,000') >= 0,
      back.split('\n').filter(function (l) { return l.indexOf('Collected') >= 0; }));

var mmr = w.sandbox.previewMultiMonthReceipts();
check('the two-month receipt is listed for checking',
      mmr.indexOf('SS-1') >= 0 && mmr.indexOf('READ AS SEVERAL MONTHS') >= 0, 'see log');
check('  ...and the single-month one is not',
      mmr.indexOf('SS-2') < 0, 'a one-month receipt was listed');

console.log('');
console.log('--- a late fee is not a month paid ---');
// Every note below is real, copied from the Receipts tab.
var say = function (primary, note) {
  var r = rp(primary, note);
  return { got: r.periods.map(function (p) { return MN[p % 12]; }).join('+'),
           held: (r.ignored || []).map(function (p) { return MN[p % 12]; }).join(',') };
};
var JUN = P('June', 2026), JUL = P('July', 2026), AUG = P('August', 2026);

var a = say(JUN, 'Fee for June & late fee for May');
check('"Fee for June & late fee for May" pays June only',
      a.got === 'June', a);
check('  ...and says May was mentioned but not counted', a.held === 'May', a);

check('"Fee for June + late fee of May" pays June only',
      say(JUN, 'Fee for June + late fee of May').got === 'June',
      say(JUN, 'Fee for June + late fee of May'));

check('"Fee for July + 100 advanced for august" pays July only',
      say(JUL, 'Fee for July + 100 advanced for august').got === 'July',
      say(JUL, 'Fee for July + 100 advanced for august'));

check('"Fee for August + late fee of July" pays August only',
      say(AUG, 'Fee for August + late fee of July').got === 'August',
      say(AUG, 'Fee for August + late fee of July'));

// The qualifier must not swallow months that ARE being paid.
check('"Fee for May, June + late fee of May" still pays May and June',
      say(JUN, 'Fee for May, June + late fee of May').got === 'May+June',
      say(JUN, 'Fee for May, June + late fee of May'));
check('  ...and holds nothing back, May was paid in its own right',
      say(JUN, 'Fee for May, June + late fee of May').held === '',
      say(JUN, 'Fee for May, June + late fee of May'));

check('"Fee for May + june + late fees" pays both months',
      say(JUN, 'Fee for May + june + late fees').got === 'May+June',
      say(JUN, 'Fee for May + june + late fees'));

check('"Fee for May & june (100 due)" pays both — an amount owing is not a qualifier',
      say(JUN, 'Fee for May & june (100 due)').got === 'May+June',
      say(JUN, 'Fee for May & june (100 due)'));

check('"Fee for August + balance of July" pays both — a balance is that month\'s fee',
      say(AUG, 'Fee for August + balance of July').got === 'July+August',
      say(AUG, 'Fee for August + balance of July'));

check('"Fee for April, May & June" pays all three',
      say(JUN, 'Fee for April, May & June').got === 'April+May+June',
      say(JUN, 'Fee for April, May & June'));

check('"Fee for July & august" is unaffected',
      say(JUL, 'Fee for July & august').got === 'July+August',
      say(JUL, 'Fee for July & august'));

check('"Fee for August & September" reaches into next month',
      say(AUG, 'Fee for August & September').got === 'August+September',
      say(AUG, 'Fee for August & September'));

// And the student grid must show it.
console.log('');
console.log('--- the student grid reflects a two-month payment ---');
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh, cRow('SR-1', 'Paid Ahead', backLabel)]);
w.sheets.Receipts = makeSheet([rhd,
  [ 'SS-1', '01 ' + backM.substring(0, 3) + ' ' + backY, 'Paid Ahead', '9000000001',
    '4000', backM, backY, 'Cash', '', 'Monthly Fee', '',
    'Fee for ' + backM + ' & ' + prevM, 'Paid Ahead' ]
]);
var gm = w.sandbox.analyticsTabModel_(w.sandbox.buildFeeCoverage_());
var cols = gm.students.head.slice(gm.firstMonthCol - 1);
var row = gm.students.rows[0];
check('both months read PAID in the students grid',
      row[gm.firstMonthCol - 1 + cols.indexOf(backLabel)] === 'PAID' &&
      row[gm.firstMonthCol - 1 + cols.indexOf(prevLabel)] === 'PAID',
      [row[gm.firstMonthCol - 1 + cols.indexOf(backLabel)],
       row[gm.firstMonthCol - 1 + cols.indexOf(prevLabel)]]);

console.log('');
console.log('--- fee types that must not cover a month ---');
// The panel offers six fee types. Only Monthly Fee is a month's tuition; the
// rest are real money but must never make a month read as paid. A Late Fee is
// the penalty for paying late, not the payment.
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Uniform Only',  backLabel),
  cRow('SR-2', 'Late Fee Only', backLabel),
  cRow('SR-3', 'Registered',    backLabel),
  cRow('SR-4', 'Actually Paid', backLabel)
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-1', 'Uniform Only',  'Uniform Only',  prevM, prevY, 'Uniform / Costume Fee', '1200'),
  cRcpt('SS-2', 'Late Fee Only', 'Late Fee Only', prevM, prevY, 'Late Fee',              '100'),
  cRcpt('SS-3', 'Registered',    'Registered',    prevM, prevY, 'Registration Fee',      '500'),
  cRcpt('SS-4', 'Actually Paid', 'Actually Paid', prevM, prevY, 'Monthly Fee',           '2000')
]);
var ft = w.sandbox.feeCoverageForMonth(prevLabel);
var ftPaid = ft.substring(ft.indexOf('PAID ('));

check('a uniform / costume fee does not cover the month',
      ftPaid.indexOf('Uniform Only') < 0, 'uniform fee counted as tuition');
check('a late fee does not cover the month',
      ftPaid.indexOf('Late Fee Only') < 0, 'late fee counted as tuition');
check('a registration fee does not cover the month',
      ftPaid.indexOf('Registered') < 0, 'registration fee counted as tuition');
check('a monthly fee does',
      ftPaid.indexOf('Actually Paid') >= 0, 'monthly fee not counted');
check('so three of the four students read as unpaid',
      ft.indexOf('No receipt                       3') >= 0,
      ft.split('\n').filter(function (l) { return l.indexOf('No receipt') >= 0; }));
check('  ...and the non-tuition receipts are counted as skipped, not lost',
      w.sandbox.buildFeeCoverage_().counts.otherType === 3,
      w.sandbox.buildFeeCoverage_().counts.otherType);

// The dropdown is the only list of fee types anywhere.
var adminHtml = fs.readFileSync('sriramstudio_admin.html', 'utf8');
var opts = (adminHtml.match(/<select id="r-feetype"[\s\S]*?<\/select>/) || [''])[0];
['Monthly Fee', 'Registration Fee', 'Uniform / Costume Fee', 'Late Fee', 'Workshop', 'Other']
  .forEach(function (t) {
    check('the panel offers "' + t + '"', opts.indexOf('>' + t + '<') >= 0, opts);
  });
// Code.gs must not branch on particular fee types — it only ever asks whether
// one starts with "Monthly". Comments may name them by way of example, so
// strip those before looking. 'Monthly Fee' is exempt: it is the documented
// fallback when a receipt arrives without a fee type at all.
var gsCode = fs.readFileSync(SRC, 'utf8')
  .replace(/\/\*[\s\S]*?\*\//g, ' ')
  .replace(/(^|[^:])\/\/[^\n]*/g, '$1 ');
// 'Workshop' and 'Other' are left out: they are also Enrollments *Type*
// values, which Code.gs legitimately does name, and the words collide.
['Registration Fee', 'Uniform / Costume Fee', 'Late Fee']
  .forEach(function (t) {
    check('Code.gs does not branch on "' + t + '"',
          gsCode.indexOf("'" + t + "'") < 0 && gsCode.indexOf('"' + t + '"') < 0,
          'Code.gs now enumerates fee types - it should only test for "monthly"');
  });

console.log('');
console.log('--- receipts column alignment ---');
w = freshWorld('1234');
w.sheets.Receipts = makeSheet([rhd]);
var al = w.sandbox.auditReceiptColumns();
check('the current Receipts headers are reported as aligned',
      al.indexOf('ALIGNED') >= 0, al.split('\n').filter(function (l) {
        return l.indexOf('WRONG') >= 0; }));

// A receipt written now must land under the headings the audit claims.
call(w, { action: 'addReceipt', pin: '1234', data: enc({
  studentName: 'Align Kid', students: ['Align Kid'], guardianPhone: '9000000001',
  amount: '2000', month: 'August', year: '2026', payMode: 'Cash', upiRef: '',
  feeType: 'Uniform / Costume Fee', dateReceived: '01 Aug 2026', note: 'costume' }) });
var hdr = w.sheets.Receipts._data[0];
var row = w.sheets.Receipts._data[1];
var at = function (name) {
  for (var i = 0; i < hdr.length; i++) if (String(hdr[i]).indexOf(name) === 0) return row[i];
  return '(no such column)';
};
check('the fee type lands under Fee Type',
      at('Fee Type') === 'Uniform / Costume Fee', at('Fee Type'));
check('the amount lands under Amount', at('Amount') === '2000', at('Amount'));
check('the note lands under Note', at('Note') === 'costume', at('Note'));
check('the student lands under Student Name', at('Student Name') === 'Align Kid',
      at('Student Name'));

// Now move a column, as deleting or inserting one in the sheet would.
var shuffled = rhd.slice();
shuffled.splice(3, 0, 'Somebody Inserted This');   // a new column before Contact
w.sheets.Receipts = makeSheet([shuffled]);
var bad = w.sandbox.auditReceiptColumns();
check('an inserted column is caught',
      bad.indexOf('COLUMN(S) OUT OF PLACE') >= 0,
      bad.split('\n').filter(function (l) { return l.indexOf('OUT OF PLACE') >= 0; }));
check('  ...and it says not to issue receipts until it is fixed',
      bad.indexOf('Do NOT issue receipts') >= 0, 'no warning given');
check('  ...and says receipts already issued are unaffected',
      bad.indexOf('already issued are unaffected') >= 0, 'no reassurance given');

// A column added at the END is the safe way, and must not be flagged.
var appended = rhd.concat(['Fee Breakdown']);
w.sheets.Receipts = makeSheet([appended]);
var okExtra = w.sandbox.auditReceiptColumns();
check('a column added at the end is not flagged',
      okExtra.indexOf('ALIGNED') >= 0, 'a trailing column was called a mismatch');
check('  ...but is listed so it is visible',
      okExtra.indexOf('Fee Breakdown') >= 0, 'trailing column not shown');

console.log('');
console.log('--- one receipt, several students, several fee types ---');
w = freshWorld('1234');
var SPLIT = [
  { student: 'Riya Sen', fees: [{ type: 'Monthly Fee', amount: 2000 },
                                { type: 'Uniform / Costume Fee', amount: 800 }] },
  { student: 'Diya Sen', fees: [{ type: 'Uniform / Costume Fee', amount: 800 },
                                { type: 'Late Fee', amount: 100 }] }
];
var text = w.sandbox.serialiseFeeSplit_(SPLIT);
check('the split writes as one readable line',
      text === 'Riya Sen: Monthly Fee 2000; Uniform / Costume Fee 800 | ' +
               'Diya Sen: Uniform / Costume Fee 800; Late Fee 100', text);
check('  ...and reads back to the same thing',
      JSON.stringify(w.sandbox.parseFeeSplit_(text)) === JSON.stringify(SPLIT),
      w.sandbox.parseFeeSplit_(text));
check('a fee type containing a slash survives the round trip',
      w.sandbox.parseFeeSplit_(text)[0].fees[1].type === 'Uniform / Costume Fee',
      w.sandbox.parseFeeSplit_(text)[0].fees[1]);
check('the total is the sum of every line',
      w.sandbox.feeSplitTotal_(SPLIT) === 3700, w.sandbox.feeSplitTotal_(SPLIT));
check('per-type totals add across students',
      w.sandbox.feeTotalsByType_(SPLIT)['Uniform / Costume Fee'] === 1600,
      w.sandbox.feeTotalsByType_(SPLIT));
check('only the sibling with a monthly fee is a monthly payer',
      JSON.stringify(Object.keys(w.sandbox.monthlyPayersIn_(SPLIT))) ===
      JSON.stringify(['riya sen']), w.sandbox.monthlyPayersIn_(SPLIT));
check('an empty split reads as nothing rather than throwing',
      w.sandbox.parseFeeSplit_('').length === 0 &&
      w.sandbox.parseFeeSplit_(null).length === 0, 'threw or returned junk');
check('garbage in the column is ignored, not half-parsed',
      w.sandbox.parseFeeSplit_('who knows what this is').length === 0,
      w.sandbox.parseFeeSplit_('who knows what this is'));

console.log('');
console.log('--- what the sheet ends up holding ---');
w = freshWorld('1234');
w.sheets.Receipts = makeSheet([rhd]);
var res = call(w, { action: 'addReceipt', pin: '1234', data: enc({
  studentName: 'Riya Sen & Diya Sen', students: ['Riya Sen', 'Diya Sen'],
  guardianPhone: '9830012345', month: 'August', year: '2026',
  payMode: 'UPI', upiRef: '4477xx', dateReceived: '24 Aug 2026', split: SPLIT }) });
check('the receipt is created', res.success === true && !!res.receiptNo, res);

var hd = w.sheets.Receipts._data[0], rw = w.sheets.Receipts._data[1];
var col = function (name) {
  for (var i = 0; i < hd.length; i++) if (String(hd[i]).indexOf(name) === 0) return rw[i];
  return '(no column)';
};
check('the total is computed from the split, not taken on trust',
      col('Amount') === 3700, col('Amount'));
check('Fee Type summarises what was paid',
      col('Fee Type') === 'Monthly Fee + Uniform / Costume Fee + Late Fee',
      col('Fee Type'));
check('the Fee Split column holds the per-student detail',
      col('Fee Split') === text, col('Fee Split'));
check('a per-type column is created and totalled',
      col('Uniform / Costume Fee ₹') === 1600, col('Uniform / Costume Fee ₹'));
check('  ...for each type on the receipt',
      col('Monthly Fee ₹') === 2000 && col('Late Fee ₹') === 100,
      [col('Monthly Fee ₹'), col('Late Fee ₹')]);
check('the original thirteen columns have not moved',
      hd.slice(0, 13).join(',') === rhd.join(','), hd.slice(0, 13));
check('  ...so the new ones sit after them',
      hd.length === 17 && String(hd[13]).indexOf('Monthly Fee') === 0, hd);
check('the alignment audit still passes',
      w.sandbox.auditReceiptColumns().indexOf('ALIGNED') >= 0, 'audit broke');

console.log('');
console.log('--- the sibling who only paid a costume fee ---');
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Riya Sen', backLabel),
  cRow('SR-2', 'Diya Sen', backLabel)
]);
// Same receipt, but booked against last month so the month is a closed one.
w.sheets.Receipts = makeSheet([rhd.concat(['Monthly Fee ₹','Uniform / Costume Fee ₹','Late Fee ₹','Fee Split']),
  ['SS-2026-0801', '01 ' + prevM.substring(0,3) + ' ' + prevY, 'Riya Sen & Diya Sen',
   '9830012345', 3700, prevM, prevY, 'UPI', '4477xx',
   'Monthly Fee + Uniform / Costume Fee + Late Fee', '', '', 'Riya Sen | Diya Sen',
   2000, 1600, 100, text]
]);
var cm = w.sandbox.feeCoverageForMonth(prevLabel);
var cmPaid = cm.substring(cm.indexOf('PAID ('));
check('the sister who paid a monthly fee is covered',
      cmPaid.indexOf('Riya Sen') >= 0, 'see log');
check('the sister who only paid a costume and late fee is NOT',
      cmPaid.indexOf('Diya Sen') < 0, 'a costume fee covered the month');
check('  ...and she is listed as owing the month',
      cm.indexOf('Diya Sen') > cm.indexOf('NO RECEIPT FOR'),
      cm.split('\n').filter(function (l) { return l.indexOf('Diya') >= 0; }));
check('the money is still counted in full',
      cm.indexOf('Rs. 3,700') >= 0,
      cm.split('\n').filter(function (l) { return l.indexOf('Collected') >= 0; }));

console.log('');
console.log('--- receipts without a split are unaffected ---');
w2 = freshWorld('1234');
w2.sheets.Enrollments = makeSheet([dh, cRow('SR-1', 'Old Kid', backLabel)]);
w2.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-9', 'Old Kid', 'Old Kid', prevM, prevY, 'Monthly Fee', '2000')
]);
var old = w2.sandbox.feeCoverageForMonth(prevLabel);
check('an older receipt with no Fee Split still covers the month',
      old.substring(old.indexOf('PAID (')).indexOf('Old Kid') >= 0, 'see log');

// A different division of the same total is a different payment, not a retry.
w2 = freshWorld('1234');
var base = { studentName: 'A & B', students: ['A Kid', 'B Kid'],
             month: 'August', year: '2026' };
base.split = [{ student: 'A Kid', fees: [{ type: 'Monthly Fee', amount: 2000 }] },
              { student: 'B Kid', fees: [{ type: 'Monthly Fee', amount: 1000 }] }];
var r1 = call(w2, { action: 'addReceipt', pin: '1234', data: enc(base) });
base.split = [{ student: 'A Kid', fees: [{ type: 'Monthly Fee', amount: 1000 }] },
              { student: 'B Kid', fees: [{ type: 'Monthly Fee', amount: 2000 }] }];
var r2 = call(w2, { action: 'addReceipt', pin: '1234', data: enc(base) });
check('the same total split differently is not swallowed as a retry',
      r1.receiptNo !== r2.receiptNo && !r2.duplicate, [r1, r2]);
var same = call(w2, { action: 'addReceipt', pin: '1234', data: enc(base) });
check('  ...but an identical resend still is',
      same.receiptNo === r2.receiptNo && same.duplicate === true, same);

console.log('');
console.log('--- the fees tab ---');
w = freshWorld('1234');
w.sheets.Enrollments = makeSheet([dh, cRow('SR-1', 'Riya Sen', backLabel),
                                      cRow('SR-2', 'Diya Sen', backLabel)]);
var splitTxt = 'Riya Sen: Monthly Fee 2000; Uniform / Costume Fee 800 | ' +
               'Diya Sen: Uniform / Costume Fee 800; Late Fee 100';
w.sheets.Receipts = makeSheet([
  rhd.concat(['Fee Split']),
  ['SS-1', '01 ' + prevM.substring(0,3) + ' ' + prevY, 'Riya Sen & Diya Sen', '98300',
   3700, prevM, prevY, 'UPI', '', 'Monthly Fee + Uniform / Costume Fee + Late Fee',
   '', '', 'Riya Sen | Diya Sen', splitTxt],
  // A plain registration fee, no split - its money must still be counted.
  ['SS-2', '01 ' + prevM.substring(0,3) + ' ' + prevY, 'Riya Sen', '98300',
   500, prevM, prevY, 'Cash', '', 'Registration Fee', '', '', 'Riya Sen', '']
]);
var cv = w.sandbox.buildFeeCoverage_();
var fm = w.sandbox.analyticsTabModel_(cv);

check('every fee type on the receipts becomes a column',
      ['Monthly Fee','Uniform / Costume Fee','Late Fee','Registration Fee']
        .every(function (t) { return fm.fees.head.indexOf(t) > 0; }), fm.fees.head);
check('the last column is the row total', fm.fees.head[fm.fees.head.length - 1] === 'Total',
      fm.fees.head);

var prevRow = fm.fees.rows.filter(function (r) { return r[0] === prevLabel; })[0];
var colOf = function (name) { return fm.fees.head.indexOf(name); };
check('a split receipt is broken up by type',
      prevRow[colOf('Monthly Fee')] === 2000 &&
      prevRow[colOf('Uniform / Costume Fee')] === 1600 &&
      prevRow[colOf('Late Fee')] === 100,
      [prevRow[colOf('Monthly Fee')], prevRow[colOf('Uniform / Costume Fee')],
       prevRow[colOf('Late Fee')]]);
check('a non-tuition receipt still contributes its money',
      prevRow[colOf('Registration Fee')] === 500, prevRow[colOf('Registration Fee')]);
check('the row total is every fee type added up',
      prevRow[prevRow.length - 1] === 4200, prevRow[prevRow.length - 1]);
check('  ...which is the full amount of both receipts',
      prevRow[prevRow.length - 1] === 3700 + 500, prevRow[prevRow.length - 1]);

check('a month with no receipts reads zero, not blank',
      fm.fees.rows.every(function (r) {
        return r.slice(1).every(function (v) { return typeof v === 'number'; }); }),
      'a non-numeric cell would break the column total');

var prev = w.sandbox.previewAnalyticsRefresh();
check('the fees tab is listed among the tabs a refresh would build',
      prev.indexOf('Analytics Fees') >= 0,
      prev.split('\n').filter(function (l) { return l.indexOf('Analytics') >= 0; }));
check('  ...and the preview still writes nothing',
      prev.indexOf('Nothing has been written') >= 0, 'preview lost its promise');
check('  ...and still reports the data tabs as untouched',
      prev.indexOf('read only, never written') >= 0, prev);

// Counting revenue must not have changed which receipts cover a month.
var cmf = w.sandbox.feeCoverageForMonth(prevLabel);
check('the registration fee still does not cover the month',
      cmf.substring(cmf.indexOf('PAID (')).indexOf('Diya Sen') < 0,
      'coverage changed when revenue counting was added');
check('  ...and the monthly payer still is covered',
      cmf.substring(cmf.indexOf('PAID (')).indexOf('Riya Sen') >= 0, 'see log');
check('receipts skipped for coverage are still counted as skipped',
      cv.counts.otherType === 1, cv.counts.otherType);

console.log('');
console.log('--- months stated on the receipt beat months guessed from a note ---');
w = freshWorld('1234');
var rpm = w.sandbox.receiptPeriods_;
var JUL = P('July', 2026), AUG = P('August', 2026), DEC = P('December', 2026);

check('a stated span covers every month in it',
      rpm(JUL, '', 'July 2026 | August 2026').periods.join(',') === [JUL, AUG].join(','),
      rpm(JUL, '', 'July 2026 | August 2026').periods);
check('  ...and is flagged as spanning months',
      rpm(JUL, '', 'July 2026 | August 2026').multi === true, 'not flagged');
check('  ...and says so, rather than being an inference',
      rpm(JUL, '', 'July 2026 | August 2026').stated === true, 'not marked as stated');

check('a stated span wins over a note that disagrees',
      rpm(JUL, 'Fee for July + late fee of June', 'July 2026 | August 2026')
        .periods.join(',') === [JUL, AUG].join(','),
      rpm(JUL, 'Fee for July + late fee of June', 'July 2026 | August 2026').periods);
check('  ...and nothing is held back when the months are stated',
      rpm(JUL, 'Fee for July + late fee of June', 'July 2026 | August 2026')
        .ignored.length === 0, 'a qualifier was applied to a stated span');

check('a stated span crossing a year is kept as written',
      rpm(DEC, '', 'December 2026 | January 2027').periods.join(',') ===
      [DEC, P('January', 2027)].join(','),
      rpm(DEC, '', 'December 2026 | January 2027').periods);

check('a single stated month is not called a span',
      rpm(AUG, '', 'August 2026').multi === false,
      rpm(AUG, '', 'August 2026'));

check('an empty Fee Months falls back to reading the note',
      rpm(JUL, 'Fee for July & August', '').periods.join(',') === [JUL, AUG].join(','),
      rpm(JUL, 'Fee for July & August', '').periods);
check('  ...and so does an unreadable one',
      rpm(JUL, 'Fee for July & August', 'nonsense').periods.join(',') === [JUL, AUG].join(','),
      rpm(JUL, 'Fee for July & August', 'nonsense').periods);
check('the note-only path still holds back a late fee month',
      rpm(JUL, 'Fee for July + late fee of June', '').periods.length === 1,
      rpm(JUL, 'Fee for July + late fee of June', ''));

console.log('');
console.log('--- a stated span in the coverage numbers ---');
w.sheets.Enrollments = makeSheet([dh, cRow('SR-1', 'Span Kid', backLabel)]);
w.sheets.Receipts = makeSheet([rhd.concat(['Fee Months']),
  ['SS-1', '01 ' + backM.substring(0,3) + ' ' + backY, 'Span Kid', '9', 4000,
   backM, backY, 'Cash', '', 'Monthly Fee', '', '', 'Span Kid',
   backLabel + ' | ' + prevLabel]
]);
var sp1 = w.sandbox.feeCoverageForMonth(backLabel);
var sp2 = w.sandbox.feeCoverageForMonth(prevLabel);
check('the month it was filed under is covered',
      sp1.substring(sp1.indexOf('PAID (')).indexOf('Span Kid') >= 0, 'see log');
check('and so is the second month it states',
      sp2.substring(sp2.indexOf('PAID (')).indexOf('Span Kid') >= 0, 'see log');
check('  ...with the receipt still booked to the first month only',
      sp2.indexOf('Receipts booked to this month    0') >= 0,
      sp2.split('\n').filter(function (l) { return l.indexOf('Receipts booked') >= 0; }));

console.log('');
console.log('--- only the dashboard stays on the tab strip ---');
w = freshWorld('1234');
var hidden = {}, active = null;
var mkTab = function () {
  var s = makeSheet([['x']]);
  s.hideSheet = function () { hidden[this._name] = true; };
  s.showSheet = function () { hidden[this._name] = false; };
  return s;
};
['Analytics Dashboard','Analytics Coverage','Analytics Students',
 'Analytics Fees','Analytics Names','Analytics Gaps'].forEach(function (n) {
  var s = mkTab(); s._name = n; w.sheets[n] = s;
});
// The stub spreadsheet needs to remember which sheet is active, because a
// hidden sheet cannot be the active one.
w.sandbox.SpreadsheetApp.getActiveSpreadsheet = function () {
  return {
    getSheetByName: function (n) { return w.sheets[n] || null; },
    insertSheet: function (n) { return (w.sheets[n] = mkTab()); },
    setActiveSheet: function (s) { active = s._name; }
  };
};

var n = w.sandbox.hideAnalyticsWorkingTabs();
check('the five working tabs are hidden',
      ['Analytics Coverage','Analytics Students','Analytics Fees',
       'Analytics Names','Analytics Gaps'].every(function (t) { return hidden[t]; }),
      hidden);
check('the dashboard is NOT hidden', !hidden['Analytics Dashboard'], hidden);
check('  ...and is made active first, since a hidden sheet cannot be active',
      active === 'Analytics Dashboard', active);
check('the data tabs are never touched',
      !hidden['Enrollments'] && !hidden['Receipts'] && !hidden['Config'], hidden);
check('it says the dashboard still works',
      n.indexOf('still reads them') >= 0, n);

w.sandbox.showAnalyticsWorkingTabs();
check('showing them again reveals every one',
      ['Analytics Coverage','Analytics Students','Analytics Fees',
       'Analytics Names','Analytics Gaps'].every(function (t) { return !hidden[t]; }),
      hidden);

// A missing tab must not throw — somebody may have deleted one by hand.
delete w.sheets['Analytics Fees'];
var threw = false;
try { w.sandbox.hideAnalyticsWorkingTabs(); } catch (e) { threw = true; }
check('a tab that has been deleted by hand does not break hiding', !threw, 'it threw');

console.log('');
console.log('--- one bad date must not blow the tabs up ---');
w = freshWorld('1234');
// A joining date years in the past, of the sort that turns up in real rosters.
w.sheets.Enrollments = makeSheet([dh,
  cRow('SR-1', 'Normal Kid', backLabel),
  cRow('SR-2', 'Ancient Record', 'January 2010')
]);
w.sheets.Receipts = makeSheet([rhd,
  cRcpt('SS-1', 'Normal Kid', 'Normal Kid', prevM, prevY, 'Monthly Fee', '2000')
]);
var cov = w.sandbox.buildFeeCoverage_();
var mdl = w.sandbox.analyticsTabModel_(cov);

check('the coverage model still spans the whole history',
      cov.months.length > 24, cov.months.length + ' months');
check('but the tabs show a window of at most 24 months',
      mdl.monthNames.length <= 24, mdl.monthNames.length + ' month columns');
check('  ...so the students grid stays a sensible size',
      mdl.students.head.length <= 14 + 24,
      mdl.students.head.length + ' columns');
check('  ...and it is the RECENT months that are kept',
      mdl.monthNames[mdl.monthNames.length - 1] === curM + ' ' + curY,
      mdl.monthNames[mdl.monthNames.length - 1]);
check('the fees tab is windowed the same way',
      mdl.fees.rows.length <= 24 + 1, mdl.fees.rows.length + ' rows');
check('the coverage tab too', mdl.coverage.rows.length <= 24,
      mdl.coverage.rows.length + ' rows');
check('the dashboard still opens on a month that exists in the window',
      mdl.monthNames.indexOf(mdl.defaultMonth) >= 0 || mdl.defaultMonth === '',
      [mdl.defaultMonth, mdl.monthNames.slice(-3)]);

// The recent month must still be counted correctly after windowing.
var recent = mdl.students.rows.filter(function (r) { return r[2] === 'Normal Kid'; })[0];
var prevIdx = mdl.firstMonthCol - 1 + mdl.monthNames.indexOf(prevLabel);
check('a paid month still reads PAID after windowing',
      recent[prevIdx] === 'PAID', recent[prevIdx]);

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);
