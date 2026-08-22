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
    Utilities: { formatDate: () => '01 Jan 2026, 10:00 AM' },
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
r = call(w, { action: 'previewLegacyStudents' });
check('previewLegacyStudents with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'previewLegacyStudents', pin: '1234' });
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

console.log('\n--- legacy roster: two students, one name ---');
w = freshWorld('1234');
// One 'Krisha Agarwal' already enrolled, with no phone on record.
call(w, { action: 'addEnrollment', pin: '1234',
          data: enc({ mode: 'legacy', studentName: 'Krisha Agarwal' }) });
// The roster lists two of them, told apart by phone.
w.sheets['Legacy Students'] = makeSheet([
  ['Student Name', 'Contact', 'Center'],
  ['Krisha Agarwal', '9111111111', 'Salt Lake'],
  ['Krisha Agarwal', '9222222222', 'Bhawanipur'],
  ['Solo Student',   '',           'discontinue']
]);
var roster = w.sandbox.buildLegacyRoster_();
var krishaAdds = roster.toAdd.filter(function (x) { return x.name === 'Krisha Agarwal'; }).length;
var krishaUpds = roster.toUpdate.filter(function (u) { return u.rec.name === 'Krisha Agarwal'; }).length;
check('both Krishas are accounted for', krishaAdds + krishaUpds === 2,
      { add: krishaAdds, update: krishaUpds });
check('  ...one fills the blank row, one is added new',
      roster.toUpdate.length === 1 && roster.toAdd.filter(function (x) {
        return x.name === 'Krisha Agarwal'; }).length === 1,
      { add: roster.toAdd.map(function (x) { return x.name; }),
        update: roster.toUpdate.length });
check('  ...and neither is silently dropped',
      roster.blocked.length === 0 && roster.ambiguous.length === 0,
      { blocked: roster.blocked.length, ambiguous: roster.ambiguous.length });
var solo = roster.toAdd.filter(function (x) { return x.name === 'Solo Student'; })[0];
check('a leaving word in Center becomes the status, not a branch',
      solo && solo.status === 'Left' && !solo.centre, solo);

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

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);
