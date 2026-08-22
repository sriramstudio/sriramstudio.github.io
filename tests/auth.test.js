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
    getRange(r, c) {
      return {
        setValue: v => { while (data.length < r) data.push([]); data[r - 1][c - 1] = v; },
        setFontWeight: () => chain, setBackground: () => chain, setFontColor: () => chain
      };
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
       'Workshop Date', 'Workshop Fee', 'Heard From', 'Notes'],
      ['SR-2026-0101', '01 Jan 2026', 'New Admission', 'Test Child', '2015-01-01', 'F', '', '',
       'Guardian', 'Mother', '9999', '9999', 'a@b.c', '12 Somewhere Rd', '', 'Bhawanipur',
       '', '', '', '', '', '', '', '']
    ]),
    Receipts: makeSheet([['Receipt No', 'Issued At', 'Student Name', 'Contact', 'Amount',
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
r = call(w, { action: 'previewNameCleanup' });
check('previewNameCleanup with NO pin is refused', r.auth === true, r);
r = call(w, { action: 'previewNameCleanup', pin: '1234' });
check('  ...and is not a routed action even WITH the pin',
      typeof r.error === 'string' && r.error.indexOf('Unknown action') === 0, r);

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);
