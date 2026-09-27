// The Contact box on the receipt form, tested without a browser.
//
// The rules it must keep:
//   - it shows the number on the selected student's record
//   - it is blank when no student is selected, typed or not
//   - a number Anjali types survives while those students stay selected
//   - a number is only ever offered for saving into a BLANK record
//   - two students with two different numbers are warned about, never merged
//
// Usage:  node tests/contact.test.js sriramstudio_admin.html

const fs = require('fs');
const file = process.argv[2] || 'sriramstudio_admin.html';
const src  = fs.readFileSync(file, 'utf8');

const NEEDED = ['selectedPhoneList', 'syncContact', 'flagContactMismatch',
                'phoneFillCandidates'];
const bodies = NEEDED.map(function (fn) {
  const m = src.match(new RegExp('\\nfunction ' + fn + '\\(.*?\\n\\}\\n', 's'));
  if (!m) { console.log('FAIL  ' + fn + ' is not in ' + file); process.exit(1); }
  return m[0];
});

const els = { 'r-phone': { value: '' },
              'r-other-contacts': { textContent: '', style: {} } };
const harness = `
  let selectedStudents = [], phoneTyped = false;
  const document = { getElementById: function (id) { return els[id] || null; } };
  ${bodies.join('\n')}
  return { set: function (list) { selectedStudents = list; },
           typed: function (v) { els['r-phone'].value = v; phoneTyped = true; },
           sync: function () { phoneTyped = false; syncContact(); },
           syncKeepTyped: syncContact,
           offer: phoneFillCandidates };
`;
const api = new Function('els', harness)(els);

let fails = 0;
function is(label, got, want) {
  if (JSON.stringify(got) !== JSON.stringify(want)) {
    fails++;
    console.log('FAIL  ' + label + '\n      got  ' + JSON.stringify(got) +
                '\n      want ' + JSON.stringify(want));
  } else {
    console.log('  PASS  ' + label);
  }
}
const S = (name, phone, id) => ({ uid: name, name: name, phone: phone,
                                  id: id === undefined ? 'SR-' + name : id });

api.set([S('Riya', '9830792201')]); api.sync();
is('a student on record fills the box', els['r-phone'].value, '9830792201');
is('  ...and nothing is offered for saving', api.offer(), null);

api.set([]); api.sync();
is('no student selected leaves it blank', els['r-phone'].value, '');

api.set([S('Ira', '')]); api.sync();
is('a student with no number gives a blank box', els['r-phone'].value, '');
api.typed('98765 43210');
let p = api.offer();
is('a typed number is offered for that student',
   p && p.students.map(function (x) { return x.name; }), ['Ira']);
is('  ...not flagged as coming from a sibling', p && p.fromSibling, false);

api.set([S('Ira', ''), S('Vir', '')]); api.syncKeepTyped();
is('a typed number survives a second blank sibling',
   els['r-phone'].value, '98765 43210');
p = api.offer();
is('  ...and is offered for both',
   p && p.students.map(function (x) { return x.name; }), ['Ira', 'Vir']);

api.set([]); api.syncKeepTyped();
is('a typed number dies with the last chip', els['r-phone'].value, '');

api.set([S('Anya', '9830792201'), S('Dev', '')]); api.sync();
is('a sibling with a number fills the box', els['r-phone'].value, '9830792201');
p = api.offer();
is('  ...offered only for the sibling who has none',
   p && p.students.map(function (x) { return x.name; }), ['Dev']);
is('  ...and flagged as coming from the sibling', p && p.fromSibling, true);

api.set([S('Krisha A', '9830792201'), S('Krisha B', '9000000000')]); api.sync();
is('two different numbers: the first is kept',
   els['r-phone'].value, '9830792201');
is('  ...Anjali is warned',
   els['r-other-contacts'].textContent.indexOf('different numbers') > -1, true);
is('  ...and nothing is offered for saving', api.offer(), null);

api.set([S('A', '9830792201'), S('B', '98307 92201')]); api.sync();
is('the same number written differently is one number, not a clash',
   els['r-other-contacts'].textContent, '');

api.set([{ uid: 'x', name: 'Walk-in', phone: '', id: '' }]);
api.typed('9876543210');
is('a student who is not on the roster is never written to', api.offer(), null);

api.set([S('Ira', '')]); api.typed('98765');
is('a half-typed number is not offered', api.offer(), null);


// ── Clearing the last student clears the rest of the form ──────
// A note or a fee period left over from an abandoned receipt must not ride
// onto the next family's. Removing one sibling of several must not disturb
// anything Anjali has typed.
(function () {
  const fEls = {
    'r-note':   { value: 'late fee for May' },
    'r-upiref': { value: 'REF123' },
    'r-month':  { value: 'March' },
    'r-year':   { value: '2020' },
    'r-daterecv': { value: '2020-01-01' },
    'amount-words': { textContent: 'x' },
    'r-phone':  { value: '' },
    'r-other-contacts': { textContent: '', style: {} },
    'r-students': { style: {}, innerHTML: '', querySelectorAll: function () { return []; } }
  };
  const want = ['resetReceiptInputs', 'renderStudentChips'].map(function (fn) {
    const m = src.match(new RegExp('\\nfunction ' + fn + '\\(.*?\\n\\}\\n', 's'));
    if (!m) { console.log('FAIL  ' + fn + ' is not in ' + file); process.exit(1); }
    return m[0];
  });
  const form = new Function('els', `
    let selectedStudents = [], phoneTyped = false, hadStudents = false;
    const document = { getElementById: function (id) { return els[id] || null; } };
    function resetPeriodRange() { els.periodReset = true; }
    function syncContact() { if (!selectedStudents.length) els['r-phone'].value = ''; }
    function renderFeeEditor() {}
    ${want.join('\n')}
    return { set: function (l) { selectedStudents = l; }, render: renderStudentChips };
  `)(fEls);

  const thisMonth = new Date().toLocaleString('en-US', { month: 'long' });

  form.set([S('A', ''), S('B', '')]); form.render();
  is('a note survives while students are selected', fEls['r-note'].value, 'late fee for May');

  form.set([S('A', '')]); form.render();
  is('  ...and survives removing one of two', fEls['r-note'].value, 'late fee for May');

  form.set([]); form.render();
  is('removing the last student clears the note', fEls['r-note'].value, '');
  is('  ...and the UPI reference', fEls['r-upiref'].value, '');
  is('  ...and puts the fee month back to this month', fEls['r-month'].value, thisMonth);
  is('  ...and the year', String(fEls['r-year'].value), String(new Date().getFullYear()));
  is('  ...and closes the multi-month range', fEls.periodReset, true);

  fEls['r-note'].value = 'new note'; form.render();
  is('a note typed on an already-empty form is left alone', fEls['r-note'].value, 'new note');
})();

console.log(fails ? '\n' + fails + ' CHECK(S) FAILED' : '\nALL CONTACT CHECKS PASSED');
process.exit(fails ? 1 : 0);
