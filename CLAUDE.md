# Sriram Studio — working notes for Claude

Admin system for a dance school in Kolkata run by one non-technical person
(Anjali). `README.md` is the operational runbook and is kept current — read it
before changing anything.

## This is live data about children

`Enrollments` holds ~350 students' names, dates of birth, phone numbers and
home addresses. The Apps Script web app is reachable by anyone who reads
`register.html`, so the PIN checked in `doGet` is the only thing protecting it.

- Every new `doGet` action is **protected by default**. Only the public
  registration path (`addEnrollment` with mode `app-admission` / `app-workshop`)
  and `verifyPin` may skip the gate.
- Maintenance functions run **from the Apps Script editor only** — never route
  them through `doGet`.

## Always run these

```
node tests/auth.test.js Code.gs                    # after any Code.gs change
node tests/check-refs.js sriramstudio_admin.html   # after any HTML change
```

`node --check` only catches syntax. `check-refs.js` exists because deleting a
function and leaving its call sites behind passed every check and then failed
at runtime inside a `catch`, surfacing as a misleading "could not load"
message.

## Deploying

HTML is served by GitHub Pages — `git push` and wait 1–3 minutes.

`Code.gs` goes through clasp:

```
clasp push
clasp redeploy AKfycbzRB6XQrT8lx2lgcgv9nPQ18akiRCU8x58ritJIgbjNRkKxE3FZAifMREpZ4f7ZgSWD -d "what changed"
```

**Never `clasp deploy`** — it creates a *new* deployment with a *new* URL and
breaks `register.html`, which has the current one hardcoded. Redeploying to the
existing id keeps the URL. Roll back with `-V <older version>`.

Push the HTML **before** redeploying the backend. The reverse order breaks the
admin panel in the window between the two.

## Never write to the sheet without a preview first

Every bulk operation gets a read-only `preview*` function that reports what it
*would* do. Saurav runs it and pastes the output back. Three separate
correctness bugs in the roster importer were caught this way and none reached
the data.

Real data has repeatedly contradicted reasonable assumptions — a `Center`
column also holding status words, sibling names merged into one receipt, a name
appearing twice in one sheet and once in another. Do not infer the shape of the
sheets; have Apps Script report it. The Drive connector cannot read them.

Prefer refusing to guess over silently merging or splitting records: surface the
ambiguous rows and let Saurav resolve them.

## Conventions that already caught bugs

- **Look columns up by name, never by position.** Deleting three columns from
  the sheet silently misaligned 18 rows written by positional `appendRow`.
  Match headers on prefix — `Joining Date` is really
  `Joining Date/Approx Joining Month`.
- **Add columns at the end**, via `ensureColumn_`, so existing data never
  shifts.
- **Fill blanks; don't overwrite.** Imports and backfills are re-run as data
  arrives over time.
- Issued receipts are historical records — **do not rewrite them**. Fix data
  going forward instead.

## Style

Saurav wants a recommendation, not a menu of options, and will overrule with
domain knowledge — treat that as authoritative. Say plainly when something is
unverified.
