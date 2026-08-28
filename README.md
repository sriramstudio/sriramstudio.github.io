# Sriram Studio — Digital Administration System

**Operational Runbook & Change Guide**
Last updated: August 2026

This document explains what each file does, how the system fits together, and — most importantly — the **exact steps to follow whenever you make any change**, so nothing is ever missed.

---

## 1. System Overview

Sriram Studio's system has **4 files** working together:

| File | Where it lives | Who uses it | Purpose |
|------|---------------|-------------|---------|
| `index.html` | GitHub Pages (public) | Everyone | The main studio website — `sriramstudio.github.io` |
| `register.html` | GitHub Pages (public) | Students / Parents | Public self-registration form |
| `sriramstudio_admin.html` | GitHub Pages (private URL) | You & Anjali | Admin panel — enroll, receipts, records |
| `Code.gs` | Google Apps Script (inside the Google Sheet) | Runs automatically | The backend / database engine |

**The database** is a Google Sheet with 3 tabs: `Enrollments`, `Receipts`, `Config`.

**How data flows:**
```
register.html  ─┐
                ├─→  Code.gs (Apps Script)  ─→  Google Sheet (database)
sriramstudio_admin.html ─┘                          │
                                                     └─→  Email notification to Anjali
```

---

## 2. The Golden Rules (read before ANY change)

1. **After downloading any fresh file, always re-check the URL / script settings before uploading to GitHub.** Fresh copies may reset placeholders.

2. **`register.html` needs the `SCRIPT_URL` set manually every time it is freshly downloaded.** This is the #1 thing people forget. See Section 4.

3. **A change in one place usually needs the same change in several places.** Adding a centre, a dance style, or changing contact details touches MULTIPLE files and MULTIPLE sections within each file. This document lists every spot.

4. **After changing `Code.gs`, you must REDEPLOY (new version), not just Save.** See Section 7.

5. **After uploading to GitHub, always hard-refresh** (Ctrl+Shift+R) or add `?v=2` to the URL to bypass browser cache.

---

## 3. Current Studio Details (single source of truth)

Keep this section updated. When any of these change, follow the matching checklist in Section 5.

**Centres (4):**
- Bhawanipur – 53A Girish Mukherjee Road, Kolkata – 700025
- Wood Street – 6/4 Govind Mahal, 3 Wood Street, Kolkata – 700016
- Kankurgachi – P-328 CIT Road, Kolkata – 700054
- Salt Lake – 34, DA Block, Sector 1, Salt Lake, Kolkata – 700064

**Dance styles (4):** Odissi (Classical), Semi-Classical, Bollywood, Western

**Contact:** +91 98307 92201 · sriramstudio.anjali@gmail.com

**Social:** Instagram/YouTube @sriramstudio.official · Facebook sriramstudio.dance

**Notification email (receives new applications):** prasad.anjali.91@gmail.com

**Apps Script URL (current):**
`https://script.google.com/macros/s/AKfycbzRB6XQrT8lx2lgcgv9nPQ18akiRCU8x58ritJIgbjNRkKxE3FZAifMREpZ4f7ZgSWD/exec`

**Script ID** (for clasp — see Section 7):
`13g1OBFlt9ffsyQn_XfAw87W-MFbLE2CiOlllmcTMBzq7Zz5Sx7RfeO0_`

**Deployment ID** (the live one; redeploy to THIS id to keep the URL):
`AKfycbzRB6XQrT8lx2lgcgv9nPQ18akiRCU8x58ritJIgbjNRkKxE3FZAifMREpZ4f7ZgSWD`

---

## 4. How to Upload / Replace a File on GitHub

1. Go to **github.com** → sign in → open your repository
2. Click **Add file → Upload files**
3. Drag the file in (filename MUST match exactly, e.g. `register.html`)
4. Click **Commit changes**
5. Wait 1–3 minutes for GitHub Pages to publish
6. Open the page and **hard-refresh** (Ctrl+Shift+R) or add `?v=2` to the URL

**⚠ For `register.html` ONLY — before uploading:**
1. Open the file in Notepad (or Notepad++ / VS Code)
2. Find the line (near the top of the `<script>` section):
   `const SCRIPT_URL = '...';`
3. Make sure it contains the current Apps Script URL (see Section 3), NOT `YOUR_APPS_SCRIPT_URL_HERE`
4. Save, then upload

*Note: `sriramstudio_admin.html` does NOT need this — it asks for the URL once on its setup screen and remembers it in the browser.*

---

## 5. Change Checklists

### ➤ 5A. Adding (or editing) a CENTRE / BRANCH

A centre appears in **3 files across 7 places**. Miss one and the centre won't show everywhere.

**File 1 — `index.html`** (the website) — 4 places:
- [ ] **Meta description** (near top, line ~6): update the "Four locations" count word
- [ ] **About text** (~line 325): update "across four locations" count word
- [ ] **Stats pill** (~line 332): update the number in `<span class="stat-num">4</span> Locations`
- [ ] **Branches section** (~line 353): add a new `<div class="branch-card">` block. Copy an existing card, change the name, address, and the Google Maps `query=` link

**File 2 — `register.html`** (public form) — 3 places:
- [ ] **Admission location dropdown** (~line 278): add `<option>New Centre – Address</option>`
- [ ] **Workshop location dropdown** (~line 298): add the SAME `<option>` here too
- [ ] **Footer strip** (~line 371): add centre name to `Bhawanipur · Wood Street · ...`

**File 3 — `sriramstudio_admin.html`** (admin panel) — 3 places:
- [ ] **Admission/Existing dropdown** (~line 477): add `<option>New Centre – Address</option>`
- [ ] **Workshop dropdown** (~line 517): add the SAME `<option>` here too
- [ ] **Receipt footer** (~line 1197): add centre name to `📍 Bhawanipur · Wood Street · ...`

**Then:** upload all 3 files to GitHub (remember the SCRIPT_URL check for register.html).
**Update Section 3 of this README** with the new centre.

> **Important:** the exact `<option>` text must be IDENTICAL in register.html and admin, because receipts and records match on this text. Copy-paste, don't retype.

---

### ➤ 5B. Adding (or editing) a DANCE STYLE

Dance styles appear in **1 file, 1 place** (they were removed from the forms earlier):
- [ ] **`index.html`** Dance Styles section (~line 341): add `<span class="style-tag">New Style</span>`
- [ ] If the count changes, update the About text if it mentions a number

Upload `index.html` to GitHub.

---

### ➤ 5C. Changing CONTACT DETAILS (phone / email / social)

Phone, email, and social handles appear in **multiple files**:

**`index.html`** — Contact section (~line 389 onwards): phone link, email link, Instagram, YouTube, Facebook rows, AND the "Save Contact" vCard block (~line 429)

**`register.html`** — thank-you screen contact strip and error message (search for `98307`)

**`sriramstudio_admin.html`** — receipt footer + WhatsApp copy text (search for `98307`)

**`Code.gs`** — email notification footer (search for `98307` or `sriramstudio.github.io`)

- [ ] Search each file for the old value and replace everywhere
- [ ] Update Section 3 of this README
- [ ] Upload changed files; redeploy Code.gs if it changed

---

### ➤ 5D. Changing the NOTIFICATION EMAIL

This is stored in the Google Sheet, NOT in code:
- [ ] Open Google Sheet → **Config** tab
- [ ] Find the `notify_email` row → change the value in column B
- [ ] Done. No file upload, no redeploy needed.

*(Same for `notify_whatsapp` and `callmebot_key`.)*

---

### ➤ 5E. Changing FEE TYPES (receipt dropdown)

- [ ] **`sriramstudio_admin.html`** — find `id="r-feetype"` (~line 650s), edit the `<option>` list
- [ ] Update the list below, and Section 11A if the change affects fee coverage
- [ ] Upload admin file to GitHub

Current fee types: **Monthly Fee**, **Registration Fee**,
**Uniform / Costume Fee**, **Late Fee**, **Workshop**, **Other**

The dropdown is the only list — nothing else in the code enumerates fee types,
so adding one is a single edit. `Code.gs` never hardcodes the names; it asks
only whether a fee type *starts with* "Monthly".

> **What that means for the analytics.** Only `Monthly Fee` (or a blank fee
> type, for older receipts) makes a month count as paid. Registration, uniform,
> late and workshop fees are money, but they are not a month's tuition, so they
> are counted in revenue and deliberately excluded from fee coverage. A late fee
> in particular must never mark a month paid — the late fee is the penalty, the
> tuition is separate. If you ever add a fee type that *should* cover a month,
> its name must begin with "Monthly", or `buildFeeCoverage_` needs changing.

---

## 6. How to Use the System (daily operations)

### Public Registration (`sriramstudio.github.io/register.html`)
Share this link on Instagram bio, WhatsApp, Google Business. Students/parents fill it themselves. Each submission:
- Lands in the Google Sheet `Enrollments` tab (Type = "Application – …")
- Triggers an email to the notification address
- Shows the applicant a thank-you screen with a Reference ID

### Admin Panel (`sriramstudio.github.io/sriramstudio_admin.html`)
Asks for the **PIN as soon as it loads** — the PIN now guards the whole panel,
not just the Receipt tab, because Records and Enroll are authenticated too.
The PIN is verified **on the server**: the web app URL is public (it is embedded
in `register.html`), so the PIN is the only thing between a visitor and the
student database. Ten wrong attempts locks the endpoint for 15 minutes.

- **Enroll tab** → 3 modes: New Admission, Workshop, Existing Student (for adding your current roster)
- **Receipt tab** → Type a name and pick from the autocomplete; the student becomes a chip. **Pick more than one to club siblings onto a single receipt** — one amount, one receipt number, all names printed. Remove a chip with its ×. Typing a name without picking still works for someone not yet in the database. Then amount, fee type, date received, note → Generate → Copy for WhatsApp or Print/Save PDF
- **Records tab** → view/search/filter all enrollments, applications, and receipts

### ID formats
- **Enrollment ID:** `SR-YYYY-MMDDHHmmssSSS` (timestamp-based, never duplicates, auto-updates year)
- **Receipt No:** `SS-YYYY-0001` (sequential; counter is `receipt_seq` in Config sheet — change it there to reset)

---

## 7. How to Update `Code.gs` (the backend)

### The quick way — `deploy.cmd` (from the project folder)

Open PowerShell in the project folder and run one of:

```
.\deploy.cmd
.\deploy.cmd "what changed"
```

- **No description** → pushes `Code.gs` only. This is all you need for anything
  run from the Apps Script editor or the **Analytics** menu, because the editor
  always runs the code you just pushed.
- **With a description** → pushes, then redeploys the live web app. Use this
  when the admin panel or the public form needs the change.

It **runs the tests first and refuses to push if any fail**, so a broken
`Code.gs` cannot reach the sheet. It redeploys to the **existing** deployment
id, which keeps the URL, so no HTML needs touching.

To roll back: `clasp.cmd redeploy <same-id> -V <older version number>`

> **Why `.cmd` and not the bare `clasp`?** PowerShell's execution policy blocks
> npm's `clasp.ps1` shim — you get *"running scripts is disabled on this
> system"*. `clasp.cmd` is the very same program and is not blocked, so nothing
> needs loosening. The same applies to `npm.cmd`. Also, PowerShell 5.1 has no
> `&&`, so run commands one at a time.
>
> If `clasp.cmd` itself is not found, install it:
> `npm.cmd install -g @google/clasp`, then `clasp.cmd login`.

`.claspignore` restricts the push to `Code.gs` and `appsscript.json`, so neither
the GitHub Pages HTML nor `deploy.cmd` is ever uploaded into the script project.

Never use `clasp deploy` — that makes a *new* deployment with a *new* URL and
breaks `register.html`. `deploy.cmd` only ever uses `clasp redeploy`.

### The manual way (if clasp is unavailable)

1. Open your Google Sheet → **Extensions → Apps Script**
2. Select all existing code (Ctrl+A) → delete → paste the new `Code.gs` → **Save** (Ctrl+S)
3. **Deploy → Manage deployments**
4. Click the **pencil/edit icon** on the existing deployment
5. Under "Version" select **New version**
6. Click **Deploy**

**⚠ Saving alone does NOT update the live app. You MUST redeploy a new version.**

**To verify it updated:** in the editor, check the function you changed shows the new code. Under Manage deployments, confirm the version number went up.

**The URL stays the same** across redeployments — no need to update the HTML files (unless you create a brand-new deployment instead of a new version, which generates a new URL — avoid that).

---

## 7A. If something goes wrong — rolling back

Every working state is tagged before a risky change, so going back is one
command. Nothing here rewrites history or deletes anything.

```
.\rollback.cmd            put Code.gs and the admin panel back
.\rollback.cmd --list     show which states you can go back to
```

It restores the two files into your folder, runs the tests against them, and
stops. **It does not push.** It then prints the two commands that make the
restore live — `.\deploy.cmd` for the backend, `git push` for the admin panel.

Changed your mind before making it live?
`git checkout HEAD -- Code.gs sriramstudio_admin.html` puts the newer code
straight back.

**The Google Sheet is not touched by any of this.** No rollback deletes a
column or a row. Where a change adds columns they go at the *end*, so leaving
them behind after a rollback is harmless — the older code simply ignores them.

**Before any change that touches receipts**, take a copy of the data too:
right-click the `Receipts` tab → **Duplicate**. That is the real safety net and
it takes two seconds.

To roll back the *deployed* web app rather than the code, see the end of
Section 7: `clasp.cmd redeploy <same-id> -V <older version number>`.

**Tags so far:**

| Tag | State |
|-----|-------|
| `pre-fee-split` | Everything working before the per-student, per-fee-type receipt split. |

---

## 8. Google Sheet Structure (reference)

**Enrollments tab columns:**
ID · Enrolled At · Type · Student Name · Date of Birth · Gender · Blood Group · School/College · Guardian Name · Relation · Phone · WhatsApp · Email · Address · Program · Location · Batch · Joining Date · Pracheen Kala Kendra · Workshop Name · Workshop Date · Workshop Fee · Heard From · Notes · **Status** · **Left On** · **Review**

`Status` is `Active` or `Left`. Blank counts as active. To mark someone as left,
type `Left` in their `Status` cell — `left`, `discontinued`, `inactive`,
`stopped` and `dropped` are accepted too. Record the last fee month and year in
`Left On`, e.g. `March 2027`.

**A student marked `Left` disappears from the receipt autocomplete.** Their row
and receipt history stay intact and they are still visible in Records — they
simply cannot be picked for a new receipt.

**If they rejoin, they get a new row.** Do not reactivate the old one: the old
record stays `Left` as the history of that spell, and the new row is the live
one. Because left students are hidden from autocomplete, there is no ambiguity
about which to pick.

`Review` is set automatically when a new enrolment matches an existing name —
`Rejoining? Earlier record SR-... is marked Left` when the match has left, or
`Possible duplicate of active record SR-...` when it has not. Records shows a
**REVIEW** badge. Check it, then clear the cell. A genuinely new student who
happens to share a name is never blocked, only flagged.

**Receipts tab columns** — the first thirteen, in the order a receipt is
written: Receipt No · Issued At · Student Name · Contact · Amount (₹) ·
Fee Month · Fee Year · Payment Mode · UPI Reference · Fee Type · Date Received ·
Note · Students

Then, added at the end as they are first needed: **`Fee Split`** and one
**`<fee type> ₹`** column per fee type. Section 11A explains what they hold.

> **Never insert, delete or reorder the first thirteen.** A receipt is written
> by position, so a moved column would file the amount or the fee type under
> the wrong heading — silently, and only for *new* receipts. **Add columns at
> the END only.** Run `auditReceiptColumns` after any change to this tab; it
> says plainly whether new receipts will still land correctly.

**Config tab keys:**
- `pin` — admin panel PIN (default 1234)
- `receipt_seq` — next receipt number (change to reset numbering)
- `notify_email` — where new-application emails go
- `notify_whatsapp` — WhatsApp number for CallMeBot (optional)
- `callmebot_key` — CallMeBot API key (optional)

---

## 9. Common Problems & Fixes

| Problem | Cause | Fix |
|---------|-------|-----|
| "Form not configured" on register page | SCRIPT_URL not set in register.html | Section 4 — set the URL, re-upload |
| Change not showing after upload | Browser cache / GitHub Pages delay | Hard-refresh (Ctrl+Shift+R) or add `?v=2` to URL |
| Code.gs change not taking effect | Only saved, not redeployed | Section 7 — redeploy new version |
| Autocomplete not finding students | — | Fixed; uses safe column lookup. If it breaks, a deleted sheet column is likely |
| Email not received | Permission not granted, or wrong email in Config | Run any function in Apps Script once to grant permission; check Config `notify_email` |
| Receipt prints blank / multiple pages | Old file version | Use latest admin file (print CSS fixed) |
| Setup screen asks every time (admin) | File opened locally, not via GitHub URL | Always open via `sriramstudio.github.io/...` URL |
| "Unauthorised." from the panel | Browser is running an old cached copy that does not send the PIN | Hard-refresh, or add `?v=2` (bump the number) to the URL |
| Autocomplete finds nothing on one device | Same cause as above — stale cached page | Hard-refresh that device. On iPhone use the `?v=` trick; there is no Ctrl+Shift+R |
| Locked out for 15 minutes | Ten failed PIN attempts. The counter is script-wide, so a phone quietly retrying an old cached page can lock everyone out | Wait it out, and hard-refresh any device still on the old copy |

---

## 10. Ownership & Access

- The Google Sheet + Apps Script are owned by **Saurav's** personal Google account.
- **Anjali (prasad.anjali.91@gmail.com)** has **Editor** access to the sheet — she can view and edit data directly, but the script runs under Saurav's account.
- Notification emails are sent from the owner's Gmail.
- To fully hand over to Anjali later: transfer sheet ownership, then she re-authorises and redeploys the script under her account (generates a new URL that must be updated in the HTML files).

---

---

## 11. Maintenance functions (Apps Script editor only)

These are **not** reachable over the web — they only run from the editor.
Open the Sheet → Extensions → Apps Script, pick the function from the dropdown,
click Run, then read the **Execution log**.

| Function | What it does |
|----------|--------------|
| `addStatusColumn` | Creates the `Status`, `Left On` and `Review` columns if missing. |
| `auditEnrollments` | Reports what each column actually holds, and flags values sitting under the wrong header. Writes nothing. |
| `auditEnrollmentRow` | Dumps single rows header-by-header. Pass a row number, or nothing for a spread. Writes nothing. |
| `auditReceiptColumns` | Checks the `Receipts` headers still match the order a new receipt is written in. **Run after any change to the Receipts columns.** Writes nothing. |
| `previewColumnRepair` | Shows which rows are misaligned and exactly what would move. Writes nothing. |
| `repairColumnAlignment` | Realigns them. Backs the tab up first and refuses if any value would be lost. |
| `dropEmptyOverflowColumns` | Removes unheadered columns past `Notes`, but only once they are empty. |
| `previewJoiningBackfill` | Shows which enrolments would get a joining month from their enrolment date. Writes nothing. |
| `backfillJoiningMonth` | Fills those in. Blank cells only; never overwrites. |

Every one of these was written for a specific one-off job and left in place in
case the same problem recurs. The `preview` half of each pair always writes
nothing — run it, read the log, and only then run its counterpart.

---

## 11A. Analytics — has everyone paid?

Also editor-only, and every one of them writes nothing. None is reachable over
the web: they read the whole student database, so they stay off `doGet`.

| Function | What it answers |
|----------|-----------------|
| `analyticsReport` | The overview — roster, centres, ages, receipts, revenue, and who is behind. |
| `feeCoverageByMonth` | **Month by month: how many students were due, how many a receipt covers, and whether the month is fully paid.** Start here. |
| `feeCoverageForMonth('August 2026')` | One month in full: every student with no receipt, with phone and centre, ready to chase. Run with no argument for the latest month that has receipts. |
| `previewReceiptNameMatching` | How every name printed on a receipt was tied back to a row in `Enrollments`. **Read this before believing the other three.** |
| `previewMultiMonthReceipts` | Every receipt read as settling more than one month, and why. |
| `feeGapsByStudent` | Per student, which months have no receipt naming them. |

**How a month is counted.** `Enrollments` is the golden source. A student is
expected to pay from their `Joining Date` month until today, or until the month
in `Left On` if they have left. Workshop rows and applications nobody acted on
are not expected to pay a monthly fee — unless a monthly receipt names them, in
which case they are. Only `Monthly Fee` receipts (and receipts with the fee
type left blank) cover a month; registration and workshop fees do not.

**Several students and several fee types on one receipt.** A family settles
everything at once — a monthly fee each for two sisters, a costume fee for one,
a late fee for the other. Anjali enters one card per student with as many fee
lines as she needs; the total is computed from the lines and cannot disagree
with them. **Typing a fee for one student copies it to the others**, because
siblings usually pay the same, and it only ever fills a blank — anything
already typed stands.

The sheet keeps this in two places. `Fee Split` holds the per-student detail:

```
Riya Sen: Monthly Fee 2000; Uniform / Costume Fee 800 | Diya Sen: Late Fee 100
```

and a column per fee type (`Monthly Fee ₹`, `Uniform / Costume Fee ₹` …) holds
that type's total across the receipt, so a column can simply be summed. Both
are created on first use and always **at the end**, so the thirteen columns a
receipt has always used never move. Receipts issued before this leave them
blank and keep the older whole-receipt rule.

**Why the split matters:** fee coverage judges a split receipt *per student*.
The sister who paid a monthly fee is covered for the month; the one who paid
only a costume fee is not, and stays on the collections list. Without it, one
receipt would mark both siblings paid and the second would silently disappear.

**Two months on one receipt.** Anjali can tick a span on the receipt form — a
first month and an "up to" month — and it is recorded in a **`Fee Months`**
column. Where that column is filled it is the answer, and the note is not
consulted: stated beats inferred. `Fee Month` / `Fee Year` still hold the FIRST
month, so every older report reads the same.

For receipts issued before that existed, the months are still read from the
note. Families often settle two months together. The
`Fee Month` column holds only one month, so the second is read from the **note**
— `July and August fees`, `2 months fee - Jul & Aug`. Both months then count as
paid for every child on that receipt. `December and January` correctly rolls
into the next year.

Naming a month is not the same as paying for it. `Fee for June & late fee for
May` settles **June only** — May's is a penalty, and May's own tuition may still
be outstanding. Same for `100 advanced for August`: an advance is not August's
fee. `Balance of July` **is** counted, because that money is July's fee. Every
month held back this way is listed by `previewMultiMonthReceipts`.

Three deliberate limits. A note naming a *single* month different from `Fee Month`
does **not** add a month — that is a mis-keyed `Fee Month`, which
`findDuplicateReceipts` already reports. And a note saying `2 months fee`
without naming which two is **not** guessed at; it is listed by
`previewMultiMonthReceipts` so the note can be corrected. Naming the months in
the note is all it takes.

The money stays in the month the receipt is filed under — that is when the cash
came in. Only the *coverage* spans both months.

**Siblings.** One receipt covering `Riya Sen | Diya Sen` counts for both
children, and so does an older receipt with `Riya & Diya Sen` typed into one
cell — the last child on the receipt carries the surname for the rest. Names
run together with no separator at all are handled too: `Anshika Anvika Shome`
is read as two sisters, and `Jeena Jia` names one child from her first name and
lends her surname to the other. This is only accepted when every piece lands on
a *different* child, so a single mangled name is never split by accident.

**Spelling.** A receipt name one or two characters away from exactly one roster
name is credited to that student. Anything further apart, or equally close to
two children, is credited to nobody and listed instead — so a month may show as
short until you look. A sibling who is not on the roster at all is never
credited to the sibling who is.

**First name only.** A receipt reading just `Hiral` is credited when exactly one
child answers to that name. When two do, it is reported instead — and the report
names them both, with their rows.

**Two children, one name.** Where the name cannot be told apart on its own, the
`Contact` number on the receipt is matched against `Phone` in `Enrollments` and
settles it. **This is why filling in phone numbers matters**: two children
called Krisha Agarwal were unattributable until their numbers were on the
roster. For an exact shared name, a child being the only one enrolled that
month will also settle it; for a near-miss spelling that is *not* enough, and
the report says so rather than picking one.

**Names that need telling by hand** live at the top of the fee-coverage section
of `Code.gs`:

```
const NAME_EQUIVALENTS = [
  ['Vani Maskara', 'Vaani Maskara'],
  ['Krishanya Kanoria', 'Krishnaya Kanodia'],
  ['Jisha Desai', 'Josh’s Desai', "Josh's Desai"]
];
const DISTINCT_STUDENTS = [
  ['Jia Bhimani', 'Jeena Bhimani'],
  ['Jahnvi Jain', 'Janhvi Jain']
];
```

`NAME_EQUIVALENTS` is one child spelled two ways — whichever spelling is in
`Enrollments` wins, and the other is treated as that child on receipts.
`DISTINCT_STUDENTS` is the opposite: two real children with similar names,
pinned to their own rows so the spelling matcher can never merge them.
`previewReceiptNameMatching` prints the state of both lists, including whether
each configured name was actually found in `Enrollments` — check that first if
a pair is not behaving. Add to either list as new cases turn up.

---

## 11B. The dashboard tabs (for Anjali)

Everything in 11A also exists as ordinary tabs in the Sheet, so nobody has to
open the Apps Script editor.

**The `Analytics` menu** sits in the Sheet's own menu bar, next to Help. It
appears when the Sheet is opened — after pushing a new `Code.gs`, reload the
Sheet once for it to show up.

| Menu item | What it does |
|-----------|--------------|
| **Refresh dashboard** | Rebuilds all five tabs from the current data. |
| **Preview refresh** | Says what a refresh would do. Writes nothing. |
| **Name matching report** | Section 11A's name report, in a window. |
| **Two-month receipts report** | Every receipt read as settling more than one month. |
| **Fee gaps report** | Who is behind, in a window. |

**The five tabs:**

| Tab | What it holds |
|-----|---------------|
| `Analytics Dashboard` | Pick a month; the counts, the fee-type breakdown and the unpaid list update instantly. |
| `Analytics Coverage` | One row per month — due, paid, unpaid, collected — and a chart. |
| `Analytics Students` | One row per student, one column per month, PAID/UNPAID colour-coded. Filter it. |
| `Analytics Fees` | What was collected each month, split by fee type, with a Total column. The fee types are read from the receipts, so a new one appears here by itself. |
| `Analytics Names` | Every name printed on a receipt and the row it was credited to. Anything needing a decision sorts to the top. |
| `Analytics Gaps` | Students missing at least one month, worst first, with phone numbers. |

**The month picker is a formula, not a script.** Changing the month on the
Dashboard updates it at once — no waiting, no permission prompt. Only a change
to `Enrollments` or `Receipts` needs **Refresh dashboard**.

**Safety.** A refresh writes only to those five tabs. `Enrollments`, `Receipts`
and `Config` are opened read-only and a write to them is refused outright.
Every generated tab is stamped with a hidden marker in the note on cell A1; if
a tab of the same name exists without that marker, the refresh **stops and
changes nothing** rather than overwriting somebody's work. Deleting a generated
tab is safe — the next refresh rebuilds it.

Do not type into the generated tabs. Edits there are lost on the next refresh.

---

## 12. Testing

`tests/auth.test.js` runs `Code.gs` against stubbed Apps Script globals — no
network, no sheet. From the project folder:

```
node tests/auth.test.js Code.gs
```

It covers the PIN gate, the lockout, the public registration path staying open,
and the roster import's matching rules. Run it after any change to `Code.gs`.

---

---

*End of runbook. Keep Section 3 updated as the single source of truth, and follow the matching checklist in Section 5 for every change.*
