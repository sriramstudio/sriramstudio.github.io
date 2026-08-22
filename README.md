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

**The database** is a Google Sheet with 4 tabs: `Enrollments`, `Receipts`, `Config`,
and `Legacy Students` (the canonical list of students who joined before the
registration form existed — see Section 11).

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

- [ ] **`sriramstudio_admin.html`** — find `id="r-feetype"` (~line 640s), edit the `<option>` list
- [ ] Upload admin file to GitHub

Current fee types: Monthly Fee, Registration Fee, Workshop, Other

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
- **Receipt tab** → PIN-protected. Type a name (autocompletes from records), enter amount, fee type, date received, note → Generate → Copy for WhatsApp or Print/Save PDF
- **Records tab** → view/search/filter all enrollments, applications, and receipts

### ID formats
- **Enrollment ID:** `SR-YYYY-MMDDHHmmssSSS` (timestamp-based, never duplicates, auto-updates year)
- **Receipt No:** `SS-YYYY-0001` (sequential; counter is `receipt_seq` in Config sheet — change it there to reset)

---

## 7. How to Update `Code.gs` (the backend)

### The quick way — clasp (from the project folder)

The project folder is a git repository wired to GitHub, and `clasp` pushes
`Code.gs` straight to Apps Script. Two commands:

```
clasp push
clasp redeploy AKfycbzRB6XQrT8lx2lgcgv9nPQ18akiRCU8x58ritJIgbjNRkKxE3FZAifMREpZ4f7ZgSWD -d "what changed"
```

Redeploying to that **existing** deployment id creates a new version at the
**same URL**, so no HTML needs touching. Never use `clasp deploy` — that makes
a *new* deployment with a *new* URL and breaks `register.html`.

`.claspignore` restricts the push to `Code.gs` and `appsscript.json`, so the
GitHub Pages HTML is never uploaded into the script project.

To roll back: `clasp redeploy <same-id> -V <older version number>`

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

## 8. Google Sheet Structure (reference)

**Enrollments tab columns:**
ID · Enrolled At · Type · Student Name · Date of Birth · Gender · Blood Group · School/College · Guardian Name · Relation · Phone · WhatsApp · Email · Address · Program · Location · Batch · Joining Date · Pracheen Kala Kendra · Workshop Name · Workshop Date · Workshop Fee · Heard From · Notes · **Status**

`Status` is `Active` or `Left`. Blank counts as active. A student marked `Left`
still appears in the receipt autocomplete (so old receipts can be reprinted) but
is flagged there and in Records, so nobody bills them by accident. To mark
someone as left, type `Left` in their `Status` cell.

**Legacy Students tab columns:**
Student Name · Contact · Center

The canonical list of pre-form students. `Center` has also been used for status
words like "discontinue"; the import understands that and files it as `Status`
rather than as a branch. Add a `Status` column if you ever need to record both a
centre and a status for the same student.

*(Program, Batch, Pracheen columns are legacy — no longer filled by the forms but kept for old records.)*

**Receipts tab columns:**
Receipt No · Issued At · Student Name · Contact · Amount (₹) · Fee Month · Fee Year · Payment Mode · UPI Reference · Fee Type · Date Received · Note

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
| `previewLegacyStudents` | Reports what the roster import would do. Changes no rows (it does create the `Status` column header if missing). |
| `importLegacyStudents` | Adds students from `Legacy Students` to `Enrollments` as `Existing Student`. Safe to re-run. |
| `addStatusColumn` | Creates the `Status` column on its own. |

The import is an **upsert**: a name already in `Enrollments` is not added again,
and blank `Phone` / `Location` / `Status` cells are filled from the roster when
it later gains that detail. Existing values are never overwritten — so you can
add phone numbers and centres to the roster over time and re-run it.

It **refuses to guess** in two cases, listing them instead: a name repeated in
the roster with no phone to tell those students apart, and a name already
enrolled more than once where the roster row has no phone to say which is meant.
Add a phone number to those rows and run it again.

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
