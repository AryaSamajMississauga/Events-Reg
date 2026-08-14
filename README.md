# Arya Samaj — Geet Sangeet Registration & Check-In

Two web apps that share one Google Sheet as their live database, plus an Excel master file for your own tracking.

## Files

- **`index.html`** — guest registration form. Collects details, calculates totals ($20 adult / $5 kids under 12), and issues a QR ticket once payment is confirmed.
- **`checkin.html`** — volunteer tool for the door. Scans a guest's QR code (or looks them up by code/name) and marks them checked in.
- **`test.html`** — a diagnostic page. Open it to check whether the apps can reach your Google Sheet, and see the exact server response if something's wrong.
- **`Code.gs`** — the Google Apps Script backend (runs inside the Google Sheet, deployed separately — not hosted on GitHub).
- **`GeetSangeet-Master.xlsx`** — an Excel workbook for organizing registrations, tracking verified payments, and marking check-ins by hand. See "Excel" section below.

---

## 1. Publish the web pages with GitHub Pages

1. Create a new repo at [github.com/new](https://github.com/new) (e.g. `geet-sangeet-2026`). Public repos get free Pages hosting.
2. On the repo page: **Add file → Upload files** → drag in `index.html`, `checkin.html`, `test.html` (and `Code.gs`, `README.md` if you want them versioned) → **Commit changes**.
3. **Settings → Pages** → Source: **Deploy from a branch** → Branch: `main`, folder `/ (root)` → **Save**.
4. After a minute you get a live URL like `https://YOUR-USERNAME.github.io/geet-sangeet-2026/`.
   - Registration form: that URL (loads `index.html` automatically)
   - Volunteer check-in: add `checkin.html`
   - Connection test: add `test.html`

Command-line alternative:
```bash
git init
git add .
git commit -m "Geet Sangeet apps"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/geet-sangeet-2026.git
git push -u origin main
```

---

## 2. If registrations aren't saving — check the connection

Open **`test.html`** on your live site and press "Test Get All."

- See JSON (like `[]` or a list) → **it works.** ✅
- See an HTML / "Sign in" page → your Apps Script deployment access is wrong. Fix: Apps Script → **Deploy → Manage deployments → ✏️ Edit → "Who has access" = Anyone → Version: New version → Deploy.**
- See "Failed to fetch" / a network error → the `/exec` URL is wrong or the script isn't deployed. Re-copy the `/exec` URL (not `/dev`) and paste it into `API_URL` at the top of the `<script>` in `index.html`, `checkin.html`, and `test.html`.

**Note:** the apps only work from a real hosted page (like GitHub Pages), not from inside a preview window. The preview sandbox blocks the call to Google.

---

## 3. Excel master file

`GeetSangeet-Master.xlsx` has four tabs:
- **Instructions** — a legend.
- **Registrations** — one row per booking. Pale-yellow cells are yours to fill; the green row is an example to overwrite. Paid / Payment Verified / Checked In use Yes/No dropdowns.
- **Attendees** — one row per person (Price fills in automatically from Ticket Type).
- **Summary** — automatic totals: guests, adults, kids, expected revenue, collected (verified only), outstanding, and check-ins.

**Why Excel is a master file, not the live database:** connecting a web app to *live* two-way Excel storage needs Microsoft Graph or Office Scripts + Power Automate — a lot of setup, often a paid plan. Google Sheets + Apps Script (already built) is far simpler and free, so the recommended setup is: **apps write to Google Sheets live → whenever you want, File → Download → Microsoft Excel (.xlsx) from the Google Sheet, or paste rows into this master file.** You get Excel for organizing without the heavy live-sync plumbing.

---

## Notes

- Payment is a manual e-transfer to `Arya.samaj.miss@gmail.com`. "I've Sent the Payment" in the app is self-reported — always confirm against the real inbox (use the Payment Verified column in the Excel file).
- The check-in page can see the whole guest list, so share that link only with volunteers, not on public flyers.
