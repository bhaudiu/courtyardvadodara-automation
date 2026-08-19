# Courtyard Vadodara — Daily Automation (setup)

This is the same mechanism as Stoneriver: a **GitHub Action** runs on a schedule,
fetches the daily reports from Gmail, rebuilds the dashboard data, and commits it
back — Cloudflare Pages then auto-deploys. No server, no fixed time.

## What's in this folder
- `.github/workflows/deploy.yml` — the hourly workflow (checks every hour, 24/7)
- `fetch_email.py` — pulls the 3 reports from the Gmail inbox (IMAP)
- `build_site.py` — folds the day into the store and rebuilds `data.json`
- `parse_reports.py` — parses Manager Flash + Trial Balance + Daily Operations
- `store.json` — persistent history (24 days seeded) + budget/YTD anchor
- `data.json` — the current dashboard data
- `requirements.txt` — Python deps (openpyxl, pdfplumber)

## One-time setup (≈5 min)

### 1. Add these files to the `courtyardvadodara-automation` repo
Keep the folder structure — especially `.github/workflows/deploy.yml`.
(Claude can push them for you when your Chrome is connected.)

### 2. Create a Gmail App Password for cyv.ai.automation@gmail.com
- Turn on **2-Step Verification** for that Google account.
- Google Account → **Security → App passwords** → create one → copy the 16-character password.

### 3. Add 3 repository Secrets
Repo → **Settings → Secrets and variables → Actions → New repository secret**:
| Name | Value |
|------|-------|
| `IMAP_HOST` | `imap.gmail.com` |
| `IMAP_USER` | `cyv.ai.automation@gmail.com` |
| `IMAP_PASS` | *(the 16-char app password)* |

### 4. Allow the workflow to commit
Repo → **Settings → Actions → General → Workflow permissions** → select
**Read and write permissions** → Save.

### 5. Forward the reports
Have the property team forward (or auto-forward) the 3 daily reports —
**Manager Flash**, **Trial Balance**, **Daily Operations** — to
**cyv.ai.automation@gmail.com**.

## How it behaves
- Runs **every hour**, scanning the last ~60 emails. Whenever all 3 of the day's
  reports have arrived, it parses them, appends the day, and publishes — so it's
  tied to *when reports land*, not a fixed clock time.
- If reports haven't arrived yet, it does nothing and checks again next hour.
- Idempotent: re-running on the same reports produces no change / no redeploy.
- You can also trigger it manually anytime: repo → **Actions → Refresh Courtyard DBR → Run workflow**.

## Notes / limits
- The full Budget / MTD / YTD report is reconstructed from the daily history plus
  your budget anchor. Last-year comparisons use the 2025 daily data currently in
  the sheet (Aug 1–12); add more 2025 history anytime to extend year-over-year.
- The dashboard is public unless you add Cloudflare Access (login gate) — still recommended.
