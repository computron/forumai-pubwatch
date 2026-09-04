# forumai-pubwatch

Finds FORUM-AI publications without waiting for PIs to self-report, and feeds them into the
**FORUM-AI Paper Tracking** Google Sheet for review.

Every Monday a GitHub Action:

1. sweeps **arXiv** (author search, same-day) and **OpenAlex** (every Crossref DOI: ChemRxiv, journals,
   proceedings) for the ten FORUM-AI PIs, from 2025-09-01 onward;
2. skips anything already in the tracker workbook (any tab, matched by arXiv ID, DOI, or title) or in
   `seen.json`, and drops records whose metadata already rules them out (non-DOE funders, off-topic,
   datasets, meeting abstracts);
3. reads the full text of the few survivors and appends rows to two tabs of the sheet:
   - `new_papers` — acknowledges FORUM-AI, not yet on the main list
   - `no_funding` — no funding statement found, or full text not reachable (the note column says which)
4. commits `seen.json` so no paper is ever surfaced twice.

Reviewers only touch the Google Sheet: move a row to the main list, or delete it. Nothing else.
A deleted row does not come back. To resurface a paper on purpose, remove its ID from `seen.json`.

## One-time Google setup (about five minutes, needs your Google login)

1. <https://console.cloud.google.com/> → pick or create a project → **APIs & Services → Library** →
   enable **Google Sheets API** and **Google Drive API**.
2. **IAM & Admin → Service Accounts → Create service account** (any name, no roles needed) →
   open it → **Keys → Add key → JSON**. A `.json` file downloads.
3. Open the tracker Google Sheet → **Share** → paste the service account's e-mail
   (`…@…iam.gserviceaccount.com`) → **Editor**. If LBL's Workspace refuses to share outside lbl.gov,
   see the fallback below.
4. From a terminal, in a clone of this repo:

       gh secret set GSHEET_URL --body "https://docs.google.com/spreadsheets/d/<the sheet id>/edit"
       gh secret set GSHEET_SA_JSON < ~/Downloads/<the downloaded key>.json

5. **Actions → forumai-watch → Run workflow** to do the first live run. It creates the two tabs.

Until step 4 is done the workflow still runs, but in report-only mode: the digest shows up under the
run's **Summary** and the ledger is not advanced.

**Fallback if sharing to a service account is blocked:** run the script on a Mac instead of in Actions,
with your own login: `uv run forumai_watch.py --gsheet <url> --days 14` (no `--creds`) opens a browser
consent once via gspread's OAuth flow; schedule it with launchd or cron.

## Running locally

    uv run forumai_watch.py --days 14 --sheet "FORUM-AI Paper Tracking.xlsx"      # report against a local xlsx
    uv run forumai_watch.py --since 2026-01-01 --all --dry-run                     # what would be swept, fetch nothing
    uv run forumai_watch.py --help

Dependencies are declared inline (PEP 723); `uv` installs them. `pdftotext` (poppler) is optional but faster.
Full-text access to paywalled journals is the residual gap: those rows land in `no_funding` with a note
listing the Crossref funders, for a by-hand check. ChemRxiv blocks scripted downloads and is handled the same way.

## Tuning

Everything configurable is at the top of `forumai_watch.py`: the PI table (arXiv name forms, OpenAlex author
IDs), the acknowledgement regex, the funding-statement and topic vocabularies, `PROGRAM_START`, tab names.
