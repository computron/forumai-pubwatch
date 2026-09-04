# forumai-pubwatch

Finds FORUM-AI publications without waiting for PIs to self-report, and feeds them into the
**FORUM-AI Paper Tracking** Google Sheet for review. Reviewers only ever touch the sheet.

## How it works

Every Monday morning a GitHub Action (`forumai_watch.py`):

1. sweeps **arXiv** (author search, same-day) and **OpenAlex** (every Crossref DOI: ChemRxiv, journals,
   proceedings) for the ten FORUM-AI PIs, from 2025-09-01 onward;
2. skips everything already in `seen.json`, and drops records whose metadata already rules them out
   (non-DOE funders, off-topic, datasets, meeting abstracts) -- typically a handful of records survive;
3. reads the full text of the survivors and sorts them:
   - acknowledges FORUM-AI -> `new_papers`
   - explicit funding statement without FORUM-AI -> dropped silently (not our business)
   - no funding statement, or text not reachable -> `no_funding` (the note says which);
4. appends those to **`candidates.json`** (public, cumulative, bibliographic data only) and commits it
   together with `seen.json`.

Every Tuesday morning a small Apps Script inside the Google Sheet (`Code.gs`) fetches `candidates.json` and
appends every item not already in the workbook to the `new_papers` / `no_funding` tabs, remembering what it
appended in a hidden tab. Reviewing a row = move it to the main list, or delete it. It never comes back.

No credentials anywhere: the Action only reads public APIs and writes to this repo; the script in the sheet
runs as the sheet's owner and can touch only that sheet.

## One-time install in the sheet (two minutes)

1. Open the tracker Google Sheet -> **Extensions -> Apps Script**.
2. Delete the placeholder `function myFunction() {}` and paste the whole of `Code.gs`. Click the save icon.
3. In the toolbar, pick `setup` in the function dropdown and click **Run**. Google asks for permission once
   ("this script wants to edit this spreadsheet and connect to an external service"): **Review permissions
   -> choose your account -> Allow**. If it says "Google hasn't verified this app", click Advanced -> Go to
   ... (unsafe) -> Allow; it is your own script, in your own sheet.
4. Back in the sheet: two new tabs, `new_papers` and `no_funding`, filled from the backfill. A
   **FORUM-AI watch** menu appears for manual pulls. The weekly pull is now scheduled.

## Local use

    uv run forumai_watch.py --days 14 --sheet "FORUM-AI Paper Tracking.xlsx"      # markdown report against a local xlsx
    uv run forumai_watch.py --since 2026-01-01 --dry-run                           # what would be swept; fetch nothing
    uv run forumai_watch.py --help

Dependencies are declared inline (PEP 723); `uv` installs them. `pdftotext` (poppler) is optional but faster.

## Known gaps

- Paywalled journal articles cannot be read; they land in `no_funding` with the Crossref funder list for a
  by-hand check (about 17 in the whole first year). FORUM-AI has no award number of its own, so funder
  metadata cannot settle them.
- ChemRxiv blocks scripted downloads (Cloudflare); same treatment.
- To resurface a paper on purpose, delete its row from the hidden `_forumai_seen` tab (View -> Hidden sheets).

## Tuning

Everything configurable is at the top of `forumai_watch.py`: the PI table (arXiv name forms, OpenAlex author
IDs), the acknowledgement regex, the funding-statement and topic vocabularies, `PROGRAM_START`, tab names.
