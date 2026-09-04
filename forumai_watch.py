#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = ["feedparser", "requests", "openpyxl", "pypdf", "gspread>=6", "google-auth"]
# ///
"""
forumai_watch.py -- keep the FORUM-AI publication list alive without relying on self-reporting.

Two channels, same triage:
  A. arXiv        author sweep via the arXiv API (same-day; every FORUM-AI paper so far was here first).
  B. OpenAlex     author sweep by pinned OpenAlex author IDs. Catches everything with a Crossref DOI:
                  ChemRxiv / Research Square preprints, journal articles, proceedings. A day or two behind.

For every new record: get the full text (arXiv HTML -> arXiv PDF -> open-access PDF from OpenAlex ->
`fulltext-download <doi>` if that tool is installed and we are on a subscribing network), then

  ACK         text names FORUM-AI, paper not on the main list         -> tab "new_papers"
  NO-FUNDING  text readable, no funding statement of any kind         -> tab "no_funding"
  UNREADABLE  no text; Crossref funder metadata absent or lists DOE   -> tab "no_funding", note says why
  skip        explicit funding statement (or Crossref funders) that never mentions FORUM-AI,
              e.g. D2S2, ASCR, BES, TRI.  Not our business (Anubhav, 2026-09-04).  Silent.
  skip        datasets, errata, corrections, paratext.

Never show the same paper twice: a paper is processed if its arXiv ID, DOI, or title appears in ANY
cell of ANY tab of the tracker workbook, or in seen.json (every ID this script ever wrote). So Lien can
move a row to the main list or delete it outright; either way it will not come back.

Usage (living list in the Google Sheet):
  uv run forumai_watch.py --gsheet <spreadsheet URL or ID> --creds service_account.json --days 14
Usage (local xlsx copy, markdown report to stdout):
  uv run forumai_watch.py --sheet "FORUM-AI Paper Tracking.xlsx" --days 30
  uv run forumai_watch.py --since 2026-01-01 --all --no-update-seen     # recall check, changes nothing
One-time backfill (seeds seen.json; the only run that reads more than a handful of full texts):
  uv run forumai_watch.py --since 2025-09-01 --sheet "FORUM-AI Paper Tracking.xlsx"
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
from pathlib import Path

import feedparser
import requests

# --- configuration -------------------------------------------------------------------------

# display name, arXiv au: tokens, OpenAlex author IDs, affiliation regex (annotation only).
# arXiv author search tolerates middle initials ("Chan_Maria" matches "Maria K. Y. Chan"); list every
# first-name form a PI actually uses (Zwart is "Petrus H. Zwart" on arXiv). OpenAlex splits some people
# across several author IDs; list them all. Affiliation is checked for names with known collisions
# ("Anubhav Jain" also publishes computer-vision papers from NYU) and only annotates, never drops.
PIS = [
    ("Anubhav Jain",             ["Jain_Anubhav"],             ["A5003640520", "A5084934954"], r"Lawrence Berkeley|LBNL|LBL\b"),
    ("Kristin Persson",          ["Persson_Kristin"],          ["A5037535334", "A5103952818"], None),
    ("Gerbrand Ceder",           ["Ceder_Gerbrand"],           ["A5014983956"],                None),
    ("Tirthankar Ghosal",        ["Ghosal_Tirthankar"],        ["A5081072666"],                None),
    ("Massimiliano Lupo Pasini", ["Lupo_Pasini_Massimiliano"], ["A5019108731"],                None),
    ("Khaled Ibrahim",           ["Ibrahim_Khaled"],           ["A5062016419"],                r"Lawrence Berkeley|LBNL|LBL\b|NERSC"),
    ("Hanqi Guo",                ["Guo_Hanqi"],                ["A5054749881"],                None),
    ("Markus Buehler",           ["Buehler_Markus"],           ["A5011504360"],                None),
    ("Peter Zwart",              ["Zwart_Peter", "Zwart_Petrus"], ["A5081770773", "A5028559746"], None),
    ("Maria Chan",               ["Chan_Maria"],               ["A5036691276"],                r"Argonne|ANL\b"),
]
FIRST_NAMES = {"Peter Zwart": {"peter", "petrus"}}   # accepted first-name forms beyond the display name
AFFIL = {d: a for d, _t, _o, a in PIS}

ACK_RE = re.compile(r"\bFORUM[\s\-‐-―]?AI\b", re.IGNORECASE)
# Evidence that the paper carries an explicit funding statement of some kind.
FUNDING_RE = re.compile(
    r"(supported (in part )?by|funded (in part )?by|funding (from|was|is|provided)|financial support|"
    r"under (contract|grant|award)|grant (no|number|#)|award (no|number|#)|\bDE-(AC|SC|FG|EE|AR)\d|"
    r"National Science Foundation|\bNSF\b|Department of Energy|\bDOE\b|Office of Science|"
    r"Research Council|Horizon (Europe|2020)|\bERC\b|\bNIH\b|\bDARPA\b|\bONR\b|\bAFOSR\b|\bARO\b)",
    re.IGNORECASE)
DOE_RE = re.compile(r"Department of Energy|\bDOE\b|Office of Science|Basic Energy Sciences|"
                    r"Advanced Scientific Computing", re.IGNORECASE)
SKIP_TYPES = {"dataset", "erratum", "paratext", "retraction", "peer-review", "editorial", "letter",
              "conference-abstract", "other", "supplementary-materials"}
SKIP_SOURCES = re.compile(r"Meeting Abstracts|Zenodo|figshare|eScholarship", re.I)   # abstracts and mirrors
# FORUM-AI is AI/ML for science. A record whose title+abstract has none of this vocabulary (and, on arXiv,
# no cs.* / stat.ML category) is not fetched at all. Checked against every paper on the tracking sheet.
TOPIC_RE = re.compile(
    r"(machine[- ]learning|deep[- ]learning|artificial intelligence|\bAI\b|\bML\b|\bLLMs?\b|language models?|"
    r"foundation models?|neural|agents?\b|agentic|generative|transformer|graph neural|reinforcement learning|"
    r"fine[- ]tun|benchmark|knowledge graph|conformal|uncertainty quantif|surrogate|autonomous|self-driving|"
    r"data[- ]driven|multimodal|large language|GPT|hypothesis generation|literature mining|text mining|"
    r"interatomic potential|active learning|Bayesian optimi|inverse design|representation learning)", re.I)
ARXIV_CS_RE = re.compile(r"^(cs\.|stat\.ML)")

PROGRAM_START = dt.date(2025, 9, 1)   # FORUM-AI began Sept 2025; anything earlier is definitely not ours

ARXIV_ID_RE = re.compile(r"\b(\d{4}\.\d{4,5})(?:v\d+)?\b")
DOI_RE = re.compile(r"\b(10\.\d{4,9}/[^\s\"'<>]+)", re.IGNORECASE)
ARXIV_API = "https://export.arxiv.org/api/query"
OPENALEX_API = "https://api.openalex.org/works"
UA = {"User-Agent": "forumai-watch/0.2 (mailto:ajain@lbl.gov)"}


def pi_on_paper(display: str, authors: list[str]) -> bool:
    """Surname + first-name check on returned author strings, so arXiv fuzz-matches don't count."""
    first, *_, last = display.split()
    firsts = FIRST_NAMES.get(display, {first.lower()})
    for a in authors:
        toks = a.replace(".", " ").split()
        if toks and toks[-1].lower() == last.lower() and toks[0].lower() in firsts:
            return True
        if last.lower() == "pasini" and "pasini" in a.lower() and first.lower() in a.lower():
            return True
    return False


def norm_title(t: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (t or "").lower()).strip()


def norm_doi(d: str | None) -> str | None:
    if not d:
        return None
    d = d.strip().lower()
    d = re.sub(r"^https?://(dx\.)?doi\.org/", "", d)
    return d.rstrip(".,;)")


def doi_family(d: str | None) -> str | None:
    """ChemRxiv / Research Square mint a DOI per version (.../v1, /v2): one paper, one key."""
    return re.sub(r"/v\d+$", "", d) if d else d


# --- channel A: arXiv ----------------------------------------------------------------------

def arxiv_sweep(since: dt.date, until: dt.date, only: set[str] | None) -> dict[str, dict]:
    window = f"submittedDate:[{since:%Y%m%d}0000 TO {until:%Y%m%d}2359]"
    papers: dict[str, dict] = {}
    for display, tokens, _oa, _af in PIS:
        if only and display not in only:
            continue
        for token in tokens:
            start = 0
            while True:
                q = urllib.parse.urlencode({"search_query": f"au:{token} AND {window}", "start": start,
                                            "max_results": 100, "sortBy": "submittedDate", "sortOrder": "descending"})
                feed = feedparser.parse(f"{ARXIV_API}?{q}", request_headers=UA)
                for e in feed.entries:
                    aid = ARXIV_ID_RE.search(e.id).group(1)
                    authors = [a.name for a in e.authors]
                    if not pi_on_paper(display, authors):
                        continue
                    p = papers.setdefault(aid, {
                        "key": aid, "arxiv": aid, "doi": None, "title": " ".join(e.title.split()),
                        "authors": authors, "date": e.published[:10], "link": f"https://arxiv.org/abs/{aid}",
                        "source": "arXiv", "type": "preprint", "pis": [], "oa_pdf": None, "funders": [],
                        "abstract": " ".join(e.summary.split()), "cats": [t.term for t in e.tags],
                    })
                    if display not in p["pis"]:
                        p["pis"].append(display)
                total = int(feed.feed.get("opensearch_totalresults", 0))
                start += 100
                if start >= total or not feed.entries:
                    break
                time.sleep(3)
            time.sleep(3)  # arXiv API etiquette
    return papers


# --- channel B: OpenAlex -------------------------------------------------------------------

def inv_abstract(inv: dict | None) -> str:
    if not inv:
        return ""
    pos = sorted((i, w) for w, idxs in inv.items() for i in idxs)
    return " ".join(w for _, w in pos)


def prefilter(p: dict) -> str | None:
    """Reason to drop the record WITHOUT fetching full text, or None to keep it."""
    if p["funders"] and not any(DOE_RE.search(f) for f in p["funders"]):
        return "funders-not-DOE"
    if TOPIC_RE.search(p["title"] + " " + p.get("abstract", "")) or any(ARXIV_CS_RE.match(c) for c in p["cats"]):
        return None
    return "off-topic"

def openalex_sweep(since: dt.date, only: set[str] | None) -> dict[str, dict]:
    """Works by any PI with publication_date >= since, excluding anything that has an arXiv location
    (channel A owns those). Keyed by DOI. Publication dates are unreliable for 'newness'; the
    seen-ledger is what makes an item new, so callers pass a generous `since`."""
    papers: dict[str, dict] = {}
    for display, _t, oa_ids, _af in PIS:
        if (only and display not in only) or not oa_ids:
            continue
        cursor = "*"
        while cursor:
            r = requests.get(OPENALEX_API, headers=UA, timeout=60, params={
                "filter": f"authorships.author.id:{'|'.join(oa_ids)},from_publication_date:{since}",
                "per-page": 200, "cursor": cursor,
                "select": "id,doi,title,publication_date,type,primary_location,best_oa_location,locations,"
                          "authorships,funders,awards,abstract_inverted_index"})
            r.raise_for_status()
            j = r.json()
            for w in j["results"]:
                if (w.get("publication_date") or "9999") < str(PROGRAM_START):
                    continue
                if any("arxiv" in json.dumps(l).lower() for l in (w.get("locations") or [])):
                    continue
                if (w.get("type") or "") in SKIP_TYPES:
                    continue
                src = ((w.get("primary_location") or {}).get("source") or {}).get("display_name") or "?"
                if SKIP_SOURCES.search(src):
                    continue
                doi = norm_doi(w.get("doi"))
                key = doi_family(doi) or w["id"]
                oa = w.get("best_oa_location") or {}
                p = papers.setdefault(key, {
                    "key": key, "arxiv": None, "doi": doi, "title": " ".join((w.get("title") or "").split()),
                    "authors": [a["author"]["display_name"] for a in w.get("authorships") or []],
                    "date": w.get("publication_date") or "", "link": f"https://doi.org/{doi}" if doi else w["id"],
                    "source": src, "type": w.get("type") or "", "pis": [],
                    "oa_pdf": oa.get("pdf_url") or oa.get("landing_page_url"),
                    "funders": [f["display_name"] for f in w.get("funders") or []],
                    "abstract": inv_abstract(w.get("abstract_inverted_index")), "cats": [],
                })
                if display not in p["pis"]:
                    p["pis"].append(display)
            cursor = j["meta"].get("next_cursor") if j["results"] else None
            time.sleep(0.2)
    return papers


# --- full text -----------------------------------------------------------------------------

def pdf_to_text(data: bytes) -> str:
    with tempfile.TemporaryDirectory() as d:
        pdf = Path(d, "p.pdf"); pdf.write_bytes(data)
        if shutil.which("pdftotext"):
            subprocess.run(["pdftotext", str(pdf), str(Path(d, "p.txt"))], check=False, capture_output=True)
            if Path(d, "p.txt").exists():
                return Path(d, "p.txt").read_text(errors="ignore")
        try:
            from pypdf import PdfReader
            return "\n".join((pg.extract_text() or "") for pg in PdfReader(str(pdf)).pages)
        except Exception:
            return ""


def fetch_pdf_text(url: str) -> str:
    try:
        # preprint servers behind Cloudflare refuse non-browser user agents for PDFs
        hdr = {"User-Agent": "Mozilla/5.0 (Macintosh) forumai-watch/0.2 (mailto:ajain@lbl.gov)", "Accept": "application/pdf,*/*"}
        r = requests.get(url, headers=hdr, timeout=120, allow_redirects=True)
    except requests.RequestException:
        return ""
    if r.ok and r.content[:5] == b"%PDF-":
        return pdf_to_text(r.content)
    return ""


def server_pdf_url(p: dict) -> str | None:
    """Direct PDF for preprint servers whose OpenAlex record has no pdf_url."""
    doi = p["doi"] or ""
    if "biorxiv" in p["source"].lower() or "medrxiv" in p["source"].lower():
        host = "medrxiv" if "medrxiv" in p["source"].lower() else "biorxiv"
        return f"https://www.{host}.org/content/{doi_family(doi)}v1.full.pdf"
    m = re.match(r"10\.21203/rs\.3\.(rs-\d+)(?:/v(\d+))?$", doi)          # Research Square
    if m:
        return f"https://www.researchsquare.com/article/{m.group(1)}/v{m.group(2) or 1}.pdf"
    return None


def fetch_fulltext(p: dict, use_tool: bool) -> tuple[str, str]:
    """(text, how). how in {'arxiv-html','arxiv-pdf','oa-pdf','server-pdf','fulltext-download',''}."""
    if p["arxiv"]:
        r = requests.get(f"https://arxiv.org/html/{p['arxiv']}", headers=UA, timeout=60)
        if r.ok and "ltx_page_main" in r.text:
            return re.sub(r"<[^>]+>", " ", r.text), "arxiv-html"
        t = fetch_pdf_text(f"https://arxiv.org/pdf/{p['arxiv']}")
        return (t, "arxiv-pdf") if t else ("", "")
    if p["oa_pdf"]:
        t = fetch_pdf_text(p["oa_pdf"])
        if t:
            return t, "oa-pdf"
    if (u := server_pdf_url(p)):
        t = fetch_pdf_text(u)
        if t:
            return t, "server-pdf"
    if use_tool and p["doi"] and shutil.which("fulltext-download"):
        with tempfile.TemporaryDirectory() as d:
            try:
                subprocess.run(["fulltext-download", p["doi"], d, "paper.pdf"], check=False,
                               capture_output=True, timeout=300)
            except subprocess.TimeoutExpired:
                pass
            f = Path(d, "paper.pdf")
            if f.exists() and f.read_bytes()[:5] == b"%PDF-":
                return pdf_to_text(f.read_bytes()), "fulltext-download"
    return "", ""


def snippet(text: str, m: re.Match, before=160, after=80) -> str:
    return " ".join(text[max(0, m.start() - before): m.end() + after].split())


def classify(p: dict, text: str) -> None:
    p["ack"] = ""
    if text:
        m = ACK_RE.search(text)
        if m:
            p["tier"], p["ack"] = "ACK", snippet(text, m)
        elif FUNDING_RE.search(text):
            p["tier"] = "skip"
        else:
            p["tier"] = "NO-FUNDING"
        p["unconfirmed"] = [d for d in p["pis"] if AFFIL.get(d) and not re.search(AFFIL[d], text, re.I)]
        return
    p["unconfirmed"] = []
    # No text. Fall back on the Crossref funder metadata OpenAlex carries.
    if p["funders"] and not any(DOE_RE.search(f) for f in p["funders"]):
        p["tier"] = "skip"                       # explicitly funded by someone who is not DOE
    else:
        p["tier"] = "UNREADABLE"
        if "chemrxiv" in p["source"].lower():
            p["note"] = "ChemRxiv blocks automated download (Cloudflare); open the link and read the acknowledgement."
        else:
            p["note"] = ("full text not reachable; Crossref funders: " + (", ".join(p["funders"]) or "none listed")
                         + ". Check the acknowledgement by hand.")


# --- the tracker workbook -------------------------------------------------------------------

def processed_from_cells(cells) -> tuple[set[str], set[str]]:
    """Every arXiv ID, DOI, and long-cell title found anywhere -> already processed."""
    ids: set[str] = set(); titles: set[str] = set()
    for c in cells:
        if not isinstance(c, str) or not c:
            continue
        ids.update(ARXIV_ID_RE.findall(c))
        ids.update(norm_doi(d) for d in DOI_RE.findall(c))
        if len(c) > 25:
            titles.add(norm_title(c))
    return ids, titles


def load_xlsx(path: Path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    return processed_from_cells(c for ws in wb.worksheets for row in ws.iter_rows(values_only=True) for c in row)


def gsheet_open(ref: str, creds: Path | None):
    import gspread
    gc = gspread.service_account(filename=str(creds)) if creds else gspread.oauth()
    return gc.open_by_url(ref) if ref.startswith("http") else gc.open_by_key(ref)


def load_gsheet(sh):
    return processed_from_cells(c for ws in sh.worksheets() for row in ws.get_all_values() for c in row)


TAB_HEADER = ["first_seen", "id", "link", "date", "source", "type", "title", "FORUM-AI PIs",
              "all authors", "acknowledgement text", "note"]
TAB_FOR_TIER = {"ACK": "new_papers", "NO-FUNDING": "no_funding", "UNREADABLE": "no_funding"}


def gsheet_append(sh, papers: list[dict], today: dt.date) -> dict[str, int]:
    counts: dict[str, int] = {}
    for tab in sorted(set(TAB_FOR_TIER.values())):
        rows = [p for p in papers if TAB_FOR_TIER.get(p["tier"]) == tab]
        if not rows:
            continue
        try:
            ws = sh.worksheet(tab)
        except Exception:
            ws = sh.add_worksheet(title=tab, rows=200, cols=len(TAB_HEADER))
        if not ws.get_all_values():
            ws.append_row(TAB_HEADER)
        ws.append_rows([[str(today), p["key"], p["link"], p["date"], p["source"], p["type"], p["title"],
                         ", ".join(p["pis"]), ", ".join(p["authors"]), p.get("ack", ""), p.get("note", "")]
                        for p in rows], value_input_option="RAW")
        counts[tab] = len(rows)
    return counts


# --- main ----------------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--days", type=int, default=30, help="arXiv look-back window (default 30)")
    ap.add_argument("--since", type=dt.date.fromisoformat, help="explicit start date, overrides --days")
    ap.add_argument("--openalex-slack", type=int, default=45,
                    help="OpenAlex window starts this many days before --since, because publication dates "
                         "lag record creation (default 45); seen.json keeps this from repeating items")
    ap.add_argument("--sheet", type=Path, help="local copy of the tracker .xlsx (report-only mode)")
    ap.add_argument("--gsheet", help="Google Sheet URL or ID of the tracker (living-list mode: appends rows)")
    ap.add_argument("--creds", type=Path, help="service-account JSON for --gsheet; omit to use browser OAuth")
    ap.add_argument("--seen", type=Path, default=Path("seen.json"), help="ledger of already-handled IDs")
    ap.add_argument("--all", action="store_true", help="also show tracked/seen items (recall check)")
    ap.add_argument("--no-update-seen", action="store_true")
    ap.add_argument("--no-openalex", action="store_true", help="arXiv channel only")
    ap.add_argument("--no-fulltext-tool", action="store_true", help="never call fulltext-download")
    ap.add_argument("--pi", action="append", help="restrict to this PI display name (repeatable)")
    ap.add_argument("--dry-run", action="store_true", help="sweep and prefilter only; fetch nothing, write nothing")
    args = ap.parse_args()

    today = dt.date.today()
    since = max(args.since or (today - dt.timedelta(days=args.days)), PROGRAM_START)
    oa_since = max(since - dt.timedelta(days=args.openalex_slack), PROGRAM_START)
    only = set(args.pi) if args.pi else None
    sh = gsheet_open(args.gsheet, args.creds) if args.gsheet else None
    tracked_ids, tracked_titles = load_gsheet(sh) if sh else (load_xlsx(args.sheet) if args.sheet else (set(), set()))
    seen = set(json.loads(args.seen.read_text())) if args.seen.exists() else set()

    log = lambda *a: print(*a, file=sys.stderr)
    log(f"# FORUM-AI watch {today}: arXiv since {since}; OpenAlex since {oa_since}")
    papers = arxiv_sweep(since, today, only)
    log(f"arXiv: {len(papers)} candidates")
    if not args.no_openalex:
        oa = openalex_sweep(oa_since, only)
        # a DOI record whose title we already hold from arXiv is the same paper
        have = {norm_title(p["title"]) for p in papers.values()}
        oa = {k: v for k, v in oa.items() if norm_title(v["title"]) not in have}
        log(f"OpenAlex: {len(oa)} non-arXiv candidates")
        papers.update(oa)

    stage = {"swept": len(papers), "already known": 0, "funders-not-DOE": 0, "off-topic": 0, "full text fetched": 0}
    for p in papers.values():
        p["tracked"] = (p["key"] in tracked_ids or (p["arxiv"] or "") in tracked_ids
                        or norm_title(p["title"]) in tracked_titles)
        p["seen"] = p["key"] in seen
        p["tier"] = "skip"
        if not args.all and (p["tracked"] or p["seen"]):
            stage["already known"] += 1; continue
        why = prefilter(p)
        if why:
            stage[why] += 1
            if args.dry_run and p["tracked"]:
                log(f"  WOULD DROP a tracked paper ({why}): {p['title'][:80]}")
            continue
        stage["full text fetched"] += 1
        if args.dry_run:
            continue
        text, how = fetch_fulltext(p, use_tool=not args.no_fulltext_tool)
        p["how"] = how
        classify(p, text)
        time.sleep(1)
    log("stages: " + ", ".join(f"{k} {v}" for k, v in stage.items()))
    if args.dry_run:
        return 0

    order = {"ACK": 0, "NO-FUNDING": 1, "UNREADABLE": 2}
    shown = sorted((p for p in papers.values() if p["tier"] != "skip"), key=lambda p: (order[p["tier"]], p["date"]))
    out = [f"# FORUM-AI watch -- {today}",
           f"{stage['swept']} records swept; {stage['already known']} already known; "
           f"{stage['funders-not-DOE'] + stage['off-topic']} dropped on metadata; "
           f"{stage['full text fetched']} full texts read; {len(shown)} to look at.", ""]
    for tier, heading in [("ACK", "## Acknowledges FORUM-AI, not on the main list"),
                          ("NO-FUNDING", "## No funding statement found"),
                          ("UNREADABLE", "## Full text not reachable, funding unknown")]:
        group = [p for p in shown if p["tier"] == tier]
        if not group:
            continue
        out += [heading, ""]
        for p in group:
            flag = " *(already on sheet)*" if p["tracked"] else (" *(reported before)*" if p["seen"] else "")
            out.append(f"- **{p['title']}**{flag}  ")
            out.append(f"  {p['date']} · {p['source']} ({p['type']}) · {p['link']} · PIs: {', '.join(p['pis'])} · "
                       f"{len(p['authors'])} authors · text: {p.get('how') or 'none'}")
            if p.get("unconfirmed"):
                out.append(f"  affiliation not found in text for: {', '.join(p['unconfirmed'])} (name collision?)")
            if p.get("note"):
                out.append(f"  {p['note']}")
            if p.get("ack"):
                out.append(f"  > …{p['ack']}…")
            out.append("")
    if not shown:
        out.append("Nothing new.")
    if sh and shown and not args.all:
        counts = gsheet_append(sh, shown, today)
        out.append("Appended to the tracker: " + ", ".join(f"{n} rows -> {t}" for t, n in counts.items()))
    print("\n".join(out))

    if not args.no_update_seen:
        seen.update(p["key"] for p in papers.values())
        args.seen.write_text(json.dumps(sorted(seen), indent=0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
