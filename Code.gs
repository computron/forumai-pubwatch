/**
 * FORUM-AI publication watch -- the sheet side.
 *
 * Lives inside the "FORUM-AI Paper Tracking" Google Sheet (Extensions -> Apps Script). Once a week it pulls
 * the public candidates feed that the GitHub Action (github.com/computron/forumai-pubwatch) maintains and
 * appends anything not already in this workbook to two tabs:
 *   new_papers   acknowledges FORUM-AI, not yet on the main list
 *   no_funding   no funding statement found, or full text not reachable (see the note column)
 *
 * Review = move the row to the main list, or delete it. Either way it never comes back: what has been
 * appended is remembered in a hidden tab (_forumai_seen), and anything whose arXiv ID, DOI or title appears
 * anywhere in the workbook is skipped too.
 *
 * Install: paste this file, click Run on `setup` once, click Allow. That installs the weekly timer and does a
 * first pull. "FORUM-AI watch -> Pull new papers now" appears in the sheet's menu for manual pulls.
 */
const FEED_URL = 'https://raw.githubusercontent.com/computron/forumai-pubwatch/main/candidates.json';
const SEEN_TAB = '_forumai_seen';
const HEADER = ['first_seen', 'id', 'link', 'date', 'source', 'type', 'title', 'FORUM-AI PIs',
                'all authors', 'acknowledgement text', 'note'];

function setup() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'pull') ScriptApp.deleteTrigger(t);
  });
  // The Action runs Monday morning (US Pacific); pull Tuesday morning in the sheet's time zone.
  ScriptApp.newTrigger('pull').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(8).create();
  const added = pull();
  SpreadsheetApp.getUi().alert('FORUM-AI watch installed. Weekly pull every Tuesday 8am.\nFirst pull added: '
                               + JSON.stringify(added));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('FORUM-AI watch').addItem('Pull new papers now', 'pull').addToUi();
}

function pull() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = UrlFetchApp.fetch(FEED_URL, {muteHttpExceptions: true});
  if (resp.getResponseCode() !== 200) throw new Error('Feed fetch failed: HTTP ' + resp.getResponseCode());
  const feed = JSON.parse(resp.getContentText());
  const known = knownKeys_(ss);
  const seen = seenSheet_(ss);
  const seenIds = new Set(seen.getLastRow()
      ? seen.getRange(1, 1, seen.getLastRow(), 1).getValues().flat().map(String) : []);
  const added = {};
  feed.forEach(c => {
    const key = String(c.id).toLowerCase();
    if (seenIds.has(key) || known.ids.has(key) || known.titles.has(norm_(c.title))) return;
    tab_(ss, c.tab).appendRow([c.first_seen, c.id, c.link, c.date, c.source, c.type, c.title,
                               c.pis, c.authors, c.ack, c.note]);
    seen.appendRow([key, new Date()]);
    seenIds.add(key);
    added[c.tab] = (added[c.tab] || 0) + 1;
  });
  Logger.log('added: ' + JSON.stringify(added));
  return added;
}

/** Every arXiv ID, DOI (version-stripped) and long-cell title anywhere in the workbook. */
function knownKeys_(ss) {
  const ids = new Set(), titles = new Set();
  const arx = /\b(\d{4}\.\d{4,5})(?:v\d+)?\b/g;
  const doi = /\b(10\.\d{4,9}\/[^\s"'<>]+)/gi;
  ss.getSheets().forEach(sh => {
    if (sh.getName() === SEEN_TAB) return;
    sh.getDataRange().getValues().forEach(row => row.forEach(c => {
      if (typeof c !== 'string' || !c) return;
      let m;
      while ((m = arx.exec(c)) !== null) ids.add(m[1]);
      while ((m = doi.exec(c)) !== null) ids.add(m[1].toLowerCase().replace(/\/v\d+$/, '').replace(/[.,;)]+$/, ''));
      if (c.length > 25) titles.add(norm_(c));
    }));
  });
  return {ids, titles};
}

function norm_(t) {
  return String(t || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function tab_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADER);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(7, 420);
  }
  return sh;
}

function seenSheet_(ss) {
  let sh = ss.getSheetByName(SEEN_TAB);
  if (!sh) {
    sh = ss.insertSheet(SEEN_TAB);
    sh.hideSheet();
  }
  return sh;
}
