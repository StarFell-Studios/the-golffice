/** @OnlyCurrentDoc */
/**
 * Content review vote store.
 * Lives inside a Google Sheet (Extensions > Apps Script). Deploy as a web app:
 * Execute as: Me. Who has access: Anyone.
 *
 * Sheet "Votes" keeps one row per reviewer per piece (the latest choice).
 * Sheet "History" keeps every change with a timestamp.
 */

const KEY = 'tee-time-2026';           // must match key in config.js
const VOTES = 'Votes';
const HISTORY = 'History';
const HEAD = ['Updated', 'Batch', 'Reviewer', 'Item ID', 'Item', 'Vote', 'Note'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== KEY) return json({ ok: false, error: 'bad_key' });
  try {
    if (p.action === 'vote') return json(saveVote(p));
    if (p.action === 'list') return json(listVotes(p.batch || ''));
    return json({ ok: true, ping: 'ready' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function saveVote(p) {
  const vote = String(p.vote || '');
  if (['approve', 'maybe', 'no', ''].indexOf(vote) < 0) return { ok: false, error: 'bad_vote' };
  const batch = clip(p.batch, 60), reviewer = clip(p.reviewer, 40).toLowerCase();
  const item = clip(p.item, 80), title = clip(p.title, 120), note = clip(p.note, 1500);
  if (!batch || !reviewer || !item) return { ok: false, error: 'missing' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const sh = sheet(VOTES);
    const data = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === batch && data[i][2] === reviewer && data[i][3] === item) { row = i + 1; break; }
    }
    const values = [now, batch, reviewer, item, title, vote, note];
    if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
    else sh.appendRow(values);
    sheet(HISTORY).appendRow(values);
    color(sh);
    return { ok: true, saved: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function listVotes(batch) {
  const data = sheet(VOTES).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!batch || data[i][1] === batch) {
      out.push({ updated: data[i][0], reviewer: data[i][2], item: data[i][3], vote: data[i][5], note: data[i][6] });
    }
  }
  return { ok: true, votes: out };
}

function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEAD);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function color(sh) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const votes = sh.getRange(2, 6, n - 1, 1).getValues();
  const bg = votes.map(function (r) {
    return [r[0] === 'approve' ? '#d9ead3' : r[0] === 'maybe' ? '#fff2cc' : r[0] === 'no' ? '#f4cccc' : '#ffffff'];
  });
  sh.getRange(2, 6, n - 1, 1).setBackgrounds(bg);
}

function clip(v, n) { return String(v == null ? '' : v).slice(0, n); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create both sheets. */
function setup() { sheet(VOTES); sheet(HISTORY); }
