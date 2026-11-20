/** @OnlyCurrentDoc */
/**
 * Weber Mortgage onboarding questionnaire answer store.
 * Lives inside its own Google Sheet (Extensions > Apps Script). Deploy as a web app:
 * Execute as: Me. Who has access: Anyone.
 *
 * Sheet "Answers" keeps the latest answer to each question, in questionnaire order.
 * Sheet "Team" keeps the agents Michael hired, including custom ones.
 * Sheet "History" keeps every save with a timestamp and the full answers as JSON.
 */

const KEY = 'weber-home-2026';          // must match key in config.js
const ANSWERS = 'Answers', TEAM = 'Team', HISTORY = 'History', RAW = 'Raw';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== KEY) return json({ ok: false, error: 'bad_key' });
  try {
    if (p.action === 'load') return json(load());
    return json({ ok: true, ping: 'ready' });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json({ ok: false, error: 'bad_json' }); }
  if (body.key !== KEY) return json({ ok: false, error: 'bad_key' });
  try { return json(save(body.data || {})); } catch (err) { return json({ ok: false, error: String(err) }); }
}

function save(d) {
  const raw = JSON.stringify(d);
  if (raw.length > 200000) return { ok: false, error: 'too_big' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const now = new Date(), ans = d.answers || {}, qs = d.questions || [];
    // Answers: one row per question, in order
    const rows = qs.map(function (q) {
      const v = ans[q.id];
      return [q.section || '', clip(q.label, 300), Array.isArray(v) ? v.join(', ') : (v == null ? '' : String(v))];
    });
    const a = sheet(ANSWERS, ['Section', 'Question', 'Answer']);
    a.getRange(2, 1, Math.max(a.getLastRow() - 1, 1), 3).clearContent();
    if (rows.length) a.getRange(2, 1, rows.length, 3).setValues(rows);
    a.getRange(1, 5).setValue('Status: ' + (d.status || 'draft') + ' · last saved ' + now.toLocaleString());
    // Team
    const t = sheet(TEAM, ['Job title', 'Name', 'What they do', 'Great job after 90 days', 'Custom']);
    t.getRange(2, 1, Math.max(t.getLastRow() - 1, 1), 5).clearContent();
    const team = (d.team || []).map(function (m) { return [clip(m.title, 120), clip(m.name, 80), clip(m.job, 1500), clip(m.success, 1500), m.custom ? 'yes' : '']; });
    if (team.length) t.getRange(2, 1, team.length, 5).setValues(team);
    // History and the raw copy the page reloads from
    sheet(HISTORY, ['Saved', 'Status', 'Answers JSON']).appendRow([now, d.status || 'draft', clip(raw, 49000)]);
    const r = sheet(RAW, ['Latest JSON']); r.getRange(2, 1).setValue(clip(raw, 49000));
    return { ok: true, saved: now.toISOString() };
  } finally { lock.releaseLock(); }
}

function load() {
  const r = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAW);
  const v = r ? r.getRange(2, 1).getValue() : '';
  return { ok: true, data: v ? JSON.parse(v) : null };
}

function sheet(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name); sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); sh.setFrozenRows(1);
  }
  return sh;
}
function clip(v, n) { return String(v == null ? '' : v).slice(0, n); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Run once from the editor to create the sheets. */
function setup() { sheet(ANSWERS, ['Section', 'Question', 'Answer']); sheet(TEAM, ['Job title', 'Name', 'What they do', 'Great job after 90 days', 'Custom']); sheet(HISTORY, ['Saved', 'Status', 'Answers JSON']); sheet(RAW, ['Latest JSON']); }
