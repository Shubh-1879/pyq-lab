// PYQ Lab sync backend. Paste into Extensions > Apps Script of a new Google Sheet.
// 1) Change KEY below to any secret word.  2) Deploy > New deployment > Web app,
//    Execute as: Me, Who has access: Anyone.  3) Copy the web app URL into the app's Sync settings.
// Tabs: "attempts" (every answer you give), "questions" (questions you add), "keyfixes" (answer-key corrections).
// You can also type questions straight into the "questions" tab: fill exam..remember, leave id and json blank.
const KEY = 'CHANGE_ME';
const ACOLS = ['id','date','subject','topic','year','qno','result','chosen','conf','eliminated','elimBad','reason','note','sec','source','json'];
const FCOLS = ['id','date','answer','note','json'];
const QCOLS = ['id','exam','source','subject','topic','year','question','a','b','c','d','answer','explanation','remember','json'];

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function rows_(sh) { const n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues(); }

function doGet(e) {
  if ((e.parameter.key || '') !== KEY) return out_({ok:false, error:'bad key'});
  const attempts = [];
  rows_(sheet_('attempts', ACOLS)).forEach(r => { try { attempts.push(JSON.parse(r[ACOLS.indexOf('json')])); } catch (err) {} });
  const qsh = sheet_('questions', QCOLS), questions = [];
  const rows = rows_(qsh);
  rows.forEach((r, i) => {
    const g = c => r[QCOLS.indexOf(c)];
    let id = String(g('id') || '');
    if (!id) { id = 's' + (i + 2) + '-' + Utilities.getUuid().slice(0, 6); qsh.getRange(i + 2, 1).setValue(id); }  // typed directly in the sheet
    const ans = 'abcd'.indexOf(String(g('answer')).toLowerCase().replace(/[()\s]/g, ''));
    if (!g('question') || ans < 0) return;
    questions.push({id, exam: g('exam') || 'Custom', source: g('source') || '', subject: g('subject') || 'General', topic: g('topic') || 'General',
      year: Number(g('year')) || null, question: String(g('question')), options: [g('a'), g('b'), g('c'), g('d')].map(String),
      answer: ans, summary: g('explanation') || '', keyfact: g('remember') || '', mode: 'text'});
  });
  const fixes = [];
  rows_(sheet_('keyfixes', FCOLS)).forEach(r => { try { fixes.push(JSON.parse(r[FCOLS.indexOf('json')])); } catch (err) {} });
  return out_({ok:true, attempts, questions, fixes});
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  if (body.key !== KEY) return out_({ok:false, error:'bad key'});
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ash = sheet_('attempts', ACOLS), have = new Set(rows_(ash).map(r => String(r[0])));
    const arows = [];
    (body.attempts || []).forEach(a => {
      if (!a || !a.id || have.has(String(a.id))) return; have.add(String(a.id));
      arows.push([a.id, new Date(a.ts || Date.now()), a.subject || '', a.topic || '', a.year || '', a.qno || '',
        a.result || '', a.chosen == null || a.chosen < 0 ? '' : 'abcd'[a.chosen], a.conf || '',
        (a.elim || []).map(i => 'abcd'[i]).join(' '), a.elimBad ? 'YES' : '', a.reason || '', a.note || '',
        a.sec || '', a.source || '', JSON.stringify(a)]);
    });
    if (arows.length) ash.getRange(ash.getLastRow() + 1, 1, arows.length, ACOLS.length).setValues(arows);
    const qsh = sheet_('questions', QCOLS), qhave = new Set(rows_(qsh).map(r => String(r[0])));
    const qrows = [];
    (body.questions || []).forEach(q => {
      if (!q || !q.id || qhave.has(String(q.id))) return; qhave.add(String(q.id));
      const o = q.options || [];
      qrows.push([q.id, q.exam || '', q.source || '', q.subject || '', q.topic || '', q.year || '', q.question || '',
        o[0] || '', o[1] || '', o[2] || '', o[3] || '', q.answer == null ? '' : 'abcd'[q.answer], q.summary || '', q.keyfact || '', JSON.stringify(q)]);
    });
    if (qrows.length) qsh.getRange(qsh.getLastRow() + 1, 1, qrows.length, QCOLS.length).setValues(qrows);
    const fsh = sheet_('keyfixes', FCOLS), frows = [];
    (body.fixes || []).forEach(f => {
      if (!f || !f.id) return;
      frows.push([f.id, new Date(f.ts || Date.now()), f.undo ? 'UNDO (official key)' : (f.answer == null ? 'dropped' : 'abcd'[f.answer]), f.note || '', JSON.stringify(f)]);
    });
    if (frows.length) fsh.getRange(fsh.getLastRow() + 1, 1, frows.length, FCOLS.length).setValues(frows);
    return out_({ok:true, added: arows.length, addedQuestions: qrows.length, addedFixes: frows.length});
  } finally { lock.releaseLock(); }
}
