// Paste this into Extensions > Apps Script of a new Google Sheet.
// Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
// Copy the web app URL into CONFIG.ENDPOINT in public/index.html.
// To update: paste the new code, then Deploy > Manage deployments > Edit >
// Version: New version > Deploy. This keeps the same URL. If COLS gains a column,
// add its name to the header row of an existing "responses" sheet by hand.
// Answers go to "responses" and "raw". Emails for findings go to "findings_requests",
// with no participant ID, so emails cannot be matched to answers through the data.

const COLS = ['received_at','participant_id','is_test','started_at','submitted_at','context_order',
  'role','experience','industry','involvement','agent_experience',
  'vignette_id','context','context_position','position','agent',
  'reversible','impact','novelty','track_record','tier','tier_label',
  'hardest','hardest_why','context_seconds',
  'accountability','accountability_other','model_upgrade','approval_review','evidence','involvement_other'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    if (d.type === 'findings') {
      // Date only, so the request time cannot be matched to a response time.
      const day = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      sheet_(ss, 'findings_requests', ['date','email','is_test'])
        .appendRow([day, String(d.email || '').slice(0, 200), d.is_test === true]);
      return ok_();
    }

    sheet_(ss, 'raw', ['received_at','participant_id','json'])
      .appendRow([now, d.participant_id, JSON.stringify(d)]);

    const b = d.background || {}, c = d.closing || {};
    const rows = (d.cases || []).map(x => [now, d.participant_id, d.is_test === true, d.started_at, d.submitted_at, d.context_order,
      b.role, b.experience, b.industry, b.involvement, b.agent_experience,
      x.vignette_id, x.context, x.context_position, x.position, x.agent,
      x.reversible, x.impact, x.novelty, x.track_record, x.tier, x.tier_label,
      x.hardest, x.hardest_why, x.context_seconds,
      c.accountability || '', c.accountability_other || '', c.model_upgrade || '',
      c.approval_review || '', c.evidence || '',
      String(b.involvement_other || '').slice(0, 100)]);
    if (rows.length) {
      const sh = sheet_(ss, 'responses', COLS);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, COLS.length).setValues(rows);
    }
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
