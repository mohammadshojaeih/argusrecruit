/**
 * ArgusRecruit ATS \u2014 Apps Script
 *
 * What it does
 * ------------
 * 1. Watches Gmail for [Application <jobId>] emails and adds each candidate
 *    as a row in the Applications sheet (auto-filling all fields from the
 *    structured payload in the email body).
 * 2. Saves the CV attachment to:
 *       <ROOT_FOLDER_ID>/<jobId> - <jobTitle>/00-New/<date>_<name>_<title>.<ext>
 * 3. Watches the Stage dropdown column. When you change a row's stage:
 *    - moves the CV to the new stage folder
 *    - sends the candidate the email template for that stage (3 languages)
 *    - logs the change in History
 * 4. Duplicates: if a candidate re-applies for a job they're already on,
 *    the new CV replaces the old (active stages only \u2014 skip if Rejected).
 *
 * One-time setup
 * --------------
 * Just call `setup()` from the editor \u2014 it provisions the Sheet (Applications,
 * Email Templates, Settings, History) with dropdowns, formulas, and conditional
 * formatting. Then add a time-trigger on `tick` every 5 minutes.
 */

// ----------------------------------------------------------------------
// CONFIG
// ----------------------------------------------------------------------
// ROOT_FOLDER_ID is the ORIGINAL account's CV folder, kept only as a
// fallback. When you run setup() in a fresh Google account, it auto-creates
// a candidates folder and records its ID in the Settings sheet (see
// getRootFolder_), so you do NOT need to edit this to migrate accounts.

const ROOT_FOLDER_ID = '1Ssla0RAD7XXY81o3R2TV5r1_UQRbUhL4';
const APP_SHEET      = 'Applications';
const TPL_SHEET      = 'Email Templates';
const CFG_SHEET      = 'Settings';
const PROCESSED_LBL  = 'ats-processed';
const SEARCH_QUERY   = 'subject:"[Application" has:attachment -label:ats-processed newer_than:90d';

// Telegram intake bot \u2014 see installTelegramWebhook() for setup
const TG_BOT_TOKEN     = '8706377970:AAEtKkn1Cl68PSmX65wFtFMJAL3XFHaegeo'; // @ArgusIntakeBot
const TG_ADMIN_CHAT_ID = 814437645;  // Mohammad \u2014 only this account is allowed to use the intake bot

const SOURCES = [
  'web-apply',
  'manual-intake',
  'sourced-telegram',
  'sourced-linkedin',
  'referral',
  'sourced-other'
];

const STAGES = [
  '00-New',
  '00-Pre-Contact',
  '01-Queued \u00B7 Reviewed',
  '01-Reviewed',
  '02-Queued \u00B7 Shortlist',
  '02-Shortlist',
  '03-Queued \u00B7 Interview',
  '03-Interview',
  '04-Queued \u00B7 Offer',
  '04-Offer',
  '99-Queued \u00B7 Reject',
  '99-Rejected'
];

// Stages that should send an email when entered (and the template key to use).
const STAGE_EMAIL_TEMPLATE = {
  '01-Reviewed':  'reviewed',
  '02-Shortlist': 'shortlist',
  '03-Interview': 'interview',
  '04-Offer':     'offer',
  '99-Rejected':  'rejected'
};

// Column indexes in the Applications sheet (1-based).
const COL = {
  date:        1,   // A  Date Applied
  jobId:       2,   // B  Job ID
  source:      3,   // C  Source
  jobTitle:    4,   // D  Job Title
  cvLink:      5,   // E  CV
  name:        6,   // F  Name
  phone:       7,   // G  Phone
  linkedin:    8,   // H  LinkedIn
  rating:      9,   // I  \u2190 DROPDOWN, 1\u20135 stars (manual)
  stage:       10,  // J  \u2190 DROPDOWN, the user edits this
  notes:       11,  // K  Notes
  email:       12,  // L  Email
  lang:        13,  // M  Lang
  otherApps:   14,  // N  (auto formula)
  lastChange:  15,  // O  Last Change
  history:     16,  // P  History
  followUp:    17,  // Q  \u2190 DATE, next candidate follow-up (manual)
  lastStage:   18,  // R  (hidden \u2014 script uses to detect changes)
  fileId:      19   // S  (hidden \u2014 CV file ID for moves)
};

// Resume rating dropdown options (1\u20135 stars).
const RATINGS = ['\u2605', '\u2605\u2605', '\u2605\u2605\u2605', '\u2605\u2605\u2605\u2605', '\u2605\u2605\u2605\u2605\u2605'];


// ----------------------------------------------------------------------
// MAIN ENTRY \u2014 run via 5-minute trigger
// ----------------------------------------------------------------------

function tick() {
  try { processInbox_(); } catch (e) { console.error('processInbox error: ' + e); }
  try { processStageChanges_(); } catch (e) { console.error('processStageChanges error: ' + e); }
}


// ----------------------------------------------------------------------
// STEP 1 \u2014 Process new application emails
// ----------------------------------------------------------------------

function processInbox_() {
  const label = getOrCreateLabel_(PROCESSED_LBL);
  const threads = GmailApp.search(SEARCH_QUERY, 0, 50);
  if (threads.length === 0) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(APP_SHEET);
  const root = getRootFolder_();

  for (const thread of threads) {
    try {
      const msg = thread.getMessages()[thread.getMessages().length - 1];
      const payload = parseAtsPayload_(msg);
      if (!payload) { console.warn('Skipped (no ATS payload): ' + msg.getSubject()); continue; }

      const attachments = msg.getAttachments({ includeInlineImages: false });
      const cvs = attachments.filter(a => /\.(pdf|docx?|odt|rtf)$/i.test(a.getName()));
      if (cvs.length === 0) { console.warn('No CV: ' + msg.getSubject()); continue; }
      const cv = cvs[0];

      const { jobId, jobTitle, name, email, phone, linkedin, lang, source, submittedAt } = payload;
      const jobFolder   = findOrCreateFolder_(root, `${jobId} - ${jobTitle}`);
      const stageFolder = findOrCreateFolder_(jobFolder, '00-New');

      // De-dup
      const existing = findRowByEmailAndJob_(sheet, email, jobId);
      if (existing) {
        const stage = sheet.getRange(existing, COL.stage).getValue();
        if (stage === '99-Rejected') {
          console.log('Already rejected, skipping: ' + email + ' / ' + jobId);
          thread.addLabel(label); thread.markRead(); thread.moveToArchive();
          continue;
        }
        // Replace CV
        const fileName = buildFilename_(submittedAt, name, jobTitle, cv.getName());
        // Move old file aside as _v1
        const oldId = sheet.getRange(existing, COL.fileId).getValue();
        if (oldId) {
          try {
            const oldFile = DriveApp.getFileById(oldId);
            oldFile.setName(oldFile.getName().replace(/(\.[^.]+)$/, '_v' + Date.now() + '$1'));
          } catch (_) { /* old file might be gone */ }
        }
        // Save new file into the current stage folder of this row
        const currentStageFolder = findOrCreateFolder_(jobFolder, stage);
        const newFile = currentStageFolder.createFile(cv).setName(fileName);
        sheet.getRange(existing, COL.cvLink).setValue('=HYPERLINK("' + newFile.getUrl() + '","Open")');
        sheet.getRange(existing, COL.fileId).setValue(newFile.getId());
        sheet.getRange(existing, COL.lastChange).setValue(new Date());
        appendHistory_(sheet, existing, `CV re-uploaded \u2192 ${stage}`);
        thread.addLabel(label); thread.markRead(); thread.moveToArchive();
        continue;
      }

      // New row
      const fileName = buildFilename_(submittedAt, name, jobTitle, cv.getName());
      const file = stageFolder.createFile(cv).setName(fileName);

      const row = [
        new Date(submittedAt || msg.getDate()),     // A date
        jobId || '',                                 // B Job ID
        source || 'web-apply',                       // C Source
        jobTitle || '',                              // D Job Title
        '=HYPERLINK("' + file.getUrl() + '","Open")', // E CV
        name || '',                                  // F Name
        phone || '',                                 // G Phone
        linkedin || '',                              // H LinkedIn
        '',                                          // I rating (manual)
        '00-New',                                    // J Stage
        '',                                          // K notes
        (email || '').toLowerCase(),                 // L Email
        lang || 'en',                                // M Lang
        '',                                          // N will become formula below
        new Date(),                                  // O last change
        'created \u2192 00-New',                          // P history
        '',                                          // Q follow-up (manual)
        '00-New',                                    // R lastStage
        file.getId()                                 // S fileId
      ];
      sheet.appendRow(row);
      const r = sheet.getLastRow();
      sheet.getRange(r, COL.otherApps).setFormula(
        `=IFERROR(COUNTIF($L$2:$L, $L${r})-1, 0)`
      );
      thread.addLabel(label); thread.markRead(); thread.moveToArchive();
      console.log('Added: ' + name + ' / ' + jobId);
    } catch (e) {
      console.error('Thread error: ' + e + '\n' + (e.stack || ''));
    }
  }
}


// ----------------------------------------------------------------------
// STEP 2 \u2014 Detect Stage changes and react (move file + send email)
// ----------------------------------------------------------------------

function processStageChanges_() {
  // Serialize runs. Without this, an overlapping trigger (or a second installed
  // trigger) reads the same pre-update snapshot and sends the stage email twice.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.warn('processStageChanges_ skipped: another run holds the lock');
    return;
  }
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(APP_SHEET);
    const root = getRootFolder_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, COL.fileId).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const r = data[i];
      const currentStage = r[COL.stage - 1];
      const lastStage    = r[COL.lastStage - 1];
      if (!currentStage || currentStage === lastStage) continue;

      const jobId    = r[COL.jobId - 1];
      const jobTitle = r[COL.jobTitle - 1];
      const fileId   = r[COL.fileId - 1];
      const email    = r[COL.email - 1];
      const lang     = r[COL.lang - 1] || 'en';
      const name     = r[COL.name - 1];

      // Advance the idempotency guard FIRST and flush it to the sheet, BEFORE
      // any side effect. If the email send (or Drive move) later throws, the
      // guard is already recorded, so the next tick will NOT re-process this
      // transition and re-email the candidate. Over-sending is worse than a
      // rare dropped notification (which is logged to History for manual retry).
      sheet.getRange(row, COL.lastStage).setValue(currentStage);
      sheet.getRange(row, COL.lastChange).setValue(new Date());
      SpreadsheetApp.flush();

      try {
        // Move the CV
        if (fileId) {
          const jobFolder = findOrCreateFolder_(root, `${jobId} - ${jobTitle}`);
          const newStageFolder = findOrCreateFolder_(jobFolder, currentStage);
          moveFileToFolder_(fileId, newStageFolder);
        }
        // Maybe send email
        const tplKey = STAGE_EMAIL_TEMPLATE[currentStage];
        if (tplKey && email) {
          sendStageEmail_(tplKey, lang, {
            name, jobTitle, jobId, email
          });
        }
        appendHistory_(sheet, row, `${lastStage || '?'} \u2192 ${currentStage}`);
      } catch (e) {
        // Guard is already advanced, so we won't loop. Surface the failure in
        // History instead of silently retrying (and spamming) every tick.
        console.error(`Row ${row} stage transition failed: ${e}`);
        appendHistory_(sheet, row, `${lastStage || '?'} \u2192 ${currentStage} (action FAILED - not retried: ${e})`);
      }
    }
  } finally {
    lock.releaseLock();
  }
}


// ----------------------------------------------------------------------
// PARSING + DRIVE HELPERS
// ----------------------------------------------------------------------

function parseAtsPayload_(msg) {
  const html = msg.getBody();
  const m = html.match(/ATS_PAYLOAD_START([\s\S]*?)ATS_PAYLOAD_END/);
  if (!m) return null;
  try { return JSON.parse(m[1].trim()); } catch (_) { return null; }
}

function findOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function moveFileToFolder_(fileId, destFolder) {
  const file = DriveApp.getFileById(fileId);
  // Remove from all current parents, add to destination
  const parents = file.getParents();
  while (parents.hasNext()) parents.next().removeFile(file);
  destFolder.addFile(file);
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function findRowByEmailAndJob_(sheet, email, jobId) {
  if (!email) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const emails = sheet.getRange(2, COL.email, last - 1, 1).getValues();
  const jobs   = sheet.getRange(2, COL.jobId, last - 1, 1).getValues();
  const target = (email || '').toLowerCase().trim();
  for (let i = 0; i < emails.length; i++) {
    if ((emails[i][0] || '').toString().toLowerCase().trim() === target &&
        (jobs[i][0] || '').toString().trim() === (jobId || '').trim()) {
      return i + 2;
    }
  }
  return null;
}

function sanitizeSegment_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036F]/g, '')
    .replace(/[^A-Za-z0-9 ]+/g, '')
    .trim().replace(/\s+/g, '_').slice(0, 60) || 'unknown';
}

function buildFilename_(submittedAt, name, jobTitle, originalName) {
  const date = submittedAt ? new Date(submittedAt) : new Date();
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  const ext = (originalName.match(/\.[A-Za-z0-9]{1,8}$/) || ['.pdf'])[0].toLowerCase();
  return `${y}-${m}-${d}_${sanitizeSegment_(name)}_${sanitizeSegment_(jobTitle)}${ext}`;
}

function appendHistory_(sheet, row, entry) {
  const cell = sheet.getRange(row, COL.history);
  const old = cell.getValue() || '';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  cell.setValue(old + (old ? ' | ' : '') + stamp + ' ' + entry);
}


// ----------------------------------------------------------------------
// EMAIL SENDING
// ----------------------------------------------------------------------

function sendStageEmail_(templateKey, lang, ctx) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TPL_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const langCol = headers.indexOf(lang) > -1 ? headers.indexOf(lang) : headers.indexOf('en');
  if (langCol < 0) { console.warn('No language column in templates'); return; }

  const row = data.find(r => r[0] === templateKey);
  if (!row) { console.warn('Template not found: ' + templateKey); return; }

  const subjectTpl = row[langCol + 0];   // The lang column itself = subject
  const bodyTpl    = row[langCol + 1];   // Next col = body (each lang occupies 2 cols: subj, body)
  if (!subjectTpl || !bodyTpl) { console.warn('Empty template for ' + templateKey + '/' + lang); return; }

  const fill = (s) => s
    .replace(/\{name\}/g, ctx.name || '')
    .replace(/\{jobTitle\}/g, ctx.jobTitle || '')
    .replace(/\{jobId\}/g, ctx.jobId || '');

  const enabled = String(getSetting_('Send Emails Enabled') || 'TRUE').toUpperCase() === 'TRUE';
  if (!enabled) {
    console.log('Email send disabled by Settings; would have sent to ' + ctx.email);
    return;
  }

  const subject = fill(subjectTpl);
  const html = fill(bodyTpl);
  const fromName = getSetting_('From Name') || 'ArgusRecruit';
  const replyTo = getSetting_('Reply To') || 'headhunter@argusrecruit.com';

  // Prefer Resend: it sends from the verified argusrecruit.com domain
  // (SPF/DKIM), so mail is delivered reliably. Gmail's own outbound filter
  // blocks automated HTML mail from the ATS account (bounce error 69585).
  // If no Resend key is configured, fall back to Gmail.
  const resendKey = String(getSetting_('Resend API Key') || '').trim();
  if (resendKey) {
    const fromEmail = String(getSetting_('From Email') || 'headhunter@argusrecruit.com').trim();
    sendViaResend_(resendKey, { fromName, fromEmail, to: ctx.email, subject, html, replyTo });
  } else {
    GmailApp.sendEmail(ctx.email, subject, '', { htmlBody: html, name: fromName, replyTo });
  }
}

// Send one email via the Resend API (same service the website uses).
// Retries briefly on transient failures (429 rate-limit, 5xx) so a momentary
// blip doesn't drop a stage notification. Permanent errors (4xx other than 429)
// throw immediately — retrying them would never succeed.
function sendViaResend_(apiKey, m) {
  const payload = JSON.stringify({
    from: `${m.fromName} <${m.fromEmail}>`,
    to: [m.to],
    subject: m.subject,
    html: m.html,
    reply_to: m.replyTo
  });
  const maxAttempts = 3;
  let lastCode = 0, lastBody = '';
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch('https://api.resend.com/emails', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: payload,
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) return;      // delivered
    lastCode = code; lastBody = res.getContentText();
    const transient = (code === 429 || code >= 500);
    if (!transient || attempt === maxAttempts) break;
    Utilities.sleep(1200 * attempt);            // simple linear backoff
  }
  throw new Error('Resend send failed (' + lastCode + '): ' + lastBody);
}

function getSetting_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG_SHEET);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  for (const row of data) {
    if ((row[0] || '').toString().trim() === key) return row[1];
  }
  return null;
}

function setSetting_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG_SHEET);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

/**
 * Resolve the Drive folder where candidate CVs live, scoped to the account
 * the script runs under.
 *   1. Use the ID recorded in the Settings sheet (written by setup()).
 *   2. Else use the ROOT_FOLDER_ID constant (the original account's folder).
 *   3. Else \u2014 a fresh account where neither is accessible \u2014 create a folder
 *      and remember its ID, so a brand-new Google account works with no
 *      manual folder-ID juggling.
 */
function getRootFolder_() {
  const stored = getSetting_('Root Folder ID');
  if (stored) {
    try { return DriveApp.getFolderById(String(stored).trim()); } catch (_) {}
  }
  if (ROOT_FOLDER_ID) {
    try { return DriveApp.getFolderById(ROOT_FOLDER_ID); } catch (_) {}
  }
  const folder = DriveApp.createFolder('Resumes');
  setSetting_('Root Folder ID', folder.getId());
  return folder;
}


// ----------------------------------------------------------------------
// SHEET PROVISIONING (run once from the editor)
// ----------------------------------------------------------------------

/**
 * setup() is IDEMPOTENT and SAFE TO RE-RUN.
 * It only creates missing sheets / sets headers, dropdowns, formatting.
 * It NEVER deletes existing rows.
 *
 * For destructive operations use:
 *   - resetTemplatesOnly()  \u2192 wipes only Email Templates sheet and re-seeds it
 *   - resetApplications_DANGER() \u2192 wipes Applications data (requires manual edit to enable)
 */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // -- 1. Applications sheet (non-destructive) ------------------------
  let app = ss.getSheetByName(APP_SHEET);
  const isNewApp = !app;
  if (!app) app = ss.insertSheet(APP_SHEET);

  // Write headers for a brand-new sheet, a header-less sheet, OR an existing
  // sheet that has only headers and no data yet (safe to re-lay-out to the
  // current column order). A sheet that already holds candidate rows is left
  // untouched to avoid mislabeling columns — recreate it to adopt a new order.
  const APP_HEADERS = [
    'Date Applied', 'Job ID', 'Source', 'Job Title', 'CV',
    'Name', 'Phone', 'LinkedIn', 'Rating', 'Stage',
    'Notes', 'Email', 'Lang', '# Other Apps', 'Last Change',
    'History', 'Follow-up', 'Last Known Stage (auto)', 'File ID (auto)'
  ];
  const firstCell = app.getRange(1, 1).getValue();
  if (isNewApp || !firstCell || app.getLastRow() <= 1) {
    app.getRange(1, 1, 1, 19).setValues([APP_HEADERS]);
  }
  app.setFrozenRows(1);
  app.getRange('A1:S1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  // Widths follow the new order: A date, B jobId, C source, D jobTitle, E CV,
  // F name, G phone, H linkedin, I rating, J stage, K notes, L email, M lang,
  // N #other, O last change, P history, Q follow-up.
  app.setColumnWidth(1, 100); app.setColumnWidth(2, 80);  app.setColumnWidth(3, 90);
  app.setColumnWidth(4, 180); app.setColumnWidth(5, 60);  app.setColumnWidth(6, 140);
  app.setColumnWidth(7, 110); app.setColumnWidth(8, 200); app.setColumnWidth(9, 90);
  app.setColumnWidth(10, 180); app.setColumnWidth(11, 200); app.setColumnWidth(12, 200);
  app.setColumnWidth(13, 60); app.setColumnWidth(14, 110); app.setColumnWidth(15, 110);
  app.setColumnWidth(16, 280); app.setColumnWidth(17, 110);
  app.hideColumns(COL.lastStage, 2);

  // Stage validation
  const stageRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STAGES, true)
    .setAllowInvalid(false)
    .build();
  app.getRange(2, COL.stage, 1000, 1).setDataValidation(stageRule);

  // Source validation
  const sourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SOURCES, true)
    .setAllowInvalid(true)
    .build();
  app.getRange(2, COL.source, 1000, 1).setDataValidation(sourceRule);

  // Rating validation (1\u20135 stars dropdown)
  const ratingRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(RATINGS, true)
    .setAllowInvalid(false)
    .build();
  app.getRange(2, COL.rating, 1000, 1)
    .setDataValidation(ratingRule)
    .setHorizontalAlignment('center');

  // Follow-up date column \u2014 calendar picker + date display format
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  app.getRange(2, COL.followUp, 1000, 1)
    .setDataValidation(dateRule)
    .setNumberFormat('yyyy-mm-dd');

  // Highlight rows where # Other Apps > 0  (column N in the new order)
  const condRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$N2>0')
    .setBackground('#FFF7E0')
    .setRanges([app.getRange(2, 1, 1000, 19)])
    .build();
  // Highlight follow-ups that are due (date today or earlier, and not empty)
  const dueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($Q2<>"", $Q2<=TODAY())')
    .setBackground('#FCE4E4')
    .setRanges([app.getRange(2, COL.followUp, 1000, 1)])
    .build();
  app.setConditionalFormatRules([condRule, dueRule]);

  // -- 2. Email Templates sheet (non-destructive) ---------------------
  let tpl = ss.getSheetByName(TPL_SHEET);
  const isNewTpl = !tpl;
  if (!tpl) tpl = ss.insertSheet(TPL_SHEET);
  if (isNewTpl || !tpl.getRange(1, 1).getValue()) {
    tpl.getRange(1, 1, 1, 7).setValues([[
      'Key', 'en', 'en (body)', 'ru', 'ru (body)', 'hy', 'hy (body)'
    ]]);
    const rowsTpl = stageTemplates_();
    tpl.getRange(2, 1, rowsTpl.length, 7).setValues(rowsTpl);
    tpl.setRowHeights(2, rowsTpl.length, 90);
  }
  tpl.setFrozenRows(1);
  tpl.getRange('A1:G1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  tpl.setColumnWidth(1, 110);
  for (let c = 2; c <= 7; c++) tpl.setColumnWidth(c, 300);

  // -- 3. Settings sheet (non-destructive) ----------------------------
  let cfg = ss.getSheetByName(CFG_SHEET);
  const isNewCfg = !cfg;
  if (!cfg) cfg = ss.insertSheet(CFG_SHEET);
  if (isNewCfg || !cfg.getRange(1, 1).getValue()) {
    cfg.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    cfg.getRange(2, 1, 7, 2).setValues([
      ['Send Emails Enabled', 'TRUE'],
      ['From Name',           'ArgusRecruit'],
      ['From Email',          'headhunter@argusrecruit.com'],
      ['Reply To',            'headhunter@argusrecruit.com'],
      ['Resend API Key',      ''],   // paste your Resend key for reliable delivery
      ['Root Folder ID',      ''],   // auto-filled below by getRootFolder_()
      ['Trigger Interval',    'every 5 minutes (set via Triggers panel)']
    ]);
  }
  // Add newer settings to Settings sheets created before they existed.
  if (getSetting_('From Email') === null) setSetting_('From Email', 'headhunter@argusrecruit.com');
  if (getSetting_('Resend API Key') === null) setSetting_('Resend API Key', '');
  cfg.getRange('A1:B1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  cfg.setColumnWidth(1, 200);
  cfg.setColumnWidth(2, 400);

  // -- 4. Ensure THIS account has an accessible CV root folder --------
  //    On a fresh Google account the constant ROOT_FOLDER_ID is not
  //    accessible, so this creates the folder and records its real ID in
  //    Settings. On the original account it just re-confirms the existing one.
  try {
    setSetting_('Root Folder ID', getRootFolder_().getId());
  } catch (e) {
    console.error('Root folder provisioning failed: ' + e);
  }

  ss.toast('ATS structure verified. Existing data was preserved.', 'Setup complete', 6);
}


/** Wipes and re-seeds ONLY the Email Templates sheet. Safe \u2014 Applications data is untouched. */
function resetTemplatesOnly() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Reset Email Templates?',
    'This will replace the Email Templates sheet with the latest branded defaults.\nApplications data is NOT touched.\n\nContinue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  let tpl = ss.getSheetByName(TPL_SHEET);
  if (!tpl) tpl = ss.insertSheet(TPL_SHEET);
  tpl.clear();
  tpl.getRange(1, 1, 1, 7).setValues([[
    'Key', 'en', 'en (body)', 'ru', 'ru (body)', 'hy', 'hy (body)'
  ]]);
  tpl.setFrozenRows(1);
  tpl.getRange('A1:G1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  const rowsTpl = stageTemplates_();
  tpl.getRange(2, 1, rowsTpl.length, 7).setValues(rowsTpl);
  tpl.setColumnWidth(1, 110);
  for (let c = 2; c <= 7; c++) tpl.setColumnWidth(c, 300);
  tpl.setRowHeights(2, rowsTpl.length, 90);
  ss.toast('Email Templates reset.', 'Done', 4);
}


/**
 * One-time migration: reorder an EXISTING Applications sheet (with data) into
 * the current column layout, carrying every column's data with it. Matches
 * columns by header label, so it works whether the sheet is in the old 17- or
 * 19-column order. A timestamped backup tab is created first. Safe to re-run —
 * it no-ops once the columns are already in the new order.
 */
function reorderApplicationsColumns() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(APP_SHEET);
  if (!sh) { ui.alert('No "' + APP_SHEET + '" sheet found.'); return; }

  const NEW_HEADERS = [
    'Date Applied', 'Job ID', 'Source', 'Job Title', 'CV',
    'Name', 'Phone', 'LinkedIn', 'Rating', 'Stage',
    'Notes', 'Email', 'Lang', '# Other Apps', 'Last Change',
    'History', 'Follow-up', 'Last Known Stage (auto)', 'File ID (auto)'
  ];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const curHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  if (curHeaders.length === NEW_HEADERS.length && NEW_HEADERS.every((h, i) => curHeaders[i] === h)) {
    ui.alert('Columns are already in the new order. Nothing to do.');
    return;
  }

  const resp = ui.alert('Reorder columns?',
    'This rearranges the Applications columns into the new order, keeping each column\'s data with it. A backup tab is created first.\n\nContinue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // Safety backup.
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sh.copyTo(ss).setName('Applications (backup ' + stamp + ')');

  // Read values + formulas so CV hyperlinks (and any other formulas) survive.
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const formulas = sh.getRange(1, 1, lastRow, lastCol).getFormulas();
  const cell = (r, c) => (formulas[r][c] !== '' ? formulas[r][c] : values[r][c]);

  // For each target column, find its source by header label (-1 = not present).
  const srcIndex = NEW_HEADERS.map(h => curHeaders.indexOf(h));

  // Build the reordered matrix. # Other Apps is rebuilt later, so leave it blank.
  const otherAppsTarget = NEW_HEADERS.indexOf('# Other Apps');
  const out = [];
  for (let r = 0; r < lastRow; r++) {
    const row = [];
    for (let c = 0; c < NEW_HEADERS.length; c++) {
      const s = srcIndex[c];
      row.push((c === otherAppsTarget || s === -1) ? '' : cell(r, s));
    }
    out.push(row);
  }
  out[0] = NEW_HEADERS.slice();

  // Wipe old content/formats/validations, then write the reordered data.
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.clear();
  sh.getRange(1, 1, out.length, NEW_HEADERS.length).setValues(out);

  // Re-apply headers styling, widths, dropdowns, date format, conditional formatting.
  setup();

  // Rebuild the # Other Apps formula per data row (now keyed off Email at column L).
  for (let r = 2; r <= lastRow; r++) {
    sh.getRange(r, COL.otherApps).setFormula(`=IFERROR(COUNTIF($L$2:$L, $L${r})-1, 0)`);
  }

  ss.toast('Columns reordered (' + (lastRow - 1) + ' rows). A backup tab was created.', 'Done', 6);
}


// ----------------------------------------------------------------------
// CUSTOM MENU \u2014 for manual ops
// ----------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ATS')
    .addItem('Sync now', 'tick')
    .addItem('Verify setup (safe to re-run)', 'setup')
    .addSeparator()
    .addItem('Preview selected email template', 'previewSelectedTemplate')
    .addItem('Show intake form URL', 'showIntakeUrl')
    .addSeparator()
    .addItem('Install Telegram intake bot webhook', 'installTelegramWebhook')
    .addItem('Uninstall Telegram intake bot webhook', 'uninstallTelegramWebhook')
    .addSeparator()
    .addItem('Reorder columns to new layout (keeps data)', 'reorderApplicationsColumns')
    .addItem('Reset email templates only', 'resetTemplatesOnly')
    .addToUi();
}

/**
 * Preview the email template under the currently-selected cell.
 * Usage: open Email Templates sheet \u2192 click any body cell (en, ru, hy columns)
 *        \u2192 ATS menu \u2192 "Preview selected email template".
 * Renders the cell's HTML with sample placeholders filled in.
 */
function previewSelectedTemplate() {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getActiveRange();
  const sheet = range.getSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== TPL_SHEET) {
    ui.alert('Open the Email Templates sheet first, then click on a body cell.');
    return;
  }
  const col = range.getColumn();
  const row = range.getRow();
  if (row < 2) { ui.alert('Pick a body cell (rows 2+).'); return; }

  // Body columns are: C(3)=en body, E(5)=ru body, G(7)=hy body.
  let html = '';
  let langLabel = '';
  if (col === 3) { html = sheet.getRange(row, 3).getValue(); langLabel = 'EN'; }
  else if (col === 5) { html = sheet.getRange(row, 5).getValue(); langLabel = 'RU'; }
  else if (col === 7) { html = sheet.getRange(row, 7).getValue(); langLabel = 'HY'; }
  else {
    ui.alert('Select a cell in the "en (body)", "ru (body)", or "hy (body)" column.');
    return;
  }
  if (!html) { ui.alert('That cell is empty.'); return; }

  const key = sheet.getRange(row, 1).getValue();
  const filled = String(html)
    .replace(/\{name\}/g, 'Sara Mohammadi')
    .replace(/\{jobTitle\}/g, 'Senior QA Engineer')
    .replace(/\{jobId\}/g, 'AR-STE01');

  const out = HtmlService.createHtmlOutput(filled)
    .setWidth(720)
    .setHeight(720)
    .setTitle('Preview: ' + key + ' (' + langLabel + ')');
  ui.showModalDialog(out, 'Preview: ' + key + ' (' + langLabel + ')');
}

function showIntakeUrl() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Deploy first',
      'You need to deploy the script as a Web app.\n\n' +
      '1. In the Apps Script editor: Deploy \u2192 New deployment\n' +
      '2. Type: Web app\n' +
      '3. Execute as: Me (you)\n' +
      '4. Who has access: Anyone with the link  (or Anyone within your domain)\n' +
      '5. Click Deploy and copy the Web app URL.\n\n' +
      'Then run this menu item again to see the URL.',
      ui.ButtonSet.OK);
    return;
  }
  ui.alert('Intake form URL', url + '\n\nBookmark this. Open it on phone or desktop to add a candidate from any source.', ui.ButtonSet.OK);
}


// ----------------------------------------------------------------------
// INTAKE FORM (Apps Script Web App)
// ----------------------------------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutput(intakeFormHtml_())
    .setTitle('ArgusRecruit \u00B7 Add Candidate')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function listActiveJobs_() {
  const root = getRootFolder_();
  const folders = root.getFolders();
  const jobs = [];
  while (folders.hasNext()) {
    const f = folders.next();
    const m = f.getName().match(/^([A-Z]+-[A-Z0-9]+)\s+-\s+(.+)$/);
    if (m) jobs.push({ jobId: m[1], jobTitle: m[2], folderName: f.getName() });
  }
  jobs.sort((a, b) => a.jobId.localeCompare(b.jobId));
  return jobs;
}

function intakeFormHtml_() {
  const jobs = listActiveJobs_();
  const jobOptions = jobs.map(j =>
    `<option value="${j.jobId}|${j.jobTitle}">${j.jobId} \u2014 ${j.jobTitle}</option>`
  ).join('');
  const sourceOptions = SOURCES
    .filter(s => s !== 'web-apply') // can't be web-apply from here
    .map(s => `<option value="${s}">${s}</option>`).join('');

  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Add candidate \u00B7 ArgusRecruit ATS</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/intl-tel-input@24/build/css/intlTelInput.css">
<style>.iti { width: 100%; }</style>
<style>
  :root { --navy:#0E2440; --gold:#D4AF37; --soft:#F5EFE3; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
         margin:0; padding:24px; background: var(--soft); color:#222; }
  .card { max-width: 560px; margin: 0 auto; background:#fff; border-radius:14px;
          box-shadow: 0 6px 24px rgba(0,0,0,0.07); padding:28px; }
  h1 { color: var(--navy); margin: 0 0 6px; font-size: 22px; }
  .sub { color:#666; font-size: 13px; margin-bottom: 22px; }
  label { display:block; font-size: 12px; color:#444; margin-top: 14px; font-weight:600;
          letter-spacing: 0.5px; }
  input, select, textarea { width:100%; padding: 10px 12px; border:1px solid #d2c9b6;
          border-radius:8px; font-size:14px; font-family:inherit; background:#fff; margin-top:5px; }
  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--gold); }
  .req::after { content:" *"; color:#b00; }
  button { background: var(--gold); color: var(--navy); border:0; padding: 14px 22px;
           border-radius:10px; font-weight:700; font-size:14px; letter-spacing:.5px;
           cursor:pointer; margin-top:22px; width:100%; }
  button:hover { background:#b88f1f; color:#fff; }
  .msg { margin-top:14px; padding:12px; border-radius:8px; font-size:14px; display:none; }
  .ok  { display:block; background:#e7f5e8; color:#216c2c; border:1px solid #b5e2bb; }
  .err { display:block; background:#fdecea; color:#7a1f15; border:1px solid #f3b9b1; }
</style>
</head><body>
<div class="card">
  <h1>Add a candidate</h1>
  <div class="sub">For passive sourcing or anyone who didn't come through the website.</div>
  <form id="f">
    <label class="req">Job</label>
    <select name="job" required>
      <option value="">\u2014 select a job \u2014</option>
      ${jobOptions}
    </select>
    <label class="req">Source</label>
    <select name="source" required>${sourceOptions}</select>
    <label class="req">Candidate name</label>
    <input name="name" required maxlength="120">
    <label>Email</label>
    <input name="email" type="email" maxlength="200">
    <label>Phone</label>
    <input name="phone" type="tel" maxlength="60">
    <label>LinkedIn URL</label>
    <input name="linkedin" type="url" maxlength="300">
    <label>CV file (optional \u2014 PDF, DOC, DOCX, max 10MB)</label>
    <input name="cv" type="file" accept=".pdf,.doc,.docx,.odt,.rtf">
    <label>Lang for outgoing emails</label>
    <select name="lang"><option value="en">English</option><option value="ru">Russian</option><option value="hy">Armenian</option></select>
    <label>Notes (private)</label>
    <textarea name="notes" rows="3" maxlength="2000" placeholder="e.g. ex-CTO at FinTech in Yerevan; open to next move; speaks RU+FA"></textarea>
    <button type="submit">Add candidate</button>
    <div id="msg" class="msg"></div>
  </form>
</div>
<script src="https://cdn.jsdelivr.net/npm/intl-tel-input@24/build/js/intlTelInput.min.js"></script>
<script>
const form = document.getElementById('f');
const msg = document.getElementById('msg');
const phoneInput = form.querySelector('input[name="phone"]');
let iti = null;
if (window.intlTelInput) {
  iti = window.intlTelInput(phoneInput, {
    initialCountry: 'am',
    preferredCountries: ['am','ir','ru','gb','ae','ca','us'],
    separateDialCode: true,
    utilsScript: 'https://cdn.jsdelivr.net/npm/intl-tel-input@24/build/js/utils.js'
  });
}
form.addEventListener('submit', async (e) => {
  e.preventDefault();
  msg.className = 'msg'; msg.textContent = '';
  const submitBtn = form.querySelector('button');

  // Validate + normalize phone if filled in
  let phoneE164 = '';
  if (iti && phoneInput.value.trim()) {
    if (!iti.isValidNumber()) {
      msg.className = 'msg err';
      msg.textContent = 'Please enter a valid phone number (or leave it empty).';
      return;
    }
    phoneE164 = iti.getNumber();
  }
  submitBtn.disabled = true; submitBtn.textContent = 'Saving\u2026';

  const f = new FormData(form);
  const cv = f.get('cv');
  const payload = {
    job: f.get('job'),
    source: f.get('source'),
    name: (f.get('name')||'').trim(),
    email: (f.get('email')||'').trim().toLowerCase(),
    phone: phoneE164,
    linkedin: (f.get('linkedin')||'').trim(),
    lang: f.get('lang') || 'en',
    notes: (f.get('notes')||'').trim(),
    cvName: cv && cv.size ? cv.name : '',
    cvType: cv && cv.size ? cv.type : '',
    cvData: ''
  };
  if (cv && cv.size > 10 * 1024 * 1024) {
    msg.className = 'msg err'; msg.textContent = 'CV file too large (max 10MB).';
    submitBtn.disabled = false; submitBtn.textContent = 'Add candidate';
    return;
  }
  if (cv && cv.size > 0) {
    payload.cvData = await new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result.split(',')[1]);
      r.onerror = rej;
      r.readAsDataURL(cv);
    });
  }
  google.script.run
    .withSuccessHandler(r => {
      if (r.ok) {
        msg.className = 'msg ok';
        msg.textContent = '\u2713 Added: ' + r.name + ' \u2192 ' + r.jobId;
        form.reset();
      } else {
        msg.className = 'msg err';
        msg.textContent = '\u26A0 ' + (r.error || 'Could not save.');
      }
      submitBtn.disabled = false; submitBtn.textContent = 'Add candidate';
    })
    .withFailureHandler(err => {
      msg.className = 'msg err';
      msg.textContent = '\u26A0 ' + err.message;
      submitBtn.disabled = false; submitBtn.textContent = 'Add candidate';
    })
    .submitIntake_(payload);
});
</script>
</body></html>`;
}

function submitIntake_(payload) {
  try {
    if (!payload || !payload.job || !payload.name) {
      return { ok: false, error: 'Missing required fields' };
    }
    const [jobId, jobTitle] = payload.job.split('|');
    const sheet = SpreadsheetApp.getActive().getSheetByName(APP_SHEET);
    if (!sheet) return { ok: false, error: 'Applications sheet not found. Run setup() first.' };

    const root = getRootFolder_();
    const jobFolder = findOrCreateFolder_(root, `${jobId} - ${jobTitle}`);

    // Determine initial stage:
    //   if no CV at all \u2192 00-Pre-Contact
    //   else            \u2192 00-New
    const hasCv = payload.cvData && payload.cvName;
    const initialStage = hasCv ? '00-New' : '00-Pre-Contact';
    const stageFolder = findOrCreateFolder_(jobFolder, initialStage);

    let fileId = '';
    let cvHyperlink = '';
    if (hasCv) {
      const bytes = Utilities.base64Decode(payload.cvData);
      const blob = Utilities.newBlob(bytes, payload.cvType || 'application/pdf', payload.cvName);
      const fileName = buildFilename_(new Date().toISOString(), payload.name, jobTitle, payload.cvName);
      const file = stageFolder.createFile(blob).setName(fileName);
      fileId = file.getId();
      cvHyperlink = '=HYPERLINK("' + file.getUrl() + '","Open")';
    }

    const row = [
      new Date(),                                          // A Date Applied
      jobId,                                               // B Job ID
      payload.source || 'manual-intake',                   // C Source
      jobTitle,                                            // D Job Title
      cvHyperlink,                                         // E CV
      payload.name,                                        // F Name
      payload.phone || '',                                 // G Phone
      payload.linkedin || '',                              // H LinkedIn
      '',                                                  // I rating (manual)
      initialStage,                                        // J Stage
      payload.notes || '',                                 // K Notes
      (payload.email || '').toLowerCase(),                 // L Email
      payload.lang || 'en',                                // M Lang
      '',                                                  // N otherApps (formula below)
      new Date(),                                          // O Last Change
      'created (' + (payload.source || 'manual') + ') \u2192 ' + initialStage, // P History
      '',                                                  // Q follow-up (manual)
      initialStage,                                        // R lastStage
      fileId                                               // S fileId
    ];
    sheet.appendRow(row);
    const r = sheet.getLastRow();
    sheet.getRange(r, COL.otherApps).setFormula(
      `=IFERROR(COUNTIF($L$2:$L, $L${r})-1, 0)`
    );
    return { ok: true, name: payload.name, jobId: jobId };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// ----------------------------------------------------------------------
// BRANDED EMAIL TEMPLATES
// ----------------------------------------------------------------------
//   Each cell stores a complete HTML document with ArgusRecruit branding
//   (navy + gold). Placeholders: {name}, {jobTitle}, {jobId}.
//   To edit copy: open Email Templates sheet and change the text \u2014 the
//   HTML structure is in the cell, edit between the tags.

function stageTemplates_() {
  const stages = [
    {
      key: 'reviewed',
      copy: {
        en: {
          subject: 'Your application is being reviewed \u2014 {jobTitle}',
          eyebrow: '\u2022 Application Under Review \u2022',
          h1: 'Thanks \u2014 we\'re looking at your profile.',
          greeting: 'Hi {name},',
          body: [
            'Thank you for applying to <strong style="color:#D4AF37;">{jobTitle}</strong> at ArgusRecruit. Our team is now carefully reviewing your application, and we\'ll get back to you soon.',
            'If your background matches the role, we will be in touch within 1\u20133 business days. If we don\'t see a fit for this specific role, your profile will stay in our network and we may reach out about future openings that better match your background.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: '\u0412\u0430\u0448\u0430 \u0437\u0430\u044F\u0432\u043A\u0430 \u043D\u0430 \u0440\u0430\u0441\u0441\u043C\u043E\u0442\u0440\u0435\u043D\u0438\u0438 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u0417\u0430\u044F\u0432\u043A\u0430 \u0440\u0430\u0441\u0441\u043C\u0430\u0442\u0440\u0438\u0432\u0430\u0435\u0442\u0441\u044F \u2022',
          h1: '\u0421\u043F\u0430\u0441\u0438\u0431\u043E \u2014 \u043C\u044B \u0440\u0430\u0441\u0441\u043C\u0430\u0442\u0440\u0438\u0432\u0430\u0435\u043C \u0432\u0430\u0448 \u043F\u0440\u043E\u0444\u0438\u043B\u044C.',
          greeting: '\u0417\u0434\u0440\u0430\u0432\u0441\u0442\u0432\u0443\u0439\u0442\u0435, {name},',
          body: [
            '\u0421\u043F\u0430\u0441\u0438\u0431\u043E \u0437\u0430 \u0432\u0430\u0448\u0443 \u0437\u0430\u044F\u0432\u043A\u0443 \u043D\u0430 \u0440\u043E\u043B\u044C <strong style="color:#D4AF37;">{jobTitle}</strong>. \u041D\u0430\u0448\u0430 \u043A\u043E\u043C\u0430\u043D\u0434\u0430 \u0442\u0449\u0430\u0442\u0435\u043B\u044C\u043D\u043E \u0440\u0430\u0441\u0441\u043C\u0430\u0442\u0440\u0438\u0432\u0430\u0435\u0442 \u0432\u0430\u0448 \u043F\u0440\u043E\u0444\u0438\u043B\u044C \u0438 \u0441\u0432\u044F\u0436\u0435\u0442\u0441\u044F \u0441 \u0432\u0430\u043C\u0438 \u0432 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0435\u0435 \u0432\u0440\u0435\u043C\u044F.',
            '\u0415\u0441\u043B\u0438 \u0432\u0430\u0448 \u043E\u043F\u044B\u0442 \u043F\u043E\u0434\u0445\u043E\u0434\u0438\u0442 \u0440\u043E\u043B\u0438, \u043C\u044B \u0441\u0432\u044F\u0436\u0435\u043C\u0441\u044F \u0441 \u0432\u0430\u043C\u0438 \u0432 \u0442\u0435\u0447\u0435\u043D\u0438\u0435 1\u20133 \u0440\u0430\u0431\u043E\u0447\u0438\u0445 \u0434\u043D\u0435\u0439. \u0415\u0441\u043B\u0438 \u0441\u043E\u0432\u043F\u0430\u0434\u0435\u043D\u0438\u044F \u0434\u043B\u044F \u044D\u0442\u043E\u0439 \u043A\u043E\u043D\u043A\u0440\u0435\u0442\u043D\u043E\u0439 \u0440\u043E\u043B\u0438 \u043D\u0435 \u0431\u0443\u0434\u0435\u0442, \u043C\u044B \u0441\u043E\u0445\u0440\u0430\u043D\u0438\u043C \u0432\u0430\u0448 \u043F\u0440\u043E\u0444\u0438\u043B\u044C \u0432 \u043D\u0430\u0448\u0435\u0439 \u0441\u0435\u0442\u0438 \u0438 \u043C\u043E\u0436\u0435\u043C \u0441\u0432\u044F\u0437\u0430\u0442\u044C\u0441\u044F \u0441 \u0432\u0430\u043C\u0438 \u043F\u043E \u043F\u043E\u0432\u043E\u0434\u0443 \u0431\u0443\u0434\u0443\u0449\u0438\u0445 \u0432\u0430\u043A\u0430\u043D\u0441\u0438\u0439.'
          ],
          team: '\u041A\u043E\u043C\u0430\u043D\u0434\u0430 ArgusRecruit'
        },
        hy: {
          subject: '\u0541\u0565\u0580 \u0564\u056B\u0574\u0578\u0582\u0574\u0568 \u0584\u0576\u0576\u0561\u0580\u056F\u057E\u0578\u0582\u0574 \u0567 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u0534\u056B\u0574\u0578\u0582\u0574\u0568 \u0584\u0576\u0576\u0561\u0580\u056F\u0574\u0561\u0576 \u0583\u0578\u0582\u056C\u0578\u0582\u0574 \u2022',
          h1: '\u0547\u0576\u0578\u0580\u0570\u0561\u056F\u0561\u056C\u0578\u0582\u0569\u0575\u0578\u0582\u0576 \u2014 \u0574\u0565\u0576\u0584 \u0578\u0582\u057D\u0578\u0582\u0574\u0576\u0561\u057D\u056B\u0580\u0578\u0582\u0574 \u0565\u0576\u0584 \u0571\u0565\u0580 \u057A\u0580\u0578\u0586\u056B\u056C\u0568:',
          greeting: '\u0532\u0561\u0580\u0587, {name},',
          body: [
            '<strong style="color:#D4AF37;">{jobTitle}</strong> \u057A\u0561\u0577\u057F\u0578\u0576\u056B \u0570\u0561\u0574\u0561\u0580 \u0564\u056B\u0574\u0565\u056C\u0578\u0582 \u0570\u0561\u0574\u0561\u0580 \u0577\u0576\u0578\u0580\u0570\u0561\u056F\u0561\u056C\u0578\u0582\u0569\u0575\u0578\u0582\u0576: \u0544\u0565\u0580 \u0569\u056B\u0574\u0568 \u0574\u0561\u0576\u0580\u0561\u056F\u0580\u056F\u056B\u057F \u0578\u0582\u057D\u0578\u0582\u0574\u0576\u0561\u057D\u056B\u0580\u0578\u0582\u0574 \u0567 \u0571\u0565\u0580 \u0564\u056B\u0574\u0578\u0582\u0574\u0568 \u0587 \u0577\u0578\u0582\u057F\u0578\u057E \u056F\u056F\u0561\u057A\u057E\u0565\u0576\u0584 \u0571\u0565\u0566 \u0570\u0565\u057F:',
            '\u0540\u0561\u0574\u0561\u057A\u0561\u057F\u0561\u057D\u056D\u0561\u0576 \u056C\u056B\u0576\u0565\u056C\u0578\u0582 \u0564\u0565\u057A\u0584\u0578\u0582\u0574 \u0574\u0565\u0576\u0584 \u056F\u056F\u0561\u057A\u057E\u0565\u0576\u0584 \u0571\u0565\u0566 \u0570\u0565\u057F 1\u20133 \u0561\u0577\u056D\u0561\u057F\u0561\u0576\u0584\u0561\u0575\u056B\u0576 \u0585\u0580\u057E\u0561 \u0568\u0576\u0569\u0561\u0581\u0584\u0578\u0582\u0574: \u0540\u0561\u056F\u0561\u057C\u0561\u056F \u0564\u0565\u057A\u0584\u0578\u0582\u0574 \u0571\u0565\u0580 \u057A\u0580\u0578\u0586\u056B\u056C\u0568 \u056F\u057A\u0561\u0570\u057A\u0561\u0576\u057E\u056B \u0574\u0565\u0580 \u0581\u0561\u0576\u0581\u0578\u0582\u0574 \u0561\u057A\u0561\u0563\u0561 \u0570\u0576\u0561\u0580\u0561\u057E\u0578\u0580\u0578\u0582\u0569\u0575\u0578\u0582\u0576\u0576\u0565\u0580\u056B \u0570\u0561\u0574\u0561\u0580:'
          ],
          team: 'ArgusRecruit-\u056B \u0569\u056B\u0574\u0568'
        }
      }
    },
    {
      key: 'shortlist',
      copy: {
        en: {
          subject: 'You\'ve been shortlisted \u2014 {jobTitle}',
          eyebrow: '\u2022 Shortlisted \u2022',
          h1: 'You\'re on the shortlist.',
          greeting: 'Hi {name},',
          body: [
            'Great news \u2014 you have been shortlisted for the <strong style="color:#D4AF37;">{jobTitle}</strong> role. Our client will now review your profile directly.',
            'We\'ll get back to you with the next steps within a few business days. In the meantime, please keep your calendar flexible for the coming week in case interview slots are offered.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: '\u0412\u044B \u0432 \u0448\u043E\u0440\u0442-\u043B\u0438\u0441\u0442\u0435 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u0412 \u0448\u043E\u0440\u0442-\u043B\u0438\u0441\u0442\u0435 \u2022',
          h1: '\u0412\u044B \u043F\u043E\u043F\u0430\u043B\u0438 \u0432 \u0448\u043E\u0440\u0442-\u043B\u0438\u0441\u0442.',
          greeting: '\u0417\u0434\u0440\u0430\u0432\u0441\u0442\u0432\u0443\u0439\u0442\u0435, {name},',
          body: [
            '\u0425\u043E\u0440\u043E\u0448\u0438\u0435 \u043D\u043E\u0432\u043E\u0441\u0442\u0438 \u2014 \u0432\u044B \u0432\u043E\u0448\u043B\u0438 \u0432 \u0448\u043E\u0440\u0442-\u043B\u0438\u0441\u0442 \u043F\u043E \u043F\u043E\u0437\u0438\u0446\u0438\u0438 <strong style="color:#D4AF37;">{jobTitle}</strong>. \u041A\u043B\u0438\u0435\u043D\u0442 \u0441\u0435\u0439\u0447\u0430\u0441 \u0440\u0430\u0441\u0441\u043C\u0430\u0442\u0440\u0438\u0432\u0430\u0435\u0442 \u0432\u0430\u0448 \u043F\u0440\u043E\u0444\u0438\u043B\u044C.',
            '\u041C\u044B \u0441\u0432\u044F\u0436\u0435\u043C\u0441\u044F \u0441 \u0432\u0430\u043C\u0438 \u043F\u043E \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0438\u043C \u0448\u0430\u0433\u0430\u043C \u0432 \u0442\u0435\u0447\u0435\u043D\u0438\u0435 \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u0438\u0445 \u0440\u0430\u0431\u043E\u0447\u0438\u0445 \u0434\u043D\u0435\u0439. \u041F\u043E\u0441\u0442\u0430\u0440\u0430\u0439\u0442\u0435\u0441\u044C \u043E\u0441\u0442\u0430\u0432\u0438\u0442\u044C \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C \u0433\u0438\u0431\u043A\u0438\u043C \u043D\u0430 \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0443\u044E \u043D\u0435\u0434\u0435\u043B\u044E \u2014 \u0432\u043E\u0437\u043C\u043E\u0436\u043D\u044B \u0441\u043B\u043E\u0442\u044B \u0434\u043B\u044F \u0441\u043E\u0431\u0435\u0441\u0435\u0434\u043E\u0432\u0430\u043D\u0438\u044F.'
          ],
          team: '\u041A\u043E\u043C\u0430\u043D\u0434\u0430 ArgusRecruit'
        },
        hy: {
          subject: '\u0534\u0578\u0582\u0584 \u0568\u0576\u057F\u0580\u057E\u0565\u056C \u0565\u0584 \u056F\u0561\u0580\u0573 \u0581\u0578\u0582\u0581\u0561\u056F\u0578\u0582\u0574 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u053F\u0561\u0580\u0573 \u0581\u0578\u0582\u0581\u0561\u056F\u0578\u0582\u0574 \u2022',
          h1: '\u0534\u0578\u0582\u0584 \u056F\u0561\u0580\u0573 \u0581\u0578\u0582\u0581\u0561\u056F\u0578\u0582\u0574 \u0565\u0584:',
          greeting: '\u0532\u0561\u0580\u0587, {name},',
          body: [
            '\u0540\u0561\u0573\u0565\u056C\u056B \u056C\u0578\u0582\u0580 \u2014 \u0564\u0578\u0582\u0584 \u0568\u0576\u057F\u0580\u057E\u0565\u056C \u0565\u0584 <strong style="color:#D4AF37;">{jobTitle}</strong> \u057A\u0561\u0577\u057F\u0578\u0576\u056B \u056F\u0561\u0580\u0573 \u0581\u0578\u0582\u0581\u0561\u056F\u0578\u0582\u0574: \u0544\u0565\u0580 \u0570\u0561\u0573\u0561\u056D\u0578\u0580\u0564\u0576 \u0561\u0575\u056A\u0574 \u0578\u0582\u0572\u0572\u0561\u056F\u056B\u0578\u0580\u0565\u0576 \u0578\u0582\u057D\u0578\u0582\u0574\u0576\u0561\u057D\u056B\u0580\u0578\u0582\u0574 \u0567 \u0571\u0565\u0580 \u057A\u0580\u0578\u0586\u056B\u056C\u0568:',
            '\u0544\u0565\u0576\u0584 \u056F\u057F\u0565\u0572\u0565\u056F\u0561\u0581\u0576\u0565\u0576\u0584 \u0571\u0565\u0566 \u0570\u0561\u057B\u0578\u0580\u0564 \u0584\u0561\u0575\u056C\u0565\u0580\u056B \u0574\u0561\u057D\u056B\u0576 \u0574\u056B \u0584\u0561\u0576\u056B \u0561\u0577\u056D\u0561\u057F\u0561\u0576\u0584\u0561\u0575\u056B\u0576 \u0585\u0580\u057E\u0561 \u0568\u0576\u0569\u0561\u0581\u0584\u0578\u0582\u0574: \u053D\u0576\u0564\u0580\u0578\u0582\u0574 \u0565\u0576\u0584 \u057A\u0561\u0570\u0565\u056C \u0571\u0565\u0580 \u0585\u0580\u0561\u0581\u0578\u0582\u0575\u0581\u0568 \u0573\u056F\u0578\u0582\u0576 \u0570\u0561\u057B\u0578\u0580\u0564 \u0577\u0561\u0562\u0561\u0569\u057E\u0561 \u0570\u0561\u0574\u0561\u0580:'
          ],
          team: 'ArgusRecruit-\u056B \u0569\u056B\u0574\u0568'
        }
      }
    },
    {
      key: 'interview',
      copy: {
        en: {
          subject: 'Employer has approved your profile \u2014 {jobTitle}',
          eyebrow: '\u2022 Employer Approved \u2022',
          h1: 'Good news \u2014 the employer wants to interview you.',
          greeting: 'Hi {name},',
          body: [
            'After reviewing your profile, the employer has approved you for the next stage of the <strong style="color:#D4AF37;">{jobTitle}</strong> hiring process.',
            'They will coordinate directly with you in the next few days to schedule the interview. Please keep your calendar flexible for the coming week.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: '\u0420\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u044C \u043E\u0434\u043E\u0431\u0440\u0438\u043B \u0432\u0430\u0448 \u043F\u0440\u043E\u0444\u0438\u043B\u044C \u2014 {jobTitle}',
          eyebrow: '\u2022 \u041E\u0434\u043E\u0431\u0440\u0435\u043D\u043E \u0440\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u0435\u043C \u2022',
          h1: '\u0425\u043E\u0440\u043E\u0448\u0438\u0435 \u043D\u043E\u0432\u043E\u0441\u0442\u0438 \u2014 \u0440\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u044C \u0445\u043E\u0447\u0435\u0442 \u043F\u0440\u043E\u0432\u0435\u0441\u0442\u0438 \u0441 \u0432\u0430\u043C\u0438 \u0438\u043D\u0442\u0435\u0440\u0432\u044C\u044E.',
          greeting: '\u0417\u0434\u0440\u0430\u0432\u0441\u0442\u0432\u0443\u0439\u0442\u0435, {name},',
          body: [
            '\u041F\u043E\u0441\u043B\u0435 \u0440\u0430\u0441\u0441\u043C\u043E\u0442\u0440\u0435\u043D\u0438\u044F \u0432\u0430\u0448\u0435\u0433\u043E \u043F\u0440\u043E\u0444\u0438\u043B\u044F \u0440\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u044C \u043E\u0434\u043E\u0431\u0440\u0438\u043B \u0432\u0430\u0448\u0443 \u043A\u0430\u043D\u0434\u0438\u0434\u0430\u0442\u0443\u0440\u0443 \u0434\u043B\u044F \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0435\u0433\u043E \u044D\u0442\u0430\u043F\u0430 \u043E\u0442\u0431\u043E\u0440\u0430 \u043D\u0430 \u0440\u043E\u043B\u044C <strong style="color:#D4AF37;">{jobTitle}</strong>.',
            '\u0412 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0438\u0435 \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E \u0434\u043D\u0435\u0439 \u0440\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u044C \u043D\u0430\u043F\u0440\u044F\u043C\u0443\u044E \u0441\u043E\u0433\u043B\u0430\u0441\u0443\u0435\u0442 \u0441 \u0432\u0430\u043C\u0438 \u0432\u0440\u0435\u043C\u044F \u0438\u043D\u0442\u0435\u0440\u0432\u044C\u044E. \u041F\u043E\u0441\u0442\u0430\u0440\u0430\u0439\u0442\u0435\u0441\u044C \u043E\u0441\u0442\u0430\u0432\u0438\u0442\u044C \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C \u0433\u0438\u0431\u043A\u0438\u043C \u043D\u0430 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0443\u044E \u043D\u0435\u0434\u0435\u043B\u044E.'
          ],
          team: '\u041A\u043E\u043C\u0430\u043D\u0434\u0430 ArgusRecruit'
        },
        hy: {
          subject: '\u0533\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u0576 \u0570\u0561\u057D\u057F\u0561\u057F\u0565\u056C \u0567 \u0571\u0565\u0580 \u057A\u0580\u0578\u0586\u056B\u056C\u0568 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u0540\u0561\u057D\u057F\u0561\u057F\u057E\u0561\u056E \u0567 \u0563\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u056B \u056F\u0578\u0572\u0574\u056B\u0581 \u2022',
          h1: '\u0540\u0561\u0573\u0565\u056C\u056B \u056C\u0578\u0582\u0580 \u2014 \u0563\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u0576 \u0581\u0561\u0576\u056F\u0561\u0576\u0578\u0582\u0574 \u0567 \u0570\u0561\u0580\u0581\u0561\u0566\u0580\u0578\u0582\u0575\u0581 \u057E\u0565\u0580\u0581\u0576\u0565\u056C \u0571\u0565\u0566 \u0570\u0565\u057F:',
          greeting: '\u0532\u0561\u0580\u0587, {name},',
          body: [
            '\u0541\u0565\u0580 \u057A\u0580\u0578\u0586\u056B\u056C\u0568 \u0578\u0582\u057D\u0578\u0582\u0574\u0576\u0561\u057D\u056B\u0580\u0565\u056C\u0578\u0582\u0581 \u0570\u0565\u057F\u0578 \u0563\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u0576 \u0570\u0561\u057D\u057F\u0561\u057F\u0565\u056C \u0567 \u0571\u0565\u0580 \u0569\u0565\u056F\u0576\u0561\u056E\u0578\u0582\u0569\u0575\u0578\u0582\u0576\u0568 <strong style="color:#D4AF37;">{jobTitle}</strong> \u057A\u0561\u0577\u057F\u0578\u0576\u056B \u0570\u0561\u057B\u0578\u0580\u0564 \u0583\u0578\u0582\u056C\u056B \u0570\u0561\u0574\u0561\u0580:',
            '\u0540\u0561\u057B\u0578\u0580\u0564 \u0574\u056B \u0584\u0561\u0576\u056B \u0585\u0580\u057E\u0561 \u0568\u0576\u0569\u0561\u0581\u0584\u0578\u0582\u0574 \u0563\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u0576 \u0578\u0582\u0572\u0572\u0561\u056F\u056B\u0578\u0580\u0565\u0576 \u056F\u0570\u0561\u0574\u0561\u056F\u0561\u0580\u0563\u056B \u0571\u0565\u0566 \u0570\u0565\u057F\u055D \u0570\u0561\u0580\u0581\u0561\u0566\u0580\u0578\u0582\u0575\u0581\u0568 \u057A\u056C\u0561\u0576\u0561\u057E\u0578\u0580\u0565\u056C\u0578\u0582 \u0570\u0561\u0574\u0561\u0580: \u053D\u0576\u0564\u0580\u0578\u0582\u0574 \u0565\u0576\u0584 \u057A\u0561\u0570\u0565\u056C \u0571\u0565\u0580 \u0585\u0580\u0561\u0581\u0578\u0582\u0575\u0581\u0568 \u0573\u056F\u0578\u0582\u0576:'
          ],
          team: 'ArgusRecruit-\u056B \u0569\u056B\u0574\u0568'
        }
      }
    },
    {
      key: 'offer',
      copy: {
        en: {
          subject: 'An offer is on its way \u2014 {jobTitle}',
          eyebrow: '\u2022 Offer Incoming \u2022',
          h1: 'Excellent \u2014 an offer is being prepared.',
          greeting: 'Hi {name},',
          body: [
            'Excellent news \u2014 the employer has approved you for a formal offer on the <strong style="color:#D4AF37;">{jobTitle}</strong> role.',
            'Please wait to receive the formal offer \u2014 our team or the employer will share it with you soon. It will cover the package, relocation, and start date.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: '\u041E\u0444\u0444\u0435\u0440 \u0441\u043A\u043E\u0440\u043E \u0443 \u0432\u0430\u0441 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u041E\u0444\u0444\u0435\u0440 \u0441\u043A\u043E\u0440\u043E \u2022',
          h1: '\u041E\u0442\u043B\u0438\u0447\u043D\u043E \u2014 \u043E\u0444\u0444\u0435\u0440 \u0433\u043E\u0442\u043E\u0432\u0438\u0442\u0441\u044F.',
          greeting: '\u0417\u0434\u0440\u0430\u0432\u0441\u0442\u0432\u0443\u0439\u0442\u0435, {name},',
          body: [
            '\u041E\u0442\u043B\u0438\u0447\u043D\u044B\u0435 \u043D\u043E\u0432\u043E\u0441\u0442\u0438 \u2014 \u0440\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u044C \u043E\u0434\u043E\u0431\u0440\u0438\u043B \u0432\u0430\u0448\u0443 \u043A\u0430\u043D\u0434\u0438\u0434\u0430\u0442\u0443\u0440\u0443 \u0438 \u0433\u043E\u0442\u043E\u0432\u0438\u0442 \u043E\u0444\u0438\u0446\u0438\u0430\u043B\u044C\u043D\u044B\u0439 \u043E\u0444\u0444\u0435\u0440 \u043F\u043E \u0440\u043E\u043B\u0438 <strong style="color:#D4AF37;">{jobTitle}</strong>.',
            '\u041F\u043E\u0436\u0430\u043B\u0443\u0439\u0441\u0442\u0430, \u043E\u0436\u0438\u0434\u0430\u0439\u0442\u0435 \u043E\u0444\u0438\u0446\u0438\u0430\u043B\u044C\u043D\u043E\u0433\u043E \u043E\u0444\u0444\u0435\u0440\u0430 \u2014 \u043D\u0430\u0448\u0430 \u043A\u043E\u043C\u0430\u043D\u0434\u0430 \u0438\u043B\u0438 \u0440\u0430\u0431\u043E\u0442\u043E\u0434\u0430\u0442\u0435\u043B\u044C \u0441\u0432\u044F\u0436\u0443\u0442\u0441\u044F \u0441 \u0432\u0430\u043C\u0438 \u0432 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0435\u0435 \u0432\u0440\u0435\u043C\u044F. \u041E\u0444\u0444\u0435\u0440 \u043E\u0445\u0432\u0430\u0442\u0438\u0442 \u043F\u0430\u043A\u0435\u0442, \u0440\u0435\u043B\u043E\u043A\u0430\u0446\u0438\u044E \u0438 \u0434\u0430\u0442\u0443 \u0432\u044B\u0445\u043E\u0434\u0430.'
          ],
          team: '\u041A\u043E\u043C\u0430\u043D\u0434\u0430 ArgusRecruit'
        },
        hy: {
          subject: '\u0531\u057C\u0561\u057B\u0561\u0580\u056F\u0568 \u0571\u0565\u0580 \u0573\u0561\u0576\u0561\u057A\u0561\u0580\u0570\u056B\u0576 \u0567 \u2014 {jobTitle}',
          eyebrow: '\u2022 \u0544\u0578\u057F\u0561\u056C\u0578\u0582\u057F \u0561\u057C\u0561\u057B\u0561\u0580\u056F \u2022',
          h1: '\u0540\u056B\u0561\u0576\u0561\u056C\u056B \u2014 \u0561\u057C\u0561\u057B\u0561\u0580\u056F\u0568 \u057A\u0561\u057F\u0580\u0561\u057D\u057F\u057E\u0578\u0582\u0574 \u0567:',
          greeting: '\u0532\u0561\u0580\u0587, {name},',
          body: [
            '\u0533\u0565\u0580\u0561\u0566\u0561\u0576\u0581 \u056C\u0578\u0582\u0580 \u2014 \u0563\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u0576 \u0570\u0561\u057D\u057F\u0561\u057F\u0565\u056C \u0567 \u0571\u0565\u0580 \u0569\u0565\u056F\u0576\u0561\u056E\u0578\u0582\u0569\u0575\u0578\u0582\u0576\u0568 <strong style="color:#D4AF37;">{jobTitle}</strong> \u057A\u0561\u0577\u057F\u0578\u0576\u056B \u0570\u0561\u0574\u0561\u0580 \u0587 \u057A\u0561\u057F\u0580\u0561\u057D\u057F\u0578\u0582\u0574 \u0567 \u057A\u0561\u0577\u057F\u0578\u0576\u0561\u056F\u0561\u0576 \u0561\u057C\u0561\u057B\u0561\u0580\u056F\u0568:',
            '\u053D\u0576\u0564\u0580\u0578\u0582\u0574 \u0565\u0576\u0584 \u057D\u057A\u0561\u057D\u0565\u056C \u057A\u0561\u0577\u057F\u0578\u0576\u0561\u056F\u0561\u0576 \u0561\u057C\u0561\u057B\u0561\u0580\u056F\u056B\u0576 \u2014 \u0574\u0565\u0580 \u0569\u056B\u0574\u0568 \u056F\u0561\u0574 \u0563\u0578\u0580\u056E\u0561\u057F\u0578\u0582\u0576 \u0577\u0578\u0582\u057F\u0578\u057E \u056F\u056F\u0561\u057A\u057E\u0565\u0576 \u0571\u0565\u0566 \u0570\u0565\u057F: \u0531\u057C\u0561\u057B\u0561\u0580\u056F\u0568 \u056F\u0576\u0565\u0580\u0561\u057C\u056B \u0583\u0561\u0569\u0565\u0569\u0568, \u057F\u0565\u0572\u0561\u0583\u0578\u056D\u0578\u0582\u0574\u0568 \u0587 \u0574\u0565\u056F\u0576\u0561\u0580\u056F\u056B \u0585\u0580\u0568:'
          ],
          team: 'ArgusRecruit-\u056B \u0569\u056B\u0574\u0568'
        }
      }
    },
    {
      key: 'rejected',
      copy: {
        en: {
          subject: 'Update on your {jobTitle} application',
          eyebrow: '\u2022 Application Update \u2022',
          h1: 'Update on your application.',
          greeting: 'Hi {name},',
          body: [
            'Thank you again for your interest in the <strong style="color:#D4AF37;">{jobTitle}</strong> role and for the time you put into your application.',
            'After careful consideration, we will not be moving forward with your candidacy for this specific role. This was not an easy decision \u2014 the bar was very high and the final shortlist was small.',
            'Your profile will stay in our network. If we see future roles that better match your background, we will reach out. Wishing you the very best in your search.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: '\u041E\u0431\u043D\u043E\u0432\u043B\u0435\u043D\u0438\u0435 \u043F\u043E \u0432\u0430\u0448\u0435\u0439 \u0437\u0430\u044F\u0432\u043A\u0435 \u043D\u0430 {jobTitle}',
          eyebrow: '\u2022 \u041E\u0431\u043D\u043E\u0432\u043B\u0435\u043D\u0438\u0435 \u043F\u043E \u0437\u0430\u044F\u0432\u043A\u0435 \u2022',
          h1: '\u041E\u0431\u043D\u043E\u0432\u043B\u0435\u043D\u0438\u0435 \u043F\u043E \u0432\u0430\u0448\u0435\u0439 \u0437\u0430\u044F\u0432\u043A\u0435.',
          greeting: '\u0417\u0434\u0440\u0430\u0432\u0441\u0442\u0432\u0443\u0439\u0442\u0435, {name},',
          body: [
            '\u0421\u043F\u0430\u0441\u0438\u0431\u043E \u0437\u0430 \u0438\u043D\u0442\u0435\u0440\u0435\u0441 \u043A \u0440\u043E\u043B\u0438 <strong style="color:#D4AF37;">{jobTitle}</strong> \u0438 \u0437\u0430 \u0432\u0440\u0435\u043C\u044F, \u043F\u043E\u0442\u0440\u0430\u0447\u0435\u043D\u043D\u043E\u0435 \u043D\u0430 \u0437\u0430\u044F\u0432\u043A\u0443.',
            '\u041F\u043E\u0441\u043B\u0435 \u0442\u0449\u0430\u0442\u0435\u043B\u044C\u043D\u043E\u0433\u043E \u0440\u0430\u0441\u0441\u043C\u043E\u0442\u0440\u0435\u043D\u0438\u044F \u043C\u044B \u0440\u0435\u0448\u0438\u043B\u0438 \u043D\u0435 \u043F\u0440\u043E\u0434\u043E\u043B\u0436\u0430\u0442\u044C \u0440\u0430\u0441\u0441\u043C\u043E\u0442\u0440\u0435\u043D\u0438\u0435 \u0432\u0430\u0448\u0435\u0439 \u043A\u0430\u043D\u0434\u0438\u0434\u0430\u0442\u0443\u0440\u044B \u043F\u043E \u044D\u0442\u043E\u0439 \u043A\u043E\u043D\u043A\u0440\u0435\u0442\u043D\u043E\u0439 \u043F\u043E\u0437\u0438\u0446\u0438\u0438. \u042D\u0442\u043E \u0431\u044B\u043B\u043E \u043D\u0435\u043F\u0440\u043E\u0441\u0442\u043E\u0435 \u0440\u0435\u0448\u0435\u043D\u0438\u0435 \u2014 \u043A\u043E\u043D\u043A\u0443\u0440\u0441 \u0431\u044B\u043B \u043E\u0447\u0435\u043D\u044C \u0432\u044B\u0441\u043E\u043A\u0438\u0439, \u0438 \u0444\u0438\u043D\u0430\u043B\u044C\u043D\u044B\u0439 \u0448\u043E\u0440\u0442-\u043B\u0438\u0441\u0442 \u043D\u0435\u0431\u043E\u043B\u044C\u0448\u043E\u0439.',
            '\u0412\u0430\u0448 \u043F\u0440\u043E\u0444\u0438\u043B\u044C \u043E\u0441\u0442\u0430\u043D\u0435\u0442\u0441\u044F \u0432 \u043D\u0430\u0448\u0435\u0439 \u0441\u0435\u0442\u0438. \u0415\u0441\u043B\u0438 \u0443 \u043D\u0430\u0441 \u043F\u043E\u044F\u0432\u044F\u0442\u0441\u044F \u0431\u0443\u0434\u0443\u0449\u0438\u0435 \u0440\u043E\u043B\u0438, \u043A\u043E\u0442\u043E\u0440\u044B\u0435 \u043B\u0443\u0447\u0448\u0435 \u043F\u043E\u0434\u0445\u043E\u0434\u044F\u0442 \u0432\u0430\u043C, \u043C\u044B \u0441\u0432\u044F\u0436\u0435\u043C\u0441\u044F. \u0416\u0435\u043B\u0430\u0435\u043C \u0443\u0434\u0430\u0447\u0438 \u0432 \u043F\u043E\u0438\u0441\u043A\u0435.'
          ],
          team: '\u041A\u043E\u043C\u0430\u043D\u0434\u0430 ArgusRecruit'
        },
        hy: {
          subject: '\u0539\u0561\u0580\u0574\u0561\u0581\u0578\u0582\u0574 {jobTitle} \u057A\u0561\u0577\u057F\u0578\u0576\u056B \u0564\u056B\u0574\u0578\u0582\u0574\u056B \u057E\u0565\u0580\u0561\u0562\u0565\u0580\u0575\u0561\u056C',
          eyebrow: '\u2022 \u0534\u056B\u0574\u0578\u0582\u0574\u056B \u0569\u0561\u0580\u0574\u0561\u0581\u0578\u0582\u0574 \u2022',
          h1: '\u0539\u0561\u0580\u0574\u0561\u0581\u0578\u0582\u0574 \u0571\u0565\u0580 \u0564\u056B\u0574\u0578\u0582\u0574\u056B \u057E\u0565\u0580\u0561\u0562\u0565\u0580\u0575\u0561\u056C:',
          greeting: '\u0532\u0561\u0580\u0587, {name},',
          body: [
            '\u053F\u0580\u056F\u056B\u0576 \u0577\u0576\u0578\u0580\u0570\u0561\u056F\u0561\u056C\u0578\u0582\u0569\u0575\u0578\u0582\u0576 <strong style="color:#D4AF37;">{jobTitle}</strong> \u057A\u0561\u0577\u057F\u0578\u0576\u056B \u0576\u056F\u0561\u057F\u0574\u0561\u0574\u0562 \u0571\u0565\u0580 \u0570\u0565\u057F\u0561\u0584\u0580\u0584\u0580\u0578\u0582\u0569\u0575\u0561\u0576 \u0587 \u0564\u056B\u0574\u0578\u0582\u0574\u056B \u057E\u0580\u0561 \u056E\u0561\u056D\u057D\u057E\u0561\u056E \u056A\u0561\u0574\u0561\u0576\u0561\u056F\u056B \u0570\u0561\u0574\u0561\u0580:',
            '\u0544\u0561\u0576\u0580\u0561\u056F\u0580\u056F\u056B\u057F \u0564\u056B\u057F\u0561\u0580\u056F\u0578\u0582\u0574\u056B\u0581 \u0570\u0565\u057F\u0578 \u0574\u0565\u0576\u0584 \u0578\u0580\u0578\u0577\u0565\u056C \u0565\u0576\u0584 \u0579\u0577\u0561\u0580\u0578\u0582\u0576\u0561\u056F\u0565\u056C \u0571\u0565\u0580 \u0569\u0565\u056F\u0576\u0561\u056E\u0578\u0582\u0569\u0575\u0578\u0582\u0576\u0568 \u0561\u0575\u057D \u056F\u0578\u0576\u056F\u0580\u0565\u057F \u0564\u0565\u0580\u056B \u0570\u0561\u0574\u0561\u0580: \u054D\u0561 \u0570\u0565\u0577\u057F \u0578\u0580\u0578\u0577\u0578\u0582\u0574 \u0579\u0567\u0580 \u2014 \u0574\u0580\u0581\u0561\u056F\u0581\u0578\u0582\u0569\u0575\u0578\u0582\u0576\u0568 \u0577\u0561\u057F \u0562\u0561\u0580\u0571\u0580 \u0567\u0580:',
            '\u0541\u0565\u0580 \u057A\u0580\u0578\u0586\u056B\u056C\u0568 \u056F\u057A\u0561\u0570\u057A\u0561\u0576\u057E\u056B \u0574\u0565\u0580 \u0581\u0561\u0576\u0581\u0578\u0582\u0574: \u0531\u057A\u0561\u0563\u0561\u0575\u0578\u0582\u0574 \u0561\u057E\u0565\u056C\u056B \u0570\u0561\u0574\u0561\u057A\u0561\u057F\u0561\u057D\u056D\u0561\u0576 \u0564\u0565\u0580\u0565\u0580 \u0578\u0582\u0576\u0565\u0576\u0561\u056C\u0578\u0582 \u0564\u0565\u057A\u0584\u0578\u0582\u0574 \u0574\u0565\u0576\u0584 \u056F\u056F\u0561\u057A\u057E\u0565\u0576\u0584 \u0571\u0565\u0566 \u0570\u0565\u057F: \u0532\u0561\u0580\u0565\u0574\u0561\u0572\u0569\u0561\u0576\u0584\u0576\u0565\u0580\u0578\u057E \u0571\u0565\u0580 \u0578\u0580\u0578\u0576\u0574\u0561\u0576 \u0574\u0565\u057B:'
          ],
          team: 'ArgusRecruit-\u056B \u0569\u056B\u0574\u0568'
        }
      }
    }
  ];

  return stages.map(s => [
    s.key,
    s.copy.en.subject,
    brandedEmailHtml_(s.copy.en, 'en'),
    s.copy.ru.subject,
    brandedEmailHtml_(s.copy.ru, 'ru'),
    s.copy.hy.subject,
    brandedEmailHtml_(s.copy.hy, 'hy')
  ]);
}

const CTA_TEXT = {
  en: 'Browse All Open Roles',
  ru: '\u0412\u0441\u0435 \u043E\u0442\u043A\u0440\u044B\u0442\u044B\u0435 \u0432\u0430\u043A\u0430\u043D\u0441\u0438\u0438',
  hy: '\u0532\u0578\u056C\u0578\u0580 \u0562\u0561\u0581 \u0564\u0565\u0580\u0565\u0580\u0568'
};
const CTA_URL = {
  en: 'https://argusrecruit.com/jobs/',
  ru: 'https://argusrecruit.com/ru/jobs/',
  hy: 'https://argusrecruit.com/hy/jobs/'
};

function brandedEmailHtml_(c, lang) {
  lang = lang || 'en';
  const ctaText = CTA_TEXT[lang] || CTA_TEXT.en;
  const ctaUrl = CTA_URL[lang] || CTA_URL.en;
  const paragraphs = c.body
    .map(p => `<p style="margin:0 0 14px;color:rgba(255,255,255,0.85);">${p}</p>`)
    .join('');
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ArgusRecruit</title></head>
<body style="margin:0;padding:0;background:#0E2440;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;color:#ffffff;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0E2440;padding:32px 16px;">
    <tr><td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#1E4170;border:1px solid rgba(212,175,55,0.2);border-radius:14px;overflow:hidden;">
        <tr><td align="center" style="padding:36px 32px 24px;background:#16345A;border-bottom:1px solid rgba(212,175,55,0.2);">
          <img src="https://argusrecruit.com/logo.png" alt="ArgusRecruit" width="64" style="display:block;height:64px;width:auto;">
          <div style="margin-top:14px;font-size:11px;letter-spacing:2.5px;color:#D4AF37;text-transform:uppercase;font-weight:700;">Many Eyes. One Purpose.</div>
        </td></tr>
        <tr><td align="center" style="padding:30px 32px 0;">
          <div style="display:inline-block;font-size:10px;letter-spacing:4px;color:#D4AF37;text-transform:uppercase;font-weight:700;padding:6px 16px;border:1px solid rgba(212,175,55,0.4);border-radius:999px;">${c.eyebrow}</div>
        </td></tr>
        <tr><td align="center" style="padding:18px 32px 0;">
          <h1 style="margin:0;font-size:24px;font-weight:700;letter-spacing:0.5px;color:#ffffff;text-transform:none;line-height:1.25;">${c.h1}</h1>
        </td></tr>
        <tr><td style="padding:28px 36px 18px;color:rgba(255,255,255,0.85);font-size:15px;line-height:1.75;">
          <p style="margin:0 0 16px;">${c.greeting}</p>
          ${paragraphs}
        </td></tr>
        <tr><td align="center" style="padding:10px 32px 32px;">
          <a href="${ctaUrl}" style="display:inline-block;background:#D4AF37;color:#0E2440;font-weight:700;font-size:13px;letter-spacing:1px;padding:13px 28px;border-radius:999px;text-decoration:none;">${ctaText}</a>
        </td></tr>
        <tr><td style="padding:0 32px 28px;color:rgba(255,255,255,0.7);font-size:13px;line-height:1.6;text-align:center;font-style:italic;">\u2014 ${c.team}</td></tr>
        <tr><td style="padding:22px 32px;background:#0E2440;border-top:1px solid rgba(212,175,55,0.15);text-align:center;">
          <div style="color:rgba(255,255,255,0.55);font-size:12px;line-height:1.65;">
            <a href="https://argusrecruit.com" style="color:#D4AF37;text-decoration:none;">argusrecruit.com</a>
            &nbsp;\u00B7&nbsp;
            <a href="mailto:contact@argusrecruit.com" style="color:#D4AF37;text-decoration:none;">contact@argusrecruit.com</a>
          </div>
          <div style="margin-top:12px;color:rgba(255,255,255,0.4);font-size:11px;letter-spacing:1px;text-transform:uppercase;">\u00A9 2026 ArgusRecruit \u00B7 Yerevan, Armenia</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}


// ----------------------------------------------------------------------
// TELEGRAM INTAKE BOT
//
// Setup (one-time, run installTelegramWebhook() after deployment):
//   1. Deploy this script as a Web App (Deploy \u2192 New deployment).
//      Execute as: Me. Who has access: Anyone (so Telegram can POST).
//   2. Run installTelegramWebhook() once \u2014 it tells Telegram to send all
//      bot updates to this web app.
//   3. Open DM with @Sonicbot_bot, type /start, then forward a CV.
//
// Note: the bot will only respond to TG_ADMIN_CHAT_ID. Group messages and
// other users are silently ignored.
//
// Drive OCR: the bot converts each PDF to a temporary Google Doc to read
// the text. This uses the Drive Advanced Service which is automatically
// available \u2014 no manual enable needed.
// ----------------------------------------------------------------------

function installTelegramWebhook() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('Deploy the script as a Web App first, then run this.');
    return;
  }
  const res = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + TG_BOT_TOKEN + '/setWebhook?url=' + encodeURIComponent(url),
    { muteHttpExceptions: true }
  );
  SpreadsheetApp.getUi().alert('Telegram webhook response:\n' + res.getContentText());
}

function uninstallTelegramWebhook() {
  const res = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + TG_BOT_TOKEN + '/deleteWebhook',
    { muteHttpExceptions: true }
  );
  SpreadsheetApp.getUi().alert('Webhook removed:\n' + res.getContentText());
}

// doPost handles BOTH:
//   - the intake form HTML web app submissions (via google.script.run, no HTTP body)
//   - Telegram webhook updates (HTTP POST with JSON body)
// They are routed by inspecting the parameters.
function doPost(e) {
  try {
    if (e && e.postData && e.postData.type === 'application/json') {
      const update = JSON.parse(e.postData.contents);
      handleTelegramUpdate_(update);
      return ContentService.createTextOutput('ok');
    }
  } catch (err) {
    console.error('doPost error: ' + err + '\n' + (err.stack || ''));
  }
  return ContentService.createTextOutput('ok');
}

function handleTelegramUpdate_(update) {
  // Callback queries \u2014 inline-button taps
  if (update.callback_query) {
    const cq = update.callback_query;
    if (cq.from && cq.from.id !== TG_ADMIN_CHAT_ID) {
      tgAnswerCallback_(cq.id, 'Not authorized.');
      return;
    }
    handleCallbackQuery_(cq);
    return;
  }
  const msg = update.message || update.edited_message;
  if (!msg) return;
  const chatId = msg.chat && msg.chat.id;
  const fromId = msg.from && msg.from.id;
  if (fromId !== TG_ADMIN_CHAT_ID) return;

  if (msg.text) handleTgText_(chatId, msg.text.trim());
  if (msg.document) handleTgDocument_(chatId, msg.document, msg);
}

function tgSend_(chatId, text, extra) {
  const payload = Object.assign({
    chat_id: chatId,
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  }, extra || {});
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_BOT_TOKEN + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  try { return JSON.parse(res.getContentText()); } catch (_) { return null; }
}

function tgSendKb_(chatId, text, keyboard) {
  return tgSend_(chatId, text, { reply_markup: { inline_keyboard: keyboard } });
}

function tgEdit_(chatId, messageId, text, keyboard) {
  const payload = {
    chat_id: chatId,
    message_id: messageId,
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };
  if (keyboard) payload.reply_markup = { inline_keyboard: keyboard };
  UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_BOT_TOKEN + '/editMessageText', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function tgAnswerCallback_(callbackId, text) {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_BOT_TOKEN + '/answerCallbackQuery', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ callback_query_id: callbackId, text: text || '' }),
    muteHttpExceptions: true
  });
}

function tgState_(chatId) {
  const props = PropertiesService.getUserProperties();
  const key = 'tg_state_' + chatId;
  const raw = props.getProperty(key);
  return raw ? JSON.parse(raw) : {};
}
function tgSaveState_(chatId, state) {
  PropertiesService.getUserProperties()
    .setProperty('tg_state_' + chatId, JSON.stringify(state));
}
function tgClearState_(chatId) {
  PropertiesService.getUserProperties().deleteProperty('tg_state_' + chatId);
}

function showHomeMenu_(chatId) {
  tgSendKb_(chatId,
    '<b>ArgusRecruit Intake Bot</b>\n\nPick a job, choose how the candidate reached you, then send their CV \u2014 I\'ll read it, show what I found, and ask you to confirm.',
    [
      [{ text: '\uD83D\uDCCB Pick a job', callback_data: 'pickjob' }],
      [{ text: '\uD83D\uDD0D State', callback_data: 'state' }, { text: '\uD83C\uDD98 Help', callback_data: 'help' }]
    ]
  );
}

function showJobsKb_(chatId, messageId) {
  const jobs = listActiveJobs_();
  if (jobs.length === 0) {
    const text = 'No active job folders found in Drive root.';
    if (messageId) tgEdit_(chatId, messageId, text, [[{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]]);
    else tgSendKb_(chatId, text, [[{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]]);
    return;
  }
  const rows = jobs.map(j => [{
    text: `${j.jobId} \u2014 ${j.jobTitle}`,
    callback_data: 'job:' + j.jobId
  }]);
  rows.push([{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]);
  const text = '<b>Pick a job:</b>';
  if (messageId) tgEdit_(chatId, messageId, text, rows);
  else tgSendKb_(chatId, text, rows);
}

function showSourcesKb_(chatId, messageId, jobId, jobTitle) {
  const sources = SOURCES.filter(s => s !== 'web-apply');
  const rows = sources.map(s => [{
    text: s,
    callback_data: 'src:' + jobId + '|' + s
  }]);
  rows.push([{ text: '\u21A9\uFE0F Back to jobs', callback_data: 'pickjob' }]);
  const text = `<b>Job:</b> <code>${jobId}</code> \u2014 ${escHtml_(jobTitle)}\n\nHow did you find this candidate?`;
  if (messageId) tgEdit_(chatId, messageId, text, rows);
  else tgSendKb_(chatId, text, rows);
}

function showReadyKb_(chatId, messageId, jobId, jobTitle, source) {
  const text =
    `\u2705 <b>Ready.</b>\n\n` +
    `Job: <code>${jobId}</code> \u2014 ${escHtml_(jobTitle)}\n` +
    `Source: <code>${source}</code>\n\n` +
    `Now <b>send or forward</b> the candidate's CV (PDF, DOC, DOCX).`;
  const kb = [[
    { text: '\uD83D\uDD04 Change job', callback_data: 'pickjob' },
    { text: '\u21A9\uFE0F Cancel', callback_data: 'home' }
  ]];
  if (messageId) tgEdit_(chatId, messageId, text, kb);
  else tgSendKb_(chatId, text, kb);
}

function handleCallbackQuery_(cq) {
  const chatId = cq.message && cq.message.chat && cq.message.chat.id;
  const messageId = cq.message && cq.message.message_id;
  const data = cq.data || '';
  tgAnswerCallback_(cq.id);

  if (data === 'home') {
    tgEdit_(chatId, messageId,
      '<b>ArgusRecruit Intake Bot</b>\n\nPick a job, choose how the candidate reached you, then send their CV.',
      [
        [{ text: '\uD83D\uDCCB Pick a job', callback_data: 'pickjob' }],
        [{ text: '\uD83D\uDD0D State', callback_data: 'state' }, { text: '\uD83C\uDD98 Help', callback_data: 'help' }]
      ]
    );
    return;
  }
  if (data === 'pickjob') { showJobsKb_(chatId, messageId); return; }
  if (data === 'help') {
    tgEdit_(chatId, messageId,
      '<b>Help</b>\n\n1. Tap <b>Pick a job</b> and choose the role.\n2. Choose how the candidate reached you.\n3. Send or forward the CV (PDF/DOC/DOCX).\n4. I\'ll read it, show what I found, and you can <b>Confirm</b> or <b>Cancel</b>.\n\nYou can also type <code>/start</code> anytime to get back to the home menu.',
      [[{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]]
    );
    return;
  }
  if (data === 'state') {
    const st = tgState_(chatId);
    tgEdit_(chatId, messageId,
      'Current state:\n' +
      'job: ' + (st.jobId || '\u2014') + '\n' +
      'title: ' + (st.jobTitle || '\u2014') + '\n' +
      'source: ' + (st.source || '\u2014') + '\n' +
      'pending: ' + (st.pending ? 'yes' : 'no'),
      [[{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]]
    );
    return;
  }
  if (data.indexOf('job:') === 0) {
    const jobId = data.slice(4);
    const jobs = listActiveJobs_();
    const job = jobs.find(j => j.jobId === jobId);
    if (!job) { tgEdit_(chatId, messageId, 'Job not found.', [[{ text: '\u21A9\uFE0F Back', callback_data: 'pickjob' }]]); return; }
    showSourcesKb_(chatId, messageId, job.jobId, job.jobTitle);
    return;
  }
  if (data.indexOf('src:') === 0) {
    const rest = data.slice(4);
    const sep = rest.indexOf('|');
    const jobId = rest.slice(0, sep);
    const source = rest.slice(sep + 1);
    const jobs = listActiveJobs_();
    const job = jobs.find(j => j.jobId === jobId);
    if (!job) { tgEdit_(chatId, messageId, 'Job not found.', [[{ text: '\u21A9\uFE0F Back', callback_data: 'pickjob' }]]); return; }
    tgSaveState_(chatId, { jobId: job.jobId, jobTitle: job.jobTitle, source: source });
    showReadyKb_(chatId, messageId, job.jobId, job.jobTitle, source);
    return;
  }
  if (data === 'confirm') {
    const st = tgState_(chatId);
    if (!st.pending) { tgEdit_(chatId, messageId, 'Nothing pending. Pick a job and send a CV first.', [[{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]]); return; }
    let res;
    try { res = submitIntake_(st.pending); }
    catch (err) { res = { ok: false, error: String(err) }; }
    if (res.ok) {
      tgEdit_(chatId, messageId,
        '\u2705 Added <b>' + escHtml_(res.name) + '</b> \u2192 <code>' + res.jobId + '</code>.',
        [[
          { text: '\u2795 Another to same job', callback_data: 'again:' + st.jobId + '|' + (st.source || 'sourced-other') },
          { text: '\uD83C\uDFE0 Home', callback_data: 'home' }
        ]]
      );
    } else {
      tgEdit_(chatId, messageId,
        '\u26A0 Save failed: ' + escHtml_(res.error || 'unknown'),
        [[{ text: '\u21A9\uFE0F Back', callback_data: 'home' }]]
      );
    }
    const after = tgState_(chatId);
    delete after.pending;
    tgSaveState_(chatId, after);
    return;
  }
  if (data === 'cancel') {
    const st = tgState_(chatId);
    delete st.pending;
    tgSaveState_(chatId, st);
    tgEdit_(chatId, messageId, 'Cancelled. The CV was not added.',
      [[{ text: '\uD83C\uDFE0 Home', callback_data: 'home' }]]);
    return;
  }
  if (data.indexOf('again:') === 0) {
    const rest = data.slice(6);
    const sep = rest.indexOf('|');
    const jobId = rest.slice(0, sep);
    const source = rest.slice(sep + 1);
    const jobs = listActiveJobs_();
    const job = jobs.find(j => j.jobId === jobId);
    if (!job) { tgEdit_(chatId, messageId, 'Job not found.', [[{ text: '\u21A9\uFE0F Back', callback_data: 'pickjob' }]]); return; }
    tgSaveState_(chatId, { jobId: job.jobId, jobTitle: job.jobTitle, source: source });
    showReadyKb_(chatId, messageId, job.jobId, job.jobTitle, source);
    return;
  }
}

function handleTgText_(chatId, text) {
  if (text === '/start' || text === '/help') {
    showHomeMenu_(chatId);
    return;
  }
  if (text === '/jobs') {
    const jobs = listActiveJobs_();
    if (jobs.length === 0) { tgSend_(chatId, 'No active job folders found in Drive root.'); return; }
    const list = jobs.map(j => `\u2022 <code>${j.jobId}</code> \u2014 ${j.jobTitle}`).join('\n');
    tgSend_(chatId, '<b>Active jobs:</b>\n' + list);
    return;
  }
  if (text === '/state') {
    const st = tgState_(chatId);
    tgSend_(chatId,
      'Current state:\n' +
      '<code>job</code>: ' + (st.jobId || '\u2014') + '\n' +
      '<code>source</code>: ' + (st.source || '\u2014') + '\n' +
      '<code>pending</code>: ' + (st.pending ? 'yes' : 'no')
    );
    return;
  }
  if (text === '/cancel') {
    tgClearState_(chatId);
    tgSend_(chatId, '\u2713 Cleared. Set a new job with <code>/job AR-XXX source</code>.');
    return;
  }
  if (text === '/confirm') {
    const st = tgState_(chatId);
    if (!st.pending) { tgSend_(chatId, 'Nothing pending. Send a CV first.'); return; }
    try {
      const res = submitIntake_(st.pending);
      if (res.ok) {
        tgSend_(chatId, '\u2705 Added <b>' + escHtml_(res.name) + '</b> \u2192 <code>' + res.jobId + '</code>.\nReady for next \u2014 send another CV or <code>/cancel</code>.');
      } else {
        tgSend_(chatId, '\u26A0 Save failed: ' + escHtml_(res.error || 'unknown'));
      }
    } catch (err) {
      tgSend_(chatId, '\u26A0 Error: ' + escHtml_(String(err)));
    }
    const after = tgState_(chatId);
    delete after.pending;
    tgSaveState_(chatId, after);
    return;
  }
  // /job AR-PYT01 [source]
  const jobMatch = text.match(/^\/job\s+(\S+)(?:\s+(\S+))?\s*$/);
  if (jobMatch) {
    const jobId = jobMatch[1];
    const source = (jobMatch[2] || 'sourced-other').toLowerCase();
    // Look up the job folder
    const jobs = listActiveJobs_();
    const job = jobs.find(j => j.jobId.toLowerCase() === jobId.toLowerCase());
    if (!job) {
      tgSend_(chatId, '\u26A0 No job folder for <code>' + escHtml_(jobId) + '</code>. Use /jobs to see available IDs.');
      return;
    }
    tgSaveState_(chatId, { jobId: job.jobId, jobTitle: job.jobTitle, source: source });
    tgSend_(chatId,
      '\u2713 Set: <code>' + job.jobId + '</code> \u2014 ' + escHtml_(job.jobTitle) +
      '\nSource: <code>' + source + '</code>\n\nNow send/forward the CV.'
    );
    return;
  }
  tgSend_(chatId, 'I didn\'t understand. Send /help for commands.');
}

function handleTgDocument_(chatId, document, msg) {
  const state = tgState_(chatId);
  if (!state.jobId) {
    tgSend_(chatId, '\u26A0 Set a job first: <code>/job AR-XXX source</code>');
    return;
  }
  const filename = document.file_name || 'cv.pdf';
  if (!/\.(pdf|docx?|odt|rtf)$/i.test(filename)) {
    tgSend_(chatId, '\u26A0 Unsupported file type. Send PDF, DOC, or DOCX.');
    return;
  }
  tgSend_(chatId, '\u23F3 Reading <code>' + escHtml_(filename) + '</code>\u2026');

  let cvText = '';
  let blob = null;
  try {
    blob = tgDownloadFile_(document.file_id);
    cvText = extractTextFromCv_(blob, filename);
  } catch (err) {
    tgSend_(chatId, '\u26A0 Could not read file: ' + escHtml_(String(err)));
    return;
  }

  const fields = extractCandidateFields_(cvText, msg);

  const pending = {
    job: state.jobId + '|' + state.jobTitle,
    source: state.source || 'sourced-other',
    name: fields.name,
    email: fields.email,
    phone: fields.phone,
    linkedin: fields.linkedin,
    lang: 'en',
    notes: '(via Telegram bot)',
    cvName: filename,
    cvType: blob.getContentType(),
    cvData: Utilities.base64Encode(blob.getBytes())
  };
  tgSaveState_(chatId, Object.assign({}, state, { pending }));

  tgSendKb_(chatId,
    '<b>Found in CV:</b>\n' +
    '\uD83D\uDC64 Name: <code>' + escHtml_(fields.name || '?') + '</code>\n' +
    '\u2709 Email: <code>' + escHtml_(fields.email || '\u2014') + '</code>\n' +
    '\uD83D\uDCF1 Phone: <code>' + escHtml_(fields.phone || '\u2014') + '</code>\n' +
    '\uD83D\uDD17 LinkedIn: <code>' + escHtml_(fields.linkedin || '\u2014') + '</code>\n' +
    '\uD83D\uDCC1 Job: <code>' + state.jobId + '</code> \u2014 ' + escHtml_(state.jobTitle) + '\n' +
    '\uD83D\uDCE5 Source: <code>' + (state.source || 'sourced-other') + '</code>',
    [[
      { text: '\u2705 Confirm', callback_data: 'confirm' },
      { text: '\uD83D\uDD04 Change job', callback_data: 'pickjob' },
      { text: '\u274C Cancel', callback_data: 'cancel' }
    ]]
  );
}

function tgDownloadFile_(fileId) {
  const getRes = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + TG_BOT_TOKEN + '/getFile?file_id=' + encodeURIComponent(fileId),
    { muteHttpExceptions: true }
  );
  const meta = JSON.parse(getRes.getContentText());
  if (!meta.ok) throw new Error('Telegram getFile failed: ' + meta.description);
  const path = meta.result.file_path;
  const fileRes = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + TG_BOT_TOKEN + '/' + path);
  return fileRes.getBlob().setName(path.split('/').pop());
}

function extractTextFromCv_(blob, filename) {
  // For PDFs: copy to Drive with OCR conversion to a Google Doc, then read text.
  // For DOC/DOCX/ODT/RTF: same \u2014 Drive can convert them.
  const file = DriveApp.createFile(blob);
  let text = '';
  try {
    // Use Advanced Drive Service if available; otherwise use a manual copy with conversion.
    if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.copy) {
      const copy = Drive.Files.copy(
        { name: '__ocr_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
        file.getId(),
        { ocrLanguage: 'en' }
      );
      const doc = DocumentApp.openById(copy.id);
      text = doc.getBody().getText();
      DriveApp.getFileById(copy.id).setTrashed(true);
    } else {
      // Fallback: try a direct Doc open. Will fail for PDFs but works for DOC/DOCX uploads.
      throw new Error('Drive Advanced Service not available. Enable it in Apps Script: Services \u2192 Drive API.');
    }
  } finally {
    file.setTrashed(true);
  }
  return text;
}

function extractCandidateFields_(text, msg) {
  text = text || '';
  // Email
  const emailMatch = text.match(/[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/i);
  const email = emailMatch ? emailMatch[0].toLowerCase() : '';
  // Phone \u2014 pick the longest "+? digit sequence with separators" candidate
  const phoneCandidates = text.match(/\+?\d[\d\s\-().]{7,18}\d/g) || [];
  const phone = phoneCandidates
    .map(s => s.replace(/[^\d+]/g, ''))
    .filter(s => s.replace(/\D/g, '').length >= 9)
    .sort((a, b) => b.length - a.length)[0] || '';
  // LinkedIn
  const liMatch = text.match(/(?:https?:\/\/)?(?:www\.)?linkedin\.com\/in\/[\w\-_%]+/i);
  const linkedin = liMatch ? ('https://' + liMatch[0].replace(/^https?:\/\//, '')) : '';
  // Name \u2014 heuristic: first non-empty line that's 2\u20134 words, mostly letters, not all upper
  let name = '';
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean).slice(0, 25);
  for (const line of lines) {
    if (line.length > 80) continue;
    if (/@|http|linkedin|cv|resume|curriculum/i.test(line)) continue;
    const words = line.split(/\s+/);
    if (words.length < 2 || words.length > 4) continue;
    if (!/^[A-Za-z'\-\u00C0-\u017F\u0531-\u0556\u0561-\u0587\s]+$/.test(line)) continue;
    if (line === line.toUpperCase() && line.length > 6) continue;
    name = line.replace(/\s+/g, ' ').trim();
    break;
  }
  // Fallback: use the Telegram-sender's first/last name
  if (!name && msg && msg.from) {
    const f = msg.from.first_name || '';
    const l = msg.from.last_name || '';
    name = (f + ' ' + l).trim();
  }
  return { name, email, phone, linkedin };
}

function escHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
