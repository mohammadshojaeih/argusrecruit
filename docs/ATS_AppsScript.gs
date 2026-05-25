/**
 * ArgusRecruit ATS — Apps Script
 *
 * What it does
 * ────────────
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
 *    the new CV replaces the old (active stages only — skip if Rejected).
 *
 * One-time setup
 * ──────────────
 * Just call `setup()` from the editor — it provisions the Sheet (Applications,
 * Email Templates, Settings, History) with dropdowns, formulas, and conditional
 * formatting. Then add a time-trigger on `tick` every 5 minutes.
 */

// ──────────────────────────────────────────────────────────────────────
// CONFIG (edit ROOT_FOLDER_ID; the script will use the active spreadsheet)
// ──────────────────────────────────────────────────────────────────────

const ROOT_FOLDER_ID = '1Ssla0RAD7XXY81o3R2TV5r1_UQRbUhL4';
const APP_SHEET      = 'Applications';
const TPL_SHEET      = 'Email Templates';
const CFG_SHEET      = 'Settings';
const PROCESSED_LBL  = 'ats-processed';
const SEARCH_QUERY   = 'subject:"[Application" has:attachment -label:ats-processed newer_than:90d';

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
  '01-Queued · Reviewed',
  '01-Reviewed',
  '02-Queued · Shortlist',
  '02-Shortlist',
  '03-Queued · Interview',
  '03-Interview',
  '04-Queued · Offer',
  '04-Offer',
  '99-Queued · Reject',
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
  date:        1,   // A
  jobId:       2,   // B
  jobTitle:    3,   // C
  source:      4,   // D
  name:        5,   // E
  email:       6,   // F
  phone:       7,   // G
  linkedin:    8,   // H
  lang:        9,   // I
  stage:       10,  // J  ← DROPDOWN, the user edits this
  cvLink:      11,  // K
  notes:       12,  // L
  otherApps:   13,  // M  (auto formula)
  lastChange:  14,  // N
  history:     15,  // O
  lastStage:   16,  // P  (hidden — script uses to detect changes)
  fileId:      17   // Q  (hidden — CV file ID for moves)
};


// ──────────────────────────────────────────────────────────────────────
// MAIN ENTRY — run via 5-minute trigger
// ──────────────────────────────────────────────────────────────────────

function tick() {
  try { processInbox_(); } catch (e) { console.error('processInbox error: ' + e); }
  try { processStageChanges_(); } catch (e) { console.error('processStageChanges error: ' + e); }
}


// ──────────────────────────────────────────────────────────────────────
// STEP 1 — Process new application emails
// ──────────────────────────────────────────────────────────────────────

function processInbox_() {
  const label = getOrCreateLabel_(PROCESSED_LBL);
  const threads = GmailApp.search(SEARCH_QUERY, 0, 50);
  if (threads.length === 0) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(APP_SHEET);
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);

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
        appendHistory_(sheet, existing, `CV re-uploaded → ${stage}`);
        thread.addLabel(label); thread.markRead(); thread.moveToArchive();
        continue;
      }

      // New row
      const fileName = buildFilename_(submittedAt, name, jobTitle, cv.getName());
      const file = stageFolder.createFile(cv).setName(fileName);

      const row = [
        new Date(submittedAt || msg.getDate()),     // A date
        jobId || '',                                 // B
        jobTitle || '',                              // C
        source || 'web-apply',                       // D
        name || '',                                  // E
        (email || '').toLowerCase(),                 // F
        phone || '',                                 // G
        linkedin || '',                              // H
        lang || 'en',                                // I
        '00-New',                                    // J Stage
        '=HYPERLINK("' + file.getUrl() + '","Open")', // K
        '',                                          // L notes
        '',                                          // M will become formula below
        new Date(),                                  // N last change
        'created → 00-New',                          // O history
        '00-New',                                    // P lastStage
        file.getId()                                 // Q fileId
      ];
      sheet.appendRow(row);
      const r = sheet.getLastRow();
      sheet.getRange(r, COL.otherApps).setFormula(
        `=IFERROR(COUNTIF($F$2:$F, $F${r})-1, 0)`
      );
      thread.addLabel(label); thread.markRead(); thread.moveToArchive();
      console.log('Added: ' + name + ' / ' + jobId);
    } catch (e) {
      console.error('Thread error: ' + e + '\n' + (e.stack || ''));
    }
  }
}


// ──────────────────────────────────────────────────────────────────────
// STEP 2 — Detect Stage changes and react (move file + send email)
// ──────────────────────────────────────────────────────────────────────

function processStageChanges_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(APP_SHEET);
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
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
      sheet.getRange(row, COL.lastStage).setValue(currentStage);
      sheet.getRange(row, COL.lastChange).setValue(new Date());
      appendHistory_(sheet, row, `${lastStage || '?'} → ${currentStage}`);
    } catch (e) {
      console.error(`Row ${row} stage transition failed: ${e}`);
    }
  }
}


// ──────────────────────────────────────────────────────────────────────
// PARSING + DRIVE HELPERS
// ──────────────────────────────────────────────────────────────────────

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
  const data = sheet.getRange(2, COL.email, last - 1, COL.jobId - COL.email + 6).getValues();
  // We only need columns F (email) and B (jobId). Adjust by re-fetching range smartly.
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
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
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


// ──────────────────────────────────────────────────────────────────────
// EMAIL SENDING
// ──────────────────────────────────────────────────────────────────────

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

  GmailApp.sendEmail(ctx.email, fill(subjectTpl), '', {
    htmlBody: fill(bodyTpl),
    name: getSetting_('From Name') || 'ArgusRecruit',
    replyTo: getSetting_('Reply To') || 'contact@argusrecruit.com'
  });
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


// ──────────────────────────────────────────────────────────────────────
// SHEET PROVISIONING (run once from the editor)
// ──────────────────────────────────────────────────────────────────────

function setup() {
  const ss = SpreadsheetApp.getActive();

  // 1. Applications sheet
  let app = ss.getSheetByName(APP_SHEET);
  if (!app) app = ss.insertSheet(APP_SHEET);
  app.clear();
  app.getRange(1, 1, 1, 17).setValues([[
    'Date Applied', 'Job ID', 'Job Title', 'Source',
    'Name', 'Email', 'Phone', 'LinkedIn', 'Lang',
    'Stage', 'CV', 'Notes', '# Other Apps',
    'Last Change', 'History', 'Last Known Stage (auto)', 'File ID (auto)'
  ]]);
  app.setFrozenRows(1);
  app.getRange('A1:Q1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  app.setColumnWidth(1, 100); app.setColumnWidth(2, 80); app.setColumnWidth(3, 180);
  app.setColumnWidth(4, 90); app.setColumnWidth(5, 140); app.setColumnWidth(6, 200);
  app.setColumnWidth(7, 110); app.setColumnWidth(8, 200); app.setColumnWidth(9, 60);
  app.setColumnWidth(10, 180); app.setColumnWidth(11, 60); app.setColumnWidth(12, 200);
  app.setColumnWidth(13, 80); app.setColumnWidth(14, 110); app.setColumnWidth(15, 280);
  app.hideColumns(16, 2);

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

  // Highlight rows where # Other Apps > 0
  const condRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$M2>0')
    .setBackground('#FFF7E0')
    .setRanges([app.getRange(2, 1, 1000, 17)])
    .build();
  app.setConditionalFormatRules([condRule]);

  // 2. Email Templates sheet
  let tpl = ss.getSheetByName(TPL_SHEET);
  if (!tpl) tpl = ss.insertSheet(TPL_SHEET);
  tpl.clear();
  tpl.getRange(1, 1, 1, 7).setValues([[
    'Key', 'en', 'en (body)', 'ru', 'ru (body)', 'hy', 'hy (body)'
  ]]);
  tpl.setFrozenRows(1);
  tpl.getRange('A1:G1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  // (Body cells hold full HTML — branded ArgusRecruit template per stage/language.)
  const rowsTpl = stageTemplates_();
  tpl.getRange(2, 1, rowsTpl.length, 7).setValues(rowsTpl);
  tpl.setColumnWidth(1, 110);
  for (let c = 2; c <= 7; c++) tpl.setColumnWidth(c, 300);
  tpl.setRowHeights(2, rowsTpl.length, 90);

  // 3. Settings sheet
  let cfg = ss.getSheetByName(CFG_SHEET);
  if (!cfg) cfg = ss.insertSheet(CFG_SHEET);
  cfg.clear();
  cfg.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
  cfg.getRange('A1:B1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  cfg.getRange(2, 1, 5, 2).setValues([
    ['Send Emails Enabled', 'TRUE'],
    ['From Name',           'ArgusRecruit'],
    ['Reply To',            'contact@argusrecruit.com'],
    ['Root Folder ID',      ROOT_FOLDER_ID],
    ['Trigger Interval',    'every 5 minutes (set via Triggers panel)']
  ]);
  cfg.setColumnWidth(1, 200);
  cfg.setColumnWidth(2, 400);

  SpreadsheetApp.getActive().toast('ATS provisioned. Now set a time-driven trigger on `tick` (every 5 minutes).', 'Setup complete', 8);
}


// ──────────────────────────────────────────────────────────────────────
// CUSTOM MENU — for manual ops
// ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ATS')
    .addItem('Sync now', 'tick')
    .addItem('Run setup (first time only)', 'setup')
    .addSeparator()
    .addItem('Show intake form URL', 'showIntakeUrl')
    .addToUi();
}

function showIntakeUrl() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Deploy first',
      'You need to deploy the script as a Web app.\n\n' +
      '1. In the Apps Script editor: Deploy → New deployment\n' +
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


// ──────────────────────────────────────────────────────────────────────
// INTAKE FORM (Apps Script Web App)
// ──────────────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutput(intakeFormHtml_())
    .setTitle('ArgusRecruit · Add Candidate')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function listActiveJobs_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
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
    `<option value="${j.jobId}|${j.jobTitle}">${j.jobId} — ${j.jobTitle}</option>`
  ).join('');
  const sourceOptions = SOURCES
    .filter(s => s !== 'web-apply') // can't be web-apply from here
    .map(s => `<option value="${s}">${s}</option>`).join('');

  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Add candidate · ArgusRecruit ATS</title>
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
      <option value="">— select a job —</option>
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
    <label>CV file (optional — PDF, DOC, DOCX, max 10MB)</label>
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
  submitBtn.disabled = true; submitBtn.textContent = 'Saving…';

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
        msg.textContent = '✓ Added: ' + r.name + ' → ' + r.jobId;
        form.reset();
      } else {
        msg.className = 'msg err';
        msg.textContent = '⚠ ' + (r.error || 'Could not save.');
      }
      submitBtn.disabled = false; submitBtn.textContent = 'Add candidate';
    })
    .withFailureHandler(err => {
      msg.className = 'msg err';
      msg.textContent = '⚠ ' + err.message;
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

    const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const jobFolder = findOrCreateFolder_(root, `${jobId} - ${jobTitle}`);

    // Determine initial stage:
    //   if no CV at all → 00-Pre-Contact
    //   else            → 00-New
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
      new Date(),
      jobId,
      jobTitle,
      payload.source || 'manual-intake',
      payload.name,
      (payload.email || '').toLowerCase(),
      payload.phone || '',
      payload.linkedin || '',
      payload.lang || 'en',
      initialStage,
      cvHyperlink,
      payload.notes || '',
      '',
      new Date(),
      'created (' + (payload.source || 'manual') + ') → ' + initialStage,
      initialStage,
      fileId
    ];
    sheet.appendRow(row);
    const r = sheet.getLastRow();
    sheet.getRange(r, COL.otherApps).setFormula(
      `=IFERROR(COUNTIF($F$2:$F, $F${r})-1, 0)`
    );
    return { ok: true, name: payload.name, jobId: jobId };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// ──────────────────────────────────────────────────────────────────────
// BRANDED EMAIL TEMPLATES
// ──────────────────────────────────────────────────────────────────────
//   Each cell stores a complete HTML document with ArgusRecruit branding
//   (navy + gold). Placeholders: {name}, {jobTitle}, {jobId}.
//   To edit copy: open Email Templates sheet and change the text — the
//   HTML structure is in the cell, edit between the tags.

function stageTemplates_() {
  const stages = [
    {
      key: 'reviewed',
      copy: {
        en: {
          subject: 'Your application is being reviewed — {jobTitle}',
          eyebrow: '• Application Under Review •',
          h1: 'Thanks — we\'re looking at your profile.',
          greeting: 'Hi {name},',
          body: [
            'Thank you for applying to <strong style="color:#D4AF37;">{jobTitle}</strong> at ArgusRecruit. A senior recruiter is now reviewing your application personally.',
            'If your background matches the role, we will be in touch within 1–3 business days. If we don\'t see a fit for this specific role, your profile will stay in our network and we may reach out about future openings that better match your background.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Ваша заявка на рассмотрении — {jobTitle}',
          eyebrow: '• Заявка рассматривается •',
          h1: 'Спасибо — мы рассматриваем ваш профиль.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Спасибо за вашу заявку на роль <strong style="color:#D4AF37;">{jobTitle}</strong>. Старший рекрутер лично рассматривает ваш профиль.',
            'Если ваш опыт подходит роли, мы свяжемся с вами в течение 1–3 рабочих дней. Если совпадения для этой конкретной роли не будет, мы сохраним ваш профиль в нашей сети и можем связаться с вами по поводу будущих вакансий.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Ձեր դիմումը քննարկվում է — {jobTitle}',
          eyebrow: '• Դիմումը քննարկման փուլում •',
          h1: 'Շնորհակալություն — մենք ուսումնասիրում ենք ձեր պրոֆիլը:',
          greeting: 'Բարև, {name},',
          body: [
            '<strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի համար դիմելու համար շնորհակալություն: Ավագ ռեկրուտերը այժմ անձամբ ուսումնասիրում է ձեր փորձառությունը:',
            'Համապատասխան լինելու դեպքում մենք կկապվենք ձեզ հետ 1–3 աշխատանքային օրվա ընթացքում: Հակառակ դեպքում ձեր պրոֆիլը կպահպանվի մեր ցանցում ապագա հնարավորությունների համար:'
          ],
          team: 'ArgusRecruit-ի թիմը'
        }
      }
    },
    {
      key: 'shortlist',
      copy: {
        en: {
          subject: 'You\'ve been shortlisted — {jobTitle}',
          eyebrow: '• Shortlisted •',
          h1: 'You\'re on the shortlist.',
          greeting: 'Hi {name},',
          body: [
            'Great news — you have been shortlisted for the <strong style="color:#D4AF37;">{jobTitle}</strong> role. Our client will now review your profile directly.',
            'We\'ll get back to you with the next steps within a few business days. In the meantime, please keep your calendar flexible for the coming week in case interview slots are offered.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Вы в шорт-листе — {jobTitle}',
          eyebrow: '• В шорт-листе •',
          h1: 'Вы попали в шорт-лист.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Хорошие новости — вы вошли в шорт-лист по позиции <strong style="color:#D4AF37;">{jobTitle}</strong>. Клиент сейчас рассматривает ваш профиль.',
            'Мы свяжемся с вами по следующим шагам в течение нескольких рабочих дней. Постарайтесь оставить календарь гибким на следующую неделю — возможны слоты для собеседования.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Դուք ընտրվել եք կարճ ցուցակում — {jobTitle}',
          eyebrow: '• Կարճ ցուցակում •',
          h1: 'Դուք կարճ ցուցակում եք:',
          greeting: 'Բարև, {name},',
          body: [
            'Հաճելի լուր — դուք ընտրվել եք <strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի կարճ ցուցակում: Մեր հաճախորդն այժմ ուղղակիորեն ուսումնասիրում է ձեր պրոֆիլը:',
            'Մենք կտեղեկացնենք ձեզ հաջորդ քայլերի մասին մի քանի աշխատանքային օրվա ընթացքում: Խնդրում ենք պահել ձեր օրացույցը ճկուն հաջորդ շաբաթվա համար:'
          ],
          team: 'ArgusRecruit-ի թիմը'
        }
      }
    },
    {
      key: 'interview',
      copy: {
        en: {
          subject: 'Interview confirmed — {jobTitle}',
          eyebrow: '• Interview Confirmed •',
          h1: 'Congratulations — interview confirmed.',
          greeting: 'Hi {name},',
          body: [
            'Your interview for <strong style="color:#D4AF37;">{jobTitle}</strong> has been confirmed. A senior member of our team and the hiring manager will meet you.',
            'We will follow up shortly with the exact time, format, and the people you will be meeting. If you have a preferred time window in the next 5 business days, reply to this email so we can prioritize it.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Собеседование подтверждено — {jobTitle}',
          eyebrow: '• Собеседование подтверждено •',
          h1: 'Поздравляем — собеседование подтверждено.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Ваше собеседование на роль <strong style="color:#D4AF37;">{jobTitle}</strong> подтверждено. Вас встретит старший член нашей команды и руководитель найма.',
            'Мы пришлём точное время, формат и список участников. Если у вас есть предпочтительные слоты в ближайшие 5 рабочих дней — ответьте на это письмо.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Հարցազրույցը հաստատվել է — {jobTitle}',
          eyebrow: '• Հարցազրույցը հաստատված •',
          h1: 'Շնորհավորում ենք — հարցազրույցը հաստատվել է:',
          greeting: 'Բարև, {name},',
          body: [
            'Ձեր հարցազրույցը <strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի համար հաստատվել է: Մեր թիմի ավագ անդամը և վարձող մենեջերը կհանդիպեն ձեզ հետ:',
            'Շուտով կուղարկենք ձեզ ճշգրիտ ժամանակը, ձևաչափը և մասնակիցների ցանկը: Եթե հաջորդ 5 աշխատանքային օրերի ընթացքում ունեք նախընտրելի ժամ, պատասխանեք այս նամակին:'
          ],
          team: 'ArgusRecruit-ի թիմը'
        }
      }
    },
    {
      key: 'offer',
      copy: {
        en: {
          subject: 'An offer is on its way — {jobTitle}',
          eyebrow: '• Offer Incoming •',
          h1: 'Excellent — an offer is being prepared.',
          greeting: 'Hi {name},',
          body: [
            'Excellent news — our client is preparing a formal offer for the <strong style="color:#D4AF37;">{jobTitle}</strong> role.',
            'We will reach out in the next 24 hours to walk you through the package, answer any questions about compensation, relocation, or timing, and align on next steps.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Оффер скоро у вас — {jobTitle}',
          eyebrow: '• Оффер скоро •',
          h1: 'Отлично — оффер готовится.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Отличные новости — клиент готовит официальный оффер по роли <strong style="color:#D4AF37;">{jobTitle}</strong>.',
            'В ближайшие 24 часа мы свяжемся, чтобы обсудить пакет, ответить на вопросы о компенсации, переезде и сроках, и согласовать следующие шаги.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Առաջարկը ձեր ճանապարհին է — {jobTitle}',
          eyebrow: '• Մոտալուտ առաջարկ •',
          h1: 'Հիանալի — առաջարկը պատրաստվում է:',
          greeting: 'Բարև, {name},',
          body: [
            'Գերազանց լուր — մեր հաճախորդը պատրաստում է պաշտոնական առաջարկ <strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի համար:',
            'Հաջորդ 24 ժամվա ընթացքում մենք կկապվենք ձեզ հետ՝ քննարկելու փաթեթը, փոխհատուցման, տեղափոխման և ժամանակացույցի մանրամասները:'
          ],
          team: 'ArgusRecruit-ի թիմը'
        }
      }
    },
    {
      key: 'rejected',
      copy: {
        en: {
          subject: 'Update on your {jobTitle} application',
          eyebrow: '• Application Update •',
          h1: 'Update on your application.',
          greeting: 'Hi {name},',
          body: [
            'Thank you again for your interest in the <strong style="color:#D4AF37;">{jobTitle}</strong> role and for the time you put into your application.',
            'After careful consideration, we will not be moving forward with your candidacy for this specific role. This was not an easy decision — the bar was very high and the final shortlist was small.',
            'Your profile will stay in our network. If we see future roles that better match your background, we will reach out. Wishing you the very best in your search.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Обновление по вашей заявке на {jobTitle}',
          eyebrow: '• Обновление по заявке •',
          h1: 'Обновление по вашей заявке.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Спасибо за интерес к роли <strong style="color:#D4AF37;">{jobTitle}</strong> и за время, потраченное на заявку.',
            'После тщательного рассмотрения мы решили не продолжать рассмотрение вашей кандидатуры по этой конкретной позиции. Это было непростое решение — конкурс был очень высокий, и финальный шорт-лист небольшой.',
            'Ваш профиль останется в нашей сети. Если у нас появятся будущие роли, которые лучше подходят вам, мы свяжемся. Желаем удачи в поиске.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Թարմացում {jobTitle} պաշտոնի դիմումի վերաբերյալ',
          eyebrow: '• Դիմումի թարմացում •',
          h1: 'Թարմացում ձեր դիմումի վերաբերյալ:',
          greeting: 'Բարև, {name},',
          body: [
            'Կրկին շնորհակալություն <strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի նկատմամբ ձեր հետաքրքրության և դիմումի վրա ծախսված ժամանակի համար:',
            'Մանրակրկիտ դիտարկումից հետո մենք որոշել ենք չշարունակել ձեր թեկնածությունը այս կոնկրետ դերի համար: Սա հեշտ որոշում չէր — մրցակցությունը շատ բարձր էր:',
            'Ձեր պրոֆիլը կպահպանվի մեր ցանցում: Ապագայում ավելի համապատասխան դերեր ունենալու դեպքում մենք կկապվենք ձեզ հետ: Բարեմաղթանքներով ձեր որոնման մեջ:'
          ],
          team: 'ArgusRecruit-ի թիմը'
        }
      }
    }
  ];

  return stages.map(s => [
    s.key,
    s.copy.en.subject,
    brandedEmailHtml_(s.copy.en),
    s.copy.ru.subject,
    brandedEmailHtml_(s.copy.ru),
    s.copy.hy.subject,
    brandedEmailHtml_(s.copy.hy)
  ]);
}

function brandedEmailHtml_(c) {
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
          <a href="https://argusrecruit.com/jobs/" style="display:inline-block;background:#D4AF37;color:#0E2440;font-weight:700;font-size:13px;letter-spacing:1px;text-transform:uppercase;padding:13px 28px;border-radius:999px;text-decoration:none;">Browse All Open Roles</a>
        </td></tr>
        <tr><td style="padding:0 32px 28px;color:rgba(255,255,255,0.7);font-size:13px;line-height:1.6;text-align:center;font-style:italic;">— ${c.team}</td></tr>
        <tr><td style="padding:22px 32px;background:#0E2440;border-top:1px solid rgba(212,175,55,0.15);text-align:center;">
          <div style="color:rgba(255,255,255,0.55);font-size:12px;line-height:1.65;">
            <a href="https://argusrecruit.com" style="color:#D4AF37;text-decoration:none;">argusrecruit.com</a>
            &nbsp;·&nbsp;
            <a href="mailto:contact@argusrecruit.com" style="color:#D4AF37;text-decoration:none;">contact@argusrecruit.com</a>
          </div>
          <div style="margin-top:12px;color:rgba(255,255,255,0.4);font-size:11px;letter-spacing:1px;text-transform:uppercase;">© 2026 ArgusRecruit · Yerevan, Armenia</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}
