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

// Telegram intake bot — see installTelegramWebhook() for setup
const TG_BOT_TOKEN     = '8706377970:AAEtKkn1Cl68PSmX65wFtFMJAL3XFHaegeo'; // @ArgusIntakeBot
const TG_ADMIN_CHAT_ID = 814437645;  // Mohammad — only this account is allowed to use the intake bot

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

/**
 * setup() is IDEMPOTENT and SAFE TO RE-RUN.
 * It only creates missing sheets / sets headers, dropdowns, formatting.
 * It NEVER deletes existing rows.
 *
 * For destructive operations use:
 *   - resetTemplatesOnly()  → wipes only Email Templates sheet and re-seeds it
 *   - resetApplications_DANGER() → wipes Applications data (requires manual edit to enable)
 */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // ── 1. Applications sheet (non-destructive) ────────────────────────
  let app = ss.getSheetByName(APP_SHEET);
  const isNewApp = !app;
  if (!app) app = ss.insertSheet(APP_SHEET);

  // Only write headers if this is a brand new sheet OR the headers are missing.
  const firstCell = app.getRange(1, 1).getValue();
  if (isNewApp || !firstCell) {
    app.getRange(1, 1, 1, 17).setValues([[
      'Date Applied', 'Job ID', 'Job Title', 'Source',
      'Name', 'Email', 'Phone', 'LinkedIn', 'Lang',
      'Stage', 'CV', 'Notes', '# Other Apps',
      'Last Change', 'History', 'Last Known Stage (auto)', 'File ID (auto)'
    ]]);
  }
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

  // ── 2. Email Templates sheet (non-destructive) ─────────────────────
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

  // ── 3. Settings sheet (non-destructive) ────────────────────────────
  let cfg = ss.getSheetByName(CFG_SHEET);
  const isNewCfg = !cfg;
  if (!cfg) cfg = ss.insertSheet(CFG_SHEET);
  if (isNewCfg || !cfg.getRange(1, 1).getValue()) {
    cfg.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    cfg.getRange(2, 1, 5, 2).setValues([
      ['Send Emails Enabled', 'TRUE'],
      ['From Name',           'ArgusRecruit'],
      ['Reply To',            'contact@argusrecruit.com'],
      ['Root Folder ID',      ROOT_FOLDER_ID],
      ['Trigger Interval',    'every 5 minutes (set via Triggers panel)']
    ]);
  }
  cfg.getRange('A1:B1').setFontWeight('bold').setBackground('#0E2440').setFontColor('#FFFFFF');
  cfg.setColumnWidth(1, 200);
  cfg.setColumnWidth(2, 400);

  ss.toast('ATS structure verified. Existing data was preserved.', 'Setup complete', 6);
}


/** Wipes and re-seeds ONLY the Email Templates sheet. Safe — Applications data is untouched. */
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


// ──────────────────────────────────────────────────────────────────────
// CUSTOM MENU — for manual ops
// ──────────────────────────────────────────────────────────────────────

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
    .addItem('Reset email templates only', 'resetTemplatesOnly')
    .addToUi();
}

/**
 * Preview the email template under the currently-selected cell.
 * Usage: open Email Templates sheet → click any body cell (en, ru, hy columns)
 *        → ATS menu → "Preview selected email template".
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
            'Thank you for applying to <strong style="color:#D4AF37;">{jobTitle}</strong> at ArgusRecruit. Our team is now carefully reviewing your application, and we\'ll get back to you soon.',
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
            'Спасибо за вашу заявку на роль <strong style="color:#D4AF37;">{jobTitle}</strong>. Наша команда тщательно рассматривает ваш профиль и свяжется с вами в ближайшее время.',
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
            '<strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի համար դիմելու համար շնորհակալություն: Մեր թիմը մանրակրկիտ ուսումնասիրում է ձեր դիմումը և շուտով կկապվենք ձեզ հետ:',
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
          subject: 'Your profile is moving forward — {jobTitle}',
          eyebrow: '• Profile Forwarded •',
          h1: 'Thanks — your profile is moving forward.',
          greeting: 'Hi {name},',
          body: [
            'Thank you for taking the time for the initial interview with us. Based on our conversation, your profile could be a strong match for the <strong style="color:#D4AF37;">{jobTitle}</strong> role.',
            'We\'re now sharing your CV with the employer. Please wait to hear from us, or directly from the employer, to schedule the next interview.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Ваш профиль передан работодателю — {jobTitle}',
          eyebrow: '• Профиль передан •',
          h1: 'Спасибо — мы передаём ваш профиль дальше.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Спасибо за время, уделённое первичному интервью с нами. По итогам разговора ваш профиль может стать отличным совпадением для роли <strong style="color:#D4AF37;">{jobTitle}</strong>.',
            'Сейчас мы передаём ваше резюме работодателю. Пожалуйста, ожидайте обратной связи от нас или напрямую от работодателя для назначения следующего интервью.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Ձեր պրոֆիլն առաջ է գնում — {jobTitle}',
          eyebrow: '• Պրոֆիլը փոխանցված է •',
          h1: 'Շնորհակալություն — ձեր պրոֆիլն առաջ է գնում:',
          greeting: 'Բարև, {name},',
          body: [
            'Շնորհակալություն մեզ հետ առաջնային հարցազրույցին հատկացրած ժամանակի համար: Մեր խոսակցության հիման վրա ձեր պրոֆիլը կարող է լավ համապատասխան լինել <strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնին:',
            'Այժմ ձեր CV-ն փոխանցում ենք գործատուին: Խնդրում ենք սպասել մեր կողմից կամ ուղղակիորեն գործատուի կողմից պատասխանին՝ հաջորդ հարցազրույցը պլանավորելու համար:'
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
            'Excellent news — the employer has approved you for a formal offer on the <strong style="color:#D4AF37;">{jobTitle}</strong> role.',
            'Please wait to receive the formal offer — our team or the employer will share it with you soon. It will cover the package, relocation, and start date.'
          ],
          team: 'The ArgusRecruit Team'
        },
        ru: {
          subject: 'Оффер скоро у вас — {jobTitle}',
          eyebrow: '• Оффер скоро •',
          h1: 'Отлично — оффер готовится.',
          greeting: 'Здравствуйте, {name},',
          body: [
            'Отличные новости — работодатель одобрил вашу кандидатуру и готовит официальный оффер по роли <strong style="color:#D4AF37;">{jobTitle}</strong>.',
            'Пожалуйста, ожидайте официального оффера — наша команда или работодатель свяжутся с вами в ближайшее время. Оффер охватит пакет, релокацию и дату выхода.'
          ],
          team: 'Команда ArgusRecruit'
        },
        hy: {
          subject: 'Առաջարկը ձեր ճանապարհին է — {jobTitle}',
          eyebrow: '• Մոտալուտ առաջարկ •',
          h1: 'Հիանալի — առաջարկը պատրաստվում է:',
          greeting: 'Բարև, {name},',
          body: [
            'Գերազանց լուր — գործատուն հաստատել է ձեր թեկնածությունը <strong style="color:#D4AF37;">{jobTitle}</strong> պաշտոնի համար և պատրաստում է պաշտոնական առաջարկը:',
            'Խնդրում ենք սպասել պաշտոնական առաջարկին — մեր թիմը կամ գործատուն շուտով կկապվեն ձեզ հետ: Առաջարկը կներառի փաթեթը, տեղափոխումը և մեկնարկի օրը:'
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
    brandedEmailHtml_(s.copy.en, 'en'),
    s.copy.ru.subject,
    brandedEmailHtml_(s.copy.ru, 'ru'),
    s.copy.hy.subject,
    brandedEmailHtml_(s.copy.hy, 'hy')
  ]);
}

const CTA_TEXT = {
  en: 'Browse All Open Roles',
  ru: 'Все открытые вакансии',
  hy: 'Բոլոր բաց դերերը'
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


// ──────────────────────────────────────────────────────────────────────
// TELEGRAM INTAKE BOT
//
// Setup (one-time, run installTelegramWebhook() after deployment):
//   1. Deploy this script as a Web App (Deploy → New deployment).
//      Execute as: Me. Who has access: Anyone (so Telegram can POST).
//   2. Run installTelegramWebhook() once — it tells Telegram to send all
//      bot updates to this web app.
//   3. Open DM with @Sonicbot_bot, type /start, then forward a CV.
//
// Note: the bot will only respond to TG_ADMIN_CHAT_ID. Group messages and
// other users are silently ignored.
//
// Drive OCR: the bot converts each PDF to a temporary Google Doc to read
// the text. This uses the Drive Advanced Service which is automatically
// available — no manual enable needed.
// ──────────────────────────────────────────────────────────────────────

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
  // Callback queries — inline-button taps
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
    '<b>ArgusRecruit Intake Bot</b>\n\nPick a job, choose how the candidate reached you, then send their CV — I\'ll read it, show what I found, and ask you to confirm.',
    [
      [{ text: '📋 Pick a job', callback_data: 'pickjob' }],
      [{ text: '🔍 State', callback_data: 'state' }, { text: '🆘 Help', callback_data: 'help' }]
    ]
  );
}

function showJobsKb_(chatId, messageId) {
  const jobs = listActiveJobs_();
  if (jobs.length === 0) {
    const text = 'No active job folders found in Drive root.';
    if (messageId) tgEdit_(chatId, messageId, text, [[{ text: '↩️ Back', callback_data: 'home' }]]);
    else tgSendKb_(chatId, text, [[{ text: '↩️ Back', callback_data: 'home' }]]);
    return;
  }
  const rows = jobs.map(j => [{
    text: `${j.jobId} — ${j.jobTitle}`,
    callback_data: 'job:' + j.jobId
  }]);
  rows.push([{ text: '↩️ Back', callback_data: 'home' }]);
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
  rows.push([{ text: '↩️ Back to jobs', callback_data: 'pickjob' }]);
  const text = `<b>Job:</b> <code>${jobId}</code> — ${escHtml_(jobTitle)}\n\nHow did you find this candidate?`;
  if (messageId) tgEdit_(chatId, messageId, text, rows);
  else tgSendKb_(chatId, text, rows);
}

function showReadyKb_(chatId, messageId, jobId, jobTitle, source) {
  const text =
    `✅ <b>Ready.</b>\n\n` +
    `Job: <code>${jobId}</code> — ${escHtml_(jobTitle)}\n` +
    `Source: <code>${source}</code>\n\n` +
    `Now <b>send or forward</b> the candidate's CV (PDF, DOC, DOCX).`;
  const kb = [[
    { text: '🔄 Change job', callback_data: 'pickjob' },
    { text: '↩️ Cancel', callback_data: 'home' }
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
        [{ text: '📋 Pick a job', callback_data: 'pickjob' }],
        [{ text: '🔍 State', callback_data: 'state' }, { text: '🆘 Help', callback_data: 'help' }]
      ]
    );
    return;
  }
  if (data === 'pickjob') { showJobsKb_(chatId, messageId); return; }
  if (data === 'help') {
    tgEdit_(chatId, messageId,
      '<b>Help</b>\n\n1. Tap <b>Pick a job</b> and choose the role.\n2. Choose how the candidate reached you.\n3. Send or forward the CV (PDF/DOC/DOCX).\n4. I\'ll read it, show what I found, and you can <b>Confirm</b> or <b>Cancel</b>.\n\nYou can also type <code>/start</code> anytime to get back to the home menu.',
      [[{ text: '↩️ Back', callback_data: 'home' }]]
    );
    return;
  }
  if (data === 'state') {
    const st = tgState_(chatId);
    tgEdit_(chatId, messageId,
      'Current state:\n' +
      'job: ' + (st.jobId || '—') + '\n' +
      'title: ' + (st.jobTitle || '—') + '\n' +
      'source: ' + (st.source || '—') + '\n' +
      'pending: ' + (st.pending ? 'yes' : 'no'),
      [[{ text: '↩️ Back', callback_data: 'home' }]]
    );
    return;
  }
  if (data.indexOf('job:') === 0) {
    const jobId = data.slice(4);
    const jobs = listActiveJobs_();
    const job = jobs.find(j => j.jobId === jobId);
    if (!job) { tgEdit_(chatId, messageId, 'Job not found.', [[{ text: '↩️ Back', callback_data: 'pickjob' }]]); return; }
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
    if (!job) { tgEdit_(chatId, messageId, 'Job not found.', [[{ text: '↩️ Back', callback_data: 'pickjob' }]]); return; }
    tgSaveState_(chatId, { jobId: job.jobId, jobTitle: job.jobTitle, source: source });
    showReadyKb_(chatId, messageId, job.jobId, job.jobTitle, source);
    return;
  }
  if (data === 'confirm') {
    const st = tgState_(chatId);
    if (!st.pending) { tgEdit_(chatId, messageId, 'Nothing pending. Pick a job and send a CV first.', [[{ text: '↩️ Back', callback_data: 'home' }]]); return; }
    let res;
    try { res = submitIntake_(st.pending); }
    catch (err) { res = { ok: false, error: String(err) }; }
    if (res.ok) {
      tgEdit_(chatId, messageId,
        '✅ Added <b>' + escHtml_(res.name) + '</b> → <code>' + res.jobId + '</code>.',
        [[
          { text: '➕ Another to same job', callback_data: 'again:' + st.jobId + '|' + (st.source || 'sourced-other') },
          { text: '🏠 Home', callback_data: 'home' }
        ]]
      );
    } else {
      tgEdit_(chatId, messageId,
        '⚠ Save failed: ' + escHtml_(res.error || 'unknown'),
        [[{ text: '↩️ Back', callback_data: 'home' }]]
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
      [[{ text: '🏠 Home', callback_data: 'home' }]]);
    return;
  }
  if (data.indexOf('again:') === 0) {
    const rest = data.slice(6);
    const sep = rest.indexOf('|');
    const jobId = rest.slice(0, sep);
    const source = rest.slice(sep + 1);
    const jobs = listActiveJobs_();
    const job = jobs.find(j => j.jobId === jobId);
    if (!job) { tgEdit_(chatId, messageId, 'Job not found.', [[{ text: '↩️ Back', callback_data: 'pickjob' }]]); return; }
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
    const list = jobs.map(j => `• <code>${j.jobId}</code> — ${j.jobTitle}`).join('\n');
    tgSend_(chatId, '<b>Active jobs:</b>\n' + list);
    return;
  }
  if (text === '/state') {
    const st = tgState_(chatId);
    tgSend_(chatId,
      'Current state:\n' +
      '<code>job</code>: ' + (st.jobId || '—') + '\n' +
      '<code>source</code>: ' + (st.source || '—') + '\n' +
      '<code>pending</code>: ' + (st.pending ? 'yes' : 'no')
    );
    return;
  }
  if (text === '/cancel') {
    tgClearState_(chatId);
    tgSend_(chatId, '✓ Cleared. Set a new job with <code>/job AR-XXX source</code>.');
    return;
  }
  if (text === '/confirm') {
    const st = tgState_(chatId);
    if (!st.pending) { tgSend_(chatId, 'Nothing pending. Send a CV first.'); return; }
    try {
      const res = submitIntake_(st.pending);
      if (res.ok) {
        tgSend_(chatId, '✅ Added <b>' + escHtml_(res.name) + '</b> → <code>' + res.jobId + '</code>.\nReady for next — send another CV or <code>/cancel</code>.');
      } else {
        tgSend_(chatId, '⚠ Save failed: ' + escHtml_(res.error || 'unknown'));
      }
    } catch (err) {
      tgSend_(chatId, '⚠ Error: ' + escHtml_(String(err)));
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
      tgSend_(chatId, '⚠ No job folder for <code>' + escHtml_(jobId) + '</code>. Use /jobs to see available IDs.');
      return;
    }
    tgSaveState_(chatId, { jobId: job.jobId, jobTitle: job.jobTitle, source: source });
    tgSend_(chatId,
      '✓ Set: <code>' + job.jobId + '</code> — ' + escHtml_(job.jobTitle) +
      '\nSource: <code>' + source + '</code>\n\nNow send/forward the CV.'
    );
    return;
  }
  tgSend_(chatId, 'I didn\'t understand. Send /help for commands.');
}

function handleTgDocument_(chatId, document, msg) {
  const state = tgState_(chatId);
  if (!state.jobId) {
    tgSend_(chatId, '⚠ Set a job first: <code>/job AR-XXX source</code>');
    return;
  }
  const filename = document.file_name || 'cv.pdf';
  if (!/\.(pdf|docx?|odt|rtf)$/i.test(filename)) {
    tgSend_(chatId, '⚠ Unsupported file type. Send PDF, DOC, or DOCX.');
    return;
  }
  tgSend_(chatId, '⏳ Reading <code>' + escHtml_(filename) + '</code>…');

  let cvText = '';
  let blob = null;
  try {
    blob = tgDownloadFile_(document.file_id);
    cvText = extractTextFromCv_(blob, filename);
  } catch (err) {
    tgSend_(chatId, '⚠ Could not read file: ' + escHtml_(String(err)));
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
    '👤 Name: <code>' + escHtml_(fields.name || '?') + '</code>\n' +
    '✉ Email: <code>' + escHtml_(fields.email || '—') + '</code>\n' +
    '📱 Phone: <code>' + escHtml_(fields.phone || '—') + '</code>\n' +
    '🔗 LinkedIn: <code>' + escHtml_(fields.linkedin || '—') + '</code>\n' +
    '📁 Job: <code>' + state.jobId + '</code> — ' + escHtml_(state.jobTitle) + '\n' +
    '📥 Source: <code>' + (state.source || 'sourced-other') + '</code>',
    [[
      { text: '✅ Confirm', callback_data: 'confirm' },
      { text: '🔄 Change job', callback_data: 'pickjob' },
      { text: '❌ Cancel', callback_data: 'cancel' }
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
  // For DOC/DOCX/ODT/RTF: same — Drive can convert them.
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
      throw new Error('Drive Advanced Service not available. Enable it in Apps Script: Services → Drive API.');
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
  // Phone — pick the longest "+? digit sequence with separators" candidate
  const phoneCandidates = text.match(/\+?\d[\d\s\-().]{7,18}\d/g) || [];
  const phone = phoneCandidates
    .map(s => s.replace(/[^\d+]/g, ''))
    .filter(s => s.replace(/\D/g, '').length >= 9)
    .sort((a, b) => b.length - a.length)[0] || '';
  // LinkedIn
  const liMatch = text.match(/(?:https?:\/\/)?(?:www\.)?linkedin\.com\/in\/[\w\-_%]+/i);
  const linkedin = liMatch ? ('https://' + liMatch[0].replace(/^https?:\/\//, '')) : '';
  // Name — heuristic: first non-empty line that's 2–4 words, mostly letters, not all upper
  let name = '';
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean).slice(0, 25);
  for (const line of lines) {
    if (line.length > 80) continue;
    if (/@|http|linkedin|cv|resume|curriculum/i.test(line)) continue;
    const words = line.split(/\s+/);
    if (words.length < 2 || words.length > 4) continue;
    if (!/^[A-Za-z'\-À-ſԱ-Ֆա-և\s]+$/.test(line)) continue;
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
