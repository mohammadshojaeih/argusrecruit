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

  // 1. New format — structured JSON block
  const payloadMatch = html.match(/ATS_PAYLOAD_START([\s\S]*?)ATS_PAYLOAD_END/);
  if (payloadMatch) {
    try { return JSON.parse(payloadMatch[1].trim()); } catch (_) { /* fall through */ }
  }

  // 2. New subject format [Application <jobId>] <title> — <name>
  const subj = msg.getSubject();
  const newSubj = subj.match(/^\s*\[Application\s+([^\]\s]+)\]\s+(.+?)\s+[—\-–]\s+(.+?)\s*$/);

  // 3. Old subject format [Application] <title> — <name>  (legacy)
  const oldSubj = subj.match(/^\s*\[Application\]\s+(.+?)\s+[—\-–]\s+(.+?)\s*$/);

  if (!newSubj && !oldSubj) return null;

  // Helper: pull a field's value from the HTML body of the admin email
  const fieldRx = (label) => {
    const re = new RegExp(
      '<strong>\\s*' + label + '\\s*:\\s*</strong>([\\s\\S]*?)</p>',
      'i'
    );
    const m = html.match(re);
    if (!m) return '';
    // Strip remaining HTML and the leading dash/em-dash markers used for empty values.
    return m[1].replace(/<a[^>]*>([\s\S]*?)<\/a>/g, '$1')
               .replace(/<[^>]+>/g, '')
               .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
               .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
               .trim()
               .replace(/^—\s*$/, '');
  };

  // Pull the slug/jobId out of the Role line: "Senior QA Engineer (slug: AR-STE01)"
  // or "Senior QA Engineer (jobId: AR-STE01)" depending on apply.js version.
  let jobId    = newSubj ? newSubj[1].trim() : '';
  let jobTitle = newSubj ? newSubj[2].trim() : (oldSubj ? oldSubj[1].trim() : '');
  let name     = newSubj ? newSubj[3].trim() : (oldSubj ? oldSubj[2].trim() : '');

  const roleField = fieldRx('Role');
  if (!jobId && roleField) {
    const slugMatch = roleField.match(/\((?:slug|jobId)\s*:\s*([^)]+)\)/i);
    if (slugMatch) jobId = slugMatch[1].trim();
  }

  // Fallbacks: derive jobId from jobTitle for the most common titles we know.
  if (!jobId) {
    jobId = inferJobIdFromTitle_(jobTitle);
  }

  const email    = fieldRx('Email') || msg.getReplyTo() || msg.getFrom() || '';
  const phone    = fieldRx('Phone');
  const linkedin = fieldRx('LinkedIn');
  const lang     = (fieldRx('Language') || 'en').toLowerCase();

  return {
    jobId,
    jobTitle,
    name,
    email,
    phone,
    linkedin,
    lang,
    source: 'web-apply',
    submittedAt: msg.getDate().toISOString()
  };
}

// Quick lookup for legacy emails that didn't carry an explicit jobId.
function inferJobIdFromTitle_(title) {
  if (!title) return '';
  const map = {
    'senior qa engineer':       'AR-STE01',
    'senior test engineer':     'AR-STE01',
    'senior python developer':  'AR-PYT01',
    'senior angular developer': 'AR-ANG01',
    'senior product designer':  'AR-DES01'
  };
  return map[title.toLowerCase().trim()] || 'AR-UNKNOWN';
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
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STAGES, true)
    .setAllowInvalid(false)
    .build();
  app.getRange(2, COL.stage, 1000, 1).setDataValidation(rule);

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
  // (Body cells should hold HTML.)
  const rowsTpl = [
    ['reviewed',
      'Your application is being reviewed — {jobTitle}',
      '<p>Hi {name},</p><p>Thank you for applying to <strong>{jobTitle}</strong> at ArgusRecruit. A senior recruiter is now reviewing your application personally.</p><p>If your background matches the role, we will be in touch within 1–3 business days. If not, we will still keep your details in our network for future opportunities that may fit you better.</p><p>Warm regards,<br>The ArgusRecruit Team</p>',
      'Ваша заявка на рассмотрении — {jobTitle}',
      '<p>Здравствуйте, {name},</p><p>Спасибо за вашу заявку на роль <strong>{jobTitle}</strong>. Старший рекрутер лично рассматривает ваш профиль. Если ваш опыт подходит, мы свяжемся с вами в течение 1–3 рабочих дней.</p><p>С уважением,<br>Команда ArgusRecruit</p>',
      'Ձեր դիմումը քննարկվում է — {jobTitle}',
      '<p>Բարև, {name},</p><p>Շնորհակալություն <strong>{jobTitle}</strong> պաշտոնի համար դիմելու համար: Մեր ավագ ռեկրուտերն այժմ անձամբ քննարկում է ձեր փորձառությունը: Համապատասխան լինելու դեպքում մենք կկապվենք ձեզ հետ 1–3 աշխատանքային օրվա ընթացքում:</p><p>Հարգանքով,<br>ArgusRecruit-ի թիմը</p>'
    ],
    ['shortlist',
      'You have been shortlisted — {jobTitle}',
      '<p>Hi {name},</p><p>Great news — you have been shortlisted for the <strong>{jobTitle}</strong> role. Our client will now review your profile, and we will get back to you with next steps within a few business days.</p><p>Warm regards,<br>ArgusRecruit</p>',
      'Вы в шорт-листе — {jobTitle}',
      '<p>Здравствуйте, {name},</p><p>Хорошие новости — вы вошли в шорт-лист по позиции <strong>{jobTitle}</strong>. Клиент сейчас рассматривает ваш профиль, и мы свяжемся с вами в течение нескольких рабочих дней.</p><p>С уважением,<br>ArgusRecruit</p>',
      'Դուք ընտրվել եք կարճ ցուցակում — {jobTitle}',
      '<p>Բարև, {name},</p><p>Հաճելի լուր — դուք ընտրվել եք <strong>{jobTitle}</strong> պաշտոնի կարճ ցուցակում: Մեր հաճախորդն այժմ ուսումնասիրում է ձեր պրոֆիլը, և մենք կտեղեկացնենք ձեզ հաջորդ քայլերի մասին մի քանի աշխատանքային օրվա ընթացքում:</p><p>Հարգանքով,<br>ArgusRecruit</p>'
    ],
    ['interview',
      'Congratulations — interview confirmed for {jobTitle}',
      '<p>Hi {name},</p><p>Congratulations — your interview has been confirmed for the <strong>{jobTitle}</strong> role. We will follow up shortly with the exact time, format, and the people you will meet.</p><p>Warm regards,<br>ArgusRecruit</p>',
      'Поздравляем — собеседование подтверждено для {jobTitle}',
      '<p>Здравствуйте, {name},</p><p>Поздравляем — ваше собеседование подтверждено для роли <strong>{jobTitle}</strong>. Скоро мы пришлём вам точное время, формат и список участников.</p><p>С уважением,<br>ArgusRecruit</p>',
      'Շնորհավորում ենք — հարցազրույցը հաստատվել է {jobTitle} պաշտոնի համար',
      '<p>Բարև, {name},</p><p>Շնորհավորում ենք — ձեր հարցազրույցը հաստատվել է <strong>{jobTitle}</strong> պաշտոնի համար: Շուտով կուղարկենք ձեզ ժամանակի, ձևաչափի և մասնակիցների ճշգրիտ մանրամասները:</p><p>Հարգանքով,<br>ArgusRecruit</p>'
    ],
    ['offer',
      'An offer is coming your way — {jobTitle}',
      '<p>Hi {name},</p><p>Excellent news — our client is preparing a formal offer for the <strong>{jobTitle}</strong> role. We will reach out in the next 24 hours to walk you through the package and answer any questions.</p><p>Warm regards,<br>ArgusRecruit</p>',
      'Вам направляется оффер — {jobTitle}',
      '<p>Здравствуйте, {name},</p><p>Отличные новости — клиент готовит официальный оффер по роли <strong>{jobTitle}</strong>. В ближайшие 24 часа мы свяжемся, чтобы обсудить пакет и ответить на ваши вопросы.</p><p>С уважением,<br>ArgusRecruit</p>',
      'Ձեզ ուղարկվում է առաջարկ — {jobTitle}',
      '<p>Բարև, {name},</p><p>Գերազանց լուր — մեր հաճախորդը պատրաստում է պաշտոնական առաջարկ <strong>{jobTitle}</strong> պաշտոնի համար: Հաջորդ 24 ժամվա ընթացքում մենք կկապվենք ձեզ հետ՝ փաթեթը քննարկելու և ձեր հարցերին պատասխանելու համար:</p><p>Հարգանքով,<br>ArgusRecruit</p>'
    ],
    ['rejected',
      'Update on your {jobTitle} application',
      '<p>Hi {name},</p><p>Thank you again for your interest in the <strong>{jobTitle}</strong> role and for the time you put into your application. After careful consideration, we will not be moving forward with your candidacy for this specific role.</p><p>This was not an easy decision. We will keep your profile in our network and may reach out about future roles that better match your background.</p><p>Wishing you the very best,<br>ArgusRecruit</p>',
      'Обновление по вашей заявке на {jobTitle}',
      '<p>Здравствуйте, {name},</p><p>Спасибо за интерес к роли <strong>{jobTitle}</strong> и за время, потраченное на заявку. После тщательного рассмотрения мы решили не продолжать рассмотрение вашей кандидатуры по этой конкретной позиции.</p><p>Это было непростое решение. Мы сохраним ваш профиль и можем связаться с вами по поводу будущих ролей, лучше подходящих вам.</p><p>Желаем всего наилучшего,<br>ArgusRecruit</p>',
      'Թարմացում ձեր {jobTitle} դիմումի վերաբերյալ',
      '<p>Բարև, {name},</p><p>Կրկին շնորհակալություն <strong>{jobTitle}</strong> պաշտոնի նկատմամբ ձեր հետաքրքրության և դիմումի վրա ծախսված ժամանակի համար: Մանրակրկիտ քննարկումից հետո մենք որոշել ենք չշարունակել ձեր թեկնածությունը այս կոնկրետ դերի համար:</p><p>Սա հեշտ որոշում չէր: Մենք կպահպանենք ձեր պրոֆիլը և կարող ենք կապվել ձեզ հետ ձեր փորձառությանն ավելի համապատասխան ապագա դերերի վերաբերյալ:</p><p>Բարեմաղթանքներով,<br>ArgusRecruit</p>'
    ]
  ];
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
    .addToUi();
}
