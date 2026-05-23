// Sanity → Telegram bridge for LinkedIn post drafts.
// When a new active EN job is published, generates two ready-to-post drafts
// and DMs them to the team's Telegram so they can be pasted on LinkedIn:
//   1) EN — formal post for the ArgusRecruit company page
//   2) FA — short, casual post for the team member's personal page (he'll repost the company post)

export async function onRequestPost(context) {
  const { request, env } = context;
  const url = new URL(request.url);

  const provided = url.searchParams.get('secret') || request.headers.get('x-sanity-secret') || '';
  if (!env.BROADCAST_SECRET || provided !== env.BROADCAST_SECRET) {
    return json({ ok: false, error: 'Unauthorized' }, 401);
  }
  if (!env.TELEGRAM_BOT_TOKEN) return json({ ok: false, error: 'TELEGRAM_BOT_TOKEN missing' }, 500);

  const chatId = env.TELEGRAM_CHAT_ID || '814437645';

  try {
    const payload = await request.json();
    const doc = payload._type ? payload : payload.document || payload;

    if (doc._type !== 'job') return json({ ok: true, skipped: 'not a job' });
    if (doc.status && doc.status !== 'active') return json({ ok: true, skipped: 'job not active' });
    if (doc.language && doc.language !== 'en') return json({ ok: true, skipped: 'non-en variant' });

    const enPost = buildEnPost(doc);
    const faPost = buildFaPost(doc);

    const sendEn = await tgSend(env.TELEGRAM_BOT_TOKEN, chatId,
      `📌 *LinkedIn — Company Page (EN)* — ready to paste:\n\n${enPost}`);
    const sendFa = await tgSend(env.TELEGRAM_BOT_TOKEN, chatId,
      `📌 *LinkedIn — پست شخصی (FA)* — آماده copy/paste:\n\n${faPost}`);

    return json({ ok: true, en: sendEn.ok, fa: sendFa.ok });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

async function tgSend(token, chatId, text) {
  const res = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: 'Markdown', disable_web_page_preview: true })
  });
  return res.json();
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}

function buildEnPost(doc) {
  const title = doc.title || 'New role';
  const jobId = doc.jobId || doc.slug?.current || doc.slug || '';
  const cityCountry = [doc.locationCity, doc.locationCountry].filter(Boolean).join(', ');
  const workMode = ({ onsite: 'Onsite', remote: 'Remote', hybrid: 'Hybrid' })[doc.workplaceType] || '';
  const seniorityFromTitle = /senior|lead|principal|head|chief|director|vp|cxo|ceo|cto|cfo|cmo/i.test(title) ? 'Senior' : '';
  const employmentType = ({ FULL_TIME: 'Full-time', PART_TIME: 'Part-time', CONTRACTOR: 'Contract', TEMPORARY: 'Temporary', INTERN: 'Internship' })[doc.employmentType] || 'Full-time';

  const headerBits = [cityCountry, workMode].filter(Boolean);
  if (doc.visaSponsorship || doc.relocationAssistance) headerBits.push('Visa & Relocation Support');
  const headline = `🚨 WE'RE HIRING — ${title} (${headerBits.join(' · ')})`;

  const metaLine = [
    cityCountry ? `📍 ${cityCountry}` : '',
    workMode ? `🏢 ${workMode}` : '',
    seniorityFromTitle ? `💼 ${seniorityFromTitle}` : '',
    employmentType
  ].filter(Boolean).join(' · ');

  const visaLine = (doc.visaSponsorship || doc.relocationAssistance)
    ? `🌍 ${doc.visaSponsorship && doc.relocationAssistance ? 'Visa sponsorship and relocation support available' : doc.visaSponsorship ? 'Visa sponsorship available' : 'Relocation support available'} for international candidates`
    : '';

  const languages = (doc.languagesRequired || []).join(', ');
  const langLine = languages ? `🌐 Languages: ${languages}` : '';

  const intro = doc.excerpt
    ? doc.excerpt
    : `We're recruiting on behalf of an international client on a confidential basis. They're looking for a ${title} to join their team.`;

  const responsibilities = (doc.responsibilities || []).slice(0, 5);
  const requirements = (doc.requirements || []).slice(0, 6);

  const respBlock = responsibilities.length
    ? `*What you'll do:*\n${responsibilities.map(r => `• ${r}`).join('\n')}\n\n`
    : '';
  const reqBlock = requirements.length
    ? `*What we're looking for:*\n${requirements.map(r => `• ${r}`).join('\n')}\n\n`
    : '';

  const applyUrl = jobId ? `https://argusrecruit.com/jobs/${jobId}/` : 'https://argusrecruit.com/jobs/';

  const tagPool = ['ExecutiveSearch', 'Hiring'];
  if (cityCountry) tagPool.push(`${(doc.locationCity || '').replace(/\s+/g, '')}Jobs`);
  if (doc.visaSponsorship) tagPool.push('VisaSponsorship');
  if (doc.relocationAssistance) tagPool.push('Relocation');
  (doc.tags || []).forEach(t => { const c = String(t).replace(/[^A-Za-z0-9]/g, ''); if (c) tagPool.push(c); });
  const hashtags = [...new Set(tagPool)].slice(0, 9).map(t => '#' + t).join(' ');

  return [
    headline,
    '',
    'ArgusRecruit · Many Eyes. One Purpose.',
    '',
    intro,
    '',
    metaLine,
    visaLine,
    langLine,
    '',
    respBlock + reqBlock + '— Where ambitious talents meet exceptional employers.',
    '',
    `👉 Apply confidentially: ${applyUrl}`,
    '',
    hashtags
  ].filter(Boolean).join('\n').replace(/\n{3,}/g, '\n\n');
}

function buildFaPost(doc) {
  const title = doc.title || 'یه نقش جدید';
  const city = doc.locationCity || '';
  const workMode = ({ onsite: 'حضوری', remote: 'دورکار', hybrid: 'هیبرید' })[doc.workplaceType] || '';
  const visa = doc.visaSponsorship ? ' · ویزا اسپانسرشیپ دارد' : '';
  const reloc = doc.relocationAssistance ? ' · کمک ریلوکیشن' : '';

  const headline = `🚨 یه فرصت شغلی جالب: ${title}${city ? ` در ${city}` : ''}`;
  const lead = `با ArgusRecruit داریم برای یه شرکت بین‌المللی روی این پوزیشن کار می‌کنیم. ${workMode ? `کار به‌صورت ${workMode}` : 'کار'}${visa}${reloc}.`;
  const why = `اگه دنبال یه چالش جدی توی محیط فنی و بین‌المللی هستی، این یکی ارزش نگاه کردن داره.`;
  const cta = `👇 جزئیات کامل توی پست شرکت — لینک رو دنبال کن.`;
  const tags = '#ArgusRecruit #Hiring #ExecutiveSearch';

  return [headline, '', lead, '', why, '', cta, '', tags].join('\n');
}
