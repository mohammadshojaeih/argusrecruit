export async function onRequestPost(context) {
  const { request, env } = context;
  const debug = { env_keys: Object.keys(env), steps: [] };
  try {
    const ct = request.headers.get('content-type') || '';
    let email = '', lang = 'en';
    if (ct.includes('application/json')) {
      const body = await request.json();
      email = (body.email || '').toString().trim().slice(0, 200);
      lang = (body.lang || 'en').toString();
    } else {
      const form = await request.formData();
      email = (form.get('email') || '').toString().trim().slice(0, 200);
      lang = (form.get('lang') || 'en').toString();
    }

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json({ ok: false, error: 'Invalid email' }, 400);
    }

    const from = env.MAIL_FROM || 'ArgusRecruit <contact@argusrecruit.com>';
    const adminTo = env.MAIL_TO || 'contact@argusrecruit.com';

    // 1. Add to Resend Audience
    if (env.RESEND_API_KEY && env.RESEND_AUDIENCE_ID) {
      const audRes = await fetch(`https://api.resend.com/audiences/${env.RESEND_AUDIENCE_ID}/contacts`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, unsubscribed: false })
      });
      const audBody = await audRes.text();
      debug.steps.push({ step: 'audience', status: audRes.status, body: audBody.slice(0, 400) });
    } else {
      debug.steps.push({ step: 'audience', skipped: true, has_key: !!env.RESEND_API_KEY, has_aud: !!env.RESEND_AUDIENCE_ID });
    }

    // 2. Notify admin
    if (env.RESEND_API_KEY) {
      const adminRes = await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          from, to: [adminTo],
          subject: `[Job Alert] New subscriber: ${email}`,
          html: `<p>A new visitor subscribed to job alerts.</p><p><strong>Email:</strong> ${escapeHtml(email)}</p><p><strong>Language:</strong> ${escapeHtml(lang)}</p>`
        })
      });
      const adminBody = await adminRes.text();
      debug.steps.push({ step: 'admin_email', status: adminRes.status, body: adminBody.slice(0, 400) });
    }

    // 3. Confirmation email to subscriber
    if (env.RESEND_API_KEY) {
      const subj = { en: 'You are subscribed to ArgusRecruit job alerts', ru: 'Вы подписаны на оповещения ArgusRecruit', hy: 'Դուք բաժանորդագրված եք' };
      const bodyHtml = {
        en: `<p>Hi,</p><p>Thanks for subscribing to ArgusRecruit job alerts.</p><p>— The ArgusRecruit Team</p>`,
        ru: `<p>Здравствуйте,</p><p>Спасибо за подписку.</p><p>— Команда ArgusRecruit</p>`,
        hy: `<p>Բարև,</p><p>Շնորհակալություն բաժանորդագրվելու համար:</p><p>— ArgusRecruit-ի թիմը</p>`
      };
      const userRes = await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ from, to: [email], subject: subj[lang] || subj.en, html: bodyHtml[lang] || bodyHtml.en })
      });
      const userBody = await userRes.text();
      debug.steps.push({ step: 'user_email', status: userRes.status, body: userBody.slice(0, 400) });
    }

    return json({ ok: true, debug });
  } catch (e) {
    return json({ ok: false, error: e.message, stack: e.stack, debug }, 500);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}
function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c])); }
