export async function onRequestPost(context) {
  const { request, env } = context;
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

    // Add to Resend Audience (if configured) so Mohammad can broadcast new jobs.
    if (env.RESEND_API_KEY && env.RESEND_AUDIENCE_ID) {
      await fetch(`https://api.resend.com/audiences/${env.RESEND_AUDIENCE_ID}/contacts`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${env.RESEND_API_KEY}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ email, unsubscribed: false })
      }).catch(() => {});
    }

    // Notify admin (so Mohammad knows someone subscribed)
    if (env.RESEND_API_KEY) {
      const from = env.MAIL_FROM || 'ArgusRecruit <contact@argusrecruit.com>';
      const adminTo = env.MAIL_TO || 'contact@argusrecruit.com';
      fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          from, to: [adminTo],
          subject: `[Job Alert] New subscriber: ${email}`,
          html: `<p>A new visitor subscribed to job alerts.</p><p><strong>Email:</strong> ${escapeHtml(email)}</p><p><strong>Language:</strong> ${escapeHtml(lang)}</p>`
        })
      }).catch(() => {});

      const subj = { en: 'You are subscribed to ArgusRecruit job alerts', ru: 'Вы подписаны на оповещения о вакансиях ArgusRecruit', hy: 'Դուք բաժանորդագրված եք ArgusRecruit-ի աշխատատեղերի ծանուցումներին' };
      const body = {
        en: `<p>Hi,</p><p>Thanks for subscribing to ArgusRecruit job alerts. We will email you when we open new senior or executive roles.</p><p>You can unsubscribe any time by replying to this email.</p><p>— The ArgusRecruit Team</p>`,
        ru: `<p>Здравствуйте,</p><p>Спасибо за подписку на оповещения о вакансиях ArgusRecruit. Мы будем писать вам, когда появятся новые старшие или руководящие роли.</p><p>Вы можете отписаться в любое время, ответив на это письмо.</p><p>— Команда ArgusRecruit</p>`,
        hy: `<p>Բարև,</p><p>Շնորհակալություն ArgusRecruit-ի աշխատատեղերի ծանուցումներին բաժանորդագրվելու համար: Մենք ձեզ կտեղեկացնենք, երբ բացենք նոր ավագ կամ ղեկավար դերեր:</p><p>Կարող եք ցանկացած ժամանակ ապաբաժանորդագրվել՝ պատասխանելով այս նամակին։</p><p>— ArgusRecruit-ի թիմը</p>`
      };
      fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ from, to: [email], subject: subj[lang] || subj.en, html: body[lang] || body.en })
      }).catch(() => {});
    }

    return json({ ok: true });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}
function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c])); }
