export async function onRequestPost(context) {
  const { request, env } = context;
  try {
    const form = await request.formData();
    const hp = (form.get('hp_field') || '').toString();
    if (hp) return json({ ok: true });

    const name = (form.get('name') || '').toString().slice(0, 200);
    const email = (form.get('email') || '').toString().slice(0, 200);
    const phone = (form.get('phone') || '').toString().slice(0, 100);
    const linkedin = (form.get('linkedin') || '').toString().slice(0, 300);
    const coverNote = (form.get('coverNote') || '').toString().slice(0, 5000);
    const jobSlug = (form.get('jobSlug') || '').toString().slice(0, 200);
    const jobTitle = (form.get('jobTitle') || '').toString().slice(0, 300);
    const lang = (form.get('lang') || 'en').toString();

    if (!name || !email || !jobTitle) {
      return json({ ok: false, error: 'Missing required fields' }, 400);
    }

    const attachments = [];
    const cv = form.get('cv');
    if (cv && typeof cv === 'object' && cv.size > 0) {
      if (cv.size > 10 * 1024 * 1024) {
        return json({ ok: false, error: 'CV file too large (max 10MB)' }, 400);
      }
      const buf = await cv.arrayBuffer();
      attachments.push({ filename: cv.name || 'cv', content: arrayBufferToBase64(buf) });
    } else {
      return json({ ok: false, error: 'CV is required' }, 400);
    }

    // Structured payload for Apps Script ATS to parse reliably (kept hidden in the email).
    const atsPayload = JSON.stringify({
      jobId: jobSlug,
      jobTitle,
      name,
      email,
      phone,
      linkedin,
      lang,
      source: 'web-apply',
      submittedAt: new Date().toISOString()
    });

    const adminHtml = `
      <h2>New Application — ${esc(jobTitle)}</h2>
      <p><strong>Name:</strong> ${esc(name)}</p>
      <p><strong>Email:</strong> ${esc(email)}</p>
      <p><strong>Phone:</strong> ${esc(phone) || '—'}</p>
      <p><strong>LinkedIn:</strong> ${linkedin ? `<a href="${esc(linkedin)}">${esc(linkedin)}</a>` : '—'}</p>
      <p><strong>Role:</strong> ${esc(jobTitle)} (jobId: ${esc(jobSlug)})</p>
      <p><strong>Language:</strong> ${esc(lang)}</p>
      <p><strong>Cover Note:</strong></p>
      <p style="white-space:pre-wrap; background:#f5f5f5; padding:12px; border-radius:6px;">${esc(coverNote) || '—'}</p>
      <hr>
      <p style="color:#888; font-size:12px;">CV attached.</p>
      <!--ATS_PAYLOAD_START${atsPayload}ATS_PAYLOAD_END-->
    `;

    const subjects = {
      en: `Application received — ${jobTitle}`,
      ru: `Заявка получена — ${jobTitle}`,
      hy: `Դիմումը ստացվեց — ${jobTitle}`
    };

    function applicantHtml(lang, name, jobTitle) {
      const COPY = {
        en: {
          eyebrow: '• Application Received •',
          h1: 'Thank you for applying.',
          greeting: `Hi ${esc(name)},`,
          intro: `Thank you for applying to <strong style="color:#D4AF37;">${esc(jobTitle)}</strong> via ArgusRecruit.`,
          p1: 'We have received your CV and a senior recruiter on our team will review it personally. If your background matches the role, we will get in touch within 1–3 business days to schedule a confidential conversation.',
          p2: 'If we do not have a match for this particular role, your details will remain in our network — and we may reach out about future openings that better fit your background.',
          cta: 'Browse All Open Roles',
          team: 'The ArgusRecruit Team',
          footer: 'You received this because you applied to a role on argusrecruit.com.',
          rights: '© 2026 ArgusRecruit · Yerevan, Armenia'
        },
        ru: {
          eyebrow: '• Заявка получена •',
          h1: 'Спасибо за вашу заявку.',
          greeting: `Здравствуйте, ${esc(name)},`,
          intro: `Спасибо за заявку на роль <strong style="color:#D4AF37;">${esc(jobTitle)}</strong> через ArgusRecruit.`,
          p1: 'Мы получили ваше резюме, и старший рекрутер из нашей команды лично его рассмотрит. Если ваш опыт подходит для этой роли, мы свяжемся с вами в течение 1–3 рабочих дней для конфиденциального разговора.',
          p2: 'Если для этой конкретной роли совпадения не будет, ваши данные останутся в нашей сети — и мы можем связаться с вами по поводу будущих возможностей.',
          cta: 'Все открытые вакансии',
          team: 'Команда ArgusRecruit',
          footer: 'Вы получили это письмо, потому что подали заявку на argusrecruit.com.',
          rights: '© 2026 ArgusRecruit · Ереван, Армения'
        },
        hy: {
          eyebrow: '• Դիմումը ստացվեց •',
          h1: 'Շնորհակալություն դիմելու համար.',
          greeting: `Բարև, ${esc(name)},`,
          intro: `Շնորհակալություն <strong style="color:#D4AF37;">${esc(jobTitle)}</strong> պաշտոնի համար ArgusRecruit-ի միջոցով դիմելու համար:`,
          p1: 'Մենք ստացել ենք ձեր CV-ն, և մեր թիմի ավագ ռեկրուտերն այն անձամբ կքննարկի: Եթե ձեր փորձառությունը համապատասխանում է դերին, մենք կկապվենք ձեզ հետ 1–3 աշխատանքային օրվա ընթացքում՝ գաղտնի զրույց պլանավորելու համար:',
          p2: 'Եթե այս կոնկրետ դերի համար համապատասխանություն չլինի, ձեր տվյալները կմնան մեր ցանցում, և մենք կարող ենք կապվել ձեզ հետ ապագա հնարավորությունների վերաբերյալ:',
          cta: 'Բոլոր բաց դերերը',
          team: 'ArgusRecruit-ի թիմը',
          footer: 'Դուք ստացել եք այս նամակը, քանի որ դիմել եք argusrecruit.com-ում:',
          rights: '© 2026 ArgusRecruit · Երևան, Հայաստան'
        }
      };
      const C = COPY[lang] || COPY.en;
      const jobsPath = lang === 'en' ? '/jobs/' : `/${lang}/jobs/`;
      return `<!DOCTYPE html>
<html lang="${lang}"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ArgusRecruit</title></head>
<body style="margin:0;padding:0;background:#0E2440;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;color:#ffffff;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0E2440;padding:32px 16px;">
    <tr><td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#1E4170;border:1px solid rgba(212,175,55,0.2);border-radius:14px;overflow:hidden;">
        <tr><td align="center" style="padding:36px 32px 24px;background:#16345A;border-bottom:1px solid rgba(212,175,55,0.2);">
          <img src="https://argusrecruit.com/logo.png" alt="ArgusRecruit" width="64" style="display:block;height:64px;width:auto;">
          <div style="margin-top:14px;font-size:11px;letter-spacing:2.5px;color:#D4AF37;text-transform:uppercase;font-weight:700;">Many Eyes. One Purpose.</div>
        </td></tr>
        <tr><td align="center" style="padding:32px 32px 0;">
          <div style="display:inline-block;font-size:10px;letter-spacing:4px;color:#D4AF37;text-transform:uppercase;font-weight:700;padding:6px 16px;border:1px solid rgba(212,175,55,0.4);border-radius:999px;">${C.eyebrow}</div>
        </td></tr>
        <tr><td align="center" style="padding:18px 32px 0;">
          <h1 style="margin:0;font-size:24px;font-weight:700;letter-spacing:0.5px;color:#ffffff;text-transform:uppercase;line-height:1.2;">${C.h1}</h1>
        </td></tr>
        <tr><td style="padding:30px 36px 18px;color:rgba(255,255,255,0.85);font-size:15px;line-height:1.75;">
          <p style="margin:0 0 14px;">${C.greeting}</p>
          <p style="margin:0 0 14px;">${C.intro}</p>
          <p style="margin:0 0 14px;color:rgba(255,255,255,0.78);">${C.p1}</p>
          <p style="margin:0;color:rgba(255,255,255,0.78);">${C.p2}</p>
        </td></tr>
        <tr><td align="center" style="padding:18px 32px 36px;">
          <a href="https://argusrecruit.com${jobsPath}" style="display:inline-block;background:#D4AF37;color:#0E2440;font-weight:700;font-size:13px;letter-spacing:1px;text-transform:uppercase;padding:13px 28px;border-radius:999px;text-decoration:none;">${C.cta}</a>
        </td></tr>
        <tr><td style="padding:0 32px 30px;color:rgba(255,255,255,0.7);font-size:13px;line-height:1.6;text-align:center;font-style:italic;">— ${C.team}</td></tr>
        <tr><td style="padding:22px 32px;background:#0E2440;border-top:1px solid rgba(212,175,55,0.15);text-align:center;">
          <div style="color:rgba(255,255,255,0.55);font-size:12px;line-height:1.65;">
            ${C.footer}<br>
            <a href="https://argusrecruit.com" style="color:#D4AF37;text-decoration:none;">argusrecruit.com</a>
            &nbsp;·&nbsp;
            <a href="mailto:contact@argusrecruit.com" style="color:#D4AF37;text-decoration:none;">contact@argusrecruit.com</a>
          </div>
          <div style="margin-top:12px;color:rgba(255,255,255,0.4);font-size:11px;letter-spacing:1px;text-transform:uppercase;">${C.rights}</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
    }

    const bodies = {
      en: applicantHtml('en', name, jobTitle),
      ru: applicantHtml('ru', name, jobTitle),
      hy: applicantHtml('hy', name, jobTitle)
    };

    const from = env.MAIL_FROM || 'ArgusRecruit <contact@argusrecruit.com>';
    const adminTo = env.MAIL_TO || 'contact@argusrecruit.com';

    // Send to admin (with CV attachment)
    const adminRes = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${env.RESEND_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        from,
        to: [adminTo],
        reply_to: email,
        subject: `[Application ${jobSlug || 'AR-UNKNOWN'}] ${jobTitle} — ${name}`,
        html: adminHtml,
        attachments
      })
    });

    if (!adminRes.ok) {
      const err = await adminRes.text();
      return json({ ok: false, error: 'Email send failed', detail: err }, 500);
    }

    // Send auto-confirmation to applicant
    await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${env.RESEND_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        from,
        to: [email],
        subject: subjects[lang] || subjects.en,
        html: bodies[lang] || bodies.en
      })
    }).catch(() => {});

    return json({ ok: true });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}
