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

    const adminHtml = `
      <h2>New Application — ${esc(jobTitle)}</h2>
      <p><strong>Name:</strong> ${esc(name)}</p>
      <p><strong>Email:</strong> ${esc(email)}</p>
      <p><strong>Phone:</strong> ${esc(phone) || '—'}</p>
      <p><strong>LinkedIn:</strong> ${linkedin ? `<a href="${esc(linkedin)}">${esc(linkedin)}</a>` : '—'}</p>
      <p><strong>Role:</strong> ${esc(jobTitle)} (slug: ${esc(jobSlug)})</p>
      <p><strong>Language:</strong> ${esc(lang)}</p>
      <p><strong>Cover Note:</strong></p>
      <p style="white-space:pre-wrap; background:#f5f5f5; padding:12px; border-radius:6px;">${esc(coverNote) || '—'}</p>
      <hr>
      <p style="color:#888; font-size:12px;">CV attached.</p>
    `;

    const greetings = { en: 'Hi', ru: 'Здравствуйте,', hy: 'Բարև,' };
    const subjects = {
      en: `Application received — ${jobTitle}`,
      ru: `Заявка получена — ${jobTitle}`,
      hy: `Դիմումը ստացվեց — ${jobTitle}`
    };
    const bodies = {
      en: `<p>${greetings.en} ${esc(name)},</p>
        <p>Thank you for applying to <strong>${esc(jobTitle)}</strong> via ArgusRecruit.</p>
        <p>We have received your CV and a senior recruiter on our team will review it personally. If your background matches the role, we will get in touch within one business day to schedule a confidential conversation.</p>
        <p>If we do not have a match for this particular role, your details will remain in our network — and we may reach out about future openings that better fit your background.</p>
        <p style="margin-top:24px;">— The ArgusRecruit Team<br>
        <a href="https://argusrecruit.com">argusrecruit.com</a></p>`,
      ru: `<p>${greetings.ru} ${esc(name)},</p>
        <p>Спасибо за вашу заявку на роль <strong>${esc(jobTitle)}</strong> через ArgusRecruit.</p>
        <p>Мы получили ваше резюме, и старший рекрутер из нашей команды лично его рассмотрит. Если ваш опыт подходит для этой роли, мы свяжемся с вами в течение одного рабочего дня для конфиденциального разговора.</p>
        <p>Если для этой конкретной роли совпадения не будет, ваши данные останутся в нашей сети — и мы можем связаться с вами по поводу будущих возможностей.</p>
        <p style="margin-top:24px;">— Команда ArgusRecruit<br>
        <a href="https://argusrecruit.com">argusrecruit.com</a></p>`,
      hy: `<p>${greetings.hy} ${esc(name)},</p>
        <p>Շնորհակալություն <strong>${esc(jobTitle)}</strong> պաշտոնի համար ArgusRecruit-ի միջոցով դիմելու համար:</p>
        <p>Մենք ստացել ենք ձեր CV-ն, և մեր թիմի ավագ ռեկրուտերն այն անձամբ կքննարկի։ Եթե ձեր փորձառությունը համապատասխանում է դերին, մենք կկապվենք ձեզ հետ մեկ աշխատանքային օրվա ընթացքում՝ գաղտնի զրույց պլանավորելու համար:</p>
        <p>Եթե այս կոնկրետ դերի համար համապատասխանություն չլինի, ձեր տվյալները կմնան մեր ցանցում, և մենք կարող ենք կապվել ձեզ հետ ապագա հնարավորությունների վերաբերյալ:</p>
        <p style="margin-top:24px;">— ArgusRecruit-ի թիմը<br>
        <a href="https://argusrecruit.com">argusrecruit.com</a></p>`
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
        subject: `[Application] ${jobTitle} — ${name}`,
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
