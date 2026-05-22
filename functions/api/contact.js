export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const form = await request.formData();
    const name = (form.get('name') || '').toString().slice(0, 200);
    const email = (form.get('email') || '').toString().slice(0, 200);
    const company = (form.get('company') || '').toString().slice(0, 200);
    const intent = (form.get('intent') || '').toString().slice(0, 200);
    const message = (form.get('message') || '').toString().slice(0, 5000);
    const botcheck = (form.get('botcheck') || '').toString();

    if (botcheck) return json({ ok: true });
    if (!name || !email || !message) {
      return json({ ok: false, error: 'Missing required fields' }, 400);
    }

    const attachments = [];
    const file = form.get('attachment');
    if (file && typeof file === 'object' && file.size > 0) {
      if (file.size > 10 * 1024 * 1024) {
        return json({ ok: false, error: 'Attachment too large (max 10MB)' }, 400);
      }
      const buf = await file.arrayBuffer();
      const b64 = arrayBufferToBase64(buf);
      attachments.push({ filename: file.name || 'attachment', content: b64 });
    }

    const html = `
      <h2>New ArgusRecruit Inquiry</h2>
      <p><strong>Name:</strong> ${escape(name)}</p>
      <p><strong>Email:</strong> ${escape(email)}</p>
      <p><strong>Company:</strong> ${escape(company)}</p>
      <p><strong>Intent:</strong> ${escape(intent)}</p>
      <p><strong>Message:</strong></p>
      <p style="white-space:pre-wrap">${escape(message)}</p>
    `;

    const payload = {
      from: env.MAIL_FROM || 'ArgusRecruit <contact@argusrecruit.com>',
      to: [env.MAIL_TO || 'contact@argusrecruit.com'],
      reply_to: email,
      subject: `New inquiry from ${name}`,
      html,
      attachments
    };

    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${env.RESEND_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    if (!res.ok) {
      const err = await res.text();
      return json({ ok: false, error: 'Email send failed', detail: err }, 500);
    }
    return json({ ok: true });
  } catch (e) {
    return json({ ok: false, error: e.message }, 500);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

function escape(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
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
