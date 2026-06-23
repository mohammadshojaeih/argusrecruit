// Sanity → Resend Broadcast bridge.
// Triggered by Sanity webhook when a new Job is published.
// Builds an HTML email from the job and sends a Resend Broadcast to the Job Alerts audience.

export async function onRequestPost(context) {
  const { request, env } = context;
  const url = new URL(request.url);

  // 1. Authenticate via shared secret (?secret=... or X-Sanity-Secret header)
  const provided = url.searchParams.get('secret') || request.headers.get('x-sanity-secret') || '';
  if (!env.BROADCAST_SECRET || provided !== env.BROADCAST_SECRET) {
    return json({ ok: false, error: 'Unauthorized' }, 401);
  }

  try {
    const payload = await request.json();
    // Sanity webhook posts the document directly when "Send entire document" is on,
    // or { _id, _type, ... projection ... } when a custom projection is configured.
    const doc = payload._type ? payload : payload.document || payload;

    if (doc._type !== 'job') return json({ ok: true, skipped: 'not a job' });
    if (doc.status && doc.status !== 'active') return json({ ok: true, skipped: 'job not active' });
    if (doc.language && doc.language !== 'en') return json({ ok: true, skipped: 'non-en variant' });

    const jobId = doc.jobId || doc.slug?.current || doc.slug || doc._id;
    const title = doc.title || 'New role';
    const department = doc.department || 'Open Role';
    const workplaceType = doc.workplaceType || '';
    const location = [doc.locationCity, doc.locationCountry].filter(Boolean).join(', ');
    const excerpt = doc.excerpt || '';

    const subject = `New role: ${title}${location ? ' (' + location + ')' : ''}`;
    const html = buildEmailHtml({ title, department, workplaceType, location, excerpt, jobId });

    if (!env.RESEND_API_KEY) return json({ ok: false, error: 'RESEND_API_KEY missing' }, 500);
    if (!env.RESEND_AUDIENCE_ID) return json({ ok: false, error: 'RESEND_AUDIENCE_ID missing' }, 500);

    // 2. Don't email a link to a page that isn't built yet.
    // Publishing fires this webhook instantly, but the static rebuild lags
    // ~1-2 min, so an early click would hit the homepage fallback. Poll the
    // live job page first; if it isn't up within ~20s, return 503 so Sanity
    // retries the webhook later. We never send while the page is missing, so
    // there's no double-send and the email only goes out once the link works.
    const live = await waitForJobPageLive(jobId);
    if (!live) {
      return json({ ok: false, retry: true, reason: 'job page not deployed yet', jobId }, 503);
    }

    const from = env.MAIL_FROM || 'ArgusRecruit <contact@argusrecruit.com>';

    // 3. Create a broadcast
    const createRes = await fetch('https://api.resend.com/broadcasts', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${env.RESEND_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        audience_id: env.RESEND_AUDIENCE_ID,
        from,
        subject,
        html,
        name: `Job Alert: ${title}`
      })
    });
    const createBody = await createRes.json();
    if (!createRes.ok) {
      return json({ ok: false, step: 'create', detail: createBody }, 500);
    }
    const broadcastId = createBody.id;

    // 4. Send the broadcast immediately
    const sendRes = await fetch(`https://api.resend.com/broadcasts/${broadcastId}/send`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${env.RESEND_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({})
    });
    const sendBody = await sendRes.json().catch(() => ({}));
    if (!sendRes.ok) {
      return json({ ok: false, step: 'send', broadcastId, detail: sendBody }, 500);
    }

    return json({ ok: true, broadcastId, jobId });
  } catch (e) {
    return json({ ok: false, error: e.message, stack: e.stack }, 500);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// Poll the live job page until it has been built and deployed.
// The static job page contains its own jobId (in the apply form + share links);
// the homepage fallback never does, so jobId presence is a reliable "live" signal.
// ~7 tries x 3s stays well under Sanity's ~30s webhook timeout.
async function waitForJobPageLive(jobId, { attempts = 7, delayMs = 3000 } = {}) {
  if (!jobId) return false;
  const base = `https://argusrecruit.com/jobs/${encodeURIComponent(jobId)}/`;
  for (let i = 0; i < attempts; i++) {
    try {
      const res = await fetch(`${base}?_=${Date.now()}`, {
        headers: { 'Cache-Control': 'no-cache' },
        cf: { cacheTtl: 0, cacheEverything: false }
      });
      if (res.ok) {
        const html = await res.text();
        if (html.includes(jobId)) return true;
      }
    } catch (_) { /* transient — keep polling */ }
    if (i < attempts - 1) await sleep(delayMs);
  }
  return false;
}

function esc(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function buildEmailHtml({ title, department, workplaceType, location, excerpt, jobId }) {
  const meta = [department, workplaceType, location].filter(Boolean).map(esc).join(' &nbsp;·&nbsp; ');
  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)} — ArgusRecruit</title></head>
<body style="margin:0;padding:0;background:#0E2440;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;color:#ffffff;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0E2440;padding:32px 16px;">
    <tr><td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#1E4170;border:1px solid rgba(212,175,55,0.2);border-radius:14px;overflow:hidden;">
        <tr><td align="center" style="padding:36px 32px 24px;background:#16345A;border-bottom:1px solid rgba(212,175,55,0.2);">
          <img src="https://argusrecruit.com/logo.png" alt="ArgusRecruit" width="64" style="display:block;height:64px;width:auto;">
          <div style="margin-top:14px;font-size:11px;letter-spacing:2.5px;color:#D4AF37;text-transform:uppercase;font-weight:700;">Many Eyes. One Purpose.</div>
        </td></tr>
        <tr><td align="center" style="padding:32px 32px 0;">
          <div style="display:inline-block;font-size:10px;letter-spacing:4px;color:#D4AF37;text-transform:uppercase;font-weight:700;padding:6px 16px;border:1px solid rgba(212,175,55,0.4);border-radius:999px;">• New Opportunity •</div>
        </td></tr>
        <tr><td align="center" style="padding:18px 32px 8px;">
          <h1 style="margin:0;font-size:26px;font-weight:700;letter-spacing:0.5px;color:#ffffff;text-transform:uppercase;line-height:1.2;">${esc(title)}</h1>
        </td></tr>
        <tr><td align="center" style="padding:0 32px 24px;">
          <div style="color:rgba(255,255,255,0.75);font-size:14px;letter-spacing:0.5px;">${meta || ''}</div>
        </td></tr>
        <tr><td style="padding:0 32px;"><div style="height:1px;background:linear-gradient(90deg,transparent,rgba(212,175,55,0.4),transparent);"></div></td></tr>
        <tr><td style="padding:28px 36px;color:rgba(255,255,255,0.85);font-size:15px;line-height:1.7;">
          <p style="margin:0 0 14px;">Hi,</p>
          <p style="margin:0 0 14px;">A new role has opened in our network that we thought you would want to see.</p>
          ${excerpt ? `<p style="margin:0 0 14px;font-style:italic;color:rgba(255,255,255,0.75);">${esc(excerpt)}</p>` : ''}
          <p style="margin:0;">All applications are confidential and handled personally by a senior recruiter.</p>
        </td></tr>
        <tr><td align="center" style="padding:8px 32px 36px;">
          <a href="https://argusrecruit.com/jobs/${esc(jobId)}/" style="display:inline-block;background:#D4AF37;color:#0E2440;font-weight:700;font-size:14px;letter-spacing:1px;text-transform:uppercase;padding:14px 32px;border-radius:999px;text-decoration:none;">View Role &amp; Apply</a>
        </td></tr>
        <tr><td align="center" style="padding:0 32px 32px;">
          <a href="https://argusrecruit.com/jobs/" style="color:#D4AF37;font-size:13px;letter-spacing:1px;text-transform:uppercase;text-decoration:none;font-weight:600;">Browse All Open Roles</a>
        </td></tr>
        <tr><td style="padding:24px 32px;background:#0E2440;border-top:1px solid rgba(212,175,55,0.15);text-align:center;">
          <div style="color:rgba(255,255,255,0.55);font-size:12px;line-height:1.65;">
            You are receiving this because you subscribed to ArgusRecruit job alerts.<br>
            <a href="{{{RESEND_UNSUBSCRIBE_URL}}}" style="color:#D4AF37;text-decoration:none;">Unsubscribe</a> &nbsp;·&nbsp;
            <a href="https://argusrecruit.com" style="color:#D4AF37;text-decoration:none;">argusrecruit.com</a> &nbsp;·&nbsp;
            <a href="mailto:contact@argusrecruit.com" style="color:#D4AF37;text-decoration:none;">contact@argusrecruit.com</a>
          </div>
          <div style="margin-top:14px;color:rgba(255,255,255,0.4);font-size:11px;letter-spacing:1px;text-transform:uppercase;">© 2026 ArgusRecruit · Yerevan, Armenia</div>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}
