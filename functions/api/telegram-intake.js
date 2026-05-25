// Telegram webhook proxy → Apps Script Web App
// Telegram requires a direct 200 OK from the webhook URL.
// Apps Script Web Apps respond with a 302 to script.googleusercontent.com,
// which Telegram does not follow. So we accept the POST here, forward it
// to Apps Script (following the redirect), and respond 200 to Telegram.

export async function onRequestPost(context) {
  const { request, env } = context;
  const target = env.APPS_SCRIPT_INTAKE_URL;
  if (!target) {
    return new Response('APPS_SCRIPT_INTAKE_URL not configured', { status: 500 });
  }

  // Acknowledge to Telegram immediately; do the heavy work in waitUntil so the
  // response stays quick and we never time-out the webhook.
  const body = await request.text();
  const forwardPromise = fetch(target, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body,
    redirect: 'follow'
  }).catch(err => console.error('Apps Script forward failed:', err));

  if (context.waitUntil) context.waitUntil(forwardPromise); else await forwardPromise;
  return new Response('ok', { status: 200 });
}

export function onRequestGet() {
  return new Response('Telegram intake proxy is alive.', { status: 200 });
}
