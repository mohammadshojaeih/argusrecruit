// Diagnostic endpoint: verifies Drive setup end-to-end.
// Hit GET /api/drive-check from a browser to see what's wrong.
// Remove this file once debugging is complete.

export async function onRequestGet(context) {
  const { env } = context;
  const out = { steps: [] };

  function step(name, ok, detail) {
    out.steps.push({ name, ok, detail: detail ? String(detail).slice(0, 600) : undefined });
  }

  // 1. Check env vars exist
  if (!env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    step('env GOOGLE_SERVICE_ACCOUNT_JSON', false,
         'NOT SET — add it as a Secret in Cloudflare Pages → Settings → Environment variables (Production)');
    return json(out, 500);
  }
  step('env GOOGLE_SERVICE_ACCOUNT_JSON', true, `length=${env.GOOGLE_SERVICE_ACCOUNT_JSON.length}`);

  if (!env.DRIVE_FOLDER_ID) {
    step('env DRIVE_FOLDER_ID', false, 'NOT SET');
    return json(out, 500);
  }
  step('env DRIVE_FOLDER_ID', true, env.DRIVE_FOLDER_ID);

  // 2. Parse JSON
  let sa;
  try {
    sa = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } catch (e) {
    step('parse JSON', false, e.message);
    return json(out, 500);
  }
  step('parse JSON', true, `client_email=${sa.client_email || '?'} project=${sa.project_id || '?'}`);

  // 3. Sign + get access token
  let accessToken;
  try {
    accessToken = await getAccessToken(sa);
    step('OAuth token exchange', true, accessToken ? `token len=${accessToken.length}` : 'no token');
  } catch (e) {
    step('OAuth token exchange', false, e.message);
    return json(out, 500);
  }

  // 4. Try to GET the folder metadata (verifies share permission)
  try {
    const r = await fetch(
      `https://www.googleapis.com/drive/v3/files/${env.DRIVE_FOLDER_ID}?fields=id,name,mimeType&supportsAllDrives=true`,
      { headers: { Authorization: 'Bearer ' + accessToken } }
    );
    const body = await r.json();
    if (r.ok && body.id) {
      step('access target folder', true, `name="${body.name}" mimeType="${body.mimeType}"`);
    } else {
      step('access target folder', false,
           `HTTP ${r.status} — ${JSON.stringify(body)} (likely: folder not shared with service-account email, or wrong folder ID)`);
      return json(out, 500);
    }
  } catch (e) {
    step('access target folder', false, e.message);
    return json(out, 500);
  }

  // 5. Try to write a tiny test file
  try {
    const boundary = 'TEST-' + Math.random().toString(36).slice(2);
    const metadata = { name: `drive-check-${Date.now()}.txt`, parents: [env.DRIVE_FOLDER_ID] };
    const head = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: text/plain\r\n\r\n`;
    const tail = `\r\n--${boundary}--`;
    const body = new Blob([head, 'ArgusRecruit Drive connection test ' + new Date().toISOString(), tail]);
    const r = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
      {
        method: 'POST',
        headers: {
          Authorization: 'Bearer ' + accessToken,
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        body
      }
    );
    const j = await r.json();
    if (r.ok) {
      step('upload test file', true, `fileId=${j.id} name=${j.name}`);
    } else {
      step('upload test file', false, `HTTP ${r.status} — ${JSON.stringify(j)}`);
    }
  } catch (e) {
    step('upload test file', false, e.message);
  }

  return json(out, 200);
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj, null, 2), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' }
  });
}

async function getAccessToken(sa) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/drive.file',
    aud: sa.token_uri,
    iat: now,
    exp: now + 3600
  };
  const enc = (s) => btoa(s).replace(/=+$/g, '').replace(/\+/g, '-').replace(/\//g, '_');
  const encBytes = (b) => {
    const a = new Uint8Array(b);
    let s = '';
    for (let i = 0; i < a.length; i++) s += String.fromCharCode(a[i]);
    return btoa(s).replace(/=+$/g, '').replace(/\+/g, '-').replace(/\//g, '_');
  };
  const unsigned = `${enc(JSON.stringify(header))}.${enc(JSON.stringify(payload))}`;

  const pem = sa.private_key
    .replace(/-----BEGIN PRIVATE KEY-----/, '')
    .replace(/-----END PRIVATE KEY-----/, '')
    .replace(/\s+/g, '');
  const keyBuf = Uint8Array.from(atob(pem), c => c.charCodeAt(0)).buffer;
  const key = await crypto.subtle.importKey(
    'pkcs8', keyBuf,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false, ['sign']
  );
  const sig = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5', key, new TextEncoder().encode(unsigned)
  );
  const jwt = `${unsigned}.${encBytes(sig)}`;

  const r = await fetch(sa.token_uri, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=' + encodeURIComponent(jwt)
  });
  if (!r.ok) {
    const t = await r.text();
    throw new Error(`token endpoint HTTP ${r.status}: ${t}`);
  }
  const j = await r.json();
  if (!j.access_token) throw new Error('no access_token in response: ' + JSON.stringify(j));
  return j.access_token;
}
