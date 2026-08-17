'use strict';

/**
 * Google sign-in for the desktop build.
 *
 * The web app used Google Identity Services, which only works from a real
 * browser origin and only ever hands out a one-hour access token — no refresh
 * token, so it had to re-ask Google every hour and could only do that from
 * inside a user gesture. A bundled app cannot use that flow at all.
 *
 * This is the installed-app flow instead (RFC 8252): open the system browser,
 * catch the redirect on loopback, exchange the code with PKCE. It gives a
 * refresh token, so the session survives restarts and renews with no UI.
 *
 * The client secret in a desktop client is not really secret — it ships inside
 * the app and Google's own documentation says as much — which is exactly why
 * PKCE is required rather than optional.
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const { DATA_DIR } = require('./paths');

const AUTH_PATH = path.join(DATA_DIR, 'auth.json');
const SCOPES = [
  'https://www.googleapis.com/auth/userinfo.email',
  'https://www.googleapis.com/auth/userinfo.profile',
  'https://www.googleapis.com/auth/spreadsheets',
].join(' ');

// Refresh a little early: a token that expires mid-request is a failed save.
const EXPIRY_SKEW_MS = 60_000;

let pending = null;   // { verifier, state, createdAt } for the in-flight sign-in

/**
 * How the last sign-in attempt ended. The window that started the flow is not
 * the window Google redirects to, so the outcome has to be left somewhere the
 * app can find it — otherwise a user who presses Cancel is met with a UI that
 * keeps waiting for a callback that will never arrive.
 */
let flowState = { status: 'idle', message: '' };

function flow() {
  return flowState;
}

/* ── client config ───────────────────────────────────────────────────────── */

/**
 * The client_secret.json downloaded from Google Cloud Console. Kept out of git
 * and out of the served UI; it is read here and nowhere else.
 */
function clientConfig() {
  for (const file of [
    path.join(DATA_DIR, 'google-client.json'),
    path.join(__dirname, '..', 'google-client.json'),
  ]) {
    if (!fs.existsSync(file)) continue;
    const raw = JSON.parse(fs.readFileSync(file, 'utf8'));
    const c = raw.installed || raw.web;
    if (!c || !c.client_id) continue;
    return {
      clientId: c.client_id,
      clientSecret: c.client_secret || '',
      authUri: c.auth_uri || 'https://accounts.google.com/o/oauth2/auth',
      tokenUri: c.token_uri || 'https://oauth2.googleapis.com/token',
    };
  }
  return null;
}

function configured() {
  return !!clientConfig();
}

/* ── token storage ───────────────────────────────────────────────────────── */

function readAuth() {
  try { return JSON.parse(fs.readFileSync(AUTH_PATH, 'utf8')); } catch { return null; }
}

function writeAuth(auth) {
  fs.mkdirSync(path.dirname(AUTH_PATH), { recursive: true });
  // 0600: the refresh token is a long-lived credential for a work account.
  fs.writeFileSync(AUTH_PATH, JSON.stringify(auth, null, 2), { mode: 0o600 });
}

function clearAuth() {
  try { fs.unlinkSync(AUTH_PATH); } catch { /* already gone */ }
}

/* ── the flow ────────────────────────────────────────────────────────────── */

const b64url = (buf) => buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

function redirectUri(port) {
  // 127.0.0.1 rather than localhost: the agent binds v4 loopback only, and on
  // Windows "localhost" can resolve to ::1 first and never arrive.
  return `http://127.0.0.1:${port}/oauth/callback`;
}

/** Build the consent URL and open it in the real browser. */
function begin({ port, hint = '' }) {
  const cfg = clientConfig();
  if (!cfg) throw new Error('No Google client configured. Put google-client.json next to the agent.');

  const verifier = b64url(crypto.randomBytes(64));
  const challenge = b64url(crypto.createHash('sha256').update(verifier).digest());
  const state = b64url(crypto.randomBytes(16));
  pending = { verifier, state, createdAt: Date.now() };
  flowState = { status: 'pending', message: '' };

  const params = new URLSearchParams({
    client_id: cfg.clientId,
    redirect_uri: redirectUri(port),
    response_type: 'code',
    scope: SCOPES,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    // offline + consent is what actually yields a refresh token; without it
    // Google returns an access token only and the session dies in an hour.
    access_type: 'offline',
    prompt: 'consent',
    state,
  });
  if (hint) params.set('hd', hint);

  const url = `${cfg.authUri}?${params}`;
  openInBrowser(url);
  return url;
}

function openInBrowser(url) {
  // Inside the desktop app this is the reliable route; the agent may also run
  // standalone, where electron is simply not installed.
  try {
    const { shell } = require('electron');
    if (shell && typeof shell.openExternal === 'function') {
      shell.openExternal(url);
      return;
    }
  } catch { /* standalone agent — fall through */ }

  if (process.platform === 'win32') {
    // Deliberately NOT `cmd /c start`: cmd treats & as a command separator, so
    // an OAuth URL is cut off at the first query parameter and Google answers
    // "Required parameter is missing: response_type". rundll32 takes the URL
    // as a single argument with no shell parsing in between.
    spawn('rundll32', ['url.dll,FileProtocolHandler', url], { detached: true, stdio: 'ignore' }).unref();
  } else if (process.platform === 'darwin') {
    spawn('open', [url], { detached: true, stdio: 'ignore' }).unref();
  } else {
    spawn('xdg-open', [url], { detached: true, stdio: 'ignore' }).unref();
  }
}

async function exchange(code, state, port) {
  const cfg = clientConfig();
  if (!pending) throw new Error('No sign-in is in progress.');
  if (state !== pending.state) throw new Error('State mismatch — the callback did not come from this app.');

  const res = await fetch(cfg.tokenUri, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: cfg.clientId,
      client_secret: cfg.clientSecret,
      redirect_uri: redirectUri(port),
      grant_type: 'authorization_code',
      code_verifier: pending.verifier,
    }),
  });
  const body = await res.json();
  pending = null;
  if (!res.ok) {
    const message = body.error_description || body.error || `token exchange failed (${res.status})`;
    flowState = { status: 'error', message };
    throw new Error(message);
  }

  const auth = {
    access_token: body.access_token,
    refresh_token: body.refresh_token || null,
    expires_at: Date.now() + (Number(body.expires_in || 3600) * 1000),
    scope: body.scope || SCOPES,
  };
  auth.user = await fetchUserInfo(auth.access_token);
  writeAuth(auth);
  flowState = { status: 'ok', message: '' };
  return auth;
}

/** Google reported an error instead of a code — usually the user pressed Cancel. */
function fail(error) {
  pending = null;
  const cancelled = error === 'access_denied';
  flowState = {
    status: cancelled ? 'cancelled' : 'error',
    message: cancelled ? 'Sign-in was cancelled.' : `Google refused the sign-in: ${error}`,
  };
  return flowState;
}

async function refresh(auth) {
  const cfg = clientConfig();
  if (!auth || !auth.refresh_token) throw new Error('No refresh token — sign in again.');

  const res = await fetch(cfg.tokenUri, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: cfg.clientId,
      client_secret: cfg.clientSecret,
      refresh_token: auth.refresh_token,
      grant_type: 'refresh_token',
    }),
  });
  const body = await res.json();
  if (!res.ok) {
    // A revoked or expired grant is terminal; drop it so the UI asks again
    // rather than retrying forever.
    if (body.error === 'invalid_grant') clearAuth();
    throw new Error(body.error_description || body.error || `refresh failed (${res.status})`);
  }

  const next = {
    ...auth,
    access_token: body.access_token,
    expires_at: Date.now() + (Number(body.expires_in || 3600) * 1000),
    // Google only returns refresh_token on the first consent; keep the old one.
    refresh_token: body.refresh_token || auth.refresh_token,
  };
  writeAuth(next);
  return next;
}

async function fetchUserInfo(accessToken) {
  const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error('Could not read the Google profile.');
  const d = await res.json();
  return { name: d.name, email: d.email, picture: d.picture, hd: d.hd, email_verified: d.email_verified };
}

/** A valid access token, renewed if it is close to expiry. Null when signed out. */
async function token() {
  let auth = readAuth();
  if (!auth) return null;
  if (Date.now() < auth.expires_at - EXPIRY_SKEW_MS) return auth;
  return refresh(auth);
}

function session() {
  const auth = readAuth();
  return auth ? { user: auth.user, expires_at: auth.expires_at } : null;
}

async function signOut() {
  const auth = readAuth();
  clearAuth();
  if (!auth) return;
  // Best effort: tell Google too, so the grant does not linger.
  try {
    await fetch(`https://oauth2.googleapis.com/revoke?token=${encodeURIComponent(auth.refresh_token || auth.access_token)}`,
      { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' } });
  } catch { /* offline; the local copy is gone either way */ }
}

module.exports = { configured, begin, exchange, fail, flow, token, session, signOut, fetchUserInfo, SCOPES };
