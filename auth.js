import {
  CLIENT_ID,
  REDIRECT_URI,
  SCOPES,
  AUTH_ENDPOINT,
  TOKEN_ENDPOINT,
} from './config.js';

export async function getAccounts() {
  const data = await chrome.storage.local.get('accounts');
  return data.accounts || [];
}

async function saveAccounts(accounts) {
  await chrome.storage.local.set({ accounts });
}

/**
 * Opens an interactive OAuth2 flow (prompt=select_account) with PKCE.
 * Must be called from a user gesture context (popup).
 */
export async function addAccount() {
  const { verifier, challenge } = await generatePKCE();
  const authUrl = buildAuthUrl({ prompt: 'select_account', code_challenge: challenge, code_challenge_method: 'S256' });

  const responseUrl = await chrome.identity.launchWebAuthFlow({ url: authUrl, interactive: true });
  const code = extractCode(responseUrl);

  const tokens = await exchangeCodeForTokens(code, verifier);
  return persistAccount(tokens);
}

/**
 * Silently refreshes a token using prompt=none + login_hint + PKCE.
 * Falls back to refresh_token grant if silent flow fails.
 */
export async function refreshAccount(account) {
  let tokens;

  try {
    const { verifier, challenge } = await generatePKCE();
    const authUrl = buildAuthUrl({ prompt: 'none', login_hint: account.email, code_challenge: challenge, code_challenge_method: 'S256' });
    const responseUrl = await chrome.identity.launchWebAuthFlow({ url: authUrl, interactive: false });
    const code = extractCode(responseUrl);
    tokens = await exchangeCodeForTokens(code, verifier);
  } catch {
    if (!account.refresh_token) throw new Error(`Silent auth failed for ${account.email} and no refresh_token available.`);
    tokens = await refreshWithToken(account.refresh_token);
  }

  const accounts = await getAccounts();
  const idx = accounts.findIndex(a => a.email === account.email);
  if (idx < 0) throw new Error('Account not found in storage.');

  const updated = {
    ...accounts[idx],
    access_token: tokens.access_token,
    refresh_token: tokens.refresh_token || accounts[idx].refresh_token,
    expires_at: Date.now() + tokens.expires_in * 1000,
    needsReauth: false,
  };
  accounts[idx] = updated;
  await saveAccounts(accounts);
  return updated;
}

export async function getValidToken(account) {
  if (Date.now() < account.expires_at - 60_000) {
    return account.access_token;
  }
  const updated = await refreshAccount(account);
  return updated.access_token;
}

export async function markNeedsReauth(email) {
  const accounts = await getAccounts();
  const idx = accounts.findIndex(a => a.email === email);
  if (idx >= 0) {
    accounts[idx].needsReauth = true;
    await saveAccounts(accounts);
  }
}

export async function removeAccount(email) {
  const accounts = await getAccounts();
  await saveAccounts(accounts.filter(a => a.email !== email));
}

// ── PKCE ─────────────────────────────────────────────────────────────────────

async function generatePKCE() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  const verifier = btoa(String.fromCharCode(...array))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');

  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  const challenge = btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');

  return { verifier, challenge };
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function buildAuthUrl(extra = {}) {
  const url = new URL(AUTH_ENDPOINT);
  url.searchParams.set('client_id', CLIENT_ID);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('redirect_uri', REDIRECT_URI);
  url.searchParams.set('scope', SCOPES);
  url.searchParams.set('response_mode', 'query');
  for (const [k, v] of Object.entries(extra)) {
    url.searchParams.set(k, v);
  }
  return url.toString();
}

function extractCode(responseUrl) {
  const params = new URL(responseUrl).searchParams;
  const error = params.get('error');
  if (error) {
    const desc = params.get('error_description') || error;
    throw new Error(`Azure AD error: ${desc}`);
  }
  const code = params.get('code');
  if (!code) throw new Error('No authorization code received.');
  return code;
}

async function exchangeCodeForTokens(code, codeVerifier) {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    code,
    redirect_uri: REDIRECT_URI,
    grant_type: 'authorization_code',
    code_verifier: codeVerifier,
    scope: SCOPES,
  });
  const res = await fetch(TOKEN_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error_description || data.error);
  return data;
}

async function refreshWithToken(refreshToken) {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    refresh_token: refreshToken,
    grant_type: 'refresh_token',
    scope: SCOPES,
  });
  const res = await fetch(TOKEN_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error_description || data.error);
  return data;
}

async function persistAccount(tokens) {
  const userRes = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: { Authorization: `Bearer ${tokens.access_token}` },
  });
  const user = await userRes.json();
  const email = user.mail || user.userPrincipalName;

  const account = {
    email,
    displayName: user.displayName || email,
    access_token: tokens.access_token,
    refresh_token: tokens.refresh_token,
    expires_at: Date.now() + tokens.expires_in * 1000,
    needsReauth: false,
  };

  const accounts = await getAccounts();
  const idx = accounts.findIndex(a => a.email === email);
  if (idx >= 0) {
    accounts[idx] = account;
  } else {
    accounts.push(account);
  }
  await saveAccounts(accounts);
  return account;
}
