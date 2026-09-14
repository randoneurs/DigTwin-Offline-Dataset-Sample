// Shared cookie-signing helpers for the login gate (middleware.js + api/login.js
// + api/logout.js). Uses Web Crypto (crypto.subtle) so the same code runs
// unchanged on both the Edge runtime (middleware) and the Node runtime (api/*).

export const COOKIE_NAME = 'msight_auth';
export const TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

function toBase64Url(bytes) {
  let str = '';
  for (const b of bytes) str += String.fromCharCode(b);
  return btoa(str).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function hmac(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(message));
  return toBase64Url(new Uint8Array(sig));
}

export async function makeToken(secret) {
  const exp = String(Date.now() + TTL_MS);
  const sig = await hmac(secret, exp);
  return `${exp}.${sig}`;
}

export async function isValidToken(token, secret) {
  if (!token) return false;
  const parts = token.split('.');
  if (parts.length !== 2) return false;
  const [expStr, sig] = parts;
  const exp = Number(expStr);
  if (!Number.isFinite(exp) || Date.now() > exp) return false;
  const expected = await hmac(secret, expStr);
  if (expected.length !== sig.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ sig.charCodeAt(i);
  return diff === 0;
}
