import { COOKIE_NAME, TTL_MS, makeToken } from '../lib/auth.js';

function safeNext(next) {
  if (typeof next !== 'string' || !next.startsWith('/') || next.startsWith('//')) return '/';
  return next;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).send('Method not allowed');
    return;
  }

  const secret = process.env.AUTH_SECRET;
  const expectedUser = process.env.SITE_USERNAME;
  const expectedPass = process.env.SITE_PASSWORD;
  if (!secret || !expectedUser || !expectedPass) {
    res.status(500).send('Server misconfigured: SITE_USERNAME/SITE_PASSWORD/AUTH_SECRET not set.');
    return;
  }

  const { username, password, next } = req.body || {};
  const nextPath = safeNext(next);

  if (username !== expectedUser || password !== expectedPass) {
    res.writeHead(302, { Location: `/login.html?error=1&next=${encodeURIComponent(nextPath)}` });
    res.end();
    return;
  }

  const token = await makeToken(secret);
  res.setHeader(
    'Set-Cookie',
    `${COOKIE_NAME}=${token}; Path=/; Max-Age=${Math.floor(TTL_MS / 1000)}; HttpOnly; Secure; SameSite=Lax`
  );
  res.writeHead(302, { Location: nextPath });
  res.end();
}
