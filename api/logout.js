import { COOKIE_NAME } from '../lib/auth.js';

export default async function handler(req, res) {
  res.setHeader('Set-Cookie', `${COOKIE_NAME}=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax`);
  res.writeHead(302, { Location: '/login.html' });
  res.end();
}
