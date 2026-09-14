import { COOKIE_NAME, isValidToken } from './lib/auth.js';

export const config = { matcher: '/:path*' };

const PUBLIC_PATHS = new Set(['/login.html', '/api/login', '/api/logout']);

function getCookie(req, name) {
  const header = req.headers.get('cookie') || '';
  const match = header.match(new RegExp('(?:^|; )' + name + '=([^;]*)'));
  return match ? decodeURIComponent(match[1]) : null;
}

export default async function middleware(req) {
  const url = new URL(req.url);
  if (PUBLIC_PATHS.has(url.pathname)) return;

  const secret = process.env.AUTH_SECRET;
  if (!secret) {
    return new Response('Server misconfigured: AUTH_SECRET is not set.', { status: 500 });
  }

  const token = getCookie(req, COOKIE_NAME);
  if (await isValidToken(token, secret)) return;

  const loginUrl = new URL('/login.html', url);
  loginUrl.searchParams.set('next', url.pathname + url.search);
  return Response.redirect(loginUrl, 307);
}
