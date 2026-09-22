import { createRemoteJWKSet, jwtVerify } from 'jose';
import { DomainError } from './domain';

export type Actor = { actor: string; role: 'admin' | 'picker' };
type AuthEnv = Pick<Env, 'ECWID_MODE' | 'ACCESS_TEAM_DOMAIN' | 'ACCESS_AUD' | 'ADMIN_EMAILS'>;

export async function authenticate(request: Request, env: AuthEnv): Promise<Actor> {
  // Demo identity exists only on a loopback development URL. A deployed demo
  // configuration still fails closed; it cannot open a public inventory API.
  const hostname = new URL(request.url).hostname;
  if (env.ECWID_MODE === 'demo' && ['localhost', '127.0.0.1', '[::1]'].includes(hostname)) {
    return { actor: 'demo@local', role: 'admin' };
  }
  if (env.ECWID_MODE !== 'live' || !env.ACCESS_TEAM_DOMAIN || !env.ACCESS_AUD) {
    throw new DomainError(503, 'ACCESS_NOT_CONFIGURED', 'Staff sign-in is not configured for this deployment.');
  }
  const teamDomain = env.ACCESS_TEAM_DOMAIN.replace(/\/$/, '');
  if (!/^https:\/\/[a-z0-9-]+\.cloudflareaccess\.com$/i.test(teamDomain)) {
    throw new DomainError(503, 'ACCESS_NOT_CONFIGURED', 'Staff sign-in configuration needs attention.');
  }
  const token = request.headers.get('cf-access-jwt-assertion');
  if (!token) throw new DomainError(401, 'SIGN_IN_REQUIRED', 'Sign in through the staff access page.');
  let email: string;
  try {
    const keys = createRemoteJWKSet(new URL(`${teamDomain}/cdn-cgi/access/certs`));
    const { payload } = await jwtVerify(token, keys, {
      issuer: teamDomain,
      audience: env.ACCESS_AUD,
      algorithms: ['RS256'],
      requiredClaims: ['exp', 'iat', 'sub', 'email']
    });
    if (typeof payload.email !== 'string' || !payload.email.includes('@')) throw new Error('No staff identity');
    email = payload.email.toLowerCase();
  } catch {
    throw new DomainError(401, 'INVALID_SESSION', 'Your sign-in has expired or could not be verified. Sign in again.');
  }
  const admins = env.ADMIN_EMAILS.split(',').map(value => value.trim().toLowerCase()).filter(Boolean);
  return { actor: email, role: admins.includes(email) ? 'admin' : 'picker' };
}

export function requireAdmin(actor: Actor): void {
  if (actor.role !== 'admin') throw new DomainError(403, 'ADMIN_REQUIRED', 'Only an administrator can perform this action.');
}
