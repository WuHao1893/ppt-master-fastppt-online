import crypto from 'node:crypto';
import type { User } from '../shared/models.js';
import type { StateStore } from './store.js';
import { HttpError } from './errors.js';
import { makeId, nowIso } from './utils.js';

interface Session {
  userId: string;
  expiresAt: number;
  sessionId: string;
}

export class AuthService {
  private readonly tickets = new Map<string, Session>();
  private readonly secret: string;
  private readonly accessCode: string;
  private readonly allowedEmails: Set<string>;

  constructor(private readonly store: StateStore) {
    const configured = process.env.AUTH_SECRET?.trim();
    if (process.env.NODE_ENV === 'production' && !configured) throw new Error('AUTH_SECRET is required in production.');
    this.secret = configured || 'fastppt-online-local-secret-development-only';
    this.accessCode = process.env.AUTH_LOGIN_CODE?.trim() || '';
    this.allowedEmails = new Set((process.env.AUTH_ALLOWED_EMAILS || '').split(',').map((value) => value.trim().toLowerCase()).filter(Boolean));
    if (process.env.NODE_ENV === 'production' && (!this.accessCode || !this.allowedEmails.size)) throw new Error('AUTH_LOGIN_CODE and AUTH_ALLOWED_EMAILS are required in production.');
  }

  async login(email: string, requestedName?: string, providedAccessCode?: string): Promise<{ token: string; user: User }> {
    const normalizedEmail = email.trim().toLowerCase();
    if (process.env.NODE_ENV === 'production') {
      if (!this.allowedEmails.has(normalizedEmail)) throw new HttpError(403, 'This account is not allowed to access the workspace.');
      const provided = Buffer.from(providedAccessCode || '');
      const expected = Buffer.from(this.accessCode);
      if (provided.length !== expected.length || !crypto.timingSafeEqual(provided, expected)) throw new HttpError(401, 'Invalid access code.');
    }
    let user = this.store.state.users.find((candidate) => candidate.email.toLowerCase() === normalizedEmail);
    if (!user) {
      if (process.env.NODE_ENV !== 'production' && process.env.ALLOW_DEV_LOGIN === 'false') throw new HttpError(403, 'Development self-service login is disabled.');
      user = { userId: makeId('usr'), email: normalizedEmail, name: requestedName || normalizedEmail.split('@')[0], createdAt: nowIso() };
      await this.store.update((state) => {
        state.users.push(user!);
        state.auditLogs.push({ auditId: makeId('audit'), ownerId: user!.userId, action: 'auth.login.created_user', payload: { email, invitation: process.env.NODE_ENV === 'production' ? 'allowlist_access_code' : 'development' }, createdAt: nowIso() });
      });
    }
    const token = await this.issueToken(user.userId, 1000 * 60 * 60 * 24 * 7);
    await this.store.update((state) => state.auditLogs.push({ auditId: makeId('audit'), ownerId: user!.userId, action: 'auth.login', payload: { email: user!.email }, createdAt: nowIso() }));
    return { token, user };
  }

  verify(token: string | undefined): User | null {
    if (!token) return null;
    const session = this.decodeToken(token);
    if (!session || session.expiresAt < Date.now()) return null;
    const persisted = this.store.state.authSessions.find((candidate) => candidate.sessionId === session.sessionId);
    if (!persisted || persisted.revokedAt || new Date(persisted.expiresAt).getTime() < Date.now()) return null;
    return this.store.state.users.find((user) => user.userId === session.userId) || null;
  }

  async revoke(token: string | undefined): Promise<void> {
    if (!token) return;
    const session = this.decodeToken(token);
    if (!session) return;
    await this.store.update((state) => {
      const persisted = state.authSessions.find((candidate) => candidate.sessionId === session.sessionId);
      if (persisted) persisted.revokedAt = nowIso();
    });
  }

  issueWebSocketTicket(userId: string): string {
    const ticket = crypto.randomBytes(24).toString('base64url');
    this.tickets.set(ticket, { userId, expiresAt: Date.now() + 30_000, sessionId: ticket });
    return ticket;
  }

  consumeWebSocketTicket(ticket: string | undefined): User | null {
    if (!ticket) return null;
    const session = this.tickets.get(ticket);
    this.tickets.delete(ticket);
    if (!session || session.expiresAt < Date.now()) return null;
    return this.store.state.users.find((user) => user.userId === session.userId) || null;
  }

  private async issueToken(userId: string, lifetimeMs: number): Promise<string> {
    const sessionId = makeId('session');
    const expiresAt = Date.now() + lifetimeMs;
    await this.store.update((state) => state.authSessions.push({ sessionId, userId, expiresAt: new Date(expiresAt).toISOString(), createdAt: nowIso(), revokedAt: null }));
    const payload = Buffer.from(JSON.stringify({ sub: userId, exp: expiresAt, sid: sessionId })).toString('base64url');
    const signature = crypto.createHmac('sha256', this.secret).update(payload).digest('base64url');
    return `${payload}.${signature}`;
  }

  private decodeToken(token: string): Session | null {
    const [payload, signature] = token.split('.');
    if (!payload || !signature) return null;
    const expected = crypto.createHmac('sha256', this.secret).update(payload).digest('base64url');
    const left = Buffer.from(signature);
    const right = Buffer.from(expected);
    if (left.length !== right.length || !crypto.timingSafeEqual(left, right)) return null;
    try {
      const decoded = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')) as { sub?: string; exp?: number; sid?: string };
      if (!decoded.sub || !decoded.exp || !decoded.sid) return null;
      return { userId: decoded.sub, expiresAt: decoded.exp, sessionId: decoded.sid };
    } catch {
      return null;
    }
  }
}
