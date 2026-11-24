// Password login + server-side sessions (random 32-byte id in a Map, 12 h) + login throttle (10 / 15 min / IP).
import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import { config } from "./config.ts";

const SESSION_MS = 12 * 60 * 60 * 1000;
const LOGIN_WINDOW_MS = 15 * 60 * 1000;
const LOGIN_MAX = 10;

const sessions = new Map<string, number>(); // sid -> expires
const loginTries = new Map<string, number[]>(); // ip -> attempt times

export function passwordMatches(given: string): boolean {
  const expected = config.password();
  if (!expected) return false; // no password set = nobody gets in
  const a = createHash("sha256").update(String(given ?? ""), "utf8").digest();
  const b = createHash("sha256").update(expected, "utf8").digest();
  return timingSafeEqual(a, b);
}

/** true = this IP may try now (and the try is counted) */
export function allowLoginTry(ip: string, now = Date.now()): boolean {
  const recent = (loginTries.get(ip) ?? []).filter((t) => now - t < LOGIN_WINDOW_MS);
  if (recent.length >= LOGIN_MAX) {
    loginTries.set(ip, recent);
    return false;
  }
  recent.push(now);
  loginTries.set(ip, recent);
  return true;
}

export function createSession(now = Date.now()): string {
  const sid = randomBytes(32).toString("hex");
  sessions.set(sid, now + SESSION_MS);
  return sid;
}

export function sessionFrom(req: Request, now = Date.now()): string | null {
  const cookie = req.headers.get("cookie") ?? "";
  const m = cookie.match(/(?:^|;\s*)sid=([a-f0-9]{64})(?:;|$)/);
  if (!m) return null;
  const expires = sessions.get(m[1]);
  if (!expires) return null;
  if (expires <= now) {
    sessions.delete(m[1]);
    return null;
  }
  return m[1];
}

export function sessionCookie(sid: string, req: Request): string {
  const secure = req.headers.get("x-forwarded-proto") === "https" ? "; Secure" : "";
  return `sid=${sid}; HttpOnly; SameSite=Lax; Path=/; Max-Age=43200${secure}`;
}

export function _resetAuthForTests() {
  sessions.clear();
  loginTries.clear();
}
