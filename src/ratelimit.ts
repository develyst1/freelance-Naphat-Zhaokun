// Chat limit: RATE_LIMIT_PER_HOUR POST /api/chat per session, rolling hour.
import { config } from "./config.ts";

const HOUR_MS = 60 * 60 * 1000;
const hits = new Map<string, number[]>(); // sid -> times

function recent(sid: string, now: number): number[] {
  const list = (hits.get(sid) ?? []).filter((t) => now - t < HOUR_MS);
  hits.set(sid, list);
  return list;
}

export function remaining(sid: string, now = Date.now()): number {
  return Math.max(0, config.ratePerHour() - recent(sid, now).length);
}

/** Counts one chat if allowed. Returns null when allowed, else the time the oldest hit leaves the window. */
export function takeChat(sid: string, now = Date.now()): Date | null {
  const list = recent(sid, now);
  if (list.length >= config.ratePerHour()) return new Date(list[0] + HOUR_MS);
  list.push(now);
  return null;
}

export function _resetRateForTests() {
  hits.clear();
}
