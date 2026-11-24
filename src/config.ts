// Env values, read when asked (tests can set process.env first).
// `.env` lines may carry an inline "# comment" after the value; it is stripped here.
export function env(key: string, fallback = ""): string {
  const raw = process.env[key];
  if (raw === undefined) return fallback;
  const value = raw.replace(/\s+#.*$/, "").trim();
  return value === "" ? fallback : value;
}

export function envInt(key: string, fallback: number): number {
  const n = Number.parseInt(env(key, ""), 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export const config = {
  port: () => envInt("PORT", 3110),
  gatewayUrl: () => env("LLM_GATEWAY_URL", "https://ai.develyst.online").replace(/\/+$/, ""),
  gatewayKey: () => env("LLM_GATEWAY_KEY"),
  provider: () => env("LLM_PROVIDER"),
  model: () => env("LLM_MODEL"),
  password: () => env("DEMO_PASSWORD"),
  ratePerHour: () => envInt("RATE_LIMIT_PER_HOUR", 30),
  maxUploadMb: () => envInt("MAX_UPLOAD_MB", 10),
};
