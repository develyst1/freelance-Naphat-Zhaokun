// Bun server: static public/ + the /api/* contract of SPEC-003 §3. Listens on 127.0.0.1 only.
import { existsSync, statSync } from "node:fs";
import { join, normalize, sep } from "node:path";
import { allowLoginTry, createSession, passwordMatches, sessionCookie, sessionFrom } from "./auth.ts";
import { config } from "./config.ts";
import { chatWithFiles, GatewayError, type Msg } from "./llm.ts";
import { readOffice } from "./office/read.ts";
import { hasZipMagic, kindFromName } from "./office/zipguard.ts";
import { remaining, takeChat } from "./ratelimit.ts";
import { getFile, MIME, putFile, startSweeper, type StoredFile } from "./store.ts";

const PUBLIC_DIR = join(import.meta.dir, "..", "public");

// Far above MAX_UPLOAD_MB so the handler's JSON 413 always answers. Bun's own cap (hit first before) closed the
// connection with no JSON, which the page showed as a network error (TASK-006 D1).
export const MAX_REQUEST_BODY = 64 * 1024 * 1024;

const T = {
  badPassword: "รหัสผ่านไม่ถูกต้อง",
  tooManyLogins: "ลองผิดหลายครั้งเกินไป รออีก 15 นาที",
  needLogin: "กรุณาใส่รหัสผ่าน",
  badType: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx",
  tooBig: (mb: number) => `ไฟล์ใหญ่เกิน ${mb} MB`,
  rate: (n: number, hhmm: string) => `ใช้ครบ ${n} ข้อความต่อชั่วโมงแล้ว ลองใหม่เวลา ${hhmm}`,
  aiDown: "ระบบ AI ไม่ตอบ ลองใหม่อีกครั้ง",
  fileExpired: "ไฟล์หมดอายุแล้ว อัปโหลดใหม่",
  fileNotFound: "ไม่พบไฟล์ หรือไฟล์หมดอายุแล้ว",
  badRequest: "คำขอไม่ถูกต้อง",
};

function json(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "x-content-type-options": "nosniff", ...headers },
  });
}
const err = (status: number, message: string, extra: Record<string, unknown> = {}) => json({ error: message, ...extra }, status);

function clientIp(req: Request, server?: { requestIP(r: Request): { address: string } | null }): string {
  return req.headers.get("cf-connecting-ip") || server?.requestIP(req)?.address || "unknown";
}

function hhmm(d: Date): string {
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

// ---------- routes ----------

async function login(req: Request, server?: any): Promise<Response> {
  if (!allowLoginTry(clientIp(req, server))) return err(429, T.tooManyLogins);
  let body: any = null;
  try {
    body = await req.json();
  } catch {
    /* treated as a wrong password */
  }
  if (!passwordMatches(typeof body?.password === "string" ? body.password : "")) return err(401, T.badPassword);
  const sid = createSession();
  return json({ ok: true }, 200, { "set-cookie": sessionCookie(sid, req) });
}

function me(sid: string): Response {
  return json({ ok: true, remaining: remaining(sid), limit: config.ratePerHour(), maxUploadMb: config.maxUploadMb() });
}

async function upload(req: Request, sid: string): Promise<Response> {
  const maxMb = config.maxUploadMb();
  const declared = Number(req.headers.get("content-length") ?? 0);
  if (declared > maxMb * 1024 * 1024 + 64 * 1024) return err(413, T.tooBig(maxMb));
  let form: FormData;
  try {
    form = await req.formData();
  } catch {
    return err(400, T.badType);
  }
  const file = form.get("file");
  if (!file || typeof file === "string") return err(400, T.badType);
  if (file.size > maxMb * 1024 * 1024) return err(413, T.tooBig(maxMb));
  const kind = kindFromName(file.name);
  if (!kind) return err(400, T.badType);
  const buf = new Uint8Array(await file.arrayBuffer());
  if (!hasZipMagic(buf)) return err(400, T.badType);
  let read;
  try {
    read = await readOffice(kind, buf);
  } catch (e) {
    console.error("upload read failed:", (e as Error).message);
    return err(400, T.badType);
  }
  const stored = putFile({ name: file.name, kind, buffer: buf, text: read.text, sid });
  const body: Record<string, unknown> = { fileId: stored.fileId, name: stored.name, kind, preview: read.preview };
  if (kind === "xlsx") body.sheets = read.sheets ?? [];
  return json(body);
}

function cleanMessages(raw: unknown): Msg[] | null {
  if (!Array.isArray(raw) || raw.length === 0) return null;
  const msgs: Msg[] = [];
  for (const m of raw) {
    if (!m || (m.role !== "user" && m.role !== "assistant") || typeof m.content !== "string") return null;
    msgs.push({ role: m.role, content: m.content.slice(0, 20000) });
  }
  if (msgs[msgs.length - 1].role !== "user") return null;
  return msgs.slice(-20);
}

async function chat(req: Request, sid: string): Promise<Response> {
  let body: any;
  try {
    body = await req.json();
  } catch {
    return err(400, T.badRequest);
  }
  const messages = cleanMessages(body?.messages);
  if (!messages) return err(400, T.badRequest);
  const ids: unknown = body?.fileIds ?? [];
  if (!Array.isArray(ids) || !ids.every((x) => typeof x === "string")) return err(400, T.badRequest);
  const files: StoredFile[] = [];
  for (const id of ids as string[]) {
    const f = getFile(id, sid);
    if (!f) return err(400, T.fileExpired);
    files.push(f);
  }

  const retryAt = takeChat(sid);
  if (retryAt) return err(429, T.rate(config.ratePerHour(), hhmm(retryAt)), { retryAt: retryAt.toISOString() });

  let outcome;
  try {
    outcome = await chatWithFiles(messages, files);
  } catch (e) {
    if (e instanceof GatewayError) {
      console.error("[chat]", e.message);
      return err(502, T.aiDown);
    }
    throw e;
  }
  console.log(`[chat] provider=${outcome.provider ?? "?"} model=${outcome.model ?? "?"} files=${outcome.created.length}`);

  const outFiles: { fileId: string; name: string; kind: string; url: string }[] = [];
  const tables: { fileId: string; sheet: string; rows: (string | number | null)[][] }[] = [];
  for (const c of outcome.created) {
    const stored = putFile({ name: c.name, kind: c.kind, buffer: c.buffer, text: "", sid });
    outFiles.push({ fileId: stored.fileId, name: stored.name, kind: stored.kind, url: `/api/files/${stored.fileId}` });
    for (const t of c.tables ?? []) tables.push({ fileId: stored.fileId, sheet: t.name, rows: t.rows });
  }
  return json({ reply: outcome.reply, files: outFiles, tables, remaining: remaining(sid) });
}

function download(fileId: string, sid: string): Response {
  const f = getFile(fileId, sid);
  if (!f) return err(404, T.fileNotFound);
  return new Response(f.buffer, {
    headers: {
      "content-type": MIME[f.kind],
      "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(f.name)}`,
      "x-content-type-options": "nosniff",
    },
  });
}

function serveStatic(pathname: string): Response {
  let rel: string;
  try {
    rel = decodeURIComponent(pathname);
  } catch {
    return new Response("Not found", { status: 404 });
  }
  if (rel === "/" || rel === "") rel = "/index.html";
  const full = normalize(join(PUBLIC_DIR, rel));
  if (!full.startsWith(PUBLIC_DIR + sep)) return new Response("Not found", { status: 404 });
  if (!existsSync(full) || !statSync(full).isFile()) return new Response("Not found", { status: 404 }); // no directory listing
  return new Response(Bun.file(full));
}

// ---------- entry ----------

export async function handle(req: Request, server?: any): Promise<Response> {
  const url = new URL(req.url);
  const path = url.pathname;
  if (!path.startsWith("/api/")) {
    if (req.method !== "GET" && req.method !== "HEAD") return new Response("Method not allowed", { status: 405 });
    return serveStatic(path);
  }
  try {
    if (path === "/api/login" && req.method === "POST") return await login(req, server);
    const sid = sessionFrom(req);
    if (path === "/api/me" && req.method === "GET") return sid ? me(sid) : err(401, T.needLogin);
    if (path === "/api/upload" && req.method === "POST") return sid ? await upload(req, sid) : err(401, T.needLogin);
    if (path === "/api/chat" && req.method === "POST") return sid ? await chat(req, sid) : err(401, T.needLogin);
    const m = path.match(/^\/api\/files\/([a-f0-9]{32})$/);
    if (m && req.method === "GET") return sid ? download(m[1], sid) : err(401, T.needLogin);
    if (path.startsWith("/api/files/") && req.method === "GET") return sid ? err(404, T.fileNotFound) : err(401, T.needLogin);
    return err(404, T.badRequest);
  } catch (e) {
    console.error("[api] unexpected:", (e as Error).message);
    return err(500, "เกิดข้อผิดพลาด ลองใหม่อีกครั้ง");
  }
}

if (import.meta.main) {
  if (!config.password()) console.warn("DEMO_PASSWORD ยังว่าง — รัน `bun run setup` ก่อน (no one can log in)");
  const server = Bun.serve({
    hostname: "127.0.0.1",
    port: config.port(),
    maxRequestBodySize: MAX_REQUEST_BODY,
    fetch: handle,
  });
  startSweeper();
  console.log(`web: http://127.0.0.1:${server.port}`);
}
