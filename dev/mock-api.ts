// Dev-only mock of the SPEC-003 §3 API, for building the page without the real server.
// Run: bun dev/mock-api.ts  → http://127.0.0.1:3111  (password: demo)
// Canned data only. No gateway, no real office parsing. The BE's tests do not use this file.

import { join, normalize } from "node:path";

const PORT = 3111;
const PUBLIC_DIR = join(import.meta.dir, "..", "public");
const PASSWORD = "demo";
const LIMIT = 4; // the 5th chat in the hour gets 429
const MAX_UPLOAD_MB = 10;

type Kind = "docx" | "xlsx" | "pptx";
const MIME: Record<Kind, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

const sessions = new Map<string, number[]>(); // sid → chat timestamps
const files = new Map<string, { name: string; kind: Kind }>();
const loginFails = new Map<string, number[]>();

const json = (body: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...headers },
  });

const newId = () => crypto.randomUUID().replace(/-/g, "");

function sidOf(req: Request): string | null {
  const m = /(?:^|;\s*)sid=([^;]+)/.exec(req.headers.get("cookie") ?? "");
  return m && sessions.has(m[1]) ? m[1] : null;
}

function used(sid: string): number[] {
  const hourAgo = Date.now() - 3600_000;
  const list = sessions.get(sid)!.filter((t) => t > hourAgo);
  sessions.set(sid, list);
  return list;
}

function addFile(name: string, kind: Kind) {
  const fileId = newId();
  files.set(fileId, { name, kind });
  return { fileId, name, kind, url: `/api/files/${fileId}` };
}

const SAMPLE_ROWS: (string | number | null)[][] = [
  ["สินค้า", "จำนวน", "ราคา", "รวม"],
  ["กาแฟ", 2, 150, 300],
  ["ชาเขียว", 3, 120, 360],
  ["ขนมปัง", 5, 45, 225],
  ["รวม", null, null, 885],
];

function replyFor(text: string, sid: string) {
  const t = text.toLowerCase();
  const files: ReturnType<typeof addFile>[] = [];
  const tables: { fileId: string; sheet: string; rows: (string | number | null)[][] }[] = [];
  let reply =
    "นี่คือคำตอบจาก **mock** ครับ\n\n- ข้อแรก\n- ข้อสอง\n\nโค้ดตัวอย่าง `=SUM(D2:D4)`\n\n```\nบรรทัดหนึ่ง\nบรรทัดสอง\n```";
  if (t.includes("excel")) {
    const f = addFile("สรุปยอดขาย.xlsx", "xlsx");
    files.push(f);
    tables.push({ fileId: f.fileId, sheet: "Sales", rows: SAMPLE_ROWS });
    reply = "สร้างไฟล์ Excel ให้แล้ว ตัวเลขในตารางคำนวณด้วยสูตร";
  } else if (t.includes("ppt")) {
    files.push(addFile("นำเสนอ.pptx", "pptx"));
    reply = "สร้างไฟล์นำเสนอให้แล้ว";
  } else if (t.includes("doc")) {
    files.push(addFile("รายงาน.docx", "docx"));
    reply = "สร้างเอกสารให้แล้ว";
  } else if (/\bmd\b/.test(t)) {
    reply = [
      "# เปรียบเทียบเครื่องมือ",
      "",
      "| เครื่องมือ | จุดเด่น | หมายเหตุ |",
      "|---|:---:|---|",
      "| Word | **เอกสาร** ยาว | ใช้ง่าย<br>พิมพ์รายงาน |",
      "| Excel | ตาราง `=SUM()` | สูตร<br/>กราฟ |",
      "| PowerPoint | สไลด์ |",
      "",
      "1. ข้อหนึ่ง",
      "2. ข้อสอง",
      "",
      "- หัวข้อย่อย ก",
      "- หัวข้อย่อย ข",
      "",
      "ใช้ `inline code` ได้",
      "",
      "```",
      "| ในโค้ด | ไม่ใช่ตาราง |",
      "```",
      "",
      "<img src=x onerror=alert(1)> <b>bold?</b>",
      "",
      "| ตารางปลอม | ไม่มีเส้นคั่น |",
      "| a | b |",
    ].join("\n");
  } else if (t.includes("xss")) {
    reply = 'ทดสอบ <img src=x onerror=alert(1)> และ <script>alert(2)</script> **ตัวหนา**';
  }
  return { reply, files, tables, remaining: LIMIT - used(sid).length };
}

function hhmm(d: Date) {
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

async function api(req: Request, url: URL): Promise<Response> {
  const p = url.pathname;

  if (p === "/api/login" && req.method === "POST") {
    const ip = "local";
    const fails = (loginFails.get(ip) ?? []).filter((t) => t > Date.now() - 15 * 60_000);
    if (fails.length >= 10) return json({ error: "ลองผิดหลายครั้งเกินไป รออีก 15 นาที" }, 429);
    const body = await req.json().catch(() => ({}));
    if (body?.password !== PASSWORD) {
      loginFails.set(ip, [...fails, Date.now()]);
      return json({ error: "รหัสผ่านไม่ถูกต้อง" }, 401);
    }
    const sid = newId();
    sessions.set(sid, []);
    return json({ ok: true }, 200, {
      "Set-Cookie": `sid=${sid}; HttpOnly; SameSite=Lax; Path=/; Max-Age=43200`,
    });
  }

  const sid = sidOf(req);
  if (!sid) return json({ error: "กรุณาใส่รหัสผ่าน" }, 401);

  if (p === "/api/me" && req.method === "GET") {
    return json({ ok: true, remaining: LIMIT - used(sid).length, limit: LIMIT });
  }

  if (p === "/api/upload" && req.method === "POST") {
    const form = await req.formData().catch(() => null);
    const file = form?.get("file");
    if (!(file instanceof File)) return json({ error: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx" }, 400);
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext !== "docx" && ext !== "xlsx" && ext !== "pptx")
      return json({ error: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx" }, 400);
    if (file.size > MAX_UPLOAD_MB * 1024 * 1024)
      return json({ error: `ไฟล์ใหญ่เกิน ${MAX_UPLOAD_MB} MB` }, 413);
    const { fileId, name, kind } = addFile(file.name, ext);
    return json({
      fileId,
      name,
      kind,
      preview: "ข้อความตัวอย่างจากไฟล์ (mock) …",
      sheets: kind === "xlsx" ? [{ name: "Sheet1", rows: SAMPLE_ROWS }] : [],
    });
  }

  if (p === "/api/chat" && req.method === "POST") {
    const body = await req.json().catch(() => null);
    const messages = Array.isArray(body?.messages) ? body.messages : [];
    const fileIds: string[] = Array.isArray(body?.fileIds) ? body.fileIds : [];
    if (fileIds.some((id) => !files.has(id))) return json({ error: "ไฟล์หมดอายุแล้ว อัปโหลดใหม่" }, 400);
    const list = used(sid);
    if (list.length >= LIMIT) {
      const retry = new Date(list[0] + 3600_000);
      return json(
        { error: `ใช้ครบ ${LIMIT} ข้อความต่อชั่วโมงแล้ว ลองใหม่เวลา ${hhmm(retry)}`, retryAt: retry.toISOString() },
        429,
      );
    }
    const last = String(messages.at(-1)?.content ?? "");
    await Bun.sleep(600); // let the waiting bubble show
    if (last.toLowerCase().includes("down"))
      return new Response("<html><body>Error 530</body></html>", { status: 530, headers: { "Content-Type": "text/html" } });
    if (last.toLowerCase().includes("error")) return json({ error: "ระบบ AI ไม่ตอบ ลองใหม่อีกครั้ง" }, 502);
    list.push(Date.now());
    console.log(`chat: ${messages.length} message(s), ${fileIds.length} file(s)`);
    return json(replyFor(last, sid));
  }

  const m = /^\/api\/files\/([A-Za-z0-9]+)$/.exec(p);
  if (m && req.method === "GET") {
    const f = files.get(m[1]);
    if (!f) return json({ error: "ไม่พบไฟล์ หรือไฟล์หมดอายุแล้ว" }, 404);
    return new Response(`mock ${f.kind} file`, {
      headers: {
        "Content-Type": MIME[f.kind],
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(f.name)}`,
      },
    });
  }

  return json({ error: "not found" }, 404);
}

Bun.serve({
  hostname: "127.0.0.1",
  port: PORT,
  async fetch(req) {
    const url = new URL(req.url);
    if (url.pathname.startsWith("/api/")) return api(req, url);
    const rel = url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname.slice(1));
    const full = normalize(join(PUBLIC_DIR, rel));
    if (!full.startsWith(PUBLIC_DIR)) return new Response("not found", { status: 404 });
    const f = Bun.file(full);
    return (await f.exists()) ? new Response(f) : new Response("not found", { status: 404 });
  },
});

console.log(`mock API on http://127.0.0.1:${PORT} (password: ${PASSWORD})`);
