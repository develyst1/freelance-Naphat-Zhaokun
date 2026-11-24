// (c) the HTTP API with a mocked gateway fetch, (d) exact response keys (SPEC-003 §3). No network.
import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import { existsSync, readFileSync } from "node:fs";
import { _resetAuthForTests } from "../src/auth.ts";
import { _resetRateForTests } from "../src/ratelimit.ts";
import { handle } from "../src/server.ts";
import { FIXTURE, makeKnownFixture } from "./make-fixture.ts";

// Inline env: the tests never depend on a local .env (Bun loads it by itself).
process.env.DEMO_PASSWORD = "test-pass-123";
process.env.RATE_LIMIT_PER_HOUR = "3";
process.env.MAX_UPLOAD_MB = "1";
process.env.LLM_GATEWAY_URL = "http://gateway.test";
process.env.LLM_GATEWAY_KEY = "";
process.env.LLM_PROVIDER = "";
process.env.LLM_MODEL = "";

const realFetch = globalThis.fetch;
type Call = { url: string; body: any };
let calls: Call[] = [];
let replies: (() => Response)[] = [];

function gatewayReply(content: string): () => Response {
  return () =>
    new Response(JSON.stringify({ success: true, data: { provider: "mock", model: "mock-1", content } }), {
      status: 200,
      headers: { "content-type": "application/json" },
    });
}

beforeAll(async () => {
  if (!existsSync(FIXTURE)) await makeKnownFixture();
  globalThis.fetch = (async (input: any, init?: any) => {
    const url = String(input instanceof Request ? input.url : input);
    if (!url.startsWith("http://gateway.test/")) throw new Error("unexpected network call in a test: " + url);
    calls.push({ url, body: JSON.parse(init?.body ?? "{}") });
    const next = replies.shift();
    if (!next) throw new Error("no mocked gateway reply left");
    return next();
  }) as any;
});
afterAll(() => {
  globalThis.fetch = realFetch;
});

let ipCounter = 0;
const req = (path: string, init: RequestInit = {}) => new Request(`http://127.0.0.1:3110${path}`, init);

async function login(): Promise<string> {
  const res = await handle(req("/api/login", {
    method: "POST",
    headers: { "content-type": "application/json", "cf-connecting-ip": `10.0.0.${++ipCounter}` },
    body: JSON.stringify({ password: "test-pass-123" }),
  }));
  expect(res.status).toBe(200);
  return res.headers.get("set-cookie")!.split(";")[0];
}

async function uploadFile(cookie: string, name: string, bytes: Uint8Array) {
  const form = new FormData();
  form.append("file", new File([bytes], name));
  return handle(req("/api/upload", { method: "POST", headers: { cookie }, body: form }));
}

async function chat(cookie: string, text: string, fileIds: string[] = []) {
  return handle(req("/api/chat", {
    method: "POST",
    headers: { cookie, "content-type": "application/json" },
    body: JSON.stringify({ messages: [{ role: "user", content: text }], fileIds }),
  }));
}

const noCors = (res: Response) =>
  expect([...res.headers.keys()].filter((k) => k.toLowerCase().startsWith("access-control-allow"))).toEqual([]);

describe("(c) auth", () => {
  test("login 401 with the Thai message, 200 with cookie flags", async () => {
    const bad = await handle(req("/api/login", { method: "POST", headers: { "cf-connecting-ip": "1.1.1.1" }, body: JSON.stringify({ password: "nope" }) }));
    expect(bad.status).toBe(401);
    expect(await bad.json()).toEqual({ error: "รหัสผ่านไม่ถูกต้อง" });
    noCors(bad);

    const ok = await handle(req("/api/login", {
      method: "POST",
      headers: { "cf-connecting-ip": "1.1.1.2", "x-forwarded-proto": "https" },
      body: JSON.stringify({ password: "test-pass-123" }),
    }));
    expect(ok.status).toBe(200);
    expect(Object.keys(await ok.json())).toEqual(["ok"]);
    const cookie = ok.headers.get("set-cookie")!;
    expect(cookie).toMatch(/^sid=[a-f0-9]{64}; HttpOnly; SameSite=Lax; Path=\/; Max-Age=43200; Secure$/);
  });

  test("login throttle: the 11th try in 15 min from one IP is 429", async () => {
    let last: Response | undefined;
    for (let i = 0; i < 11; i++)
      last = await handle(req("/api/login", { method: "POST", headers: { "cf-connecting-ip": "9.9.9.9" }, body: JSON.stringify({ password: "x" }) }));
    expect(last!.status).toBe(429);
    expect(await last!.json()).toEqual({ error: "ลองผิดหลายครั้งเกินไป รออีก 15 นาที" });
  });

  test("/api/me 401 without the cookie; keys with it", async () => {
    const no = await handle(req("/api/me"));
    expect(no.status).toBe(401);
    expect(await no.json()).toEqual({ error: "กรุณาใส่รหัสผ่าน" });
    const cookie = await login();
    const yes = await handle(req("/api/me", { headers: { cookie } }));
    const body = await yes.json();
    expect(Object.keys(body)).toEqual(["ok", "remaining", "limit", "maxUploadMb"]);
    expect(body).toEqual({ ok: true, remaining: 3, limit: 3, maxUploadMb: Number(process.env.MAX_UPLOAD_MB) });
  });
});

describe("(c) upload", () => {
  test("rejects .txt and a renamed non-zip with 400", async () => {
    const cookie = await login();
    const txt = await uploadFile(cookie, "note.txt", new TextEncoder().encode("hello"));
    expect(txt.status).toBe(400);
    expect(await txt.json()).toEqual({ error: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx" });
    const fake = await uploadFile(cookie, "fake.xlsx", new TextEncoder().encode("not a zip at all"));
    expect(fake.status).toBe(400);
    expect(await fake.json()).toEqual({ error: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx" });
  });

  test("too big → 413", async () => {
    process.env.MAX_UPLOAD_MB = "1";
    const cookie = await login();
    const big = new Uint8Array(1024 * 1024 + 10);
    big.set([0x50, 0x4b, 0x03, 0x04]);
    const res = await uploadFile(cookie, "big.xlsx", big);
    expect(res.status).toBe(413);
    expect(await res.json()).toEqual({ error: "ไฟล์ใหญ่เกิน 1 MB" });
  });

  test("xlsx upload → exact keys and COMPUTED sheet values", async () => {
    const cookie = await login();
    const res = await uploadFile(cookie, "known.xlsx", new Uint8Array(readFileSync(FIXTURE)));
    expect(res.status).toBe(200);
    const body = await res.json();
    expect(Object.keys(body)).toEqual(["fileId", "name", "kind", "preview", "sheets"]);
    expect(body.kind).toBe("xlsx");
    expect(body.sheets[0].name).toBe("Data");
    expect(body.sheets[0].rows[4][3]).toBe(253.5); // D5, hand-worked in make-fixture.ts
    expect(body.sheets[1].rows[2][1]).toBeCloseTo(278.85, 9);
  });

  test("upload without the cookie → 401", async () => {
    const res = await uploadFile("", "a.xlsx", new Uint8Array([1]));
    expect(res.status).toBe(401);
  });
});

describe("(c) chat via the mocked gateway", () => {
  test("valid xlsx file-spec → file downloads; tables hold the computed numbers; exact keys", async () => {
    const cookie = await login();
    const up = await (await uploadFile(cookie, "known.xlsx", new Uint8Array(readFileSync(FIXTURE)))).json();
    calls = [];
    replies = [gatewayReply(
      "สร้างไฟล์ให้แล้วครับ\n```file-spec\n" +
        JSON.stringify({ type: "xlsx", filename: "สรุป.xlsx", sheets: [{ name: "S", rows: [["รายการ", "บาท"], ["ก", 100], ["ข", 250], ["รวม", "=SUM(B2:B3)"]] }] }) +
        "\n```",
    )];
    const res = await chat(cookie, "ทำ excel ให้หน่อย", [up.fileId]);
    expect(res.status).toBe(200);
    noCors(res);
    const body = await res.json();
    expect(Object.keys(body)).toEqual(["reply", "files", "tables", "remaining"]);
    expect(body.reply).toBe("สร้างไฟล์ให้แล้วครับ");
    expect(body.remaining).toBe(2);
    expect(Object.keys(body.files[0])).toEqual(["fileId", "name", "kind", "url"]);
    expect(body.files[0]).toMatchObject({ name: "สรุป.xlsx", kind: "xlsx" });
    expect(Object.keys(body.tables[0])).toEqual(["fileId", "sheet", "rows"]);
    expect(body.tables[0].rows[3][1]).toBe(350); // 100 + 250, hand-worked

    // the gateway got the system prompt + the uploaded file's computed text
    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe("http://gateway.test/chat");
    expect(calls[0].body).toMatchObject({ temperature: 0.3, max_tokens: 4000 });
    expect(calls[0].body.provider).toBeUndefined();
    const sys = calls[0].body.messages.filter((m: any) => m.role === "system").map((m: any) => m.content).join("\n");
    expect(sys).toContain('<file name="known.xlsx" kind="xlsx">');
    expect(sys).toContain("D5 =SUM(D2:D4) → 253.5");

    const dl = await handle(req(body.files[0].url, { headers: { cookie } }));
    expect(dl.status).toBe(200);
    expect(dl.headers.get("content-type")).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    expect(dl.headers.get("content-disposition")).toBe(`attachment; filename*=UTF-8''${encodeURIComponent("สรุป.xlsx")}`);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(await dl.arrayBuffer()) as any);
    expect((wb.getWorksheet("S")!.getCell("B4").value as any)).toMatchObject({ formula: "SUM(B2:B3)", result: 350 });

    // another session cannot download it
    const other = await login();
    expect((await handle(req(body.files[0].url, { headers: { cookie: other } }))).status).toBe(404);
  });

  test("invalid file-spec twice → one retry, then the failure line", async () => {
    const cookie = await login();
    calls = [];
    replies = [
      gatewayReply('นี่ครับ\n```file-spec\n{"type":"xlsx","filename":"a.xlsx","sheets":[}\n```'),
      gatewayReply('```file-spec\n{"type":"pdf","filename":"a"}\n```'),
    ];
    const res = await chat(cookie, "ทำไฟล์");
    expect(res.status).toBe(200);
    const body = await res.json();
    expect(calls).toHaveLength(2);
    expect(calls[1].body.messages.at(-1).content).toContain("your file-spec was invalid");
    expect(body.reply).toContain("สร้างไฟล์ไม่สำเร็จ");
    expect(body.files).toEqual([]);
  });

  test("an invalid spec fixed on the retry → the file is built", async () => {
    const cookie = await login();
    replies = [
      gatewayReply('ได้ครับ\n```file-spec\n{"type":"docx","filename":"a","blocks":[{"h1":"x","p":"y"}]}\n```'),
      gatewayReply('```file-spec\n{"type":"docx","filename":"รายงาน","blocks":[{"h1":"หัวข้อ"},{"p":"เนื้อหา"}]}\n```'),
    ];
    const body = await (await chat(cookie, "ทำ word")).json();
    expect(body.reply).toBe("ได้ครับ");
    expect(body.files[0]).toMatchObject({ name: "รายงาน.docx", kind: "docx" });
  });

  test("gateway 500 → 502 with the Thai message", async () => {
    const cookie = await login();
    replies = [() => new Response(JSON.stringify({ success: false, error: "all providers failed" }), { status: 500 })];
    const res = await chat(cookie, "สวัสดี");
    expect(res.status).toBe(502);
    expect(await res.json()).toEqual({ error: "ระบบ AI ไม่ตอบ ลองใหม่อีกครั้ง" });
  });

  test("empty content → 502", async () => {
    const cookie = await login();
    replies = [gatewayReply("   ")];
    expect((await chat(cookie, "สวัสดี")).status).toBe(502);
  });

  test("unknown fileId → 400 file expired", async () => {
    const cookie = await login();
    const res = await chat(cookie, "อ่านไฟล์", ["0".repeat(32)]);
    expect(res.status).toBe(400);
    expect(await res.json()).toEqual({ error: "ไฟล์หมดอายุแล้ว อัปโหลดใหม่" });
  });

  test("429 after RATE_LIMIT_PER_HOUR (3) chats, with retryAt", async () => {
    const cookie = await login();
    for (let i = 0; i < 3; i++) {
      replies = [gatewayReply("ตอบ " + i)];
      expect((await chat(cookie, "ข้อความ " + i)).status).toBe(200);
    }
    const res = await chat(cookie, "เกิน");
    expect(res.status).toBe(429);
    const body = await res.json();
    expect(Object.keys(body)).toEqual(["error", "retryAt"]);
    expect(body.error).toMatch(/^ใช้ครบ 3 ข้อความต่อชั่วโมงแล้ว ลองใหม่เวลา \d\d:\d\d$/);
    expect(new Date(body.retryAt).getTime()).toBeGreaterThan(Date.now());
    const me = await (await handle(req("/api/me", { headers: { cookie } }))).json();
    expect(me.remaining).toBe(0);
  });

  test("chat without the cookie → 401; unknown file download → 404", async () => {
    expect((await chat("", "x")).status).toBe(401);
    const cookie = await login();
    const res = await handle(req("/api/files/" + "a".repeat(32), { headers: { cookie } }));
    expect(res.status).toBe(404);
    expect(await res.json()).toEqual({ error: "ไม่พบไฟล์ หรือไฟล์หมดอายุแล้ว" });
  });
});

afterAll(() => {
  _resetAuthForTests();
  _resetRateForTests();
});
