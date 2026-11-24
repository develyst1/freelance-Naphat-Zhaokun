// TASK-008: big uploads get the JSON 413 (D1), /api/me exposes maxUploadMb, and the prompt names Excel rows (N1).
import { afterAll, beforeAll, expect, test } from "bun:test";
import { handle, MAX_REQUEST_BODY } from "../src/server.ts";
import { SYSTEM_PROMPT } from "../src/llm.ts";

const MB = 1024 * 1024;
let server: ReturnType<typeof Bun.serve> | undefined;

beforeAll(() => {
  process.env.DEMO_PASSWORD = "test-pass-123";
  process.env.MAX_UPLOAD_MB = "10";
});
afterAll(() => server?.stop(true));

async function cookie(ip: string): Promise<string> {
  const res = await handle(new Request("http://127.0.0.1/api/login", {
    method: "POST", headers: { "cf-connecting-ip": ip }, body: JSON.stringify({ password: "test-pass-123" }),
  }));
  return res.headers.get("set-cookie")!.split(";")[0];
}

function zipLike(size: number): Uint8Array {
  const b = new Uint8Array(size);
  b.set([0x50, 0x4b, 0x03, 0x04]);
  return b;
}

for (const mb of [11, 30]) {
  test(`handle(): a ${mb} MB upload → 413 with the JSON body`, async () => {
    process.env.MAX_UPLOAD_MB = "10";
    const form = new FormData();
    form.append("file", new File([zipLike(mb * MB)], "big.xlsx"));
    const res = await handle(new Request("http://127.0.0.1/api/upload", {
      method: "POST", headers: { cookie: await cookie(`8.8.8.${mb}`) }, body: form,
    }));
    expect(res.status).toBe(413);
    expect(await res.json()).toEqual({ error: "ไฟล์ใหญ่เกิน 10 MB" });
  });
}

test("a real Bun.serve with the server's body cap: an 11 MB multipart upload → HTTP 413 + JSON", async () => {
  process.env.MAX_UPLOAD_MB = "10";
  server = Bun.serve({ hostname: "127.0.0.1", port: 0, maxRequestBodySize: MAX_REQUEST_BODY, fetch: handle });
  const form = new FormData();
  form.append("file", new File([zipLike(11 * MB)], "big.xlsx"));
  const res = await fetch(`http://127.0.0.1:${server.port}/api/upload`, {
    method: "POST", headers: { cookie: await cookie("8.8.4.4") }, body: form,
  });
  expect(res.status).toBe(413);
  expect(await res.json()).toEqual({ error: "ไฟล์ใหญ่เกิน 10 MB" });
  expect(MAX_REQUEST_BODY).toBe(64 * MB);
});

test("/api/me carries maxUploadMb equal to env", async () => {
  process.env.MAX_UPLOAD_MB = "10";
  const res = await handle(new Request("http://127.0.0.1/api/me", { headers: { cookie: await cookie("8.8.8.1") } }));
  const body = await res.json();
  expect(Object.keys(body)).toEqual(["ok", "remaining", "limit", "maxUploadMb"]);
  expect(body.maxUploadMb).toBe(10);
});

test("the system prompt has the Excel row-numbering line (N1)", () => {
  expect(SYSTEM_PROMPT).toContain(
    "แถวแรกของ rows คือแถวที่ 1 ใน Excel — เวลาอธิบายสูตร ให้อ้างชื่อเซลล์ตามการนับนี้ / The first entry of `rows` is Excel row 1; name cells by that count.",
  );
});
