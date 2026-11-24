// TASK-014 (D4): a big file request must make a real file; raw JSON never reaches the reply; a file asked for but not
// sent gets one retry. Mocked gateway, no network.
import { afterAll, beforeAll, expect, test } from "bun:test";
import { chatWithFiles, extractSpec, FILE_FAIL_LINE, SHORTER_HINT, stripSpecText } from "../src/llm.ts";
import { readOffice } from "../src/office/read.ts";

const realFetch = globalThis.fetch;
let calls: any[] = [];
let replies: string[] = [];

beforeAll(() => {
  process.env.LLM_GATEWAY_URL = "http://gateway.test";
  process.env.LLM_PROVIDER = "";
  process.env.LLM_MODEL = "";
  globalThis.fetch = (async (_url: any, init?: any) => {
    calls.push(JSON.parse(init?.body ?? "{}"));
    const content = replies.shift();
    if (content === undefined) throw new Error("no mocked reply left");
    return new Response(JSON.stringify({ success: true, data: { provider: "mock", model: "m", content, usage: { completion_tokens: 8000 } } }));
  }) as any;
});
afterAll(() => (globalThis.fetch = realFetch));

const SHORT_PPT = JSON.stringify({ type: "pptx", filename: "นำเสนอ", slides: [{ title: "ภาพรวม", bullets: ["ยอดขายเพิ่ม"] }, { title: "สรุป" }] });
const BIG_START = '{"type":"pptx","filename":"ผู้บริหาร","slides":[{"title":"หน้าแรก","bullets":["a"]},{"ti';

function noRawJson(reply: string) {
  expect(reply).not.toContain("file-spec");
  expect(reply).not.toMatch(/\{\s*"type"\s*:\s*"(pptx|docx|xlsx)"/);
}

async function run(userText: string, answers: string[]) {
  calls = [];
  replies = [...answers];
  return chatWithFiles([{ role: "user", content: userText }], []);
}

const UNREPAIRABLE = '{"type":"pptx","filename":"x","slides":[{"title":"หน้าแ';

test("(i-a) a truncated fence that still holds complete slides → repaired in place: a file, one call, no JSON", async () => {
  const out = await run("ทำ ppt ไปเสนอผู้บริหาร", ["นี่คือสไลด์ 17 หน้า\n```file-spec\n" + BIG_START]);
  expect(calls).toHaveLength(1);
  expect(out.created).toHaveLength(1);
  expect(out.reply).toBe("นี่คือสไลด์ 17 หน้า");
  noRawJson(out.reply);
  const text = (await readOffice("pptx", out.created[0].buffer)).text;
  expect(text).toContain("Slide 1: หน้าแรก"); // the complete slide kept, the cut one dropped
  expect(text).not.toContain("Slide 2");
});

test("(i-b) a truncated fence with no complete slide → retry with a short spec → a file, no JSON", async () => {
  const out = await run("ทำ ppt ไปเสนอผู้บริหาร", ["นี่คือสไลด์\n```file-spec\n" + UNREPAIRABLE, "```file-spec\n" + SHORT_PPT + "\n```"]);
  expect(calls).toHaveLength(2);
  expect(calls[1].messages.at(-1).content).toContain("cut off/invalid");
  expect(out.created).toHaveLength(1);
  noRawJson(out.reply);
  expect((await readOffice("pptx", out.created[0].buffer)).text).toContain("Slide 1: ภาพรวม");
});

test("near-JSON with trailing commas is repaired", async () => {
  const out = await run("ทำ ppt", ['```file-spec\n{"type":"pptx","filename":"a","slides":[{"title":"ก","bullets":["x",],},],}\n```']);
  expect(calls).toHaveLength(1);
  expect(out.created).toHaveLength(1);
});

test("(ii) a bare JSON spec with no fence → a file, and the reply has no {\"type\"", async () => {
  const out = await run("ทำ ppt สั้นๆ", ["ได้ครับ " + SHORT_PPT + " เรียบร้อย"]);
  expect(calls).toHaveLength(1);
  expect(out.created[0].kind).toBe("pptx");
  expect(out.reply).toBe("ได้ครับ  เรียบร้อย".replace("  ", "  "));
  noRawJson(out.reply);
});

test("(iii) 'ทำ ppt …' answered with prose only → the retry → a file", async () => {
  const out = await run("ทำ ppt ไปเสนอผู้บริหาร", ["ผมจะทำสไลด์ 17 หน้าให้ครับ หน้า 1 …", "```file-spec\n" + SHORT_PPT + "\n```"]);
  expect(calls).toHaveLength(2);
  expect(calls[1].messages.at(-1).content).toBe("The user asked for a file. Reply ONLY with the ```file-spec block for it.");
  expect(out.created).toHaveLength(1);
  noRawJson(out.reply);
});

test("(iv) 'ไหนละ ppt' + prose → the retry → a file", async () => {
  const out = await run("ไหนละ ppt", ["ขออภัยครับ", "```file-spec\n" + SHORT_PPT + "\n```"]);
  expect(calls).toHaveLength(2);
  expect(out.created).toHaveLength(1);
});

test("(v) both calls fail → FILE_FAIL_LINE + the shorter hint, and no file-spec / {\"type\"", async () => {
  const out = await run("ทำ ppt 20 หน้า", ["```file-spec\n" + UNREPAIRABLE, UNREPAIRABLE + '"x']);
  expect(calls).toHaveLength(2);
  expect(out.created).toHaveLength(0);
  expect(out.reply).toContain(FILE_FAIL_LINE);
  expect(out.reply).toContain(SHORTER_HINT);
  noRawJson(out.reply);
});

test("(vi) a normal chat (no file words) → no retry call", async () => {
  const out = await run("สวัสดีครับ วันนี้อากาศดีไหม", ["สวัสดีครับ"]);
  expect(calls).toHaveLength(1);
  expect(out.reply).toBe("สวัสดีครับ");
});

test("(vii) the request body has max_tokens 8000", async () => {
  await run("สวัสดี", ["หวัดดี"]);
  expect(calls[0].max_tokens).toBe(8000);
});

test("extractSpec / stripSpecText shapes", () => {
  expect(extractSpec("a\n```file-spec\n" + SHORT_PPT + "\n```\nb")).toMatchObject({ found: true, text: "a\n\nb" });
  expect(extractSpec("a ```file-spec\n{\"type\":\"pptx\"")).toMatchObject({ found: true, truncated: true, text: "a" });
  expect(extractSpec(BIG_START)).toMatchObject({ found: true, repaired: true, text: "" }); // cut bare spec, salvaged
  expect(extractSpec(UNREPAIRABLE)).toMatchObject({ found: true, truncated: true, text: "" });
  expect(extractSpec("ไม่มีไฟล์")).toMatchObject({ found: false });
  expect(stripSpecText('ข้อความ {"type":"docx","filename":"a","blocks":[]} ต่อ')).toBe("ข้อความ  ต่อ");
});
