// TASK-013 (D3): a synthetic xlsx shaped like the failing real one (no customer data) — a drawing + chart on one sheet,
// a legacy comment + VML on another, with ABSOLUTE relationship targets ("/xl/drawings/drawing1.xml"), which is what made
// ExcelJS crash with `drawing.anchors`. Unsanitized ExcelJS must throw; readOffice must succeed.
import { beforeAll, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { readOffice } from "../src/office/read.ts";
import { sanitizeXlsxForRead } from "../src/office/sanitize.ts";
import { handle } from "../src/server.ts";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
let tricky: Uint8Array;

async function makeTricky(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet("Data");
  a.addRows([["ช่องทาง", "ยอด"], ["Facebook", 1200], ["LINE", 800], ["รวม", { formula: "SUM(B2:B3)" }]]);
  const b = wb.addWorksheet("Chart");
  b.addRows([["สรุป"], [1]]);
  const zip = await JSZip.loadAsync(new Uint8Array(await wb.xlsx.writeBuffer()));

  // sheet1: a legacy comment (B2) + its VML; sheet2: a drawing holding one chart — absolute targets, as in the real file
  const s1 = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
  zip.file("xl/worksheets/sheet1.xml", s1.replace("</worksheet>", `<legacyDrawing xmlns:r="${REL}" r:id="rIdV"/></worksheet>`));
  const s2 = await zip.file("xl/worksheets/sheet2.xml")!.async("string");
  zip.file("xl/worksheets/sheet2.xml", s2.replace("</worksheet>", `<drawing xmlns:r="${REL}" r:id="rIdD"/></worksheet>`));
  zip.file("xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rIdC" Type="${REL}/comments" Target="/xl/comments/comment1.xml" />` +
    `<Relationship Id="rIdV" Type="${REL}/vmlDrawing" Target="/xl/drawings/commentsDrawing1.vml" /></Relationships>`);
  zip.file("xl/worksheets/_rels/sheet2.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rIdD" Type="${REL}/drawing" Target="/xl/drawings/drawing1.xml" /></Relationships>`);
  zip.file("xl/comments/comment1.xml",
    `<?xml version="1.0" encoding="UTF-8"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<authors><author>tester</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>ตัวเลขสมมุติ</t></r></text></comment></commentList></comments>`);
  zip.file("xl/drawings/commentsDrawing1.vml", `<xml xmlns:v="urn:schemas-microsoft-com:vml"></xml>`);
  zip.file("xl/drawings/drawing1.xml",
    `<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"></xdr:wsDr>`);
  zip.file("xl/drawings/_rels/drawing1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/chart" Target="/xl/charts/chart1.xml" /></Relationships>`);
  zip.file("xl/charts/chart1.xml", `<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`);
  const ct = await zip.file("[Content_Types].xml")!.async("string");
  zip.file("[Content_Types].xml", ct.replace("</Types>",
    `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>` +
    `<Override PartName="/xl/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>` +
    `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` +
    `<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

beforeAll(async () => {
  tricky = await makeTricky();
  process.env.DEMO_PASSWORD = "test-pass-123";
  process.env.MAX_UPLOAD_MB = "10";
});

test("unsanitized ExcelJS throws on the synthetic file (the D3 crash)", async () => {
  // Same root cause as the real file (absolute rel targets). On this synthetic file ExcelJS trips on the comments link first
  // (`options.comments[rel.Target].comments`); the real file tripped on the drawing (`drawing.anchors`). Either way: it throws.
  const wb = new ExcelJS.Workbook();
  await expect(wb.xlsx.load(Buffer.from(tricky) as any)).rejects.toThrow(/comments|anchors/);
});

test("readOffice reads the cells + computes; charts 1, comments 1 with its text", async () => {
  const r = await readOffice("xlsx", tricky);
  expect(r.charts).toBe(1);
  expect(r.comments).toBe(1);
  expect(r.sheets!.map((s) => s.name)).toEqual(["Data", "Chart"]);
  expect(r.sheets![0].rows[3]).toEqual(["รวม", 2000]); // 1200 + 800, computed
  expect(r.text).toContain("(ไฟล์นี้มีกราฟ 1 อัน — ระบบอ่านเฉพาะข้อมูลในเซลล์ ไม่อ่านรูปกราฟ)");
  expect(r.text).toContain("B2: ตัวเลขสมมุติ");
});

test("the sanitized copy drops drawings/charts/comments, their rels, sheet tags and overrides; cells kept", async () => {
  const s = await sanitizeXlsxForRead(tricky);
  const z = await JSZip.loadAsync(s.buf);
  const names = Object.keys(z.files);
  expect(names.filter((n) => /drawings|charts|comments/.test(n))).toEqual([]);
  expect(await z.file("xl/worksheets/sheet2.xml")!.async("string")).not.toContain("<drawing");
  expect(await z.file("xl/worksheets/sheet1.xml")!.async("string")).not.toContain("legacyDrawing");
  expect(await z.file("xl/worksheets/_rels/sheet1.xml.rels")!.async("string")).not.toContain("comments");
  expect(await z.file("[Content_Types].xml")!.async("string")).not.toContain("/xl/charts/chart1.xml");
  expect(await z.file("xl/worksheets/sheet1.xml")!.async("string")).toContain("<f>SUM(B2:B3)</f>");
});

async function upload(name: string, bytes: Uint8Array) {
  const login = await handle(new Request("http://127.0.0.1/api/login", {
    method: "POST", headers: { "cf-connecting-ip": "6.6.6." + name.length }, body: JSON.stringify({ password: "test-pass-123" }),
  }));
  const cookie = login.headers.get("set-cookie")!.split(";")[0];
  const form = new FormData();
  form.append("file", new File([bytes], name));
  return handle(new Request("http://127.0.0.1/api/upload", { method: "POST", headers: { cookie }, body: form }));
}

test("upload: the tricky xlsx → 200 with sheets", async () => {
  const res = await upload("tricky.xlsx", tricky);
  expect(res.status).toBe(200);
  const body = await res.json();
  expect(body.sheets[0].rows[3]).toEqual(["รวม", 2000]);
  expect(body.preview).toContain("กราฟ 1 อัน");
});

test("a zip that passes the guard but has a corrupt workbook.xml → 400 readFail, not badType; .txt → badType", async () => {
  const zip = await JSZip.loadAsync(tricky);
  zip.file("xl/workbook.xml", "<workbook><sheets><sheet name=");
  const broken = await zip.generateAsync({ type: "uint8array" });
  const res = await upload("broken.xlsx", broken);
  expect(res.status).toBe(400);
  expect(await res.json()).toEqual({ error: "อ่านไฟล์นี้ไม่ได้ — ลองเปิดใน Excel แล้วบันทึกใหม่ (Save As .xlsx) แล้วอัปโหลดอีกครั้ง" });
  const txt = await upload("a.txt", new TextEncoder().encode("hi"));
  expect(await txt.json()).toEqual({ error: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx" });
});
