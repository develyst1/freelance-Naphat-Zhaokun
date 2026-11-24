// (a) known-answer formulas, (b) Thai round trip, AC-3 validity (re-open + [Content_Types].xml lists every part).
import { beforeAll, describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import mammoth from "mammoth";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { createFile, validateSpec, type FileSpec } from "../src/office/create.ts";
import { readOffice } from "../src/office/read.ts";
import { FIXTURE, makeKnownFixture } from "./make-fixture.ts";

const expected: Record<string, number | string> = JSON.parse(
  readFileSync(join(import.meta.dir, "fixtures", "known.expected.json"), "utf8"),
);

function same(actual: unknown, want: number | string) {
  if (typeof want === "number") expect(actual as number).toBeCloseTo(want, 9);
  else expect(actual).toBe(want);
}

function cellOf(computed: { name: string; rows: (string | number | null)[][] }[], ref: string) {
  const [sheet, addr] = ref.split("!");
  const m = addr.match(/^([A-Z]+)(\d+)$/)!;
  const col = m[1].charCodeAt(0) - 65;
  const row = Number(m[2]) - 1;
  return computed.find((s) => s.name === sheet)!.rows[row][col];
}

beforeAll(async () => {
  if (!existsSync(FIXTURE)) await makeKnownFixture();
});

describe("(a) known answers", () => {
  test("every computed cell of known.xlsx equals the hand-worked value", async () => {
    const buf = new Uint8Array(readFileSync(FIXTURE));
    const read = await readOffice("xlsx", buf);
    for (const [ref, want] of Object.entries(expected)) {
      if (ref.startsWith("_")) continue;
      same(cellOf(read.computed!, ref), want as number | string);
    }
    expect(read.text).toContain("D5 =SUM(D2:D4) → 253.5");
  });

  test("a created xlsx keeps each formula AND its hand-expected cached result", async () => {
    // hand-worked: D2 = 2×150 = 300 · D3 = 5×80 = 400 · D4 = SUM = 700 · D5 = AVERAGE(D2:D3) = 350 · D6 = ROUND(D4*0.07,2) = 49
    const spec: FileSpec = {
      type: "xlsx",
      filename: "ทดสอบ.xlsx",
      sheets: [{ name: "Sales", rows: [
        ["Item", "Qty", "Price", "Total"],
        ["A", 2, 150, "=B2*C2"],
        ["B", 5, 80, "=B3*C3"],
        ["รวม", "", "", "=SUM(D2:D3)"],
        ["เฉลี่ย", "", "", "=AVERAGE(D2:D3)"],
        ["VAT 7%", "", "", "=ROUND(D4*0.07,2)"],
      ] }],
    };
    const made = await createFile(spec);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(made.buffer) as any);
    const ws = wb.getWorksheet("Sales")!;
    const want: Record<string, [string, number]> = {
      D2: ["B2*C2", 300], D3: ["B3*C3", 400], D4: ["SUM(D2:D3)", 700], D5: ["AVERAGE(D2:D3)", 350], D6: ["ROUND(D4*0.07,2)", 49],
    };
    for (const [addr, [formula, result]] of Object.entries(want)) {
      const v = ws.getCell(addr).value as any;
      expect(v.formula).toBe(formula);
      expect(v.result).toBeCloseTo(result, 9);
    }
    expect(ws.getCell("A1").font?.bold).toBe(true);
    // exceljs writes calcPr but does not read it back, so check the XML itself
    const wbXml = await (await JSZip.loadAsync(made.buffer)).file("xl/workbook.xml")!.async("string");
    expect(wbXml).toMatch(/<calcPr[^>]*fullCalcOnLoad="1"/);
    expect(made.tables![0].rows[3][3]).toBe(700);
  });
});

async function contentTypesCoverEveryPart(buf: Uint8Array) {
  const zip = await JSZip.loadAsync(buf);
  const ct = await zip.file("[Content_Types].xml")!.async("string");
  const defaults = [...ct.matchAll(/<Default [^>]*Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase());
  const overrides = [...ct.matchAll(/<Override [^>]*PartName="([^"]+)"/g)].map((m) => m[1]);
  const parts = Object.values(zip.files).filter((f) => !f.dir).map((f) => f.name);
  const missing = parts.filter((p) => {
    if (p === "[Content_Types].xml") return false;
    const ext = p.split(".").pop()!.toLowerCase();
    return !defaults.includes(ext) && !overrides.includes("/" + p);
  });
  expect(missing).toEqual([]);
  return parts.length;
}

describe("(b) Thai round trip + AC-3 validity", () => {
  test("docx", async () => {
    const made = await createFile({ type: "docx", filename: "รายงาน", title: "รายงานประจำเดือน",
      blocks: [{ h1: "สรุปผล" }, { p: "ยอดขายเพิ่มขึ้นร้อยละสิบ" }, { bullets: ["ลูกค้าใหม่", "สินค้าขายดี"] }, { table: [["ชื่อ", "ยอด"], ["ก", "1"]] }] });
    expect(made.name).toBe("รายงาน.docx");
    const read = await readOffice("docx", made.buffer);
    for (const t of ["รายงานประจำเดือน", "สรุปผล", "ยอดขายเพิ่มขึ้นร้อยละสิบ", "ลูกค้าใหม่", "ชื่อ"]) expect(read.text).toContain(t);
    const raw = await mammoth.extractRawText({ buffer: Buffer.from(made.buffer) });
    expect(raw.value).toContain("สินค้าขายดี");
    expect(await contentTypesCoverEveryPart(made.buffer)).toBeGreaterThan(3);
    const zip = await JSZip.loadAsync(made.buffer);
    expect(await zip.file("word/document.xml")!.async("string")).toContain("Leelawadee UI");
  });

  test("pptx", async () => {
    const made = await createFile({ type: "pptx", filename: "นำเสนอ.pptx",
      slides: [{ title: "แผนงานปีหน้า", bullets: ["เพิ่มยอดขาย", "ลดต้นทุน"] }, { title: "สรุป" }] });
    const read = await readOffice("pptx", made.buffer);
    expect(read.text).toContain("Slide 1: แผนงานปีหน้า");
    expect(read.text).toContain("- เพิ่มยอดขาย");
    expect(read.text).toContain("Slide 2: สรุป");
    expect(await contentTypesCoverEveryPart(made.buffer)).toBeGreaterThan(5);
  });

  test("xlsx", async () => {
    const made = await createFile({ type: "xlsx", filename: "ค่าใช้จ่าย.xlsx",
      sheets: [{ name: "สรุป", rows: [["รายการ", "บาท"], ["ค่าเช่า", 5000], ["ค่าไฟ", 1200], ["รวม", "=SUM(B2:B3)"]] }] });
    const read = await readOffice("xlsx", made.buffer);
    expect(read.text).toContain("ค่าเช่า");
    expect(read.text).toContain("B4 =SUM(B2:B3) → 6200");
    expect(read.sheets![0].name).toBe("สรุป");
    expect(await contentTypesCoverEveryPart(made.buffer)).toBeGreaterThan(3);
  });
});

describe("file-spec validation", () => {
  test("rejects unknown kinds and bad shapes", () => {
    expect(validateSpec({ type: "pdf", filename: "a" })).toContain("unknown type");
    expect(validateSpec({ type: "xlsx", filename: "a", sheets: [] })).toContain("sheets");
    expect(validateSpec({ type: "docx", filename: "a", blocks: [{ h1: "x", p: "y" }] })).toContain("exactly one key");
    expect(validateSpec({ type: "pptx", filename: "a", slides: [{ title: 1 }] })).toContain("title");
    expect(validateSpec({ type: "xlsx", filename: "a", sheets: [{ name: "S", rows: [["=SUM(1,2)"]] }] })).toBeNull();
  });
});
