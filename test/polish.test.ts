// TASK-011: created xlsx are formatted (title/header/freeze/number formats/widths, % as a ratio that sums to 100%)
// and a docx carries its title once. Hand-worked values in the comments.
import { expect, test } from "bun:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { createFile } from "../src/office/create.ts";
import { SYSTEM_PROMPT } from "../src/llm.ts";

// Travel costs (hand-worked): 4500 + 3200 + 1800 + 950 = 10450 · average 10450/4 = 2612.5
// shares: 4500/10450 = 0.430622 → "43.06%" · 3200/10450 = 0.306220 · 1800/10450 = 0.172249 · 950/10450 = 0.090909 → sum exactly 1
const travel = {
  type: "xlsx" as const,
  filename: "ค่าเดินทาง",
  sheets: [{ name: "ค่าเดินทาง", rows: [
    ["สรุปค่าใช้จ่ายเดินทางประจำเดือน", "", "", ""],
    ["ลำดับ", "รายการค่าใช้จ่ายเดินทางไปประชุมต่างจังหวัด", "จำนวนเงิน (บาท)", "สัดส่วน (%)"],
    [1, "ค่าเครื่องบิน", 4500, "=C3/$C$7"],
    [2, "ค่าที่พัก", 3200, "=C4/$C$7"],
    [3, "ค่าแท็กซี่", 1800, "=C5/$C$7"],
    [4, "ค่าอาหาร", 950, "=C6/$C$7"],
    ["", "รวม", "=SUM(C3:C6)", "=SUM(D3:D6)"],
    ["", "เฉลี่ย", "=AVERAGE(C3:C6)", ""],
  ] }],
};

async function open(buffer: Uint8Array) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(buffer) as any);
  return wb.worksheets[0];
}

test("title row + header detected and styled; panes frozen under the header", async () => {
  const ws = await open((await createFile(travel)).buffer);
  expect(ws.getCell("A1").font).toMatchObject({ bold: true, size: 14 });
  const h = ws.getCell("B2");
  expect(h.font?.bold).toBe(true);
  expect((h.fill as any).fgColor.argb).toBe("FFF2F2F2");
  expect(h.border?.bottom?.style).toBe("thin");
  expect(ws.views[0]).toMatchObject({ state: "frozen", ySplit: 2 });
});

test("number formats: integer #,##0 · decimal #,##0.00 · % column 0.00% whose SUM is exactly 1 (100.00%)", async () => {
  const made = await createFile(travel);
  const ws = await open(made.buffer);
  expect(ws.getCell("C3").numFmt).toBe("#,##0"); // 4500
  expect(ws.getCell("C7").numFmt).toBe("#,##0"); // SUM = 10450
  expect(ws.getCell("C8").numFmt).toBe("#,##0.00"); // 2612.5
  expect((ws.getCell("C8").value as any).result).toBe(2612.5);
  expect(ws.getCell("D3").numFmt).toBe("0.00%");
  expect(ws.getCell("D7").numFmt).toBe("0.00%");
  expect((ws.getCell("D7").value as any).result).toBeCloseTo(1, 9);
  // the web table: % column as strings, other numbers untouched
  const rows = made.tables![0].rows;
  expect(rows[2][3]).toBe("43.06%");
  expect(rows[6][3]).toBe("100.00%");
  expect(rows[2][2]).toBe(4500);
  expect(rows[1][3]).toBe("สัดส่วน (%)");
});

test("guard: a % column the model filled with ×100 values keeps #,##0.00, not 0.00%", async () => {
  const made = await createFile({ type: "xlsx", filename: "g", sheets: [{ name: "S", rows: [
    ["รายการ", "บาท", "ร้อยละ"],
    ["ก", 300, "=ROUND(B2/B4*100,2)"], // 300/400*100 = 75
    ["ข", 100, "=ROUND(B3/B4*100,2)"], // 25
    ["รวม", "=SUM(B2:B3)", "=SUM(C2:C3)"], // 100
  ] }] });
  const ws = await open(made.buffer);
  expect(ws.getCell("C2").numFmt).toBe("#,##0"); // 75 is an integer, plain number
  expect(ws.getCell("C2").numFmt).not.toBe("0.00%");
  expect(made.tables![0].rows[1][2]).toBe(75);
  const made2 = await createFile({ type: "xlsx", filename: "g2", sheets: [{ name: "S", rows: [
    ["รายการ", "%"], ["ก", 43.06], ["ข", 56.94],
  ] }] });
  expect((await open(made2.buffer)).getCell("B2").numFmt).toBe("#,##0.00");
});

test("no title row → row 1 is the header", async () => {
  const ws = await open((await createFile({ type: "xlsx", filename: "h", sheets: [{ name: "S", rows: [["a", "b"], [1, 2]] }] })).buffer);
  expect(ws.getCell("A1").font?.bold).toBe(true);
  expect((ws.getCell("A1").fill as any).fgColor.argb).toBe("FFF2F2F2");
  expect(ws.views[0]).toMatchObject({ state: "frozen", ySplit: 1 });
});

test("widths: the long Thai column is wider than a short one, all within 8–50", async () => {
  const ws = await open((await createFile(travel)).buffer);
  const a = ws.getColumn(1).width!;
  const b = ws.getColumn(2).width!;
  expect(b).toBeGreaterThan(a);
  for (const w of [a, b, ws.getColumn(3).width!, ws.getColumn(4).width!]) {
    expect(w).toBeGreaterThanOrEqual(8);
    expect(w).toBeLessThanOrEqual(50);
  }
  expect(a).toBe(8); // "ลำดับ" = 5 chars + 2 = 7 → clamped to 8
});

async function docxXml(spec: any) {
  const made = await createFile(spec);
  return (await JSZip.loadAsync(made.buffer)).file("word/document.xml")!.async("string");
}

test("docx: a first h1 equal to the title is dropped; a different h1 is kept", async () => {
  const same = await docxXml({ type: "docx", filename: "a", title: "สรุปการประชุม", blocks: [{ h1: " สรุปการประชุม " }, { p: "เนื้อหา" }] });
  expect(same.split("สรุปการประชุม").length - 1).toBe(1);
  const diff = await docxXml({ type: "docx", filename: "b", title: "สรุปการประชุม", blocks: [{ h1: "วาระที่ 1" }] });
  expect(diff).toContain("สรุปการประชุม");
  expect(diff).toContain("วาระที่ 1");
});

test("the system prompt carries the % ratio rule and the docx single-title rule", () => {
  expect(SYSTEM_PROMPT).toContain("Percentages: write the ratio formula without ×100 and without ROUND, e.g. `=C3/$C$7`");
  expect(SYSTEM_PROMPT).toContain("docx: put the document title only in `title`; do not repeat it as the first `h1`.");
});
