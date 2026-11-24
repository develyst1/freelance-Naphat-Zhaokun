// Builds test/fixtures/known.xlsx (formulas only, no cached results, so the engine must compute every one).
// Run: bun test/make-fixture.ts   (the tests also call makeKnownFixture() when the file is missing)
//
// Inputs and the hand-worked answers (these are what known.expected.json holds; none came from the engine):
//   Data!B2..C4: Apple 3 × 25.5 · Banana 12 × 4.25 · Cherry 7 × 18
//   D2 =B2*C2            → 3 × 25.5   = 76.5
//   D3 =B3*C3            → 12 × 4.25  = 51
//   D4 =B4*C4            → 7 × 18     = 126
//   B5 =SUM(B2:B4)       → 3 + 12 + 7 = 22
//   D5 =SUM(D2:D4)       → 76.5 + 51 + 126 = 253.5
//   B6 =AVERAGE(B2:B4)   → 22 / 3 = 7.333333…
//   D6 =ROUND(AVERAGE(D2:D4),2) → 253.5 / 3 = 84.5 → 84.5
//   D7 =IF(D5>200,"สูง","ต่ำ")  → 253.5 > 200 → "สูง"
//   D8 =VLOOKUP("Banana",A2:C4,3,FALSE) → Banana's price = 4.25
//   D9 =SUMIF(B2:B4,">5",D2:D4) → Qty > 5 are Banana (51) and Cherry (126) → 177
//   D10 =COUNTIF(B2:B4,">5")    → 2
//   D11 =ROUND(D4/D5*100,2)     → 126 / 253.5 × 100 = 49.7041… → 49.7   (a percentage)
//   Summary!B1 =Data!D5          → 253.5   (cross-sheet)
//   Summary!B2 =ROUND(B1*0.1,2)  → 25.35
//   Summary!B3 =B1+B2            → 278.85
import ExcelJS from "exceljs";
import { join } from "node:path";

export const FIXTURE = join(import.meta.dir, "fixtures", "known.xlsx");

export async function makeKnownFixture(path = FIXTURE): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet("Data");
  const rows: (string | number | { formula: string } | null)[][] = [
    ["Item", "Qty", "Price", "Total"],
    ["Apple", 3, 25.5, { formula: "B2*C2" }],
    ["Banana", 12, 4.25, { formula: "B3*C3" }],
    ["Cherry", 7, 18, { formula: "B4*C4" }],
    ["รวม", { formula: "SUM(B2:B4)" }, null, { formula: "SUM(D2:D4)" }],
    ["เฉลี่ย", { formula: "AVERAGE(B2:B4)" }, null, { formula: "ROUND(AVERAGE(D2:D4),2)" }],
    ["ยอดสูงไหม", null, null, { formula: 'IF(D5>200,"สูง","ต่ำ")' }],
    ["ราคา Banana", null, null, { formula: 'VLOOKUP("Banana",A2:C4,3,FALSE)' }],
    ["รวมเมื่อ Qty>5", null, null, { formula: 'SUMIF(B2:B4,">5",D2:D4)' }],
    ["นับ Qty>5", null, null, { formula: 'COUNTIF(B2:B4,">5")' }],
    ["สัดส่วน Cherry %", null, null, { formula: "ROUND(D4/D5*100,2)" }],
  ];
  rows.forEach((r, ri) => r.forEach((v, ci) => { if (v !== null) data.getCell(ri + 1, ci + 1).value = v as any; }));
  const sum = wb.addWorksheet("Summary");
  sum.getCell("A1").value = "ยอดรวมจาก Data";
  sum.getCell("B1").value = { formula: "Data!D5" } as any;
  sum.getCell("A2").value = "ภาษี 10%";
  sum.getCell("B2").value = { formula: "ROUND(B1*0.1,2)" } as any;
  sum.getCell("A3").value = "รวมทั้งสิ้น";
  sum.getCell("B3").value = { formula: "B1+B2" } as any;
  await wb.xlsx.writeFile(path);
}

if (import.meta.main) {
  await makeKnownFixture();
  console.log("wrote", FIXTURE);
}
