// file-spec (SPEC §5) -> a real docx / xlsx / pptx. Formulas are computed by HyperFormula, never by the LLM.
import ExcelJS from "exceljs";
import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";
import PptxGenJS from "pptxgenjs";
import { computeSheets, type Cell, type OutValue } from "../formulas.ts";
import type { Kind, Sheet } from "../store.ts";

export const THAI_FONT = "Leelawadee UI";
const LIMITS = { sheets: 5, rows: 500, cols: 30, blocks: 200, slides: 30 };

export type XlsxSpec = { type: "xlsx"; filename: string; sheets: { name: string; rows: Cell[][] }[] };
export type DocxBlock = { h1: string } | { h2: string } | { p: string } | { bullets: string[] } | { table: string[][] };
export type DocxSpec = { type: "docx"; filename: string; title?: string; blocks: DocxBlock[] };
export type PptxSpec = { type: "pptx"; filename: string; slides: { title: string; bullets?: string[] }[] };
export type FileSpec = XlsxSpec | DocxSpec | PptxSpec;

export type Created = { name: string; kind: Kind; buffer: Uint8Array; tables?: Sheet[] };

const isStr = (v: unknown): v is string => typeof v === "string";
const isStrArr = (v: unknown): v is string[] => Array.isArray(v) && v.every(isStr);
const isCell = (v: unknown) => v === null || ["string", "number", "boolean"].includes(typeof v);

/** Returns the error text, or null when the spec is valid. */
export function validateSpec(spec: any): string | null {
  if (!spec || typeof spec !== "object") return "file-spec must be a JSON object";
  if (!["xlsx", "docx", "pptx"].includes(spec.type)) return `unknown type "${spec.type}" (use xlsx, docx or pptx)`;
  if (!isStr(spec.filename) || !spec.filename.trim()) return "filename must be a non-empty string";

  if (spec.type === "xlsx") {
    if (!Array.isArray(spec.sheets) || spec.sheets.length < 1 || spec.sheets.length > LIMITS.sheets)
      return `sheets must be an array of 1–${LIMITS.sheets}`;
    const names = new Set<string>();
    for (const [i, s] of spec.sheets.entries()) {
      if (!s || !isStr(s.name) || !s.name.trim() || s.name.length > 31 || /[\[\]:*?\/\\]/.test(s.name))
        return `sheets[${i}].name must be 1–31 chars without [ ] : * ? / \\`;
      if (names.has(s.name)) return `sheet name "${s.name}" is used twice`;
      names.add(s.name);
      if (!Array.isArray(s.rows) || s.rows.length > LIMITS.rows) return `sheets[${i}].rows must be an array of ≤ ${LIMITS.rows}`;
      for (const [r, row] of s.rows.entries()) {
        if (!Array.isArray(row) || row.length > LIMITS.cols) return `sheets[${i}].rows[${r}] must be an array of ≤ ${LIMITS.cols} cells`;
        if (!row.every(isCell)) return `sheets[${i}].rows[${r}] cells must be string, number, boolean or null`;
      }
    }
    return null;
  }

  if (spec.type === "docx") {
    if (spec.title !== undefined && !isStr(spec.title)) return "title must be a string";
    if (!Array.isArray(spec.blocks) || spec.blocks.length > LIMITS.blocks) return `blocks must be an array of ≤ ${LIMITS.blocks}`;
    for (const [i, b] of spec.blocks.entries()) {
      const keys = b && typeof b === "object" ? Object.keys(b) : [];
      if (keys.length !== 1) return `blocks[${i}] must have exactly one key: h1, h2, p, bullets or table`;
      const k = keys[0];
      if ((k === "h1" || k === "h2" || k === "p") && isStr(b[k])) continue;
      if (k === "bullets" && isStrArr(b.bullets)) continue;
      if (k === "table" && Array.isArray(b.table) && b.table.every(isStrArr)) continue;
      return `blocks[${i}].${k} has the wrong shape`;
    }
    return null;
  }

  if (!Array.isArray(spec.slides) || spec.slides.length < 1 || spec.slides.length > LIMITS.slides)
    return `slides must be an array of 1–${LIMITS.slides}`;
  for (const [i, s] of spec.slides.entries()) {
    if (!s || !isStr(s.title)) return `slides[${i}].title must be a string`;
    if (s.bullets !== undefined && !isStrArr(s.bullets)) return `slides[${i}].bullets must be an array of strings`;
  }
  return null;
}

export function safeFilename(name: string, kind: Kind): string {
  const base = name
    .normalize("NFC")
    .replace(/\.(docx|xlsx|pptx)$/i, "")
    .replace(/[^\p{L}\p{M}\p{N} ._-]/gu, "")
    .trim()
    .slice(0, 80);
  return `${base || "file"}.${kind}`;
}

export async function createFile(spec: FileSpec): Promise<Created> {
  const name = safeFilename(spec.filename, spec.type);
  if (spec.type === "xlsx") return { name, kind: "xlsx", ...(await buildXlsx(spec)) };
  if (spec.type === "docx") return { name, kind: "docx", buffer: await buildDocx(spec) };
  return { name, kind: "pptx", buffer: await buildPptx(spec) };
}

function cachedResult(v: OutValue): any {
  if (typeof v === "string" && v.startsWith("#")) return { error: v };
  return v ?? undefined;
}

const PERCENT_HEADER = /%|เปอร์เซ็นต์|ร้อยละ/;
export const FMT = { integer: "#,##0", decimal: "#,##0.00", percent: "0.00%" };
const HEADER_FILL = "FFF2F2F2";

const filled = (row: Cell[] | undefined) => (row ?? []).filter((c) => c !== null && c !== "" && c !== undefined).length;

/** Row index (0-based) of the header: 1 when row 1 is a one-cell title above a real header row, else 0. */
export function headerRowIndex(rows: Cell[][]): number {
  return rows.length >= 2 && filled(rows[0]) === 1 && filled(rows[1]) >= 2 ? 1 : 0;
}

/**
 * Percentage columns: the header names a percentage. The LLM is told to write the plain ratio (=C3/$C$7), shown as 0.00%.
 * Guard: when most numeric body values are > 1.5 the model multiplied by 100 anyway, so the column is NOT formatted as %.
 */
export function percentColumns(rows: Cell[][], values: OutValue[][], header: number, sheet: string): Set<number> {
  const cols = new Set<number>();
  (rows[header] ?? []).forEach((h, ci) => {
    if (typeof h !== "string" || !PERCENT_HEADER.test(h)) return;
    const nums = values.slice(header + 1).map((r) => r[ci]).filter((v): v is number => typeof v === "number");
    const big = nums.filter((v) => Math.abs(v) > 1.5).length;
    if (nums.length && big > nums.length / 2) {
      console.log(`[xlsx] ${sheet}: column ${ci + 1} "${h}" holds ×100 values, kept as plain numbers (no % format)`);
      return;
    }
    cols.add(ci);
  });
  return cols;
}

function numberFormat(v: number, percent: boolean): string {
  if (percent) return FMT.percent;
  return Number.isInteger(v) ? FMT.integer : FMT.decimal;
}

function shown(v: OutValue, percent: boolean): string {
  if (typeof v !== "number") return v === null ? "" : String(v);
  if (percent) return `${(v * 100).toFixed(2)}%`;
  const fixed = Number.isInteger(v) ? v.toFixed(0) : v.toFixed(2);
  const [int, dec] = fixed.split(".");
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (dec ? "." + dec : "");
}

async function buildXlsx(spec: XlsxSpec): Promise<{ buffer: Uint8Array; tables: Sheet[] }> {
  const computed = computeSheets(spec.sheets);
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;
  const tables: Sheet[] = [];
  for (const [si, s] of spec.sheets.entries()) {
    const ws = wb.addWorksheet(s.name);
    const values = computed[si].rows;
    const header = headerRowIndex(s.rows);
    const hasTitle = header === 1;
    const pct = percentColumns(s.rows, values, header, s.name);
    const widths: number[] = [];

    s.rows.forEach((row, ri) => {
      row.forEach((c, ci) => {
        const cell = ws.getCell(ri + 1, ci + 1);
        const v = values[ri]?.[ci] ?? null;
        if (typeof c === "string" && c.startsWith("=")) {
          cell.value = { formula: c.slice(1), result: cachedResult(v) } as any;
        } else if (c !== null && c !== "") {
          cell.value = c as any;
        }
        const isPct = ri > header && pct.has(ci);
        if (typeof v === "number" && !(hasTitle && ri === 0)) cell.numFmt = numberFormat(v, isPct);
        if (hasTitle && ri === 0) {
          cell.font = { bold: true, size: 14 };
          return; // the title does not set a column width
        }
        if (ri === header) {
          cell.font = { bold: true };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
          cell.border = { bottom: { style: "thin" } };
        }
        const len = [...shown(v, isPct)].length;
        widths[ci] = Math.max(widths[ci] ?? 0, len);
      });
    });
    widths.forEach((w, ci) => (ws.getColumn(ci + 1).width = Math.min(50, Math.max(8, (w ?? 0) + 2))));
    ws.views = [{ state: "frozen", ySplit: header + 1 }];

    // the web table: numbers stay numbers, except % columns which become "43.06%"
    const rows = values.map((r, ri) =>
      r.map((v, ci) => (ri > header && pct.has(ci) && typeof v === "number" ? `${(v * 100).toFixed(2)}%` : v)),
    );
    tables.push({ name: s.name, rows: rows.slice(0, 20).map((r) => r.slice(0, 10)) });
  }
  const buffer = new Uint8Array(await wb.xlsx.writeBuffer());
  return { buffer, tables };
}

const run = (text: string, bold = false) => new TextRun({ text, bold, font: THAI_FONT });

async function buildDocx(spec: DocxSpec): Promise<Uint8Array> {
  const children: (Paragraph | Table)[] = [];
  if (spec.title) children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [run(spec.title)] }));
  // the title only once: drop a first h1 that repeats it (the prompt also says so)
  const first = spec.blocks[0];
  const blocks =
    spec.title && first && "h1" in first && first.h1.trim() === spec.title.trim() ? spec.blocks.slice(1) : spec.blocks;
  for (const b of blocks) {
    if ("h1" in b) children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run(b.h1)] }));
    else if ("h2" in b) children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(b.h2)] }));
    else if ("p" in b) children.push(new Paragraph({ children: [run(b.p)] }));
    else if ("bullets" in b) for (const t of b.bullets) children.push(new Paragraph({ bullet: { level: 0 }, children: [run(t)] }));
    else if ("table" in b && b.table.length) {
      const width = Math.max(...b.table.map((r) => r.length));
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: b.table.map(
            (r, ri) =>
              new TableRow({
                children: Array.from({ length: width }, (_x, ci) =>
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [run(r[ci] ?? "", ri === 0)] })] }),
                ),
              }),
          ),
        }),
      );
    }
  }
  const doc = new Document({
    styles: { default: { document: { run: { font: THAI_FONT } } } },
    sections: [{ children }],
  });
  return new Uint8Array(await Packer.toBuffer(doc));
}

async function buildPptx(spec: PptxSpec): Promise<Uint8Array> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  for (const s of spec.slides) {
    const slide = pptx.addSlide();
    slide.addText(s.title, { x: 0.5, y: 0.3, w: 9, h: 1, fontFace: THAI_FONT, fontSize: 28, bold: true });
    if (s.bullets?.length) {
      slide.addText(
        s.bullets.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })),
        { x: 0.5, y: 1.4, w: 9, h: 3.8, fontFace: THAI_FONT, fontSize: 18, valign: "top" },
      );
    }
  }
  const out = (await pptx.write({ outputType: "nodebuffer" })) as Buffer;
  return new Uint8Array(out);
}
