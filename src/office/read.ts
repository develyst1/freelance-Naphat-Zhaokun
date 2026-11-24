// Read an uploaded docx / pptx / xlsx into text for the LLM, a short preview, and computed sheets for xlsx.
import ExcelJS from "exceljs";
import mammoth from "mammoth";
import type JSZip from "jszip";
import { computeSheets, type Cell, type Computed } from "../formulas.ts";
import type { Kind, Sheet } from "../store.ts";
import { checkZip } from "./zipguard.ts";
import { sanitizeXlsxForRead } from "./sanitize.ts";

export const CONTEXT_LIMIT = 25000;

export type ReadResult = {
  text: string;
  preview: string;
  sheets?: Sheet[];
  computed?: Computed[];
  charts?: number;
  comments?: number;
};

/** The upload is not a usable office zip at all (unreadable zip, zip bomb): the server answers badType. */
export class UnusableFile extends Error {}

export async function readOffice(kind: Kind, buf: Uint8Array): Promise<ReadResult> {
  let zip: JSZip;
  try {
    zip = await checkZip(buf);
  } catch (e) {
    throw new UnusableFile((e as Error).message);
  }
  if (kind === "docx") return readDocx(buf);
  if (kind === "pptx") return readPptx(zip);
  return readXlsx(buf);
}

function finish(text: string, extra: Partial<ReadResult> = {}): ReadResult {
  const clean = text.replace(/\r/g, "").replace(/\n{3,}/g, "\n\n").trim();
  const cut = clean.length > CONTEXT_LIMIT ? clean.slice(0, CONTEXT_LIMIT) + "\n[ตัดเหลือ 25,000 ตัวอักษร]" : clean;
  return { text: cut, preview: clean.slice(0, 300), ...extra };
}

async function readDocx(buf: Uint8Array): Promise<ReadResult> {
  const { value } = await mammoth.extractRawText({ buffer: Buffer.from(buf) });
  return finish(value);
}

function decodeXml(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}

/** each a:p paragraph -> its a:t runs joined */
function textsOf(xml: string): string[] {
  const paras: string[] = [];
  for (const p of xml.match(/<a:p\b[\s\S]*?<\/a:p>/g) ?? []) {
    const runs = [...p.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXml(m[1]));
    const line = runs.join("").trim();
    if (line) paras.push(line);
  }
  return paras;
}

async function readPptx(zip: JSZip): Promise<ReadResult> {
  const slideNames = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)\.xml$/)![1]) - Number(b.match(/(\d+)\.xml$/)![1]));
  const lines: string[] = [];
  for (const [i, name] of slideNames.entries()) {
    const xml = await zip.file(name)!.async("string");
    const shapes = xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? [];
    let title = "";
    const body: string[] = [];
    for (const sp of shapes) {
      const isTitle = /<p:ph\b[^>]*type="(?:title|ctrTitle)"/.test(sp);
      const texts = textsOf(sp);
      if (isTitle && !title) title = texts.join(" ");
      else body.push(...texts);
    }
    if (!title && body.length) title = body.shift()!;
    lines.push(`Slide ${i + 1}: ${title}`);
    for (const b of body) lines.push(`- ${b}`);
  }
  return finish(lines.join("\n"));
}

function cellInput(cell: ExcelJS.Cell): Cell {
  if (cell.type === ExcelJS.ValueType.Formula) {
    const f = cell.formula;
    if (f) return "=" + f;
  }
  const v = cell.value as any;
  if (v === null || v === undefined) return null;
  if (typeof v === "number" || typeof v === "string" || typeof v === "boolean") return v;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (v.richText) return v.richText.map((r: any) => r.text).join("");
  if (v.text !== undefined) return String(v.text); // hyperlink
  if (v.result !== undefined) return v.result as Cell;
  if (v.error) return String(v.error);
  return String(v);
}

export function sheetsFromWorkbook(wb: ExcelJS.Workbook): { name: string; rows: Cell[][] }[] {
  const out: { name: string; rows: Cell[][] }[] = [];
  wb.eachSheet((ws) => {
    const rows: Cell[][] = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row: Cell[] = [];
      for (let c = 1; c <= ws.columnCount; c++) row.push(cellInput(ws.getRow(r).getCell(c)));
      rows.push(row);
    }
    out.push({ name: ws.name, rows });
  });
  return out;
}

function tsv(rows: (string | number | null)[][]): string {
  return rows.map((r) => r.map((c) => (c === null ? "" : String(c))).join("\t")).join("\n");
}

async function readXlsx(buf: Uint8Array): Promise<ReadResult> {
  // TASK-013: charts / comments can crash ExcelJS, so it reads a sanitized copy (cells untouched)
  const clean = await sanitizeXlsxForRead(buf);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(clean.buf) as any);
  const computed = computeSheets(sheetsFromWorkbook(wb));
  const parts: string[] = [];
  if (clean.charts > 0) parts.push(`(ไฟล์นี้มีกราฟ ${clean.charts} อัน — ระบบอ่านเฉพาะข้อมูลในเซลล์ ไม่อ่านรูปกราฟ)`);
  if (clean.comments > 0) {
    if (clean.commentTexts.length) {
      parts.push(`(มีคอมเมนต์ในเซลล์ ${clean.commentTexts.length} รายการ)`);
      parts.push("comments:", ...clean.commentTexts);
    } else {
      parts.push(`(มีคอมเมนต์ในเซลล์ ${clean.comments} รายการ (ไม่ได้อ่าน))`);
    }
  }
  for (const s of computed) {
    parts.push(`Sheet: ${s.name}`);
    parts.push(tsv(s.rows));
    if (s.formulas.length) {
      parts.push("formulas:");
      for (const f of s.formulas) parts.push(`${f.address} ${f.formula} → ${f.value ?? ""}`);
    }
  }
  const sheets: Sheet[] = computed.map((s) => ({ name: s.name, rows: s.rows.slice(0, 20).map((r) => r.slice(0, 10)) }));
  return finish(parts.join("\n"), { sheets, computed, charts: clean.charts, comments: clean.comments });
}
