// The Excel engine: HyperFormula computes every formula. The LLM never computes a number a formula should.
import { HyperFormula } from "hyperformula";

export type Cell = string | number | boolean | null;
export type InSheet = { name: string; rows: Cell[][] };
export type OutValue = string | number | null;
export type FormulaInfo = { address: string; formula: string; value: OutValue };
export type Computed = { name: string; rows: OutValue[][]; formulas: FormulaInfo[] };

export function columnName(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function toOut(v: unknown): OutValue {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "string") return v;
  const err = v as { value?: unknown };
  if (err && typeof err.value === "string") return err.value; // CellError -> "#NAME?" etc.
  return String(v);
}

/**
 * HyperFormula reads a bare TRUE / FALSE as an unknown name (#NAME?), while Excel files use them as literals
 * (e.g. VLOOKUP(…,FALSE)). For the engine only, turn them into TRUE() / FALSE(); text in quotes is left alone.
 * The formula written to the file keeps its original text.
 */
export function engineFormula(formula: string): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => (i % 2 === 1 ? part : part.replace(/(?<![A-Za-z0-9_.!$])(TRUE|FALSE)(?![A-Za-z0-9_(])/gi, "$1()")))
    .join("");
}

/** Compute every sheet together (cross-sheet references work); returns values + the formula list. */
export function computeSheets(sheets: InSheet[]): Computed[] {
  const data: Record<string, Cell[][]> = {};
  for (const s of sheets)
    data[s.name] = s.rows.map((r) =>
      r.map((c) => (c === undefined ? null : typeof c === "string" && c.startsWith("=") ? engineFormula(c) : c)),
    );
  const hf = HyperFormula.buildFromSheets(data as any, { licenseKey: "gpl-v3" });
  try {
    return sheets.map((s) => {
      const id = hf.getSheetId(s.name)!;
      const values = hf.getSheetValues(id) as unknown[][];
      const rows: OutValue[][] = s.rows.map((r, ri) => r.map((_c, ci) => toOut(values[ri]?.[ci])));
      const formulas: FormulaInfo[] = [];
      s.rows.forEach((r, ri) =>
        r.forEach((c, ci) => {
          if (typeof c === "string" && c.startsWith("=")) {
            formulas.push({ address: `${columnName(ci)}${ri + 1}`, formula: c, value: rows[ri][ci] });
          }
        }),
      );
      return { name: s.name, rows, formulas };
    });
  } finally {
    hf.destroy();
  }
}
