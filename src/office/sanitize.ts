// TASK-013 (D3): ExcelJS crashes (`drawing.anchors` in reconcile) on some real xlsx with charts / comments.
// For READING only, strip drawings, charts, comments and their links; cells, formulas, merges and sheet names are untouched.
// The uploaded original is stored as is.
import JSZip from "jszip";

export type Sanitized = { buf: Uint8Array; charts: number; comments: number; commentTexts: string[] };

const DROP_PART = [
  /^xl\/drawings\//i,
  /^xl\/charts\//i,
  /^xl\/comments?[^/]*$/i, // xl/comments1.xml
  /^xl\/(?:.*\/)?comments?[^/]*\.xml$/i, // xl/comments/comment1.xml (a real file used the singular), xl/x/comments2.xml
  /^xl\/.*\.vml$/i,
  /^xl\/threadedComments\//i,
  /^xl\/persons\//i,
];
const DROP_REL_TYPE = /\/(drawing|vmlDrawing|comments|threadedComment|person)"/i;

function decode(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}

/** Comment texts as "A1: …" from xl/comments*.xml (legacy notes). */
function commentLines(xml: string): string[] {
  const out: string[] = [];
  for (const m of xml.matchAll(/<comment\b[^>]*\bref="([^"]+)"[^>]*>([\s\S]*?)<\/comment>/g)) {
    const text = [...m[2].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => decode(t[1])).join("").trim();
    if (text) out.push(`${m[1]}: ${text.replace(/\s+/g, " ").slice(0, 200)}`);
  }
  return out;
}

export async function sanitizeXlsxForRead(input: Uint8Array): Promise<Sanitized> {
  const zip = await JSZip.loadAsync(input);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

  const charts = names.filter((n) => /^xl\/charts\/chart[^/]*\.xml$/i.test(n)).length;
  const commentParts = names.filter((n) => /^xl\/(?:.*\/)?comments?[^/]*\.xml$/i.test(n) && !/threadedComment/i.test(n));
  const commentTexts: string[] = [];
  for (const n of commentParts) commentTexts.push(...commentLines(await zip.file(n)!.async("string")));
  const comments = commentTexts.length || commentParts.length;

  const removed = new Set<string>();
  for (const n of names) {
    if (DROP_PART.some((re) => re.test(n))) {
      zip.remove(n);
      removed.add("/" + n);
    }
  }
  // folders left empty by the removal
  for (const n of Object.keys(zip.files)) {
    if (zip.files[n]?.dir && /^xl\/(drawings|charts|comments|threadedComments|persons)\//i.test(n)) zip.remove(n);
  }

  for (const n of names) {
    if (removed.has("/" + n)) continue;
    if (/^xl\/worksheets\/[^/]+\.xml$/i.test(n)) {
      const xml = await zip.file(n)!.async("string");
      const clean = xml.replace(/<(?:\w+:)?(?:drawing|legacyDrawing|legacyDrawingHF)\b[^>]*\/>/g, "");
      if (clean !== xml) zip.file(n, clean);
    } else if (/^xl\/worksheets\/_rels\/[^/]+\.rels$/i.test(n)) {
      const xml = await zip.file(n)!.async("string");
      const clean = xml.replace(/<Relationship\b[^>]*\/>/g, (rel) => (DROP_REL_TYPE.test(rel) ? "" : rel));
      if (clean !== xml) zip.file(n, clean);
    }
  }

  const ct = zip.file("[Content_Types].xml");
  if (ct) {
    const xml = await ct.async("string");
    const clean = xml.replace(/<Override\b[^>]*PartName="([^"]+)"[^>]*\/>/g, (o, part) => (removed.has(part) ? "" : o));
    if (clean !== xml) zip.file("[Content_Types].xml", clean);
  }

  const buf = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return { buf, charts, comments, commentTexts };
}
