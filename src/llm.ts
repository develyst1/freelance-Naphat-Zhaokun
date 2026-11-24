// Chat -> the owner's LLM gateway (server side only). SPEC-003 §4.
import { config } from "./config.ts";
import { createFile, validateSpec, type Created, type FileSpec } from "./office/create.ts";
import type { StoredFile } from "./store.ts";

export type Msg = { role: "system" | "user" | "assistant"; content: string };

export class GatewayError extends Error {}

export const FILE_FAIL_LINE = "สร้างไฟล์ไม่สำเร็จ ลองพิมพ์ขอใหม่อีกครั้ง";

export const SYSTEM_PROMPT = `คุณคือผู้ช่วย AI ที่อ่านและสร้างไฟล์ Word (docx), Excel (xlsx) และ PowerPoint (pptx) ได้
You are an AI assistant that reads and creates Word (docx), Excel (xlsx) and PowerPoint (pptx) files.

- Answer in the user's language (Thai if they write Thai).
- Attached files arrive as <file name="…" kind="…">…extracted text…</file> blocks. For xlsx the values are already computed
  by a real Excel engine, and a "formulas:" list shows each formula with its computed result. Quote those numbers; never
  recompute them yourself.
- To create a file, write your normal answer and then append exactly ONE fenced block:
  \`\`\`file-spec
  { ...JSON... }
  \`\`\`
  JSON shapes:
  {"type":"xlsx","filename":"ชื่อ.xlsx","sheets":[{"name":"Sheet1","rows":[["Item","Qty","Price","Total"],["A",2,150,"=B2*C2"],["รวม","","","=SUM(D2:D2)"]]}]}
  {"type":"docx","filename":"ชื่อ.docx","title":"…","blocks":[{"h1":"…"},{"h2":"…"},{"p":"…"},{"bullets":["…","…"]},{"table":[["…","…"],["…","…"]]}]}
  {"type":"pptx","filename":"ชื่อ.pptx","slides":[{"title":"…","bullets":["…"]}]}
- In xlsx, any value that should be calculated MUST be an Excel formula string starting with "=" (SUM, AVERAGE, ROUND,
  IF, VLOOKUP, SUMIF, COUNTIF …). Never write the result of a formula yourself; the server computes it.
- Percentages: write the ratio formula without ×100 and without ROUND, e.g. \`=C3/$C$7\`, and put \`%\` in the column header.
  Optionally, a first row may be a title with one cell.
- docx: put the document title only in \`title\`; do not repeat it as the first \`h1\`.
- Limits: ≤ 5 sheets × 500 rows × 30 columns; ≤ 200 docx blocks; ≤ 30 slides. Valid JSON only (double quotes, no comments).
- แถวแรกของ rows คือแถวที่ 1 ใน Excel — เวลาอธิบายสูตร ให้อ้างชื่อเซลล์ตามการนับนี้ / The first entry of \`rows\` is Excel row 1; name cells by that count.
- Do not mention the file-spec block in your answer text.
- Size: pptx at most 12 slides with at most 6 short bullets per slide; docx at most 40 blocks; keep the JSON compact (no
  indentation). If the user asks for more slides, make 12 and say so in one line.
- Never say you have created a file unless the file-spec block is in the same message.`;

export const MAX_TOKENS = 8000;
export const SHORTER_HINT = 'ลองขอแบบสั้นลง เช่น "ทำ ppt 8 หน้า"';
const FILE_ASK_RE = /(ทำ|สร้าง|ขอ|ไหน).*(ppt|powerpoint|สไลด์|excel|xlsx|เอกสาร|word|docx|ไฟล์)|ไหนละ/i;
const RETRY_INVALID =
  "Your file-spec was cut off/invalid. Send ONLY a compact ```file-spec block, at most 8 slides / 30 blocks / 200 rows.";
const RETRY_MISSING = "The user asked for a file. Reply ONLY with the ```file-spec block for it.";

export function asksForFile(text: string): boolean {
  return FILE_ASK_RE.test(text);
}

export function fileContext(files: StoredFile[]): string {
  return files
    .map((f) => `<file name="${f.name.replace(/"/g, "'")}" kind="${f.kind}">\n${f.text}\n</file>`)
    .join("\n\n");
}

export function buildMessages(history: Msg[], files: StoredFile[]): Msg[] {
  const msgs: Msg[] = [{ role: "system", content: SYSTEM_PROMPT }];
  if (files.length) msgs.push({ role: "system", content: "Attached files:\n\n" + fileContext(files) });
  return msgs.concat(history.slice(-20));
}

export type GatewayAnswer = { content: string; provider?: string; model?: string; completionTokens?: number };

export async function callGateway(messages: Msg[]): Promise<GatewayAnswer> {
  const body: Record<string, unknown> = { messages, temperature: 0.3, max_tokens: MAX_TOKENS };
  if (config.provider()) body.provider = config.provider();
  if (config.model()) body.model = config.model();
  const headers: Record<string, string> = { "content-type": "application/json" };
  const key = config.gatewayKey();
  if (key) headers.authorization = `Bearer ${key}`; // the one place the auth header lives (skill §5: not decided yet)

  let res: Response;
  try {
    res = await fetch(`${config.gatewayUrl()}/chat`, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(90_000),
    });
  } catch (e) {
    throw new GatewayError(`gateway unreachable: ${(e as Error).message}`);
  }
  let json: any = null;
  try {
    json = await res.json();
  } catch {
    /* handled below */
  }
  const content = json?.data?.content;
  if (!res.ok || !json?.success || typeof content !== "string" || !content.trim()) {
    throw new GatewayError(`gateway HTTP ${res.status}: ${json?.error ?? "empty content"}`);
  }
  const usage = json.data.usage;
  return {
    content,
    provider: json.data.provider,
    model: json.data.model,
    completionTokens: typeof usage?.completion_tokens === "number" ? usage.completion_tokens : undefined,
  };
}

export type Extracted = {
  text: string;
  spec?: FileSpec;
  error?: string;
  found: boolean;
  truncated?: boolean;
  repaired?: boolean;
};

const BARE_START = /\{\s*"type"\s*:\s*"(?:pptx|docx|xlsx)"/;

/** End index (exclusive) of the balanced {...} starting at `start`, honouring JSON strings; -1 when it never closes. */
function balancedEnd(s: string, start: number): number {
  let depth = 0;
  let inStr = false;
  for (let i = start; i < s.length; i++) {
    const ch = s[i];
    if (inStr) {
      if (ch === "\\") i++;
      else if (ch === '"') inStr = false;
    } else if (ch === '"') inStr = true;
    else if (ch === "{") depth++;
    else if (ch === "}" && --depth === 0) return i + 1;
  }
  return -1;
}

/** The closers (`]`/`}`) needed to finish the JSON prefix s, or null when s ends inside a string. */
function closersFor(s: string): string | null {
  const stack: string[] = [];
  let inStr = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inStr) {
      if (ch === "\\") i++;
      else if (ch === '"') inStr = false;
    } else if (ch === '"') inStr = true;
    else if (ch === "{") stack.push("}");
    else if (ch === "[") stack.push("]");
    else if (ch === "}" || ch === "]") stack.pop();
  }
  return inStr ? null : stack.reverse().join("");
}

/**
 * The model sometimes stops early or writes near-JSON (seen live: the block ended far below max_tokens).
 * Repair before giving up: (1) drop trailing commas; (2) when cut off, keep everything up to the last complete
 * object/array element and close the open brackets — a 17-slide deck cut at slide 9 becomes 8 good slides.
 */
export function repairJson(raw: string): unknown | undefined {
  const s = raw.trim().replace(/```\s*$/, "").trim();
  const noTrailing = s.replace(/,\s*([}\]])/g, "$1");
  try {
    return JSON.parse(noTrailing);
  } catch {
    /* try a cut-off repair */
  }
  for (let i = noTrailing.length - 1, tries = 0; i > 0 && tries < 400; i--) {
    const ch = noTrailing[i];
    if (ch !== "}" && ch !== "]") continue;
    tries++;
    const prefix = noTrailing.slice(0, i + 1);
    const closers = closersFor(prefix);
    if (closers === null) continue;
    try {
      return JSON.parse(prefix + closers);
    } catch {
      /* keep walking back */
    }
  }
  return undefined;
}

function judge(text: string, raw: string, truncated = false): Extracted {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw.trim());
  } catch (e) {
    parsed = repairJson(raw);
    if (parsed === undefined || validateSpec(parsed))
      return { text, found: true, truncated: true, error: `cut off or invalid JSON: ${(e as Error).message}` };
    console.log(`[chat] file-spec repaired (${truncated ? "cut off" : "invalid JSON"}: ${(e as Error).message})`);
    return { text, found: true, spec: parsed as FileSpec, repaired: true };
  }
  const err = validateSpec(parsed);
  if (err) return { text, found: true, truncated, error: err };
  return { text, found: true, spec: parsed as FileSpec };
}

/**
 * Pulls a file-spec out of a model answer, in three shapes:
 * (a) a closed ```file-spec … ``` block; (b) an unterminated block (cut at max_tokens) up to the end;
 * (c) a bare {"type":"pptx|docx|xlsx",…} object with no fence. The spec text, even partial, never stays in `text`.
 */
export function extractSpec(content: string): Extracted {
  const closed = content.match(/```file-spec[^\n]*\n?([\s\S]*?)```/);
  if (closed) return judge(stripSpecText(content.replace(closed[0], "")), closed[1]);

  const open = content.indexOf("```file-spec");
  if (open >= 0) {
    const raw = content.slice(open).replace(/^```file-spec[^\n]*\n?/, "");
    return judge(stripSpecText(content.slice(0, open)), raw, true);
  }

  const bare = content.search(BARE_START);
  if (bare >= 0) {
    const end = balancedEnd(content, bare);
    const raw = end > 0 ? content.slice(bare, end) : content.slice(bare);
    const rest = content.slice(0, bare) + (end > 0 ? content.slice(end) : "");
    return judge(stripSpecText(rest), raw, end < 0);
  }
  return { text: content.trim(), found: false };
}

/** Last pass: no ```file-spec fence and no {"type":"pptx|docx|xlsx" fragment ever reaches the reply. */
export function stripSpecText(text: string): string {
  let out = text;
  const fence = out.indexOf("```file-spec");
  if (fence >= 0) out = out.slice(0, fence);
  const bare = out.search(BARE_START);
  if (bare >= 0) {
    const end = balancedEnd(out, bare);
    out = out.slice(0, bare) + (end > 0 ? out.slice(end) : "");
  }
  return out.replace(/```\s*$/, "").trim();
}

export type ChatOutcome = { reply: string; created: Created[]; provider?: string; model?: string };

export async function chatWithFiles(history: Msg[], files: StoredFile[]): Promise<ChatOutcome> {
  const messages = buildMessages(history, files);
  const first = await callGateway(messages);
  let ex = extractSpec(first.content);
  let text = ex.text;
  let spec = ex.spec;
  let provider = first.provider;
  let model = first.model;
  if (ex.truncated)
    console.log(
      `[chat] file-spec cut off · completion_tokens=${first.completionTokens ?? "?"} of max_tokens=${MAX_TOKENS}` +
        ` · chars=${first.content.length} · closing fence=${/```\s*$/.test(first.content)} · ${ex.error}`,
    );

  const lastUser = [...history].reverse().find((m) => m.role === "user")?.content ?? "";
  const wanted = ex.found || asksForFile(lastUser);

  if (!spec && wanted) {
    // one retry in total: a cut/invalid block, or a file asked for and none sent
    const ask = ex.found ? RETRY_INVALID : RETRY_MISSING;
    const retry = await callGateway([...messages, { role: "assistant", content: stripSpecText(first.content) || "…" }, { role: "user", content: ask }]);
    provider = retry.provider ?? provider;
    model = retry.model ?? model;
    const again = extractSpec(retry.content);
    if (again.truncated)
      console.log(
        `[chat] retry file-spec cut off · completion_tokens=${retry.completionTokens ?? "?"} of max_tokens=${MAX_TOKENS}` +
          ` · chars=${retry.content.length} · closing fence=${/```\s*$/.test(retry.content)} · ${again.error}`,
      );
    spec = again.spec;
    if (!text) text = again.text;
  }

  const created: Created[] = [];
  if (spec) {
    try {
      created.push(await createFile(spec));
    } catch (e) {
      console.error("createFile failed:", (e as Error).message);
      spec = undefined;
    }
  }
  if (wanted && !created.length) text = `${text}\n\n${FILE_FAIL_LINE} ${SHORTER_HINT}`.trim();
  text = stripSpecText(text);
  if (created.some((c) => c.tables?.some((t) => t.rows.some((r) => r.some((v) => typeof v === "string" && v.startsWith("#NAME"))))))
    text += "\n\nบางสูตรคำนวณในเว็บไม่ได้ แต่ Excel จะคำนวณเอง";
  return { reply: text, created, provider, model };
}
