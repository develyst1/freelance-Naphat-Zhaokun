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
- Do not mention the file-spec block in your answer text.`;

const SPEC_RE = /```file-spec\s*\n?([\s\S]*?)```/;

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

export type GatewayAnswer = { content: string; provider?: string; model?: string };

export async function callGateway(messages: Msg[]): Promise<GatewayAnswer> {
  const body: Record<string, unknown> = { messages, temperature: 0.3, max_tokens: 4000 };
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
  return { content, provider: json.data.provider, model: json.data.model };
}

/** Pulls the file-spec block out. spec=undefined when there is none; error when it does not parse/validate. */
export function extractSpec(content: string): { text: string; spec?: FileSpec; error?: string; found: boolean } {
  const m = content.match(SPEC_RE);
  if (!m) return { text: content.trim(), found: false };
  const text = content.replace(SPEC_RE, "").trim();
  let parsed: unknown;
  try {
    parsed = JSON.parse(m[1].trim());
  } catch (e) {
    return { text, found: true, error: `JSON parse error: ${(e as Error).message}` };
  }
  const err = validateSpec(parsed);
  if (err) return { text, found: true, error: err };
  return { text, found: true, spec: parsed as FileSpec };
}

export type ChatOutcome = { reply: string; created: Created[]; provider?: string; model?: string };

export async function chatWithFiles(history: Msg[], files: StoredFile[]): Promise<ChatOutcome> {
  const messages = buildMessages(history, files);
  const first = await callGateway(messages);
  let { text, spec, error, found } = extractSpec(first.content);
  let provider = first.provider;
  let model = first.model;

  if (found && !spec) {
    // one retry: ask for the corrected block only
    const retry = await callGateway([
      ...messages,
      { role: "assistant", content: first.content },
      { role: "user", content: `your file-spec was invalid: ${error}; send only the corrected \`\`\`file-spec block` },
    ]);
    provider = retry.provider ?? provider;
    model = retry.model ?? model;
    const again = extractSpec(retry.content);
    spec = again.spec;
    error = again.error ?? (again.found ? undefined : "no file-spec block in the retry");
  }

  const created: Created[] = [];
  if (found) {
    if (spec) {
      try {
        created.push(await createFile(spec));
      } catch (e) {
        console.error("createFile failed:", (e as Error).message);
        spec = undefined;
      }
    }
    if (!spec) text = `${text}\n\n${FILE_FAIL_LINE}`.trim();
  }
  if (created.some((c) => c.tables?.some((t) => t.rows.some((r) => r.some((v) => typeof v === "string" && v.startsWith("#NAME"))))))
    text += "\n\nบางสูตรคำนวณในเว็บไม่ได้ แต่ Excel จะคำนวณเอง";
  return { reply: text, created, provider, model };
}
