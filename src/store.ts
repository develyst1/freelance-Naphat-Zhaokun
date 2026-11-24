// In-memory files (uploads + generated). Each expires 1 h after it was stored; a sweep runs every 5 min.
import { randomBytes } from "node:crypto";

export type Kind = "docx" | "xlsx" | "pptx";
export type Sheet = { name: string; rows: (string | number | null)[][] };

export type StoredFile = {
  fileId: string;
  name: string;
  kind: Kind;
  buffer: Uint8Array;
  text: string; // context text for the LLM (uploads); "" for generated files
  sid: string; // the session that owns it
  expires: number;
};

const TTL_MS = 60 * 60 * 1000;
const files = new Map<string, StoredFile>();

export function putFile(f: Omit<StoredFile, "fileId" | "expires">, now = Date.now()): StoredFile {
  const stored: StoredFile = { ...f, fileId: randomBytes(16).toString("hex"), expires: now + TTL_MS };
  files.set(stored.fileId, stored);
  return stored;
}

export function getFile(fileId: string, sid: string, now = Date.now()): StoredFile | undefined {
  const f = files.get(fileId);
  if (!f || f.sid !== sid) return undefined;
  if (f.expires <= now) {
    files.delete(fileId);
    return undefined;
  }
  return f;
}

export function sweepFiles(now = Date.now()): number {
  let removed = 0;
  for (const [id, f] of files) {
    if (f.expires <= now) {
      files.delete(id);
      removed++;
    }
  }
  return removed;
}

export function startSweeper(): Timer {
  return setInterval(() => sweepFiles(), 5 * 60 * 1000);
}

export const MIME: Record<Kind, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};
