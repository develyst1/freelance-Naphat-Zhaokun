// Upload checks (SPEC §6): extension + PK magic bytes, and a zip-bomb guard.
import JSZip from "jszip";
import type { Kind } from "../store.ts";

export const MAX_ENTRIES = 2000;
export const MAX_UNZIPPED = 100 * 1024 * 1024;

export function kindFromName(name: string): Kind | null {
  const m = name.toLowerCase().match(/\.(docx|xlsx|pptx)$/);
  return m ? (m[1] as Kind) : null;
}

export function hasZipMagic(buf: Uint8Array): boolean {
  return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04;
}

/**
 * Throws when the zip is unreadable, has too many entries, or inflates past MAX_UNZIPPED.
 * It counts the REAL inflated bytes by streaming each entry, never the size the header declares,
 * and stops inflating the moment the running total passes the limit.
 */
export async function checkZip(buf: Uint8Array, maxUnzipped = MAX_UNZIPPED): Promise<JSZip> {
  const zip = await JSZip.loadAsync(buf);
  const entries = Object.values(zip.files);
  if (entries.length > MAX_ENTRIES) throw new Error("too many entries");
  let total = 0;
  for (const entry of entries) {
    if (entry.dir) continue;
    await new Promise<void>((resolve, reject) => {
      let stream: any;
      try {
        stream = (entry as any).internalStream("uint8array");
      } catch (e) {
        reject(new Error(`entry cannot be read: ${entry.name}`));
        return;
      }
      let done = false;
      const fail = (message: string) => {
        if (done) return;
        done = true;
        try {
          stream.pause();
        } catch {
          /* already stopped */
        }
        reject(new Error(message));
      };
      stream
        .on("data", (chunk: Uint8Array) => {
          total += chunk.length;
          if (total > maxUnzipped) fail("unzipped size too large");
        })
        .on("error", (e: Error) => fail(`entry cannot be read: ${entry.name}: ${e.message}`))
        .on("end", () => {
          if (!done) {
            done = true;
            resolve();
          }
        })
        .resume();
    });
  }
  return zip;
}
