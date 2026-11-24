// S1: the zip guard counts REAL inflated bytes. A 120 MB-of-zeros entry (DEFLATE 9) must be refused fast,
// without inflating it all into memory; a normal xlsx still passes.
import { expect, test } from "bun:test";
import JSZip from "jszip";
import { readFileSync, existsSync } from "node:fs";
import { checkZip } from "../src/office/zipguard.ts";
import { FIXTURE, makeKnownFixture } from "./make-fixture.ts";

process.env.DEMO_PASSWORD = "test-pass-123";
process.env.MAX_UPLOAD_MB = "10";

async function makeBomb(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
  );
  zip.file("xl/workbook.xml", new Uint8Array(120 * 1024 * 1024)); // 120 MB of zeros
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 9 } });
}

test("a 120 MB zip bomb is refused by upload (400) quickly, without inflating it all", async () => {
  // measured in a fresh process (test/bomb-probe.ts) so other test files' buffers do not mix into the reading
  const proc = Bun.spawn([process.execPath, "test/bomb-probe.ts"], { cwd: import.meta.dir + "/..", stdout: "pipe", stderr: "pipe" });
  const out = (await new Response(proc.stdout).text()).trim().split(/\r?\n/).pop()!;
  await proc.exited;
  const r = JSON.parse(out);
  console.log(`bomb: ${r.zippedKb} KB zipped → 120 MB inflated · status ${r.status} · ${r.ms} ms · rss before ${r.rssBeforeMb} MB after ${r.rssAfterMb} MB (grew ${r.grewMb} MB)`);
  expect(r.status).toBe(400);
  expect(r.body).toEqual({ error: "รองรับเฉพาะไฟล์ .docx .xlsx .pptx" });
  expect(r.ms).toBeLessThan(5000);
  expect(r.grewMb).toBeLessThan(30);
});

test("checkZip names the reason, and stops at the limit", async () => {
  const bomb = await makeBomb();
  await expect(checkZip(bomb)).rejects.toThrow("unzipped size too large");
  // a lower limit trips on the first chunks of a small but real file
  if (!existsSync(FIXTURE)) await makeKnownFixture();
  await expect(checkZip(new Uint8Array(readFileSync(FIXTURE)), 100)).rejects.toThrow("unzipped size too large");
});

test("a normal xlsx still passes the guard", async () => {
  if (!existsSync(FIXTURE)) await makeKnownFixture();
  const zip = await checkZip(new Uint8Array(readFileSync(FIXTURE)));
  expect(zip.file("xl/workbook.xml")).not.toBeNull();
});
