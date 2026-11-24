// Run by zipbomb.test.ts in a FRESH process, so the memory reading is not mixed with other test files' buffers.
// Prints one JSON line: { status, body, ms, rssBeforeMb, rssAfterMb, grewMb, zippedKb }.
import JSZip from "jszip";
import { handle } from "../src/server.ts";

process.env.DEMO_PASSWORD = "test-pass-123";
process.env.MAX_UPLOAD_MB = "10";

const zip = new JSZip();
zip.file(
  "[Content_Types].xml",
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
);
zip.file("xl/workbook.xml", new Uint8Array(120 * 1024 * 1024)); // 120 MB of zeros
const bomb = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 9 } });

const login = await handle(new Request("http://127.0.0.1/api/login", {
  method: "POST", headers: { "cf-connecting-ip": "7.7.7.7" }, body: JSON.stringify({ password: "test-pass-123" }),
}));
const cookie = login.headers.get("set-cookie")!.split(";")[0];
const form = new FormData();
form.append("file", new File([bomb], "bomb.xlsx"));

Bun.gc(true);
const before = process.memoryUsage().rss;
const t0 = performance.now();
const res = await handle(new Request("http://127.0.0.1/api/upload", { method: "POST", headers: { cookie }, body: form }));
const ms = performance.now() - t0;
const after = process.memoryUsage().rss;
const mb = (n: number) => Math.round((n / 1048576) * 10) / 10;
console.log(JSON.stringify({
  status: res.status, body: await res.json(), ms: Math.round(ms),
  rssBeforeMb: mb(before), rssAfterMb: mb(after), grewMb: mb(after - before), zippedKb: Math.round(bomb.length / 1024),
}));
