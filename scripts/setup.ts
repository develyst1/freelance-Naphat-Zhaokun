// `bun run setup`: make .env from .env.example if missing, fill DEMO_PASSWORD and SESSION_SECRET.
// Prints no value. The owner opens .env to read the password.
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { randomBytes } from "node:crypto";
import { join } from "node:path";

const root = join(import.meta.dir, "..");
const envPath = join(root, ".env");
const examplePath = join(root, ".env.example");

const ALPHABET = "abcdefghjkmnpqrstuvwxyz23456789"; // no 0/o/1/l/i
function randomWord(length: number): string {
  const bytes = randomBytes(length);
  return Array.from(bytes, (b) => ALPHABET[b % ALPHABET.length]).join("");
}

let text = existsSync(envPath) ? readFileSync(envPath, "utf8") : readFileSync(examplePath, "utf8");
const fill = (key: string, value: string) => {
  const re = new RegExp(String.raw`^(${key}=)([ \t]*)(#.*)?$`, "m"); // only an EMPTY value is filled
  if (re.test(text)) text = text.replace(re, (_m, k, _sp, comment) => `${k}${value}${comment ? "  " + comment : ""}`);
};
fill("DEMO_PASSWORD", randomWord(10));
fill("SESSION_SECRET", randomBytes(32).toString("hex"));
writeFileSync(envPath, text, "utf8");
console.log("ตั้งรหัสผ่านไว้ในไฟล์ .env แล้ว");
