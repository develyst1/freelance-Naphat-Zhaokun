// `bun run tunnel`: cloudflared quick tunnel to the local web (PORT from .env, default 3110).
import { join } from "node:path";
import { readEnvFile } from "./env-file.ts";

const port = readEnvFile(join(import.meta.dir, "..", ".env")).PORT || "3110";
const child = Bun.spawn(["cloudflared", "tunnel", "--url", `http://127.0.0.1:${port}`], {
  stdout: "inherit",
  stderr: "inherit",
});
process.exit(await child.exited);
