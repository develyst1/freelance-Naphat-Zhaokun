# freelance-Naphat-Zhaokun

แชท AI แบบ ChatGPT ที่อ่านและสร้างไฟล์ Word / Excel / PowerPoint ได้ (Excel มีสูตรจริง) · A ChatGPT-like web that reads and creates docx / xlsx / pptx with real Excel formulas.

**วิธีเปิด · How to run** (ต้องมี Bun และ cloudflared · needs Bun + cloudflared):

1. `bun install` — ลงไลบรารี · install the libraries
2. `bun run setup` — สร้าง `.env` และตั้งรหัสผ่าน (เปิดไฟล์ `.env` ดูรหัส) · creates `.env` with a password (read it in `.env`)
3. `bun run start` — เปิดเว็บที่ http://127.0.0.1:3110 · starts the web
4. `bun run tunnel` (อีกหน้าต่าง · second window) — ได้ลิงก์ `https://….trycloudflare.com` ส่งให้ลูกค้า · prints the public link
5. ลิงก์ใช้ได้เฉพาะตอนเครื่องนี้และทั้งสองหน้าต่างยังเปิดอยู่ และเปลี่ยนทุกครั้งที่รัน · the link works only while both run, and changes every run

ทดสอบ · Tests: `bun test` (no network; the AI gateway is mocked).
