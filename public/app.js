"use strict";
// Chat page for SPEC-003 §3. Plain JS, no build, no external requests.
// Model text is never put into innerHTML: the markdown renderer builds DOM nodes.

(() => {
  const $ = (id) => document.getElementById(id);
  const loginView = $("login-view");
  const chatView = $("chat-view");
  const loginForm = $("login-form");
  const passwordInput = $("password");
  const loginError = $("login-error");
  const loginBtn = $("login-btn");
  const messagesEl = $("messages");
  const composer = $("composer");
  const input = $("input");
  const sendBtn = $("send");
  const attachBtn = $("attach");
  const fileInput = $("file-input");
  const chipsEl = $("chips");
  const remainingEl = $("remaining");
  const newChatBtn = $("new-chat");
  const welcomeEl = $("welcome");
  const scrollEl = $("scroll");

  const MAX_HISTORY = 20;
  const NET_ERROR = "เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ ลองใหม่อีกครั้ง"; // thrown fetch, or non-2xx without a JSON error
  let history = []; // [{ role, content }]
  let attached = []; // [{ fileId, name, kind }]
  let busy = false;
  let maxUploadMb = null; // from GET /api/me; null (older server) = no page-side size check

  // ---------- helpers ----------

  function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  }

  async function call(path, options = {}) {
    const res = await fetch(path, { credentials: "same-origin", ...options });
    let body = {};
    try { body = await res.json(); } catch { body = {}; }
    return { status: res.status, ok: res.ok, body };
  }

  function scrollToEnd() {
    scrollEl.scrollTop = scrollEl.scrollHeight;
  }

  function noteLimits(me) {
    maxUploadMb = typeof me.maxUploadMb === "number" && me.maxUploadMb > 0 ? me.maxUploadMb : null;
  }

  // A small file glyph per kind (built with DOM calls, no markup strings).
  const SVG_NS = "http://www.w3.org/2000/svg";
  const KIND_LETTER = { docx: "W", xlsx: "X", pptx: "P" };
  function fileIcon(kind, size) {
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("viewBox", "0 0 20 24");
    svg.setAttribute("width", String(size * (20 / 24)));
    svg.setAttribute("height", String(size));
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("class", `file-icon kind-${kind}`);
    const sheet = document.createElementNS(SVG_NS, "path");
    sheet.setAttribute("class", "sheet");
    sheet.setAttribute("d", "M3 0h10l7 7v14a3 3 0 0 1-3 3H3a3 3 0 0 1-3-3V3a3 3 0 0 1 3-3z");
    const fold = document.createElementNS(SVG_NS, "path");
    fold.setAttribute("class", "fold");
    fold.setAttribute("d", "M13 0l7 7h-5a2 2 0 0 1-2-2z");
    const letter = document.createElementNS(SVG_NS, "text");
    letter.setAttribute("class", "letter");
    letter.setAttribute("x", "10");
    letter.setAttribute("y", "18.5");
    letter.setAttribute("text-anchor", "middle");
    letter.textContent = KIND_LETTER[kind] || "?";
    svg.append(sheet, fold, letter);
    return svg;
  }

  // The welcome + examples show only while the conversation is empty.
  function syncWelcome() {
    welcomeEl.hidden = messagesEl.childElementCount > 0;
  }
  new MutationObserver(syncWelcome).observe(messagesEl, { childList: true });

  function setRemaining(n) {
    if (typeof n === "number") remainingEl.textContent = `เหลือ ${n} ข้อความในชั่วโมงนี้`;
  }

  async function refreshRemaining() {
    try {
      const me = await call("/api/me");
      if (me.ok) { noteLimits(me.body); setRemaining(me.body.remaining); }
    } catch { /* the error line is already shown */ }
  }

  function systemLine(text, isError) {
    const line = el("p", isError ? "sys sys-error" : "sys", text);
    if (isError) line.setAttribute("role", "alert");
    messagesEl.append(line);
    scrollToEnd();
  }

  // ---------- views ----------

  function showLogin() {
    chatView.hidden = true;
    loginView.hidden = false;
    passwordInput.value = "";
    passwordInput.focus();
  }

  function showChat(remaining) {
    loginView.hidden = true;
    chatView.hidden = false;
    setRemaining(remaining);
    input.focus();
  }

  // ---------- safe markdown (paragraphs, **bold**, lists, `code`, ``` blocks, # headings, | tables |, <br>) ----------
  // Every node is built with DOM calls. Only <br> becomes an element; any other HTML in model text stays literal text.

  function renderInline(parent, text) {
    const re = /(\*\*[^*]+\*\*|`[^`]+`|<br\s*\/?>)/gi;
    let last = 0;
    let m;
    while ((m = re.exec(text))) {
      if (m.index > last) parent.append(document.createTextNode(text.slice(last, m.index)));
      const tok = m[0];
      if (tok.startsWith("**")) parent.append(el("strong", null, tok.slice(2, -2)));
      else if (tok.startsWith("`")) parent.append(el("code", null, tok.slice(1, -1)));
      else parent.append(el("br"));
      last = m.index + tok.length;
    }
    if (last < text.length) parent.append(document.createTextNode(text.slice(last)));
  }

  const FENCE_RE = /^\s*```/;
  const UL_RE = /^\s*[-*]\s+/;
  const OL_RE = /^\s*\d+[.)]\s+/;
  const HEADING_RE = /^\s*(#{1,3})\s+(.*)$/;
  const SEP_RE = /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$/;

  function isTableRow(line) {
    return line != null && line.includes("|") && line.trim() !== "";
  }
  function isTableStart(lines, i) {
    return isTableRow(lines[i]) && i + 1 < lines.length && SEP_RE.test(lines[i + 1]) && lines[i + 1].includes("-");
  }
  function splitRow(line) {
    let t = line.trim();
    if (t.startsWith("|")) t = t.slice(1);
    if (t.endsWith("|")) t = t.slice(0, -1);
    return t.split("|").map((c) => c.trim());
  }
  function mdTable(header, rows) {
    const wrap = el("div", "table-block");
    const scroll = el("div", "table-scroll");
    scroll.tabIndex = 0;
    const table = el("table");
    const thead = el("thead");
    const htr = el("tr");
    for (const h of header) { const th = el("th"); th.scope = "col"; renderInline(th, h); htr.append(th); }
    thead.append(htr);
    const tbody = el("tbody");
    for (const r of rows) {
      const tr = el("tr");
      for (let c = 0; c < header.length; c++) { const td = el("td"); renderInline(td, r[c] ?? ""); tr.append(td); }
      tbody.append(tr);
    }
    table.append(thead, tbody);
    scroll.append(table);
    wrap.append(scroll);
    return wrap;
  }
  function startsBlock(lines, i) {
    const l = lines[i];
    return FENCE_RE.test(l) || UL_RE.test(l) || OL_RE.test(l) || HEADING_RE.test(l) || isTableStart(lines, i);
  }

  function renderMarkdown(container, text) {
    const lines = String(text).replace(/\r\n/g, "\n").split("\n");
    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      if (FENCE_RE.test(line)) {
        const code = [];
        i++;
        while (i < lines.length && !FENCE_RE.test(lines[i])) code.push(lines[i++]);
        i++; // closing fence
        const pre = el("pre");
        pre.append(el("code", null, code.join("\n")));
        container.append(pre);
        continue;
      }
      const h = HEADING_RE.exec(line);
      if (h) {
        const node = el(`h${h[1].length + 2}`, "md-heading"); // # → h3, ## → h4, ### → h5
        renderInline(node, h[2]);
        container.append(node);
        i++;
        continue;
      }
      if (isTableStart(lines, i)) {
        const header = splitRow(lines[i]);
        i += 2;
        const rows = [];
        while (i < lines.length && isTableRow(lines[i]) && !FENCE_RE.test(lines[i])) rows.push(splitRow(lines[i++]));
        container.append(mdTable(header, rows));
        continue;
      }
      if (UL_RE.test(line) || OL_RE.test(line)) {
        const ordered = OL_RE.test(line);
        const list = el(ordered ? "ol" : "ul");
        const itemRe = ordered ? OL_RE : UL_RE;
        while (i < lines.length && itemRe.test(lines[i])) {
          const li = el("li");
          renderInline(li, lines[i].replace(itemRe, ""));
          list.append(li);
          i++;
        }
        container.append(list);
        continue;
      }
      if (line.trim() === "") { i++; continue; }
      const para = [];
      while (i < lines.length && lines[i].trim() !== "" && (para.length === 0 || !startsBlock(lines, i))) para.push(lines[i++]);
      const p = el("p");
      para.forEach((t, idx) => {
        if (idx) p.append(el("br"));
        renderInline(p, t);
      });
      container.append(p);
    }
  }

  // ---------- tables ----------

  function tableBlock(title, rows) {
    const wrap = el("div", "table-block");
    wrap.append(el("p", "table-title", title));
    const scroll = el("div", "table-scroll");
    scroll.tabIndex = 0;
    scroll.setAttribute("role", "region");
    scroll.setAttribute("aria-label", title);
    const table = el("table");
    const tbody = el("tbody");
    for (const row of rows || []) {
      const tr = el("tr");
      for (const cell of row || []) {
        const td = el("td", typeof cell === "number" ? "num" : null);
        td.textContent = cell == null ? ""
          : typeof cell === "number" ? cell.toLocaleString("en-US", { maximumFractionDigits: 4 })
          : String(cell);
        tr.append(td);
      }
      tbody.append(tr);
    }
    table.append(tbody);
    scroll.append(table);
    wrap.append(scroll);
    return wrap;
  }

  // ---------- messages ----------

  function addUser(text) {
    messagesEl.append(el("div", "msg msg-user", text));
    scrollToEnd();
  }

  function addAssistant(data) {
    const box = el("div", "msg msg-assistant");
    renderMarkdown(box, data.reply || "");
    if (Array.isArray(data.files) && data.files.length) {
      const dl = el("div", "downloads");
      for (const f of data.files) {
        const a = el("a", `file-card`);
        a.href = f.url;
        a.setAttribute("download", f.name);
        a.setAttribute("aria-label", `ดาวน์โหลด ${f.name}`);
        const text = el("span", "file-card-text");
        text.append(el("span", "file-card-name", f.name), el("span", "file-card-action", "ดาวน์โหลด"));
        a.append(fileIcon(f.kind, 28), text);
        dl.append(a);
      }
      box.append(dl);
    }
    if (Array.isArray(data.tables)) {
      for (const t of data.tables) box.append(tableBlock(`${t.sheet} (คำนวณโดยเครื่องคำนวณสูตร)`, t.rows));
    }
    messagesEl.append(box);
    scrollToEnd();
  }

  function setBusy(on) {
    busy = on;
    sendBtn.disabled = on;
    attachBtn.disabled = on;
  }

  // ---------- chips ----------

  function renderChips() {
    chipsEl.replaceChildren();
    for (const f of attached) {
      const li = el("li", "chip");
      li.append(fileIcon(f.kind, 20));
      li.append(el("span", "chip-name", f.name));
      const rm = el("button", "chip-remove", "×");
      rm.type = "button";
      rm.setAttribute("aria-label", `เอา ${f.name} ออก`);
      rm.addEventListener("click", () => {
        attached = attached.filter((x) => x.fileId !== f.fileId);
        renderChips();
        input.focus();
      });
      li.append(rm);
      chipsEl.append(li);
    }
    chipsEl.hidden = attached.length === 0;
  }

  // ---------- error handling shared by chat and upload ----------

  function handleError(res) {
    if (res.status === 401) { showLogin(); return; }
    if (res.status === 429) setRemaining(0);
    systemLine(res.body.error || NET_ERROR, true);
  }

  // ---------- actions ----------

  async function login(e) {
    e.preventDefault();
    loginError.hidden = true;
    passwordInput.removeAttribute("aria-invalid");
    loginBtn.disabled = true;
    try {
      const res = await call("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password: passwordInput.value }),
      });
      if (!res.ok) {
        loginError.textContent = res.body.error || NET_ERROR;
        loginError.hidden = false;
        passwordInput.setAttribute("aria-invalid", "true");
        passwordInput.select();
        return;
      }
      const me = await call("/api/me");
      noteLimits(me.body);
      showChat(me.body.remaining);
    } catch (err) {
      loginError.textContent = NET_ERROR;
      loginError.hidden = false;
    } finally {
      loginBtn.disabled = false;
    }
  }

  async function send(e) {
    if (e) e.preventDefault();
    stopMic(); // sending ends dictation; the text already in the box is what gets sent
    const text = input.value.trim();
    if (!text || busy) return;
    addUser(text);
    history.push({ role: "user", content: text });
    input.value = "";
    autosize();
    setBusy(true);
    const thinking = el("div", "msg msg-assistant msg-thinking", "กำลังคิด…");
    messagesEl.append(thinking);
    scrollToEnd();
    try {
      const res = await call("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: history.slice(-MAX_HISTORY),
          fileIds: attached.map((f) => f.fileId),
        }),
      });
      thinking.remove();
      if (!res.ok) {
        history.pop(); // the failed turn is not part of the conversation
        handleError(res);
        if (res.status !== 401) refreshRemaining(); // the server may count a failed chat
        return;
      }
      history.push({ role: "assistant", content: res.body.reply || "" });
      addAssistant(res.body);
      setRemaining(res.body.remaining);
    } catch (err) {
      thinking.remove();
      history.pop();
      systemLine(NET_ERROR, true);
    } finally {
      setBusy(false);
      input.focus();
    }
  }

  async function upload() {
    const file = fileInput.files && fileInput.files[0];
    fileInput.value = "";
    if (!file) return;
    if (maxUploadMb != null && file.size > maxUploadMb * 1024 * 1024) {
      systemLine(`ไฟล์ใหญ่เกิน ${maxUploadMb} MB`, true); // checked here so a too-big body is never sent (TASK-006 D1)
      return;
    }
    setBusy(true);
    try {
      const form = new FormData();
      form.append("file", file);
      const res = await call("/api/upload", { method: "POST", body: form });
      if (!res.ok) { handleError(res); return; }
      const { fileId, name, kind, sheets } = res.body;
      attached.push({ fileId, name, kind });
      renderChips();
      systemLine(`แนบ ${name} แล้ว`, false);
      if (kind === "xlsx" && Array.isArray(sheets) && sheets.length) {
        const block = tableBlock(`${sheets[0].name} (คำนวณโดยเครื่องคำนวณสูตร)`, sheets[0].rows);
        block.classList.add("standalone");
        messagesEl.append(block);
        scrollToEnd();
      }
    } catch (err) {
      systemLine(NET_ERROR, true);
    } finally {
      setBusy(false);
    }
  }

  function newChat() {
    history = [];
    attached = [];
    renderChips();
    messagesEl.replaceChildren();
    input.value = "";
    autosize();
    input.focus();
  }

  function autosize() {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight + 2, 160) + "px";
  }

  // ---------- mic: speech to text into the input (browser Web Speech API) ----------
  // Never sends by itself. Text typed before dictation is kept; only this session's words are appended.

  const micGroup = $("mic-group");
  const micBtn = $("mic");
  const micLangBtn = $("mic-lang");
  const MIC_LANGS = { th: { code: "th-TH", label: "ไทย" }, en: { code: "en-US", label: "EN" } };
  const MIC_ERRORS = {
    "not-allowed": "ไม่ได้รับอนุญาตให้ใช้ไมโครโฟน — กดอนุญาตที่แถบที่อยู่ของเบราว์เซอร์",
    "service-not-allowed": "ไม่ได้รับอนุญาตให้ใช้ไมโครโฟน — กดอนุญาตที่แถบที่อยู่ของเบราว์เซอร์",
    "no-speech": "ไม่ได้ยินเสียง ลองพูดอีกครั้ง",
    network: "แปลงเสียงไม่ได้ ต้องต่ออินเทอร์เน็ต",
  };
  const MIC_ERROR_OTHER = "ใช้ไมค์ไม่ได้ในตอนนี้";
  let micLang = "th";
  try { if (localStorage.getItem("micLang") === "en") micLang = "en"; } catch { /* storage blocked: default Thai */ }
  let recognizer = null; // the live session, or null when idle
  let micBase = ""; // what the box held when dictation started (plus committed words)

  // `?nospeech` is a test-only switch that simulates a browser without the API (TASK-012 proof).
  const NO_SPEECH = new URLSearchParams(location.search).has("nospeech");
  function speechCtor() {
    if (NO_SPEECH) return null;
    return window.SpeechRecognition || window.webkitSpeechRecognition || null;
  }

  function renderMicLang() {
    micLangBtn.textContent = MIC_LANGS[micLang].label;
  }

  function setListening(on) {
    micBtn.classList.toggle("is-listening", on);
    micBtn.setAttribute("aria-pressed", on ? "true" : "false");
    input.placeholder = on ? "กำลังฟัง… พูดได้เลย" : "";
  }

  function stopMic() {
    if (!recognizer) return;
    const r = recognizer;
    recognizer = null;
    setListening(false);
    try { r.stop(); } catch { /* already stopped */ }
  }

  function startMic() {
    const Ctor = speechCtor();
    if (!Ctor) return;
    let r;
    try { r = new Ctor(); } catch { systemLine(MIC_ERROR_OTHER, true); return; }
    r.lang = MIC_LANGS[micLang].code;
    r.interimResults = true;
    r.continuous = true;
    micBase = input.value;
    const joiner = () => (micBase && !/\s$/.test(micBase) ? " " : "");
    r.onresult = (ev) => {
      if (recognizer !== r) return;
      let interim = "";
      for (let k = ev.resultIndex; k < ev.results.length; k++) {
        const res = ev.results[k];
        const text = res[0] ? res[0].transcript : "";
        if (res.isFinal) micBase += joiner() + text.trim();
        else interim += text;
      }
      interim = interim.trim();
      input.value = interim ? micBase + joiner() + interim : micBase;
      autosize();
    };
    r.onerror = (ev) => {
      const code = ev && ev.error;
      if (code === "aborted") return; // our own stop
      systemLine(MIC_ERRORS[code] || MIC_ERROR_OTHER, true);
    };
    r.onend = () => {
      // Only a browser-side end (silence, error) touches the box. After our own stop or a send,
      // the box already holds what the person saw, or was cleared by send: leave it alone.
      if (recognizer !== r) return;
      recognizer = null;
      setListening(false);
      input.value = micBase; // drop half-heard interim words; committed text stays
      autosize();
    };
    recognizer = r;
    setListening(true);
    try { r.start(); } catch { recognizer = null; setListening(false); systemLine(MIC_ERROR_OTHER, true); }
  }

  if (speechCtor()) {
    micGroup.hidden = false;
    renderMicLang();
    micBtn.addEventListener("click", () => (recognizer ? stopMic() : startMic()));
    micLangBtn.addEventListener("click", () => {
      micLang = micLang === "th" ? "en" : "th";
      try { localStorage.setItem("micLang", micLang); } catch { /* not remembered; still works */ }
      renderMicLang();
    });
  }

  // ---------- wiring ----------

  loginForm.addEventListener("submit", login);
  composer.addEventListener("submit", send);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
      e.preventDefault();
      send();
    }
  });
  input.addEventListener("input", autosize);
  attachBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", upload);
  newChatBtn.addEventListener("click", newChat);
  // Example prompts only fill the box; the person still presses send (no surprise spend).
  for (const btn of welcomeEl.querySelectorAll(".example")) {
    btn.addEventListener("click", () => {
      input.value = btn.dataset.prompt || "";
      autosize();
      input.focus();
    });
  }

  (async () => {
    try {
      const me = await call("/api/me");
      if (me.ok) { noteLimits(me.body); showChat(me.body.remaining); }
      else showLogin();
    } catch {
      showLogin();
    }
  })();
})();
