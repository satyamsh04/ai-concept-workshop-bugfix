import "./taskpane.css";

// ===== Gateway Config =====
const GW_PORT = 18789;
const BACKOFF_BASE = 1200;
const BACKOFF_MAX = 20000;

// System instructions prepended to the first message for each email session.
const ACADEMIC_SYSTEM_PROMPT = `[Academic Email Assistant]
You help university students and staff manage academic emails professionally.
When classifying an email, respond ONLY with valid JSON using exactly these keys:
  "urgency": "high" | "medium" | "low"
  "intent": "inquiry" | "assignment" | "extension" | "grade" | "complaint" | "feedback" | "administrative" | "other"
  "summary": one sentence, max 20 words
When drafting replies, write plain professional text only — no markdown, no bullet points unless asked.
Keep all replies concise and appropriate for a university context.`;

// ===== State =====
let socket = null;
let isConnected = false;
let authToken = null;
let sessionKey = "academic:email:default";
let activeEmailId = null;        // hash of current email, used to detect email changes
let currentEmail = null;         // { subject, from, to, date, body }
let contextSent = false;         // whether ACADEMIC_SYSTEM_PROMPT + email context has been sent this session
let streamBuffer = "";           // accumulates streaming agent.delta text
let activeRunId = null;
let waitingForResponse = false;
let rpcSeq = 0;
let pendingRpc = new Map();      // id → { resolve, reject }
let historyFetching = false;
let lastShownMsgId = null;
let retryCount = 0;
let retryTimer = null;

// ===== DOM Helper =====
const $ = id => document.getElementById(id);

// ===== Boot =====
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    applyTheme();
    bindEvents();
    loadToken();
    readEmail();

    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => readEmail(),
        () => {}
      );
    }
  }
});

// ===== Theme =====
function applyTheme() {
  let dark = false;
  try {
    const t = Office.context.officeTheme;
    if (t && t.bodyBackgroundColor) {
      const hex = t.bodyBackgroundColor.replace("#", "");
      const r = parseInt(hex.slice(0, 2), 16);
      const g = parseInt(hex.slice(2, 4), 16);
      const b = parseInt(hex.slice(4, 6), 16);
      dark = (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
    }
  } catch (_) {}

  if (!dark && window.matchMedia) {
    dark = window.matchMedia("(prefers-color-scheme: dark)").matches;
    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", e => {
      document.documentElement.setAttribute("data-theme", e.matches ? "dark" : "light");
    });
  }
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
}

// ===== Token =====
function loadToken() {
  try { authToken = localStorage.getItem("acad-gateway-token") || null; } catch (_) {}
  if (!authToken) {
    showTokenPrompt();
  } else {
    connectGateway();
  }
}

function showTokenPrompt() {
  const saved = getSavedSystemPrompt();
  const div = document.createElement("div");
  div.className = "message sys-msg";
  div.id = "settings-panel";
  div.innerHTML = `<div class="msg-body" style="text-align:left">
    <strong>Setup</strong><br><br>
    <label style="font-size:11px;color:var(--text-secondary)">Gateway Token</label>
    <input type="password" id="token-field" placeholder="Paste your gateway token..."
      value="${authToken || ""}"
      style="width:100%;padding:5px 8px;border:1px solid var(--border);border-radius:4px;
             background:var(--bg-input);color:var(--text-primary);font-size:12px;
             margin-bottom:8px;font-family:monospace"/>
    <label style="font-size:11px;color:var(--text-secondary)">Custom Instructions (optional)</label>
    <textarea id="prompt-field" rows="3"
      placeholder="e.g. Always reply formally. Never use bullet points."
      style="width:100%;padding:5px 8px;border:1px solid var(--border);border-radius:4px;
             background:var(--bg-input);color:var(--text-primary);font-size:12px;
             margin-bottom:8px;font-family:var(--font);resize:vertical">${saved}</textarea>
    <button id="save-settings-btn"
      style="padding:5px 12px;background:var(--accent);color:#fff;border:none;
             border-radius:4px;cursor:pointer;font-size:12px">
      Save &amp; Connect
    </button>
    <br><small style="color:var(--text-muted)">Token location: ~/.openclaw/openclaw.json → gateway.auth.token</small>
  </div>`;
  $("chat-messages").appendChild(div);

  setTimeout(() => {
    const btn = document.getElementById("save-settings-btn");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const t = document.getElementById("token-field").value.trim();
      const p = document.getElementById("prompt-field").value.trim();
      if (t) {
        try { localStorage.setItem("acad-gateway-token", t); } catch (_) {}
        authToken = t;
      }
      try { localStorage.setItem("acad-custom-instructions", p); } catch (_) {}
      div.remove();
      addMessage("sys", "Settings saved. Connecting to gateway...");
      connectGateway();
    });
  }, 80);
}

function getSavedSystemPrompt() {
  try { return localStorage.getItem("acad-custom-instructions") || ""; } catch (_) { return ""; }
}

// ===== Email Reader =====
function readEmail() {
  const item = Office.context.mailbox.item;
  if (!item) { showNoEmail(); return; }

  try {
    const subject = item.subject || "(No subject)";
    const from    = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "Unknown";
    const date    = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : "";
    const to      = item.to ? item.to.map(r => `${r.displayName} <${r.emailAddress}>`).join(", ") : "";

    item.body.getAsync(Office.CoercionType.Text, result => {
      const body = result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
      currentEmail = { subject, from, to, date, body };

      const newId = hashString(subject + "|" + from + "|" + date);
      if (newId !== activeEmailId) {
        activeEmailId = newId;
        sessionKey = `academic:email:${newId}`;
        contextSent = false;
        lastShownMsgId = null;
        clearBadges();
        clearChatMessages();
        if (isConnected) loadHistory();
      }

      renderEmailHeader(subject, from, date);
    });
  } catch (_) {
    showNoEmail();
    addMessage("err", "Could not read email details.");
  }
}

function showNoEmail() {
  $("email-placeholder").style.display = "flex";
  $("email-info").style.display = "none";
  currentEmail = null;
}

function renderEmailHeader(subject, from, date) {
  $("email-placeholder").style.display = "none";
  $("email-info").style.display = "block";
  $("email-subject").textContent = subject;
  $("email-from").textContent = from;
  $("email-date").textContent = date;
}

// ===== Gateway Connection =====
function getGatewayUrl() {
  const loc = window.location;
  if (loc.protocol === "https:") return `wss://${loc.host}/ai-gateway`;
  return `ws://127.0.0.1:${GW_PORT}`;
}

function connectGateway() {
  if (socket && (socket.readyState === WebSocket.OPEN || socket.readyState === WebSocket.CONNECTING)) return;

  setStatus("connecting");

  try {
    socket = new WebSocket(getGatewayUrl());
  } catch (_) {
    setStatus("disconnected");
    scheduleRetry();
    return;
  }

  socket.onopen = () => {
    setStatus("connecting");
    // Give the server 2s to send connect.challenge; proceed anyway if it doesn't arrive.
    setTimeout(() => { if (!isConnected) sendHandshake(); }, 2000);
  };

  socket.onmessage = e => handleIncoming(String(e.data || ""));

  socket.onclose = () => {
    isConnected = false;
    pendingRpc.clear();
    setStatus("disconnected");
    scheduleRetry();
  };

  socket.onerror = () => {};
}

function scheduleRetry() {
  if (retryTimer) return;
  const delay = Math.min(BACKOFF_BASE * Math.pow(1.7, retryCount), BACKOFF_MAX);
  retryCount++;
  retryTimer = setTimeout(() => {
    retryTimer = null;
    connectGateway();
  }, delay);
}

// ===== Handshake (connect RPC) =====
function sendHandshake() {
  const params = {
    minProtocol: 3,
    maxProtocol: 3,
    client: {
      id: "academic-email-assistant",
      version: "1.0.0",
      platform: navigator.platform || "web",
      mode: "webchat",
      instanceId: "acad-" + Date.now(),
    },
    role: "operator",
    scopes: ["operator.admin"],
    caps: ["tool-events"],
    auth: authToken ? { token: authToken } : {},
  };

  callRpc("connect", params)
    .then(result => {
      isConnected = true;
      retryCount = 0;
      setStatus("connected");
      if (result && result.sessionKey) sessionKey = result.sessionKey;
      loadHistory();
    })
    .catch(() => setStatus("disconnected"));
}

// ===== RPC Layer =====
function callRpc(method, params) {
  return new Promise((resolve, reject) => {
    if (!socket || socket.readyState !== WebSocket.OPEN) {
      reject(new Error("Not connected"));
      return;
    }
    const id = String(++rpcSeq);
    pendingRpc.set(id, { resolve, reject });
    socket.send(JSON.stringify({ type: "req", id, method, params }));
  });
}

// ===== Message Handler =====
function handleIncoming(raw) {
  let data;
  try { data = JSON.parse(raw); } catch (_) { return; }

  if (data.type === "res" && pendingRpc.has(String(data.id))) {
    const { resolve, reject } = pendingRpc.get(String(data.id));
    pendingRpc.delete(String(data.id));
    if (data.ok === false) {
      reject(new Error(data.error?.message || data.error?.code || "RPC error"));
    } else {
      resolve(data.result || data.payload || data);
    }
    return;
  }

  if (data.type === "event") {
    handleEvent(data);
  }
}

// ===== Event Handler =====
function flushStream() {
  const text = streamBuffer.trim();
  if (text) addMessage("ai", text);
  streamBuffer = "";
}

function handleEvent(evt) {
  const event   = evt.event || "";
  const payload = evt.payload || evt.data || {};

  switch (event) {

    case "connect.challenge":
      sendHandshake();
      break;

    case "agent.run": {
      const phase = payload.phase || payload.data?.phase || "";
      if (phase === "start") {
        activeRunId = payload.runId || null;
        showTyping();
      } else if (phase === "end" || phase === "error") {
        activeRunId = null;
        flushStream();
        hideTyping();
        if (waitingForResponse) {
          waitingForResponse = false;
          fetchLatestReply();
        }
      }
      break;
    }

    case "chat": {
      const state = payload.state || "";
      if (state === "start" || state === "started") {
        flushStream();
        showTyping();
      } else if (state === "final" || state === "end" || state === "error") {
        flushStream();
        if (!activeRunId) {
          hideTyping();
          if (waitingForResponse) { waitingForResponse = false; fetchLatestReply(); }
        }
      }
      break;
    }

    case "agent.delta":
    case "chat.delta": {
      const chunk = payload.delta || payload.text || payload.content || "";
      if (chunk) {
        streamBuffer += chunk;
        renderStreamingBubble(streamBuffer);
      }
      break;
    }

    case "agent.message":
    case "chat.message": {
      flushStream();
      const content = payload.content || payload.text || payload.message || "";
      if (content) {
        const text = typeof content === "string" ? content : JSON.stringify(content);
        addMessage("ai", text);
      }
      if (!activeRunId) hideTyping();
      break;
    }

    case "agent.tool_call":
    case "tool_call":
      flushStream();
      setTypingLabel(payload.name ? `Using ${payload.name}…` : "Working…");
      showTyping();
      break;

    case "agent.tool_result":
    case "tool_result":
      setTypingLabel("Processing…");
      break;

    default:
      if (payload.content || payload.text || payload.message) {
        const text = payload.content || payload.text || payload.message;
        if (typeof text === "string" && text.trim()) {
          flushStream();
          addMessage("ai", text.trim());
          if (!activeRunId) hideTyping();
        }
      }
      break;
  }
}

// ===== History Loader =====
function loadHistory() {
  callRpc("chat.history", { sessionKey, limit: 50 })
    .then(result => {
      const msgs = extractMessages(result);
      for (const msg of msgs) {
        const text = extractText(msg);
        if (!text) continue;
        if (msg.role === "user") {
          const m = text.match(/User message:\s*([\s\S]*)/);
          addMessage("user", m ? m[1].trim() : text.trim());
        } else if (msg.role === "assistant") {
          addMessage("ai", text.trim());
          lastShownMsgId = msg.__openclaw?.id || msg.responseId || msg.timestamp || null;
          tryRenderClassification(text.trim());
        }
      }
    })
    .catch(() => {});
}

function fetchLatestReply() {
  if (historyFetching) return;
  historyFetching = true;

  callRpc("chat.history", { sessionKey, limit: 10 })
    .then(result => {
      const msgs = extractMessages(result);
      if (!msgs.length) {
        historyFetching = false;
        setTimeout(fetchLatestReply, 2000);
        return;
      }
      for (let i = msgs.length - 1; i >= 0; i--) {
        const msg = msgs[i];
        if (msg.role !== "assistant") continue;
        const text = extractText(msg);
        const msgId = msg.__openclaw?.id || msg.responseId || msg.timestamp || i;
        if (text.trim() && msgId !== lastShownMsgId) {
          historyFetching = false;
          lastShownMsgId = msgId;
          addMessage("ai", text.trim());
          tryRenderClassification(text.trim());
        } else if (msgId === lastShownMsgId) {
          historyFetching = false;
          setTimeout(fetchLatestReply, 2000);
        }
        return;
      }
      historyFetching = false;
      setTimeout(fetchLatestReply, 2000);
    })
    .catch(() => { historyFetching = false; });
}

function extractMessages(result) {
  if (!result) return [];
  if (Array.isArray(result.messages)) return result.messages;
  if (Array.isArray(result)) return result;
  if (result.history && Array.isArray(result.history)) return result.history;
  for (const k of Object.keys(result)) {
    if (Array.isArray(result[k])) return result[k];
  }
  return [];
}

function extractText(msg) {
  if (typeof msg.content === "string") return msg.content;
  if (Array.isArray(msg.content)) {
    return msg.content.filter(c => c.type === "text").map(c => c.text || "").join("\n");
  }
  return "";
}

// ===== Send Message =====
function sendMessage(text) {
  if (!isConnected) {
    addMessage("err", "Not connected to gateway. Reconnecting…");
    connectGateway();
    return;
  }

  let fullText = text;

  // On first message for this email, prepend system context
  if (currentEmail && !contextSent) {
    const body = (currentEmail.body || "").slice(0, 3000);
    const custom = getSavedSystemPrompt();
    let prefix = ACADEMIC_SYSTEM_PROMPT;
    if (custom) prefix += `\n\nAdditional instructions: ${custom}`;
    prefix += `\n\n[Email Context]\nSubject: ${currentEmail.subject}\nFrom: ${currentEmail.from}\nTo: ${currentEmail.to}\nDate: ${currentEmail.date}\n\nBody:\n${body}\n\n---\n\n`;
    fullText = prefix + `User message: ${text}`;
    contextSent = true;
  }

  callRpc("chat.send", {
    sessionKey,
    message: fullText,
    deliver: false,
    idempotencyKey: crypto.randomUUID(),
  })
    .then(() => { waitingForResponse = true; showTyping(); })
    .catch(err => { hideTyping(); addMessage("err", "Send failed: " + err.message); });
}

// ===== Classification Feature =====
function classifyEmail() {
  if (!currentEmail) { addMessage("err", "No email selected."); return; }
  addMessage("user", "Classify this email");
  showTyping();
  waitingForResponse = true;
  sendMessage('Classify this email. Respond with ONLY a JSON object with keys "urgency", "intent", and "summary".');
}

function tryRenderClassification(text) {
  const json = extractJson(text);
  if (!json) return;
  if (!json.urgency && !json.intent) return;
  renderBadges(json);
}

function extractJson(text) {
  try {
    const clean = text.trim();
    if (clean.startsWith("{")) return JSON.parse(clean);
    const match = clean.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (match) return JSON.parse(match[1].trim());
  } catch (_) {}
  return null;
}

function renderBadges(data) {
  const container = $("email-tags");
  if (!container) return;

  clearBadges();

  if (data.urgency) {
    const b = document.createElement("span");
    b.className = `badge badge-${data.urgency}`;
    b.textContent = data.urgency.toUpperCase();
    container.appendChild(b);
  }

  if (data.intent) {
    const b = document.createElement("span");
    b.className = "badge badge-info";
    b.textContent = data.intent;
    container.appendChild(b);
  }

  if (data.summary) {
    const s = document.createElement("div");
    s.style.cssText = "font-size:11px;color:var(--text-secondary);margin-top:4px;font-style:italic;";
    s.textContent = data.summary;
    container.appendChild(s);
  }

  $("email-info").style.display = "block";
  $("email-tags").style.display = "flex";
  $("email-tags").style.flexWrap = "wrap";
}

function clearBadges() {
  const t = $("email-tags");
  if (t) t.innerHTML = "";
}

// ===== Quick Actions =====
function handleQuickAction(type) {
  if (!currentEmail) { addMessage("err", "No email selected."); return; }

  const prompts = {
    summarise: "Summarise this email in 2-3 sentences.",
    extension: "Draft a professional extension request reply to this email on my behalf. Keep it polite and concise.",
  };

  if (type === "classify") {
    classifyEmail();
    return;
  }

  const prompt = prompts[type];
  if (!prompt) return;
  addMessage("user", prompt);
  showTyping();
  waitingForResponse = true;
  sendMessage(prompt);
}

// ===== Draft Reply =====
function draftReply() {
  if (!currentEmail) { addMessage("err", "No email selected."); return; }
  addMessage("user", "Draft a reply to this email");
  showTyping();
  waitingForResponse = true;
  sendMessage("Please draft a professional reply to this email. Match the tone of the original. Plain text only.");
}

function useDraft() {
  const aiMsgs = document.querySelectorAll(".ai-msg .msg-body");
  const last = aiMsgs[aiMsgs.length - 1];
  if (!last) { addMessage("err", "No draft available. Click Draft Reply first."); return; }

  const item = Office.context.mailbox.item;
  if (!item) { addMessage("err", "No email selected."); return; }

  try {
    item.displayReplyForm(last.textContent);
    addMessage("sys", "Draft opened in Outlook. Review and send when ready.");
  } catch (err) {
    addMessage("err", "Could not open draft: " + err.message);
  }
}

// ===== Chat UI =====
function addMessage(role, text) {
  const existing = document.querySelector(".streaming-bubble");
  if (existing) existing.remove();

  const container = $("chat-messages");
  const div = document.createElement("div");
  const classMap = { user: "message user-msg", ai: "message ai-msg", sys: "message sys-msg", err: "message err-msg" };
  div.className = classMap[role] || "message sys-msg";

  const body = document.createElement("div");
  body.className = "msg-body";
  body.textContent = text;
  div.appendChild(body);
  container.appendChild(div);
  scrollDown();

  if (role === "ai") $("use-draft-btn").disabled = false;
}

function renderStreamingBubble(text) {
  hideTyping();
  let el = document.querySelector(".streaming-bubble");
  if (!el) {
    const container = $("chat-messages");
    el = document.createElement("div");
    el.className = "message ai-msg streaming-bubble";
    const body = document.createElement("div");
    body.className = "msg-body";
    el.appendChild(body);
    container.appendChild(el);
  }
  el.querySelector(".msg-body").textContent = text;
  scrollDown();
}

function clearChatMessages() {
  const container = $("chat-messages");
  container.querySelectorAll(".message:not(.sys-msg)").forEach(m => m.remove());
  $("use-draft-btn").disabled = true;
}

function showTyping() { $("typing-indicator").style.display = "flex"; scrollDown(); }
function hideTyping() { $("typing-indicator").style.display = "none"; }
function setTypingLabel(text) { const l = $("typing-label"); if (l) l.textContent = text; showTyping(); }
function scrollDown() {
  const c = $("chat-messages");
  requestAnimationFrame(() => { c.scrollTop = c.scrollHeight; });
}

// ===== Status Bar =====
function setStatus(state) {
  const bar = $("status-bar");
  bar.className = "status-bar " + state;
  const labels = { connected: "Academic AI Ready", connecting: "Connecting…", disconnected: "Disconnected" };
  $("status-text").textContent = labels[state] || state;
}

// ===== Utilities =====
function hashString(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = ((h << 5) - h + str.charCodeAt(i)) | 0;
  return Math.abs(h).toString(36);
}

// ===== Event Bindings =====
function bindEvents() {
  $("send-btn").addEventListener("click", handleSend);
  $("draft-btn").addEventListener("click", draftReply);
  $("use-draft-btn").addEventListener("click", useDraft);
  $("classify-btn").addEventListener("click", () => handleQuickAction("classify"));
  $("summarise-btn").addEventListener("click", () => handleQuickAction("summarise"));
  $("extension-btn").addEventListener("click", () => handleQuickAction("extension"));

  $("settings-btn").addEventListener("click", () => {
    const existing = document.getElementById("settings-panel");
    if (existing) { existing.remove(); return; }
    showTokenPrompt();
  });

  const input = $("msg-input");
  input.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); }
  });
  input.addEventListener("input", () => {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight, 96) + "px";
  });
}

function handleSend() {
  const input = $("msg-input");
  const text = input.value.trim();
  if (!text) return;
  addMessage("user", text);
  input.value = "";
  input.style.height = "auto";
  showTyping();
  waitingForResponse = true;
  sendMessage(text);
}
