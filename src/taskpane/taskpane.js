import "./taskpane.css";

// ===== Gateway Config =====
const GW_PORT = 18789;
const BACKOFF_BASE = 1200;
const BACKOFF_MAX = 20000;
const MAX_REPLY_RETRIES = 5; // FIX #2: Max retries for fetchLatestReply

// ===== State =====
let socket = null;
let isConnected = false;
let authToken = null;
let gwPort = 18789;
let sessionKey = "agent:main:academic-email";
let activeEmailId = null;
let currentEmail = null;
let contextSentForEmail = null; // FIX #3: Will store activeEmailId instead of subject
let streamBuffer = "";
let activeRunId = null;
let waitingForResponse = false;
let rpcSeq = 0;
let pendingRpc = new Map();
let historyFetching = false;
let lastShownMsgId = null;
let retryCount = 0;
let retryTimer = null;
let replyRetryCount = 0; // FIX #2: Retry counter for fetchLatestReply

// ===== DOM Helper =====
const $ = id => document.getElementById(id);

// ===== Boot =====

/**
 * Entry point called by Office.js when the add-in environment is ready.
 * Applies the theme, binds UI events, loads the auth token, reads the current
 * email, and registers an ItemChanged handler to update the panel when the
 * user selects a different email.
 */
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

/**
 * Detects whether Outlook is using a dark theme by sampling the background
 * colour from the Office theme object. Falls back to the OS-level
 * prefers-color-scheme media query if the Office theme is unavailable.
 */
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

/**
 * Loads the gateway auth token and port from localStorage. If no token is saved,
 * shows the settings panel immediately so the user can enter their credentials
 * before connecting.
 */
function loadToken() {
  try {
    authToken = localStorage.getItem("acad-gateway-token") || "";
    gwPort = parseInt(localStorage.getItem("acad-gateway-port"), 10) || 18789;
  } catch (_) {
    authToken = "";
    gwPort = 18789;
  }

  if (!authToken) {
    showTokenPrompt();
  } else {
    connectGateway();
  }
}

/**
 * Renders an inline settings panel in the chat so the user can update their
 * gateway token and custom system instructions without leaving Outlook.
 */
function showTokenPrompt() {
  const saved = getSavedSystemPrompt();
  // FIX #4: Escape token value properly to prevent innerHTML injection
  const escapedToken = (authToken || "").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  
  const div = document.createElement("div");
  div.className = "message sys-msg";
  div.id = "settings-panel";
  div.innerHTML = `
    <div style="padding:4px 0">
      <strong>Setup</strong>
    </div>
    <div style="margin-bottom:8px">
      Gateway Token
      <input type="password" id="token-field" 
             placeholder="Paste your gateway token..." 
             value="${escapedToken}"
             style="width:100%;padding:5px 8px;border:1px solid var(--border);border-radius:4px;
                    background:var(--bg-input);color:var(--text-primary);font-size:12px;
                    margin-bottom:8px;font-family:monospace">
      Gateway Port
      <input type="number" id="port-field" 
             placeholder="18789" 
             value="${gwPort || 18789}"
             style="width:100%;padding:5px 8px;border:1px solid var(--border);border-radius:4px;
                    background:var(--bg-input);color:var(--text-primary);font-size:12px;
                    margin-bottom:8px;font-family:monospace">
      Custom Instructions (optional)
      <textarea id="prompt-field" rows="3" 
                placeholder="e.g. Always reply formally. Never use bullet points."
                style="width:100%;padding:5px 8px;border:1px solid var(--border);border-radius:4px;
                       background:var(--bg-input);color:var(--text-primary);font-size:12px;
                       margin-bottom:8px;font-family:var(--font);resize:vertical">${saved}</textarea>
      <button id="save-settings-btn" 
              style="width:100%;padding:8px;background:var(--accent);color:#fff;border:none;
                     border-radius:4px;cursor:pointer;font-weight:500">Save &amp; Connect</button>
    </div>
    <div style="font-size:11px;opacity:0.6;line-height:1.4">
      Token: ~/.openclaw/openclaw.json → gateway.auth.token · Port: default 18789
    </div>
  `;

  $("chat-messages").appendChild(div);

  setTimeout(() => {
    const btn = document.getElementById("save-settings-btn");
    if (!btn) return;

    btn.addEventListener("click", () => {
      const t = document.getElementById("token-field").value.trim();
      const port = parseInt(document.getElementById("port-field").value, 10);
      const p = document.getElementById("prompt-field").value.trim();

      if (t) {
        try { localStorage.setItem("acad-gateway-token", t); } catch (_) {}
        authToken = t;
      }
      if (port) {
        try { localStorage.setItem("acad-gateway-port", String(port)); } catch (_) {}
        gwPort = port;
      }
      try { localStorage.setItem("acad-custom-instructions", p); } catch (_) {}

      div.remove();
      addMessage("sys", "Settings saved. Connecting to gateway...");
      
      // FIX #6: Reset retryCount when manually reconnecting from settings
      retryCount = 0;
      connectGateway();
    });
  }, 80);
}

/** Returns any custom system instructions saved by the user in the settings panel. */
function getSavedSystemPrompt() {
  try {
    return localStorage.getItem("acad-custom-instructions") || "";
  } catch (_) {
    return "";
  }
}

// ===== Email Reader =====

/**
 * Reads the currently selected email using the Office.js Mailbox API.
 * Generates a unique session key per email so each conversation is stored
 * separately in OpenClaw. Clears the chat and reloads history when the
 * user switches to a different email.
 */
function readEmail() {
  const item = Office.context.mailbox.item;
  if (!item) {
    showNoEmail();
    return;
  }

  try {
    const subject = item.subject || "(No subject)";
    const from = item.from
      ? `${item.from.displayName} <${item.from.emailAddress}>`
      : "Unknown";
    const date = item.dateTimeCreated
      ? new Date(item.dateTimeCreated).toLocaleString()
      : "";
    const to = item.to
      ? item.to.map(r => `${r.displayName} <${r.emailAddress}>`).join(", ")
      : "";

    item.body.getAsync(Office.CoercionType.Text, result => {
      const body =
        result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";

      currentEmail = { subject, from, to, date, body };

      const newId = hashString(subject + "|" + from + "|" + date);

      if (newId !== activeEmailId) {
        activeEmailId = newId;
        sessionKey = `agent:main:academic-email-${newId}`;
        contextSentForEmail = null; // FIX #3: Will be set to activeEmailId on first send
        lastShownMsgId = null;
        replyRetryCount = 0; // FIX #2: Reset retry count for new email
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

/** Resets the email header to the placeholder state when no email is open. */
function showNoEmail() {
  $("email-placeholder").style.display = "flex";
  $("email-info").style.display = "none";
  $("label-row").style.display = "none";
  currentEmail = null;
}

/**
 * Populates the email header with the subject, sender, and date, then
 * refreshes the label buttons to reflect existing Outlook categories.
 */
function renderEmailHeader(subject, from, date) {
  $("email-placeholder").style.display = "none";
  $("email-info").style.display = "block";
  $("email-subject").textContent = subject;
  $("email-from").textContent = from;
  $("email-date").textContent = date;
  loadCategories();
}

// ===== Categories =====

/**
 * Reads the Outlook categories on the open email and highlights the
 * matching label buttons. Hides the label row on older Outlook versions
 * that do not support the categories API.
 */
function loadCategories() {
  const item = Office.context.mailbox.item;
  if (!item || !item.categories) {
    $("label-row").style.display = "none";
    return;
  }

  $("label-row").style.display = "flex";

  item.categories.getAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const active = (result.value || []).map(c => (c.displayName || c).toLowerCase());

    document.querySelectorAll(".btn-label").forEach(btn => {
      const cat = btn.dataset.category.toLowerCase();
      btn.classList.toggle("active", active.includes(cat));
    });
  });
}

/**
 * Ensures a category exists in the Outlook master list before applying it to
 * an item. Office.js throws "Invalid categories" if addAsync is called with a
 * name that hasn't been registered in masterCategories first.
 *
 * @param {string} name - Category display name (e.g. "Urgent")
 * @param {Function} callback - Called once the category is confirmed to exist
 */
function ensureMasterCategory(name, callback) {
  const colorMap = {
    "Urgent": Office.MailboxEnums.CategoryColor.Preset0,
    "Medium": Office.MailboxEnums.CategoryColor.Preset3,
    "Minor": Office.MailboxEnums.CategoryColor.Preset7,
  };

  Office.context.mailbox.masterCategories.getAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback();
      return;
    }

    const exists = (result.value || []).some(
      c => c.displayName.toLowerCase() === name.toLowerCase()
    );
    if (exists) {
      callback();
      return;
    }

    const color = colorMap[name] || Office.MailboxEnums.CategoryColor.Preset0;
    Office.context.mailbox.masterCategories.addAsync(
      [{ displayName: name, color }],
      () => callback()
    );
  });
}

/**
 * Toggles an Outlook category on the current email. Removes it if already
 * applied, or adds it (creating it in the master list first if needed).
 *
 * @param {string} name - Category name matching one of the label buttons
 */
function toggleCategory(name) {
  const item = Office.context.mailbox.item;
  if (!item) {
    addMessage("err", "No email selected.");
    return;
  }

  if (!item.categories) {
    addMessage("err", "Categories not supported in this Outlook version.");
    return;
  }

  item.categories.getAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      addMessage("err", "Could not read categories: " + (result.error?.message || "unknown error"));
      return;
    }

    const active = (result.value || []).map(c => (c.displayName || c).toLowerCase());

    if (active.includes(name.toLowerCase())) {
      item.categories.removeAsync([name], () => loadCategories());
    } else {
      ensureMasterCategory(name, () => {
        item.categories.addAsync([name], r => {
          if (r.status !== Office.AsyncResultStatus.Succeeded)
            addMessage("err", "Add failed: " + (r.error?.message || "unknown error"));
          loadCategories();
        });
      });
    }
  });
}

// ===== Gateway Connection =====

/** Returns the WebSocket URL for the OpenClaw Gateway. */
function getGatewayUrl() {
  return `ws://localhost:${gwPort}?token=${encodeURIComponent(authToken)}`;
}

/**
 * Opens a WebSocket connection to the OpenClaw Gateway (proxied by
 * webpack-dev-server from /ai-gateway to port 18789). Implements
 * exponential backoff on disconnect.
 */
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

  // FIX #1: Send handshake immediately on connection without arbitrary delay
  socket.onopen = () => {
    setStatus("connecting");
    sendHandshake();
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

/** Schedules a reconnection attempt using exponential backoff. */
function scheduleRetry() {
  if (retryTimer) return;
  const delay = Math.min(BACKOFF_BASE * Math.pow(1.7, retryCount), BACKOFF_MAX);
  retryCount++;
  retryTimer = setTimeout(() => {
    retryTimer = null;
    connectGateway();
  }, delay);
}

// ===== Handshake =====

/**
 * Sends the OpenClaw connect RPC to authenticate and negotiate protocol version.
 * On success, marks the connection as ready and loads chat history for the
 * current email.
 */
function sendHandshake() {
  const params = {
    minProtocol: 3,
    maxProtocol: 3,
    client: {
      id: "openclaw-control-ui",
      version: "1.0.0",
      platform: navigator.platform || "web",
      mode: "webchat",
      instanceId: "acad-" + Date.now(),
    },
    role: "operator",
    scopes: ["operator.admin"],
    caps: ["tool-events"],
    auth: { token: authToken },
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

/**
 * Sends a JSON-RPC request over the WebSocket and returns a Promise that
 * resolves with the result or rejects on error.
 *
 * @param {string} method - RPC method name (e.g. "chat.send")
 * @param {object} params - Method parameters
 * @returns {Promise}
 */
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

/**
 * Routes incoming WebSocket messages to the RPC response handler or the
 * event handler depending on the message type.
 *
 * @param {string} raw - Raw JSON string received from the WebSocket
 */
function handleIncoming(raw) {
  let data;
  try {
    data = JSON.parse(raw);
  } catch (_) {
    return;
  }

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

/**
 * Returns true if the text looks like a raw tool-call JSON object emitted by
 * the agent. These are internal operations and should not be shown to the user.
 *
 * @param {string} text
 * @returns {boolean}
 */
function isRawToolCall(text) {
  const t = text.trim();
  try {
    const obj = JSON.parse(t);
    if (obj && typeof obj.name === "string" && obj.parameters !== undefined) return true;
  } catch (_) {}
  return /^\{"name"\s*:\s*"[^"]+"\s*,\s*"parameters"\s*:/.test(t);
}

/**
 * Flushes the accumulated stream buffer to the chat as a finalised AI message
 * and clears the buffer.
 */
function flushStream() {
  const text = streamBuffer.trim();
  if (text && !isRawToolCall(text)) addMessage("ai", processAIText(text));
  streamBuffer = "";
}

/**
 * Handles server-sent events from the OpenClaw Gateway.
 * Manages the typing indicator lifecycle, assembles streamed response chunks,
 * and displays finalised messages from the AI.
 *
 * @param {object} evt - Parsed event object from the gateway
 */
function handleEvent(evt) {
  const event = evt.event || "";
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
          if (waitingForResponse) {
            waitingForResponse = false;
            fetchLatestReply();
          }
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
        if (!isRawToolCall(text)) addMessage("ai", processAIText(text));
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

/**
 * Fetches the stored conversation history for the current email session from
 * OpenClaw and renders it into the chat. User messages are stripped of the
 * email context prefix that was prepended before sending.
 */
function loadHistory() {
  callRpc("chat.history", { sessionKey, limit: 50 })
    .then(result => {
      const msgs = extractMessages(result);
      // FIX #5: Track the last message ID even if history ends on user message
      let latestId = null;
      
      for (const msg of msgs) {
        const text = extractText(msg);
        if (!text) continue;

        if (msg.role === "user") {
          const m = text.match(/User question:\s*([\s\S]*)/);
          addMessage("user", m ? m[1].trim() : text.trim());
        } else if (msg.role === "assistant") {
          addMessage("ai", processAIText(text.trim()));
          latestId = msg.__openclaw?.id || msg.responseId || msg.timestamp || null;
        }
      }
      
      // FIX #5: Set lastShownMsgId even if last message was user message
      if (latestId) {
        lastShownMsgId = latestId;
      }
    })
    .catch(() => {});
}

/**
 * Polls the chat history for the latest assistant reply after a message is sent.
 * Used as a fallback when the streamed response was not captured via delta events.
 * Retries after 2 seconds if the newest message ID hasn't changed.
 * FIX #2: Now has a maximum retry limit to prevent infinite polling.
 */
function fetchLatestReply() {
  if (historyFetching) return;
  
  // FIX #2: Stop after MAX_REPLY_RETRIES attempts
  if (replyRetryCount >= MAX_REPLY_RETRIES) {
    historyFetching = false;
    replyRetryCount = 0;
    hideTyping();
    addMessage("sys", "No response received from AI. Try sending again.");
    return;
  }

  historyFetching = true;
  replyRetryCount++;

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
          replyRetryCount = 0; // FIX #2: Reset counter on success
          lastShownMsgId = msgId;
          addMessage("ai", processAIText(text.trim()));
          hideTyping();
          return;
        } else if (msgId === lastShownMsgId) {
          historyFetching = false;
          setTimeout(fetchLatestReply, 2000);
          return;
        }
      }

      historyFetching = false;
      setTimeout(fetchLatestReply, 2000);
    })
    .catch(() => {
      historyFetching = false;
      setTimeout(fetchLatestReply, 2000);
    });
}

/**
 * Normalises the various response shapes that chat.history can return into a
 * flat array of message objects.
 *
 * @param {any} result - Raw RPC result
 * @returns {Array}
 */
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

/**
 * Extracts plain text from a message object, handling both string content
 * and the array-of-blocks format used by some models.
 *
 * @param {object} msg
 * @returns {string}
 */
function extractText(msg) {
  if (typeof msg.content === "string") return msg.content;
  if (Array.isArray(msg.content)) {
    return msg.content
      .filter(c => c.type === "text")
      .map(c => c.text || "")
      .join("\n");
  }
  return "";
}

// ===== Send Message =====

/**
 * Sends a user message to OpenClaw via the chat.send RPC.
 * Prepends the full email context on the first message of each conversation
 * so the AI knows what it is analysing. An idempotency key prevents duplicate
 * messages if the request is retried.
 *
 * @param {string} text - The message text to send
 */
function sendMessage(text) {
  if (!isConnected) {
    addMessage("err", "Not connected to gateway. Reconnecting…");
    connectGateway();
    return;
  }

  let fullText = text;

  // FIX #3: Compare contextSentForEmail against activeEmailId instead of subject
  if (currentEmail && contextSentForEmail !== activeEmailId) {
    const body = (currentEmail.body || "").slice(0, 3000);
    const custom = getSavedSystemPrompt();

    let prefix = "";
    if (custom) prefix += `[System instructions]\n${custom}\n\n`;

    prefix += `[Instructions: You are an email assistant. Answer the user's question directly. Do not repeat or quote the email body in your response. Do not create memory files, do not log email content, do not ask for confirmation, do not show next steps. When asked to assign or suggest a label/priority for this email, read the email and decide: Urgent = needs immediate action or is time-sensitive, Medium = needs attention but not critical, Minor = low priority or informational. Include exactly one of [LABEL:Urgent], [LABEL:Medium], or [LABEL:Minor] at the very end of your response.]\n\n`;

    prefix += `[Current email context]\nSubject: ${currentEmail.subject}\nFrom: ${currentEmail.from}\nTo: ${currentEmail.to}\nDate: ${currentEmail.date}\n\nBody:\n${body}\n\n---\n\n`;

    fullText = prefix + `User question: ${text}`;
    
    // FIX #3: Store activeEmailId instead of subject
    contextSentForEmail = activeEmailId;
  }

  callRpc("chat.send", {
    sessionKey,
    message: fullText,
    deliver: false,
    idempotencyKey: crypto.randomUUID(),
  })
    .then(() => {
      waitingForResponse = true;
      replyRetryCount = 0; // FIX #2: Reset retry counter when sending new message
      showTyping();
    })
    .catch(err => {
      hideTyping();
      addMessage("err", "Send failed: " + err.message);
    });
}

// ===== Actions =====

/** Asks the AI to draft a professional reply to the current email. */
function draftReply() {
  if (!currentEmail) {
    addMessage("err", "No email selected.");
    return;
  }
  addMessage("user", "Draft a reply to this email");
  showTyping();
  waitingForResponse = true;
  sendMessage(
    "Please draft a professional reply to this email. Respond in the same language as the original email."
  );
}

/**
 * Asks the AI to read the email and assign a priority label (Urgent / Medium / Minor).
 * The AI appends a [LABEL:X] tag which processAIText converts into an Outlook
 * category automatically.
 */
function autoLabel() {
  if (!currentEmail) {
    addMessage("err", "No email selected.");
    return;
  }
  addMessage("user", "Assign a priority label to this email");
  showTyping();
  waitingForResponse = true;
  sendMessage(
    "Read this email and assign a priority label based on its urgency. Reply with a brief reason for your choice."
  );
}

/**
 * Opens the last AI message as a reply draft in Outlook's native compose
 * window using the Office.js displayReplyForm API.
 */
function useDraft() {
  const aiMsgs = document.querySelectorAll(".ai-msg .msg-body");
  const last = aiMsgs[aiMsgs.length - 1];
  if (!last) {
    addMessage("err", "No draft available. Click Draft Reply first.");
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) {
    addMessage("err", "No email selected.");
    return;
  }

  try {
    item.displayReplyForm(last.textContent);
    addMessage("sys", "Draft opened in Outlook. Review and send when ready.");
  } catch (err) {
    addMessage("err", "Could not open draft: " + err.message);
  }
}

// ===== AI Label Action Parser =====

/**
 * Post-processes AI response text before displaying it:
 * strips any echoed email context block, removes trailing separators,
 * and detects a [LABEL:X] tag to automatically apply the Outlook category.
 *
 * @param {string} text - Raw AI response text
 * @returns {string} Cleaned text safe to display
 */
function processAIText(text) {
  let cleaned = text
    .replace(/\[Current email context\][\s\S]*?---\s*/g, "")
    .trim();
  cleaned = cleaned.replace(/\n---+\s*$/g, "").trim();

  const match = cleaned.match(/\[LABEL:(Urgent|Medium|Minor)\]/i);
  if (match) {
    const label =
      match[1].charAt(0).toUpperCase() + match[1].slice(1).toLowerCase();
    toggleCategory(label);
    cleaned = cleaned.replace(/\s*\[LABEL:[^\]]+\]/gi, "").trim();
  }

  return cleaned;
}

// ===== Chat UI =====

/**
 * Appends a message bubble to the chat area.
 * Roles: "user" (right-aligned), "ai" (left-aligned), "sys" (centred info),
 * "err" (centred error). Enables the Use Draft button on any AI message.
 *
 * @param {"user"|"ai"|"sys"|"err"} role
 * @param {string} text
 */
function addMessage(role, text) {
  const existing = document.querySelector(".streaming-bubble");
  if (existing) existing.remove();

  const container = $("chat-messages");
  const div = document.createElement("div");
  const classMap = {
    user: "message user-msg",
    ai: "message ai-msg",
    sys: "message sys-msg",
    err: "message err-msg"
  };
  div.className = classMap[role] || "message sys-msg";

  const body = document.createElement("div");
  body.className = "msg-body";
  body.textContent = text;
  div.appendChild(body);
  container.appendChild(div);

  scrollDown();
  if (role === "ai") $("use-draft-btn").disabled = false;
}

/**
 * Updates the live streaming bubble with the latest accumulated text as chunks
 * arrive. Creates the bubble on first call, replaces its content on subsequent
 * calls so the response appears to type out in real time.
 *
 * @param {string} text - Full accumulated stream text so far
 */
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

/** Removes all non-system messages from the chat and disables the Use Draft button. */
function clearChatMessages() {
  const container = $("chat-messages");
  container.querySelectorAll(".message:not(.sys-msg)").forEach(m => m.remove());
  $("use-draft-btn").disabled = true;
}

function showTyping() {
  $("typing-indicator").style.display = "flex";
  scrollDown();
}

function hideTyping() {
  $("typing-indicator").style.display = "none";
}

function setTypingLabel(text) {
  const l = $("typing-label");
  if (l) l.textContent = text;
  showTyping();
}

function scrollDown() {
  const c = $("chat-messages");
  requestAnimationFrame(() => {
    c.scrollTop = c.scrollHeight;
  });
}

// ===== Status Bar =====

/**
 * Updates the status bar colour and label to reflect the connection state.
 *
 * @param {"connected"|"connecting"|"disconnected"} state
 */
function setStatus(state) {
  const bar = $("status-bar");
  bar.className = "status-bar " + state;
  const labels = {
    connected: "Academic AI Ready",
    connecting: "Connecting…",
    disconnected: "Disconnected"
  };
  $("status-text").textContent = labels[state] || state;
}

// ===== Utilities =====

/**
 * Produces a short alphanumeric hash of a string using a djb2-style algorithm.
 * Used to generate stable per-email session keys.
 *
 * @param {string} str
 * @returns {string} Base-36 encoded absolute hash value
 */
function hashString(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = ((h << 5) - h + str.charCodeAt(i)) | 0;
  return Math.abs(h).toString(36);
}

// ===== Event Bindings =====

/** Wires up all button click handlers and textarea keyboard shortcuts. */
function bindEvents() {
  $("send-btn").addEventListener("click", handleSend);
  $("draft-btn").addEventListener("click", draftReply);
  $("use-draft-btn").addEventListener("click", useDraft);
  $("auto-label-btn").addEventListener("click", autoLabel);

  document.querySelectorAll(".btn-label").forEach(btn => {
    btn.addEventListener("click", () => toggleCategory(btn.dataset.category));
  });

  $("settings-btn").addEventListener("click", () => {
    const existing = document.getElementById("settings-panel");
    if (existing) {
      existing.remove();
      return;
    }
    showTokenPrompt();
  });

  const input = $("msg-input");
  input.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });

  input.addEventListener("input", () => {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight, 96) + "px";
  });
}

/** Reads the input field, renders a user bubble, and sends the message. */
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
