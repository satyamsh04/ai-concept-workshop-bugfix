# Academic Email Assistant POC

An Outlook add-in that brings a locally-hosted AI agent directly into your email sidebar. Built on the OpenClaw Gateway WebSocket protocol with Ollama for fully local inference — no data leaves your machine.

## Features

- 📧 **Email Context** — Automatically reads subject, sender, recipients, date, and body of the selected email
- 💬 **Chat Interface** — Ask questions about the email, get summaries, translations, or any AI assistance
- ✏️ **Draft Reply** — One-click reply drafting based on the email context
- 📤 **Send Reply** — Opens Outlook's native reply compose with the drafted text pre-filled
- 🌗 **Light/Dark Mode** — Auto-detects Outlook theme (Office.js + prefers-color-scheme)
- 📌 **Pinned Sidebar** — Stays open when switching between emails (VersionOverrides v1.1)
- 💾 **Per-Email Chat History** — Each email gets its own session; switching back restores the conversation
- 📎 **Smart Context** — Email body sent only with the first message per email (saves tokens)
- 🔄 **Auto-Reconnect** — WebSocket reconnects automatically with exponential backoff
- 🔒 **Token-based Auth** — Gateway token stored in browser localStorage, never in code

## Tech Stack

| Layer | Technology |
|---|---|
| Add-in Framework | Office.js (Outlook taskpane) |
| Build Tooling | Webpack 5 + webpack-dev-server |
| AI Gateway | OpenClaw Gateway (local, WebSocket RPC) |
| LLM | Ollama — qwen2.5:3b (self-hosted, ~2GB) |
| Protocol | OpenClaw Gateway Protocol v3 (WebSocket) |
| Auth | Token-based (stored in browser localStorage) |

## Architecture

```
Outlook Add-in (HTTPS :3000)
        │  wss://localhost:3000/ai-gateway
        ▼
Webpack Dev Server proxy
        │  ws://127.0.0.1:18789
        ▼
OpenClaw Gateway (local)
        │
        ▼
Ollama (qwen2.5:3b — http://127.0.0.1:11434)
```

## Prerequisites

- **Node.js 18+**
- **OpenClaw** — `npm install -g openclaw`
- **Ollama** — [ollama.com](https://ollama.com)
- **Outlook Desktop** (Windows, Classic) or Outlook Web (OWA)
- **Microsoft 365** account with sideloading enabled

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Install dev certificates (first time only)

```bash
npx office-addin-dev-certs install
```

### 3. Pull the AI model

```bash
ollama pull qwen2.5:3b
```

> `qwen2.5:3b` is recommended — fast on CPU, ~2GB, runs on most laptops. You can swap it for `llama3.1:8b` for higher quality if your machine supports it.

### 4. Configure OpenClaw

Edit `~/.openclaw/openclaw.json` and ensure it contains the following (create the file if it doesn't exist):

```json
{
  "agents": {
    "defaults": {
      "model": "ollama/qwen2.5:3b"
    }
  },
  "gateway": {
    "mode": "local",
    "auth": {
      "mode": "token"
    },
    "controlUi": {
      "allowedOrigins": ["https://localhost:3000"],
      "dangerouslyDisableDeviceAuth": true
    }
  },
  "plugins": {
    "entries": {
      "ollama": { "enabled": true }
    }
  }
}
```

### 5. Initialise the OpenClaw workspace

The OpenClaw agent uses a workspace folder at `~/.openclaw/workspace/` to store its identity and memory. On first run it will run a bootstrap workflow if this folder isn't set up — to skip that, delete `BOOTSTRAP.md` from the workspace after the gateway first starts:

```bash
# Windows
del "%USERPROFILE%\.openclaw\workspace\BOOTSTRAP.md"

# macOS/Linux
rm ~/.openclaw/workspace/BOOTSTRAP.md
```

Also add the following line to `~/.openclaw/workspace/AGENTS.md` at the top of the `## Session Startup` section:

```
Do not create or check for daily memory files automatically on session startup. Only create or update memory files when the user explicitly asks.
```

### 6. Start OpenClaw Gateway

```bash
npm run gateway
```

Note the gateway token from `~/.openclaw/openclaw.json` → `gateway.auth.token`.

### 7. Start the dev server

```bash
npm start
```

### 8. Sideload the add-in

In Outlook, go to **Get Add-ins → My Add-ins → Add a custom add-in → Add from file** and select `manifest.xml`.

### 9. Enter your gateway token

Open any email and click **Academic Assistant** in the ribbon. Click the **⚙ settings icon** in the sidebar and paste your token from `~/.openclaw/openclaw.json` → `gateway.auth.token`. Click Save — the status bar should turn green (Connected).

> Each OpenClaw installation generates a unique token. The token is saved in browser localStorage so you only need to enter it once.

## Running

With both the gateway and dev server running, open any email in Outlook and click **Academic Assistant** in the ribbon. The add-in opens in a sidebar and automatically reads the selected email.

- Type a question and press Enter or click Send
- Click **Draft Reply** to generate a professional reply
- Click **Use Draft** to open Outlook's reply compose with the draft pre-filled
- Click the 📌 pin icon to keep the sidebar open when switching emails

## Switching Models

To use a different Ollama model, update two places:

1. `~/.openclaw/openclaw.json` → `agents.defaults.model`
2. `~/.openclaw/agents/main/agent/models.json` → add the model under the `ollama.models` array

Then restart the gateway. Recommended models:

| Model | Size | Notes |
|---|---|---|
| `qwen2.5:3b` | ~2GB | Default — fast, good instruction following |
| `llama3.1:8b` | ~5GB | Higher quality, needs more RAM |
| `llama3.2:3b` | ~2GB | Fast alternative |
| `gemma3:4b` | ~3GB | Google, multimodal |

## Gateway Token

The gateway token is read from `~/.openclaw/openclaw.json` → `gateway.auth.token`. It is stored in browser `localStorage` under the key `acad-gateway-token`. To update it, click the ⚙ settings icon in the add-in sidebar.

## License

MIT
