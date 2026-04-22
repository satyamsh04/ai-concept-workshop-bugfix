# Academic Email Assistant POC

An Outlook add-in that brings a locally-hosted AI agent directly into your email sidebar. Built on the OpenClaw Gateway WebSocket protocol with Ollama for fully local inference — no data leaves your machine.

## Features

- **Email Context** — automatically reads subject, sender, recipients, date, and body of the selected email
- **Chat Interface** — ask questions, get summaries, translations, or any AI assistance
- **Draft Reply** — one-click reply drafting based on the email context
- **Send Reply** — opens Outlook's native reply compose with the drafted text pre-filled
- **Pinned Sidebar** — stays open when switching between emails (pin icon in taskpane header)
- **Per-Email Chat History** — each email gets its own session; switching back restores the conversation
- **Smart Context** — email body sent only with the first message per email (saves tokens)
- **Auto-Reconnect** — WebSocket reconnects automatically with exponential backoff
- **Light/Dark Mode** — auto-detects Outlook theme
- **Local AI only** — all inference runs on-device via Ollama; data never leaves your infrastructure

## Tech Stack

| Layer | Technology |
|---|---|
| Add-in Framework | Office.js (Outlook taskpane) |
| Build Tooling | Webpack 5 + webpack-dev-server |
| AI Gateway | OpenClaw Gateway (local, WebSocket RPC) |
| LLM | Ollama — llama3.1:8b (self-hosted) |
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
Ollama (llama3.1:8b — http://127.0.0.1:11434)
```

## Prerequisites

- **Node.js 18+**
- **OpenClaw** — `npm install -g openclaw`
- **Ollama** — [ollama.com](https://ollama.com) with `llama3.1:8b` pulled
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
ollama pull llama3.1:8b
```

### 4. Configure OpenClaw

Edit `~/.openclaw/openclaw.json` and add the following under `gateway.controlUi`:

```json
"controlUi": {
  "dangerouslyDisableDeviceAuth": true
}
```

### 5. Start OpenClaw Gateway

```bash
openclaw gateway
```

Note the gateway token from `~/.openclaw/openclaw.json` → `gateway.auth.token`.

### 6. Start the dev server

```bash
npm start
```

### 7. Sideload the add-in

In Outlook, go to **Get Add-ins → My Add-ins → Add a custom add-in → Add from file** and select `manifest.xml`.

## Running

With both the gateway and dev server running, open any email in Outlook and click **Academic Assistant** in the ribbon. The add-in opens in a sidebar and automatically reads the selected email.

- Type a question and press Enter or click Send
- Click **Draft Reply** to generate a professional reply
- Click **Use Draft** to open Outlook's reply compose with the draft pre-filled
- Click the 📌 pin icon to keep the sidebar open when switching emails

## Gateway Token

The gateway token is read from `~/.openclaw/openclaw.json` → `gateway.auth.token`. It is stored in browser `localStorage` under the key `acad-gateway-token`. To update it, click the ⚙ settings icon in the add-in sidebar.

## License

MIT
