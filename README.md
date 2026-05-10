# Academic Email Assistant

An Outlook add-in that brings a locally-hosted AI into your email sidebar. Ask questions about emails, draft replies, and auto-assign priority labels — all running on your own machine with no data sent to the cloud.

## What it does

- **Chat** — Ask the AI anything about the open email (summarise, translate, explain)
- **Draft Reply** — One click to generate a professional reply
- **Auto Label** — AI reads the email and assigns Urgent / Medium / Minor automatically
- **Manual Labels** — Apply priority labels yourself with one click
- **Per-email history** — Each email keeps its own conversation when you switch back
- **Dark/light mode** — Follows your Outlook theme automatically

---

## Before you start — install these three things

| Tool | What it does | Download |
|---|---|---|
| **Node.js 18+** | Runs the dev server | https://nodejs.org |
| **Ollama** | Runs the AI model locally | https://ollama.com |
| **Outlook Desktop** | Where the add-in lives | Included with Microsoft 365 |

> After installing Ollama, make sure it is **running** before continuing (it usually starts automatically).

---

## Setup (run once)

### 1. Clone or download this repo

```
git clone https://github.com/iammimii/AI-Agent-Concept-Workshop-Academic-Automation.git
cd AI-Agent-Concept-Workshop-Academic-Automation
```

### 2. Run the setup script

Open **PowerShell** in the project folder and run:

```powershell
npm run setup
```

This will automatically:
- Install all npm dependencies
- Install the HTTPS certificate Outlook requires
- Download the AI model (~2 GB, takes a few minutes)
- Create the OpenClaw config file
- Set up the AI workspace

> If PowerShell asks about execution policy, type `Y` and press Enter.

---

## Running the add-in

You need **two terminals open at the same time** every time you use the add-in.

### Terminal 1 — Start the AI gateway

```powershell
npm run gateway
```

Leave this running. After it starts, open this file to find your token:

```
C:\Users\<YourName>\.openclaw\openclaw.json
```

Look for the line: `"token": "abc123..."` — copy that value.

### Terminal 2 — Start the dev server

```powershell
npm start
```

Leave this running too.

---

## Sideload the add-in into Outlook (one time only)

1. Open **Outlook Desktop**
2. Click **Get Add-ins** in the ribbon
3. Go to **My Add-ins → Add a custom add-in → Add from file**
4. Select the `manifest.xml` file from this project folder
5. Click **OK**

You should now see **Academic Assistant** in your Outlook ribbon.

---

## First-time token setup

1. Open any email in Outlook
2. Click **Academic Assistant** in the ribbon — a settings panel will appear automatically
3. Paste the token you copied from `openclaw.json`
4. Click **Save & Connect**
5. The status bar at the top should turn **green** — you're connected!

> You only need to do this once. The token is saved in Outlook's local storage.

---

## Using the add-in

| What to do | How |
|---|---|
| Ask a question | Type in the chat box and press Enter |
| Summarise the email | Type "summarise this email" |
| Draft a reply | Click **Draft Reply** |
| Use the drafted reply | Click **Use Draft** — it opens Outlook's compose window |
| Auto-assign a label | Click **Auto Label** — AI decides Urgent/Medium/Minor |
| Manually label | Click **Urgent**, **Medium**, or **Minor** in the label row |
| Keep sidebar open | Click the 📌 pin icon in the top right of the sidebar |

---

## Troubleshooting

**Status bar says "Disconnected"**
- Make sure `npm run gateway` is running in a terminal
- Check that the token in the settings panel matches `~/.openclaw/openclaw.json`

**"Cannot connect" or "Not connected" error**
- Both `npm run gateway` (Terminal 1) and `npm start` (Terminal 2) must be running
- Try refreshing the add-in by closing and reopening the sidebar

**Add-in doesn't appear in the ribbon**
- Re-sideload: Get Add-ins → My Add-ins → find Academic Assistant → Remove → re-add from file

**AI replies with something unexpected**
- The model may be slow on first message — give it 10–20 seconds

**Setup script fails on certificates**
- Run PowerShell as Administrator and retry: `npm run setup`

---

## Tech stack

| Layer | Technology |
|---|---|
| Add-in UI | HTML, CSS, vanilla JavaScript |
| Office integration | Office.js (Microsoft SDK) |
| AI model | qwen2.5:3b via Ollama (local, ~2 GB) |
| AI session management | OpenClaw Gateway (WebSocket) |
| Build / dev server | Webpack 5 + webpack-dev-server |

## Architecture

```
Outlook (ribbon + taskpane)
        │
        │  wss://localhost:3000/ai-gateway
        ▼
Webpack Dev Server  (:3000)
        │
        │  ws://127.0.0.1:18789
        ▼
OpenClaw Gateway  (local)
        │
        ▼
Ollama — qwen2.5:3b  (:11434)
```
