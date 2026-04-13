# Email AI PoC

An AI-powered email processing system that integrates with Microsoft Outlook to automatically classify incoming emails and generate draft replies using a large language model.

## Features

- **Outlook Integration** — fetches emails from your inbox via Microsoft Graph API
- **AI Classification** — classifies each email by urgency, intent, and course relevance using Groq (Llama 3.3)
- **AI Draft Generation** — generates professional reply drafts with optional custom instructions
- **Draft Management** — review and edit drafts before sending
- **Send via Outlook** — sends approved drafts directly through Microsoft Graph

## Tech Stack

| Layer | Technology |
|---|---|
| API Framework | FastAPI + Uvicorn |
| Data Validation | Pydantic v2 |
| LLM | Groq (Llama 3.3 70B) |
| Email Integration | Microsoft Graph API (MSAL) |
| Testing | pytest + pytest-asyncio |
| Linting / Formatting | flake8 + black |

## Project Structure

```
app/
├── core/           # Config (pydantic-settings) and logging
├── models/         # Pydantic models — emails, drafts, API responses
├── services/       # email_parser, classifier, llm_service, draft_service, graph_service
└── api/routes/     # health, auth, emails, drafts
tests/
├── unit/           # Parser, classifier, draft service tests
└── integration/    # Email and draft route tests
```

## Getting Started

### Prerequisites

- Python 3.11+
- A [Groq API key](https://console.groq.com) (free)
- An Azure app registration with Microsoft Graph `Mail.Read` and `Mail.Send` permissions

### Installation

```bash
# Clone the repo
git clone https://github.com/iammimii/AI-Agent-Concept-Workshop-Academic-Automation.git
cd AI-Agent-Concept-Workshop-Academic-Automation

# Create and activate virtual environment
python -m venv venv
venv\Scripts\activate       # Windows
source venv/bin/activate    # macOS/Linux

# Install dependencies
pip install -r requirements.txt -r requirements-dev.txt
```

### Configuration

Copy `.env.example` to `.env` and fill in your credentials:

```bash
cp .env.example .env
```

```env
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
AZURE_TENANT_ID=your-tenant-id
OUTLOOK_USER=your@outlook.com
GROQ_API_KEY=your-groq-api-key
```

### Running the API

```bash
uvicorn app.main:app --reload
```

Open [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs) for the interactive Swagger UI.

## Workflow

The Swagger UI guides you through the full workflow step by step:

| Step | Endpoint | Description |
|---|---|---|
| 1 | `GET /auth/login` | Authenticate with Outlook via device code flow |
| 2 | `GET /emails/inbox` | Fetch and classify inbox emails |
| 3 | `POST /drafts/generate` | Generate an AI reply for an email |
| 4a | `GET /drafts` | Review pending drafts |
| 4b | `PATCH /drafts/{id}` | Edit draft body or status |
| 5 | `POST /drafts/{id}/send` | Send draft via Outlook |

## Running Tests

```bash
pytest -v
```

All 16 tests should pass:

```
tests/integration/test_draft_routes.py   5 passed
tests/integration/test_email_routes.py   2 passed
tests/unit/test_classifier.py            1 passed
tests/unit/test_draft_service.py         5 passed
tests/unit/test_email_parser.py          3 passed
```

## Environment Variables

| Variable | Description |
|---|---|
| `AZURE_CLIENT_ID` | Azure app registration client ID |
| `AZURE_CLIENT_SECRET` | Azure app registration client secret |
| `AZURE_TENANT_ID` | Azure AD tenant ID |
| `OUTLOOK_USER` | Outlook email address |
| `GROQ_API_KEY` | Groq API key |
| `GROQ_MODEL` | Model name (default: `llama-3.3-70b-versatile`) |
| `DEBUG` | Enable debug logging (default: `false`) |
| `MAX_EMAIL_BODY_CHARS` | Max email body length sent to LLM (default: `4000`) |

## License

MIT
