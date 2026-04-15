# System Logic Documentation

## AI-Powered Email Response API — Proof of Concept

**Purpose**: A document of system logic to support workshop delivery and architectural understanding.

> **Companion Document**: For quick-start setup, installation, and workflow reference, see the main [README.md](../README.md). This document provides deep technical detail that complements the README's onboarding focus.

---

## Table of Contents

1. [System Architecture Overview](#1-system-architecture-overview)
2. [Data Flow Documentation](#2-data-flow-documentation)
3. [Service-by-Service Breakdown](#3-service-by-service-breakdown)
4. [API Endpoint Reference](#4-api-endpoint-reference)
5. [Authentication Flow](#5-authentication-flow)
6. [Workshop-Friendly Summary](#6-workshop-friendly-summary)

---

## 1. System Architecture Overview

### 1.1 High-Level Architecture

The AI-Powered Email Response API is a FastAPI application that acts as an intelligent intermediary between Microsoft Outlook and a large language model (LLM). The system processes incoming emails, classifies them by urgency and intent, generates suggested replies, and allows users to review and send approved drafts through a structured REST API.

The architecture follows a layered design that separates concerns and promotes maintainability:

- **API Layer**: FastAPI routes handle HTTP requests and responses
- **Service Layer**: Business logic is encapsulated in pure Python service modules
- **Integration Layer**: External API interactions (Microsoft Graph, Groq LLM) are isolated in dedicated service modules
- **Data Layer**: Pydantic models enforce schema validation and serialisation

This separation means that each component can be developed, tested, and modified independently without affecting other parts of the system. For example, the classification logic can be improved by modifying `classifier.py` and `llm_service.py` without touching any API routes.

### 1.2 Component Diagram

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                              Client Application                             │
└─────────────────────────────────────────────────────────────────────────────┘
                                        │
                                        ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                          FastAPI Application (app/main.py)                  │
│                                                                             │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐  ┌─────────┐                         │
│  │  health │  │   auth  │  │  emails │  │ drafts  │                         │
│  │  router │  │  router │  │  router │  │  router │                         │
│  └────┬────┘  └────┬────┘  └────┬────┘  └────┬────┘                         │
└───────┼────────────┼────────────┼────────────┼──────────────────────────────┘
        │            │            │            │
        ▼            ▼            ▼            ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                              Service Layer                                  │
│                                                                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐ │
│  │email_parser  │  │  classifier  │  │  llm_service │  │   draft_service  │ │
│  │              │  │              │  │              │  │                  │ │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  └────────┬─────────┘ │
└─────────┼─────────────────┼─────────────────┼───────────────────┼───────────┘
          │                 │                 │                   │
          ▼                 ▼                 ▼                   ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                           Integration Layer                                 │
│                                                                             │
│              ┌─────────────────────┐       ┌─────────────────────┐          │
│              │    graph_service    │       │    llm_service      │          │
│              │ (Microsoft Graph)   │       │      (Groq)         │          │
│              └──────────┬──────────┘       └──────────┬──────────┘          │
└─────────────────────────┼─────────────────────────────┼─────────────────────┘
                          │                             │
                          ▼                             ▼
              ┌─────────────────────┐       ┌─────────────────────┐
              │  Microsoft Graph    │       │      Groq API       │
              │  (Outlook Mail)     │       │   (Llama 3.3 70B)   │
              └─────────────────────┘       └─────────────────────┘
```

### 1.3 Technology Stack

| Layer             | Technology                     | Purpose                                                                     |
| ----------------- | ------------------------------ | --------------------------------------------------------------------------- |
| API Framework     | FastAPI + Uvicorn              | HTTP server, async routing, request validation, auto-generated OpenAPI docs |
| Data Validation   | Pydantic v2                    | Schema enforcement, input/output serialisation, type coercion               |
| LLM Provider      | Groq (Llama 3.3 70B Versatile) | Email classification and draft reply generation                             |
| Email Integration | Microsoft Graph API + MSAL     | Outlook authentication (device code flow), mailbox read/send                |
| Configuration     | pydantic-settings              | Environment variable management with .env file support                      |
| Testing           | pytest + pytest-asyncio        | Async unit and integration tests                                            |
| Code Quality      | flake8 + black                 | Linting and auto-formatting                                                 |

### 1.4 Project Structure

```
AI-Agent-Concept-Workshop-Academic-Automation/
├── app/
│   ├── __init__.py
│   ├── main.py                 # FastAPI app initialization, router registration
│   ├── core/
│   │   ├── __init__.py
│   │   ├── config.py           # Settings (env vars, API keys, defaults)
│   │   └── logging.py          # Structured logging setup
│   ├── models/
│   │   ├── __init__.py
│   │   ├── email.py            # EmailAddress, ParsedEmail, ClassifiedEmail
│   │   ├── draft.py            # Draft, DraftCreate, DraftUpdate, DraftStatus
│   │   └── responses.py        # APIResponse[T], HealthResponse
│   ├── services/
│   │   ├── __init__.py
│   │   ├── email_parser.py     # Graph message → ParsedEmail conversion
│   │   ├── classifier.py       # Orchestrates LLM classification
│   │   ├── llm_service.py      # Groq API calls (classify + draft generation)
│   │   ├── draft_service.py    # In-memory draft CRUD operations
│   │   └── graph_service.py    # Microsoft Graph API (auth, inbox, send)
│   └── api/routes/
│       ├── __init__.py
│       ├── health.py           # GET /health
│       ├── auth.py             # GET /auth/login
│       ├── emails.py           # POST /emails/process, GET /emails/inbox
│       └── drafts.py           # Draft CRUD + send endpoints
├── docs/                       # This folder
│   └── SYSTEM_LOGIC.md         # This document
├── tests/
│   ├── unit/                   # Unit tests for services
│   │   ├── test_email_parser.py
│   │   ├── test_classifier.py
│   │   └── test_draft_service.py
│   └── integration/            # API endpoint tests
│       ├── test_email_routes.py
│       └── test_draft_routes.py
├── .env.example                # Template for environment variables
├── requirements.txt            # Production dependencies
├── requirements-dev.txt        # Development dependencies (testing, linting)
├── pytest.ini                  # Pytest configuration
└── README.md                   # Project overview and quick-start guide
```

---

## 2. Data Flow Documentation

### 2.1 Email Classification Pipeline

The classification pipeline transforms a raw email from Outlook into a structured classification containing urgency level, intent category, course relevance score, confidence score, and a plain-English summary. This pipeline is invoked when users call `GET /emails/inbox`.

```
[Microsoft Graph]
      │
      ▼
┌─────────────────────────────────────────────────────────────┐
│ 1. Raw Email (dict)                                         │
│    {                                                        │
│      "id": "message-id",                                    │
│      "subject": "Question about assignment",                │
│      "sender": {"emailAddress": {"name": "Student",         │
│                  "address": "student@university.edu"}},     │
│      "toRecipients": [...],                                 │
│      "body": {"contentType": "html", "content": "<p>..."},  │
│      "receivedDateTime": "2026-04-15T10:00:00Z",            │
│      "conversationId": "thread-id"                          │
│    }                                                        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. parse_graph_message() [email_parser.py]                  │
│    - Extracts sender, recipients, subject, body             │
│    - Strips HTML if contentType == "html"                   │
│    - Unescapes HTML entities                                │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. ParsedEmail (Pydantic model)                             │
│    message_id: str                                          │
│    subject: str                                             │
│    sender: EmailAddress(name, address)                      │
│    recipients: list[EmailAddress]                           │
│    body_text: str (plain text, max 4000 chars sent to LLM)  │
│    body_html: Optional[str]                                 │
│    received_at: Optional[datetime]                          │
│    thread_id: Optional[str]                                 │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. classify_email() [classifier.py]                         │
│    - Passes ParsedEmail to llm_classify()                   │
│    - Maps LLM response to ClassifiedEmail fields            │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. llm_classify() [llm_service.py]                          │
│    - Truncates body to max_email_body_chars (default 4000)  │
│    - Constructs prompt: "Subject: ...\n\nBody: ..."         │
│    - Sends to Groq with _CLASSIFY_SYSTEM prompt             │
│    - Parses JSON response                                   │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 6. Classification Result (dict)                             │
│    {                                                        │
│      "urgency": "medium",          # high | medium | low    │
│      "intent": "inquiry",          # inquiry | complaint |  │
│                                    # feedback | spam |      │
│                                    # administrative | other │
│      "course_relevance": 0.85,     # float 0.0-1.0          │
│      "confidence": 0.92,           # float 0.0-1.0          │
│      "summary": "Student asks about due date for            │
│                  Assignment 2 and requests extension."      │
│    }                                                        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 7. ClassifiedEmail (Pydantic model)                         │
│    - Combines original ParsedEmail with classification      │
│    - Returned to client as API response                     │
└─────────────────────────────────────────────────────────────┘
```

**Key Design Decision**: The email body is truncated to 4000 characters before being sent to the LLM. This prevents token limit issues while preserving enough context for accurate classification. The `max_email_body_chars` setting is configurable via environment variable.

### 2.2 Draft Generation Pipeline

The draft generation pipeline creates an AI-written reply draft for a specific email. This pipeline is invoked when users call `POST /drafts/generate`.

```
[Client Request: POST /drafts/generate]
      │
      │ {message_id: "original-email-id", instructions: ""}
      ▼
┌─────────────────────────────────────────────────────────────┐
│ 1. fetch_inbox_messages(top=50) [graph_service.py]          │
│    - Fetches 50 most recent emails from Outlook             │
│    - Returns list of raw Graph message objects              │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. Find matching message by ID                              │
│    - Searches inbox for message with id == message_id       │
│    - If not found, returns 404 HTTPException                │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. parse_graph_message() [email_parser.py]                  │
│    - Same parsing as classification pipeline                │
│    - Produces ParsedEmail object                            │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. llm_draft_reply(email, instructions) [llm_service.py]    │
│    - Truncates body to max_email_body_chars                 │
│    - Constructs prompt with original email                  │
│    - Appends instructions if provided                        │
│    - Sends to Groq with _DRAFT_SYSTEM prompt                │
│    - Returns plain-text reply body (no salutation/markdown) │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. Draft Creation                                           │
│    - Recipients: original sender (or recipients if group)   │
│    - Subject: "Re: {original_subject}"                      │
│    - Body: LLM-generated reply text                         │
│    - Status: pending                                        │
│    - ID: UUID                                               │
│    - created_at, updated_at: current timestamp              │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 6. Draft stored in memory [_store: dict[str, Draft]]        │
│    - Available for review via GET /drafts                   │
│    - Editable via PATCH /drafts/{id}                        │
│    - Sendable via POST /drafts/{id}/send                    │
└─────────────────────────────────────────────────────────────┘
```

**Key Design Decision**: The draft generation uses `temperature=0.7` (vs. `0.1` for classification) because reply drafting benefits from some creativity and variation, while classification requires consistent, deterministic outputs.

### 2.3 Send Pipeline

The send pipeline transmits an approved draft to the original recipient via Microsoft Outlook.

```
[Client Request: POST /drafts/{draft_id}/send]
      │
      ▼
┌─────────────────────────────────────────────────────────────┐
│ 1. Retrieve draft                                           │
│    - draft_service.get_draft(draft_id)                      │
│    - If not found: 404 HTTPException                        │
│    - If already sent: 400 HTTPException                     │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. send_message(to, subject, body) [graph_service.py]       │
│    - Gets cached access token (or errors if not authed)     │
│    - POSTs to https://graph.microsoft.com/v1.0/me/sendMail  │
│    - Authorization: Bearer {token}                          │
│    - Payload: {message: {subject, body, toRecipients}}      │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. Update draft status                                      │
│    - draft_service.update_draft(draft_id,                   │
│      DraftUpdate(status=DraftStatus.sent))                  │
│    - Returns updated Draft with status="sent"               │
└─────────────────────────────────────────────────────────────┘
```

### 2.4 Complete 5-Step Workflow

The Swagger UI (`/docs`) guides users through the complete workflow:

| Step | Endpoint                 | Description                                    |
| ---- | ------------------------ | ---------------------------------------------- |
| 1    | `GET /auth/login`        | Authenticate with Outlook via device code flow |
| 2    | `GET /emails/inbox`      | Fetch and classify inbox emails                |
| 3    | `POST /drafts/generate`  | Generate an AI reply draft for an email        |
| 4a   | `GET /drafts`            | List all pending drafts for review             |
| 4b   | `PATCH /drafts/{id}`     | Edit draft body or status before sending       |
| 5    | `POST /drafts/{id}/send` | Send approved draft via Outlook                |

This workflow enforces **human-in-the-loop** design: AI generates suggestions, humans review and approve before any email is sent.

---

## 3. Service-by-Service Breakdown

### 3.1 email_parser.py — Email Parsing Service

**File**: `app/services/email_parser.py`

**Purpose**: Convert raw Microsoft Graph message objects into structured `ParsedEmail` Pydantic models suitable for downstream processing.

**Key Function**: `parse_graph_message(msg: dict[str, Any]) → ParsedEmail`

```python
def parse_graph_message(msg: dict[str, Any]) -> ParsedEmail:
    sender_raw = msg.get("sender", {}).get("emailAddress", {})
    sender = EmailAddress(
        name=sender_raw.get("name"),
        address=sender_raw.get("address", "unknown@example.com"),
    )
    recipients = [
        EmailAddress(
            name=r["emailAddress"].get("name"),
            address=r["emailAddress"]["address"],
        )
        for r in msg.get("toRecipients", [])
    ]
    body_content_type = msg.get("body", {}).get("contentType", "text")
    body_raw = msg.get("body", {}).get("content", "")
    body_text = _strip_html(body_raw) if body_content_type == "html" else body_raw
    body_html = body_raw if body_content_type == "html" else None
    return ParsedEmail(
        message_id=msg["id"],
        subject=msg.get("subject", "(no subject)"),
        sender=sender,
        recipients=recipients,
        body_text=body_text,
        body_html=body_html,
        received_at=msg.get("receivedDateTime"),
        thread_id=msg.get("conversationId"),
    )
```

**Key Helper**: `_strip_html(raw: str) → str`

```python
def _strip_html(raw: str) -> str:
    text = re.sub(r"<[^>]+>", " ", raw)        # Remove HTML tags
    return html.unescape(re.sub(r"\s+", " ", text)).strip()  # Unescape entities, normalize whitespace
```

**Responsibilities**:

- Extract sender information (name and email address)
- Build list of recipients
- Handle both HTML and plain-text email bodies
- Normalize body content (strip HTML, unescape entities)
- Preserve thread ID for conversation tracking

**Inputs**: Microsoft Graph message object (dict)  
**Outputs**: `ParsedEmail` Pydantic model

**Dependencies**: `app.models.email.ParsedEmail`, `app.models.email.EmailAddress`

---

### 3.2 classifier.py — Email Classification Orchestrator

**File**: `app/services/classifier.py`

**Purpose**: Orchestrate the email classification process by delegating to the LLM service and mapping results to the `ClassifiedEmail` model.

**Key Function**: `async classify_email(email: ParsedEmail) → ClassifiedEmail`

```python
async def classify_email(email: ParsedEmail) -> ClassifiedEmail:
    result = await llm_classify(email)
    return ClassifiedEmail(
        email=email,
        urgency=result["urgency"],
        intent=result["intent"],
        course_relevance=result["course_relevance"],
        confidence=result["confidence"],
        summary=result["summary"],
    )
```

**Design Rationale**: This service acts as an **orchestration layer** between the API routes and the LLM integration. This separation provides two benefits: (1) the classification logic is testable in isolation by mocking `llm_classify`, and (2) future improvements to classification (e.g., adding rules-based pre-filtering, ensemble methods) can be implemented here without touching API code.

**Responsibilities**:

- Receive parsed email data
- Call LLM classification service
- Transform LLM JSON response to `ClassifiedEmail` model
- Return structured classification with original email attached

**Inputs**: `ParsedEmail`  
**Outputs**: `ClassifiedEmail` (which contains the original `ParsedEmail` plus classification fields)

**Dependencies**: `app.services.llm_service.llm_classify`

---

### 3.3 llm_service.py — LLM Integration Service

**File**: `app/services/llm_service.py`

**Purpose**: Handle all interactions with the Groq LLM API, including prompt construction, API calls, and response parsing. This is the core AI component of the system.

**Client Management**: Uses a singleton pattern for the `AsyncGroq` client to avoid recreating the client on every request.

```python
_client: AsyncGroq | None = None

def get_client() -> AsyncGroq:
    global _client
    if _client is None:
        _client = AsyncGroq(api_key=settings.groq_api_key)
    return _client
```

**System Prompts**:

```python
_CLASSIFY_SYSTEM = """
You are an academic email classification assistant. Given an email, respond ONLY
with a JSON object with exactly these keys:
"urgency": one of ["high", "medium", "low"]
"intent": one of ["inquiry", "complaint", "feedback", "spam", "administrative", "other"]
"course_relevance": float 0-1 (how relevant this email is to a university course or academic matter)
"confidence": float 0-1 (your overall classification confidence)
"summary": one-sentence plain-English summary (max 30 words)
No markdown, no explanation — raw JSON only."""

_DRAFT_SYSTEM = """
You are a professional academic email assistant. Write a helpful, concise reply
to the email provided. Respond ONLY with the plain-text body of the reply — no
salutation line, no subject, no markdown."""
```

**Key Function 1**: `async llm_classify(email: ParsedEmail) → dict[str, Any]`

````python
async def llm_classify(email: ParsedEmail) -> dict[str, Any]:
    body_excerpt = email.body_text[: settings.max_email_body_chars]
    prompt = f"Subject: {email.subject}\n\nBody:\n{body_excerpt}"

    response = await get_client().chat.completions.create(
        model=settings.groq_model,
        messages=[
            {"role": "system", "content": _CLASSIFY_SYSTEM},
            {"role": "user", "content": prompt},
        ],
        max_tokens=256,
        temperature=0.1,
    )
    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
    if raw.startswith("json"):
        raw = raw[4:]
    logger.debug("llm_classify raw response: %s", raw)
    return json.loads(raw.strip())
````

**Key Function 2**: `async llm_draft_reply(email: ParsedEmail, instructions: str = "") → str`

```python
async def llm_draft_reply(email: ParsedEmail, instructions: str = "") -> str:
    body_excerpt = email.body_text[: settings.max_email_body_chars]
    prompt = (
        f"Subject: {email.subject}\n\nOriginal email:\n{body_excerpt}"
        + (f"\n\nAdditional instructions: {instructions}" if instructions else "")
    )

    response = await get_client().chat.completions.create(
        model=settings.groq_model,
        messages=[
            {"role": "system", "content": _DRAFT_SYSTEM},
            {"role": "user", "content": prompt},
        ],
        max_tokens=512,
        temperature=0.7,
    )
    return response.choices[0].message.content.strip()
```

**Key Implementation Details**:

- `temperature=0.1` for classification ensures consistent, low-variance outputs
- `temperature=0.7` for drafting allows natural variation and creativity
- Body truncation at `max_email_body_chars` (default 4000) prevents token overflow
- JSON response parsing strips potential markdown code fences (` ```json ... ``` `)
- `max_tokens` caps response length to control costs and prevent excessive outputs

**Responsibilities**:

- Manage Groq API client lifecycle (singleton)
- Construct prompts with system instructions and email content
- Handle API calls with appropriate parameters
- Parse and validate LLM responses
- Log raw responses for debugging

**Inputs**: `ParsedEmail`, optional instructions string  
**Outputs**: Classification dict (for `llm_classify`) or draft reply string (for `llm_draft_reply`)

**Dependencies**: `groq` SDK, `app.core.config.settings`

---

### 3.4 draft_service.py — Draft Management Service

**File**: `app/services/draft_service.py`

**Purpose**: Manage the lifecycle of email draft objects including creation, retrieval, updating, and deletion. Currently uses in-memory storage; designed to be swapped for a database in production.

**Storage**: In-memory dictionary `_store: dict[str, Draft] = {}`

**Note in Code**: `# In-memory store for the PoC — swap for a DB later`

```python
def _now() -> datetime:
    return datetime.now(timezone.utc)
```

**Key Functions**:

```python
def create_draft(data: DraftCreate) -> Draft:
    draft = Draft(
        id=str(uuid.uuid4()),
        original_message_id=data.original_message_id,
        to=data.to,
        subject=data.subject,
        body=data.body,
        status=DraftStatus.pending,
        created_at=_now(),
        updated_at=_now(),
    )
    _store[draft.id] = draft
    return draft

def get_draft(draft_id: str) -> Optional[Draft]:
    return _store.get(draft_id)

def list_drafts() -> list[Draft]:
    return list(_store.values())

def update_draft(draft_id: str, data: DraftUpdate) -> Optional[Draft]:
    draft = _store.get(draft_id)
    if draft is None:
        return None
    updated = draft.model_copy(
        update={k: v for k, v in data.model_dump(exclude_none=True).items()}
    )
    updated = updated.model_copy(update={"updated_at": _now()})
    _store[draft_id] = updated
    return updated

def delete_draft(draft_id: str) -> bool:
    return _store.pop(draft_id, None) is not None
```

**DraftStatus Enum**:

```python
class DraftStatus(str, Enum):
    pending = "pending"    # Created, not yet reviewed
    approved = "approved"  # Reviewed and approved by user
    sent = "sent"          # Sent via Outlook
    rejected = "rejected"  # Rejected by user
```

**Design Decision**: Using an in-memory store simplifies the PoC by eliminating database setup. This is explicitly marked as temporary (`# In-memory store for the PoC — swap for a DB later`), and the CRUD interface is designed to be compatible with database-backed implementations in production.

**Responsibilities**:

- Generate unique draft IDs (UUID v4)
- Maintain draft state (status, timestamps)
- Provide CRUD operations on drafts
- Auto-update `updated_at` timestamp on modifications

**Inputs**: `DraftCreate`, `DraftUpdate`  
**Outputs**: `Draft` objects or collections

**Dependencies**: `app.models.draft.Draft`, `app.models.draft.DraftCreate`, `app.models.draft.DraftUpdate`, `app.models.draft.DraftStatus`

---

### 3.5 graph_service.py — Microsoft Graph API Integration

**File**: `app/services/graph_service.py`

**Purpose**: Handle all interactions with Microsoft Graph API for Outlook email access. This includes OAuth authentication via device code flow, fetching inbox messages, and sending emails.

**Authentication**: MSAL (Microsoft Authentication Library) PublicClientApplication with device code flow. No client secret is required because this authenticates as a user, not an application.

**Scopes**: `["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.Send"]`

**Token Caching**: MSAL's `SerializableTokenCache` persists tokens in memory across requests.

```python
_token_cache = msal.SerializableTokenCache()
_cached_token: str | None = None

def _get_msal_app() -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=settings.azure_client_id,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=_token_cache,
    )
```

**Key Function 1**: `initiate_device_flow() → dict[str, Any]`

Starts the OAuth device code flow. Returns a dict containing:

- `user_code`: Code the user enters at the verification URL
- `verification_uri`: URL where the user enters the code
- `message`: Human-readable message to display to the user

```python
def initiate_device_flow() -> dict[str, Any]:
    app = _get_msal_app()
    flow = app.initiate_device_flow(scopes=_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow error: {flow.get('error_description')}")
    return flow
```

**Key Function 2**: `complete_device_flow(flow: dict) → str`

Polls the device code flow until the user completes authentication. Returns the access token.

```python
def complete_device_flow(flow: dict[str, Any]) -> str:
    app = _get_msal_app()
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Auth error: {result.get('error_description')}")
    return result["access_token"]
```

**Key Function 3**: `get_access_token() → str | None`

Returns a cached token if available, otherwise returns `None`.

```python
def get_access_token() -> str | None:
    app = _get_msal_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    return None
```

**Key Function 4**: `fetch_inbox_messages(top: int = 25) → list[dict[str, Any]]`

Fetches recent emails from the authenticated user's inbox.

```python
async def fetch_inbox_messages(top: int = 25) -> list[dict[str, Any]]:
    token = get_access_token()
    if not token:
        raise RuntimeError("Not authenticated. Call GET /auth/login first.")
    url = f"{settings.graph_api_base}/me/mailFolders/inbox/messages"
    params = {
        "$top": top,
        "$select": "id,subject,sender,toRecipients,body,receivedDateTime,conversationId",
        "$orderby": "receivedDateTime desc",
    }
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=15,
        )
        resp.raise_for_status()
        return resp.json().get("value", [])
```

**Key Function 5**: `send_message(to: list[str], subject: str, body: str) → None`

Sends an email via the Graph API.

```python
async def send_message(to: list[str], subject: str, body: str) -> None:
    token = get_access_token()
    if not token:
        raise RuntimeError("Not authenticated. Call GET /auth/login first.")
    url = f"{settings.graph_api_base}/me/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        }
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=15,
        )
        resp.raise_for_status()
    logger.info("Email sent to %s", to)
```

**Design Decision**: Device code flow was chosen because it allows user authentication without requiring a server-side redirect URL. This is ideal for CLI tools, desktop applications, and PoC projects where setting up Azure AD app redirect URIs would add complexity.

**Responsibilities**:

- Manage OAuth device code flow lifecycle
- Cache and retrieve access tokens
- Construct Graph API requests with proper authorization
- Handle async HTTP calls to Graph API endpoints
- Throw descriptive errors when not authenticated

**Inputs**: Message fetch count (`top`), recipients list, subject, body  
**Outputs**: List of raw Graph message dicts, or None (for send)

**Dependencies**: `msal`, `httpx`, `app.core.config.settings`

---

### 3.6 config.py — Configuration Management

**File**: `app/core/config.py`

**Purpose**: Centralised management of all application settings, loaded from environment variables with sensible defaults.

**Implementation**: Uses `pydantic_settings.BaseSettings` to auto-load from `.env` file.

```python
class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8")

    # App
    app_name: str = "Email AI PoC"
    debug: bool = False

    # Azure / Microsoft Graph
    azure_client_id: str = ""
    azure_client_secret: str = ""
    azure_tenant_id: str = ""
    graph_api_base: str = "https://graph.microsoft.com/v1.0"
    outlook_user: str = ""

    # Groq
    groq_api_key: str = ""
    groq_model: str = "llama-3.3-70b-versatile"

    # Email processing
    max_email_body_chars: int = 4000

settings = Settings()
```

**Configuration Variables**:

| Variable               | Default                   | Description                             |
| ---------------------- | ------------------------- | --------------------------------------- |
| `AZURE_CLIENT_ID`      | (empty)                   | Azure AD app registration client ID     |
| `AZURE_CLIENT_SECRET`  | (empty)                   | Azure AD app registration client secret |
| `AZURE_TENANT_ID`      | (empty)                   | Azure AD tenant ID                      |
| `OUTLOOK_USER`         | (empty)                   | Outlook email address                   |
| `GROQ_API_KEY`         | (empty)                   | Groq API key                            |
| `GROQ_MODEL`           | `llama-3.3-70b-versatile` | Groq model name                         |
| `DEBUG`                | `false`                   | Enable debug logging                    |
| `MAX_EMAIL_BODY_CHARS` | `4000`                    | Max email body length sent to LLM       |

**Design Decision**: Using `pydantic-settings` with environment variables keeps sensitive credentials out of the codebase. The `.env` file is gitignored, and `.env.example` provides a template for required variables.

---

### 3.7 logging.py — Logging Configuration

**File**: `app/core/logging.py`

**Purpose**: Configure structured logging for the application with timestamped, formatted output.

**Implementation**:

```python
def setup_logging(debug: bool = False) -> None:
    level = logging.DEBUG if debug else logging.INFO
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)-8s | %(name)s | %(message)s")
    )
    root = logging.getLogger()
    root.setLevel(level)
    root.handlers = [handler]

def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)
```

**Log Format**: `2026-04-15 10:30:00 | INFO     | app.services.llm_service | llm_classify raw response: {...}`

**Usage**: Services create module-level loggers with `get_logger(__name__)` to enable hierarchical log routing and filtering.

---

## 4. API Endpoint Reference

### 4.1 Health Check

**`GET /health`**

Returns a simple health status response to verify the API is running.

**Response** (200 OK):

```json
{
  "status": "ok",
  "version": "0.1.0",
  "services": {
    "graph_api": "unchecked",
    "llm": "unchecked"
  }
}
```

**Purpose**: Load balancer health checks, uptime monitoring

---

### 4.2 Authentication

**`GET /auth/login`** — Step 1 of the workflow

Initiates OAuth device code flow for Microsoft Graph authentication.

**Summary**: "Step 1 — Authenticate with Outlook"

**Response** (200 OK):

```json
{
  "success": true,
  "data": {
    "message": "To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code ABC12345 to authenticate.",
    "user_code": "ABC12345",
    "verification_url": "https://microsoft.com/devicelogin"
  },
  "error": null,
  "message": null
}
```

**Workflow**:

1. Client calls this endpoint
2. User visits `verification_url` and enters `user_code`
3. User signs into Microsoft account with Outlook access
4. Token is cached automatically (background thread polls the flow)
5. Subsequent calls to `/emails/inbox` and `/drafts/generate` will work

**Errors** (500):

```json
{ "detail": "Device flow error: ..." }
```

---

### 4.3 Email Processing

**`POST /emails/process`** (hidden from OpenAPI schema)

Process a single email from a raw Graph API message object.

**Request Body**:

```json
{
  "raw_message": {
    "id": "message-id",
    "subject": "Question about assignment",
    "sender": {
      "emailAddress": { "name": "Student", "address": "student@university.edu" }
    },
    "toRecipients": [
      {
        "emailAddress": {
          "name": "Lecturer",
          "address": "lecturer@university.edu"
        }
      }
    ],
    "body": { "contentType": "text", "content": "Hi, I have a question..." },
    "receivedDateTime": "2026-04-15T10:00:00Z"
  }
}
```

**Response** (200 OK):

```json
{
  "success": true,
  "data": {
    "email": {
      "message_id": "message-id",
      "subject": "Question about assignment",
      "sender": { "name": "Student", "address": "student@university.edu" },
      "recipients": [
        { "name": "Lecturer", "address": "lecturer@university.edu" }
      ],
      "body_text": "Hi, I have a question...",
      "body_html": null,
      "received_at": "2026-04-15T10:00:00Z",
      "thread_id": null
    },
    "urgency": "medium",
    "intent": "inquiry",
    "course_relevance": 0.85,
    "confidence": 0.92,
    "summary": "Student asks about assignment deadline and requests extension."
  },
  "error": null,
  "message": null
}
```

---

**`GET /emails/inbox`** — Step 2 of the workflow

Fetch and classify emails from the authenticated user's Outlook inbox.

**Summary**: "Step 2 — Fetch & classify inbox emails"

**Query Parameters**:
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `top` | int | 10 | Number of emails to fetch (max 25) |

**Response** (200 OK):

```json
{
  "success": true,
  "data": [
    {
      "email": {"message_id": "...", "subject": "...", ...},
      "urgency": "high",
      "intent": "complaint",
      "course_relevance": 0.9,
      "confidence": 0.88,
      "summary": "Student complains about mark received on Assignment 1."
    },
    {
      "email": {"message_id": "...", "subject": "...", ...},
      "urgency": "low",
      "intent": "administrative",
      "course_relevance": 0.3,
      "confidence": 0.95,
      "summary": "University-wide announcement about library hours."
    }
  ],
  "error": null,
  "message": null
}
```

**Errors** (500):

```json
{ "detail": "Not authenticated. Call GET /auth/login first." }
```

---

### 4.4 Draft Generation

**`POST /drafts/generate`** — Step 3 of the workflow

Generate an AI reply draft for a specific email from the inbox.

**Summary**: "Step 3 — Generate AI draft reply"

**Request Body**:

```json
{
  "message_id": "original-email-id-from-inbox",
  "instructions": "Make the tone more formal"
}
```

**Response** (201 Created):

```json
{
  "success": true,
  "data": {
    "id": "draft-uuid",
    "original_message_id": "original-email-id-from-inbox",
    "to": ["student@university.edu"],
    "subject": "Re: Question about assignment",
    "body": "Thank you for reaching out regarding the assignment deadline. I understand you may be facing some challenges and I'm happy to discuss possible extensions during my office hours on Thursday.",
    "status": "pending",
    "created_at": "2026-04-15T10:05:00Z",
    "updated_at": "2026-04-15T10:05:00Z"
  },
  "error": null,
  "message": null
}
```

**Errors**:

- (404) `"Message not found in inbox"` — message_id not in fetched inbox
- (500) `"Not authenticated. Call GET /auth/login first."`

---

### 4.5 Draft Management

**`GET /drafts`** — Step 4a of the workflow

List all drafts in the system for review.

**Summary**: "Step 4a — List all pending drafts"

**Response** (200 OK):

```json
{
  "success": true,
  "data": [
    {
      "id": "draft-uuid",
      "original_message_id": "...",
      "to": ["student@university.edu"],
      "subject": "Re: Question about assignment",
      "body": "Thank you for reaching out...",
      "status": "pending",
      "created_at": "2026-04-15T10:05:00Z",
      "updated_at": "2026-04-15T10:05:00Z"
    }
  ],
  "error": null,
  "message": null
}
```

---

**`GET /drafts/{draft_id}`**

Retrieve a specific draft by ID.

**Response** (200 OK): Same structure as single draft in list above

**Errors**:

- (404) `"Draft not found"`

---

**`PATCH /drafts/{draft_id}`** — Step 4b of the workflow

Edit a draft's subject, body, or status.

**Summary**: "Step 4b — Edit draft body or status"

**Request Body** (all fields optional):

```json
{
  "subject": "Updated subject",
  "body": "Edited reply body...",
  "status": "approved"
}
```

**Status Values**: `"pending"`, `"approved"`, `"rejected"`

**Response** (200 OK): Updated draft object

**Errors**:

- (404) `"Draft not found"`

---

**`DELETE /drafts/{draft_id}`** (hidden from schema)

Delete a draft permanently.

**Response** (200 OK):

```json
{
  "success": true,
  "data": null,
  "error": null,
  "message": "Draft deleted"
}
```

---

### 4.6 Send Draft

**`POST /drafts/{draft_id}/send`** — Step 5 of the workflow

Send an approved draft via Outlook and mark it as sent.

**Summary**: "Step 5 — Send draft via Outlook"

**Response** (200 OK):

```json
{
  "success": true,
  "data": {
    "id": "draft-uuid",
    "status": "sent",
    ...
  },
  "error": null,
  "message": "Email sent successfully"
}
```

**Errors**:

- (404) `"Draft not found"`
- (400) `"Draft already sent"`
- (500) Graph API error details

---

## 5. Authentication Flow

### 5.1 OAuth Device Code Flow Overview

The system uses Microsoft's OAuth 2.0 Device Authorization Grant (device code flow) to authenticate users. This flow is designed for devices with limited input capabilities (smart TVs, CLI tools, embedded systems) where redirect-based OAuth would be impractical.

Unlike traditional OAuth where the user is redirected to a web page to log in, device code flow displays a short code and URL on the device, allowing the user to authenticate on a separate device (phone, laptop) with a full browser.

### 5.2 Step-by-Step Flow

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           Device Code Flow Sequence                         │
└─────────────────────────────────────────────────────────────────────────────┘

Client                                          Microsoft
  │                                                  │
  │  1. GET /auth/login                              │
  │  ─────────────────────────────────────────────►  │
  │                                                  │
  │  2. initiate_device_flow()                       │
  │     Returns: {                                   │
  │       user_code: "ABC12345",                     │
  │       verification_uri: "...",                   │
  │       message: "To sign in, use a web            │
  │         browser to open..."                      │
  │     }                                            │
  │  ◄─────────────────────────────────────────────  │
  │                                                  │
  │  3. Display message + code to user               │
  │                                                  │
  │  4. User opens browser → verification_uri        │
  │  ◄─────────────────────────────────────────────  │
  │     (on user's separate device)                  │
  │                                                  │
  │  5. User enters user_code "ABC12345"             │
  │  ─────────────────────────────────────────────►  │
  │                                                  │
  │  6. User signs into Microsoft account            │
  │  ─────────────────────────────────────────────►  │
  │                                                  │
  │  7. Confirmation shown                           │
  │  ◄─────────────────────────────────────────────  │
  │     (on user's separate device)                  │
  │                                                  │
  │  8. Background: complete_device_flow() polls     │
  │     for token completion                         │
  │  ─────────────────────────────────────────────►  │
  │     (polling until user completes step 6)        │
  │                                                  │
  │  9. Token received and cached in MSAL cache      │
  │  ◄─────────────────────────────────────────────  │
  │                                                  │
  │  10. GET /emails/inbox (with cached token)       │
  │  ─────────────────────────────────────────────►  │
  │                                                  │
  │  11. Graph API returns inbox messages            │
  │  ◄─────────────────────────────────────────────  │
  │                                                  │
```

### 5.3 Implementation Details

**Initiating the Flow** (`auth.py`):

```python
@router.get("/auth/login", response_model=APIResponse[dict])
async def login() -> APIResponse[dict]:
    global _active_flow
    flow = initiate_device_flow()
    _active_flow = flow
    # Kick off background polling so token gets cached automatically
    threading.Thread(target=_poll_flow, args=(flow,), daemon=True).start()
    return APIResponse(
        success=True,
        data={
            "message": flow["message"],
            "user_code": flow["user_code"],
            "verification_url": flow["verification_uri"],
        },
    )
```

**Background Polling** (`auth.py`):

```python
def _poll_flow(flow: dict[str, Any]) -> None:
    """Background thread — polls until sign-in completes and caches the token."""
    try:
        complete_device_flow(flow)
    except Exception as exc:
        pass  # Will surface as 401 on next inbox request
```

The background thread polls `complete_device_flow()` until the user completes authentication on the Microsoft login page. Once successful, the token is stored in MSAL's token cache and subsequent API calls use it automatically.

### 5.4 Token Caching Mechanism

MSAL's `SerializableTokenCache` maintains tokens in memory:

1. **First authentication**: `initiate_device_flow()` → `complete_device_flow()` → token cached
2. **Subsequent requests**: `get_access_token()` → `acquire_token_silent()` → returns cached token
3. **Token expiry**: When token expires, `acquire_token_silent()` returns `None`, and user must re-authenticate

The `token_cache` is a module-level variable in `graph_service.py`, surviving across HTTP requests within the same process.

### 5.5 Error Handling

If the user has not authenticated:

```python
if not token:
    raise RuntimeError("Not authenticated. Call GET /auth/login first.")
```

If the device code flow fails:

```python
if "user_code" not in flow:
    raise RuntimeError(f"Device flow error: {flow.get('error_description')}")
```

### 5.6 Security Considerations

- The device code has a limited lifetime (typically 15 minutes)
- The code can only be used once per authentication attempt
- No client secret is required because this is user authentication, not application authentication
- Tokens are stored in memory only (not persisted to disk in this PoC)
- All Graph API calls use HTTPS (enforced by the Graph API itself)

---

## 6. Workshop-Friendly Summary

### 6.1 What the System Does: Simple Explanation

Imagine you have a very capable administrative assistant who:

1. **Reads all your emails automatically** — without you needing to open Outlook
2. **Understands what each email is about** — is it urgent? A complaint? A simple question?
3. **Drafts replies for you** — writes professional responses that you can review
4. **Sends only when you approve** — never auto-sends; you always have final control

This assistant never gets tired, never misses an email, and always follows your guidelines. That's exactly what this system does, acting as an AI-powered layer between your Outlook inbox and your email workflow.

### 6.2 AI Agent Analogy

Think of the system as an **AI agent** with three tools:

| Tool                    | What it does                | Under the hood                                 |
| ----------------------- | --------------------------- | ---------------------------------------------- |
| **Tool 1: Read Inbox**  | Fetches emails from Outlook | `fetch_inbox_messages()` → Microsoft Graph API |
| **Tool 2: Classify**    | Understands email content   | `llm_classify()` → Groq LLM (Llama 3.3)        |
| **Tool 3: Draft Reply** | Writes suggested responses  | `llm_draft_reply()` → Groq LLM (Llama 3.3)     |

The agent follows a simple loop:

```
1. Receive message_id from user
       ↓
2. Tool 1: Read email from Outlook
       ↓
3. Tool 2: Classify (urgency? intent? relevance?)
       ↓
4. Tool 3: Draft reply
       ↓
5. Present draft to user for review
       ↓
6. If approved → Tool 1 (send): Send via Outlook
       ↓
7. Loop back to step 1
```

### 6.3 Key Design Decisions Explained

**"Why does the system classify emails before drafting replies?"**

Classification serves two purposes:

- **Triage**: Urgent emails can be prioritized automatically (high-urgency items highlighted)
- **Routing**: Different types of emails might need different reply styles (complaints vs. inquiries)

The classification output also provides a **summary** that helps the user quickly understand the email without reading the full content.

**"Why is there a human-in-the-loop? Why not auto-send?"**

AI-generated content can be inaccurate, inappropriate, or even harmful (hallucinations, bias, misinterpretation). By requiring human approval before sending:

- Errors can be caught and corrected
- The system maintains ethical standards
- Responsibility remains with the human operator
- Compliance with AI ethics frameworks (AUS AI Ethics Principles, ISO/IEC 42001)

This aligns with the **human oversight** principle required by the project's Generative AI Policy.

**"Why does the draft have a status field (pending/approved/sent/rejected)?"**

The status field tracks the draft lifecycle:

- `pending` — Just created, needs review
- `approved` — User reviewed and approved for sending
- `sent` — Successfully sent via Outlook
- `rejected` — User decided not to send

This allows the system to support workflows where drafts might be reviewed, revised, approved, and sent over time—not all in one session.

**"Why use an in-memory store for drafts instead of a database?"**

This is a **proof of concept** (PoC), not a production system. Adding a database (PostgreSQL, MongoDB) would increase complexity without adding value for the demonstration. The in-memory store clearly shows the intent: "swap for a DB later" when this becomes a real product.

### 6.4 Workshop Activity Mapping

This system logic directly supports the workshop activities:

| Workshop Activity                | Related System Component                       |
| -------------------------------- | ---------------------------------------------- |
| Build a simple AI agent workflow | The 5-step API workflow                        |
| Understand classification        | `llm_classify()` + classification fields       |
| Prompt engineering demo          | `_CLASSIFY_SYSTEM` and `_DRAFT_SYSTEM` prompts |
| Human-in-the-loop design         | Draft review before send                       |
| AI ethics discussion             | Status field, approval workflow, no auto-send  |
