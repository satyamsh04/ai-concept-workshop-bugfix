from __future__ import annotations

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional

from app.models.draft import Draft, DraftCreate, DraftUpdate
from app.models.responses import APIResponse
from app.models.draft import DraftStatus
from app.services import draft_service
from app.services.email_parser import parse_graph_message
from app.services.graph_service import fetch_inbox_messages, send_message
from app.services.llm_service import llm_draft_reply

router = APIRouter(prefix="/drafts", tags=["drafts"])


class GenerateRequest(BaseModel):
    message_id: str
    instructions: Optional[str] = ""


@router.post("/generate", response_model=APIResponse[Draft], status_code=201, include_in_schema=False)
async def generate_draft(req: GenerateRequest) -> APIResponse[Draft]:
    """Takes a message_id from the inbox, generates an AI reply via Groq, and saves it as a pending draft."""
    try:
        messages = await fetch_inbox_messages(top=50)
        msg = next((m for m in messages if m["id"] == req.message_id), None)
        if msg is None:
            raise HTTPException(status_code=404, detail="Message not found in inbox")

        parsed = parse_graph_message(msg)
        body = await llm_draft_reply(parsed, instructions=req.instructions or "")

        recipients = [r.address for r in parsed.recipients] or [parsed.sender.address]
        draft = draft_service.create_draft(
            DraftCreate(
                original_message_id=parsed.message_id,
                to=recipients,
                subject=f"Re: {parsed.subject}",
                body=body,
            )
        )
        return APIResponse(success=True, data=draft)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@router.post("", response_model=APIResponse[Draft], status_code=201, include_in_schema=False)
async def create_draft(data: DraftCreate) -> APIResponse[Draft]:
    draft = draft_service.create_draft(data)
    return APIResponse(success=True, data=draft)


@router.get("", response_model=APIResponse[list[Draft]], include_in_schema=False)
async def list_drafts() -> APIResponse[list[Draft]]:
    return APIResponse(success=True, data=draft_service.list_drafts())


@router.get("/{draft_id}", response_model=APIResponse[Draft], include_in_schema=False)
async def get_draft(draft_id: str) -> APIResponse[Draft]:
    draft = draft_service.get_draft(draft_id)
    if draft is None:
        raise HTTPException(status_code=404, detail="Draft not found")
    return APIResponse(success=True, data=draft)


@router.patch("/{draft_id}", response_model=APIResponse[Draft], include_in_schema=False)
async def update_draft(draft_id: str, data: DraftUpdate) -> APIResponse[Draft]:
    draft = draft_service.update_draft(draft_id, data)
    if draft is None:
        raise HTTPException(status_code=404, detail="Draft not found")
    return APIResponse(success=True, data=draft)


@router.delete("/{draft_id}", response_model=APIResponse[None], include_in_schema=False)
async def delete_draft(draft_id: str) -> APIResponse[None]:
    if not draft_service.delete_draft(draft_id):
        raise HTTPException(status_code=404, detail="Draft not found")
    return APIResponse(success=True, message="Draft deleted")


@router.post("/{draft_id}/send", response_model=APIResponse[Draft], include_in_schema=False)
async def send_draft(draft_id: str) -> APIResponse[Draft]:
    """Sends the draft email via Microsoft Graph and marks it as sent."""
    draft = draft_service.get_draft(draft_id)
    if draft is None:
        raise HTTPException(status_code=404, detail="Draft not found")
    if draft.status == DraftStatus.sent:
        raise HTTPException(status_code=400, detail="Draft already sent")
    try:
        await send_message(to=draft.to, subject=draft.subject, body=draft.body)
        sent = draft_service.update_draft(draft_id, DraftUpdate(status=DraftStatus.sent))
        return APIResponse(success=True, data=sent, message="Email sent successfully")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
