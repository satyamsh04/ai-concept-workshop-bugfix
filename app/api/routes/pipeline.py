from __future__ import annotations

from typing import Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from app.models.draft import Draft, DraftCreate
from app.models.email import ClassifiedEmail
from app.models.responses import APIResponse
from app.services import draft_service
from app.services.classifier import classify_email
from app.services.email_parser import parse_graph_message
from app.services.graph_service import (
    ensure_categories,
    ensure_folders,
    fetch_inbox_messages,
    move_message,
    save_draft_to_outlook,
    set_message_importance,
)
from app.services.llm_service import llm_draft_reply

router = APIRouter(prefix="/pipeline", tags=["pipeline"])


class PipelineResult(BaseModel):
    classified: list[ClassifiedEmail]
    drafts_generated: list[Draft]
    emails_moved: int


class PipelineRequest(BaseModel):
    top: Optional[int] = 10
    auto_draft: Optional[bool] = True          # generate AI drafts automatically
    confidence_threshold: Optional[float] = 0.7  # only draft if confidence >= this
    move_to_folders: Optional[bool] = True     # move emails into urgency folders


@router.post("/run", response_model=APIResponse[PipelineResult], summary="Run full email pipeline")
async def run_pipeline(req: PipelineRequest = PipelineRequest()) -> APIResponse[PipelineResult]:
    """
    Full automated pipeline:
    1. Fetch inbox emails
    2. Classify each email (urgency, intent, course relevance)
    3. Tag importance flag in Outlook
    4. Move emails into urgency folders
    5. Auto-generate AI draft replies (if confidence meets threshold)
    """
    try:
        await ensure_categories()
        await ensure_folders()

        messages = await fetch_inbox_messages(top=req.top)
        classified_list: list[ClassifiedEmail] = []
        drafts: list[Draft] = []
        moved = 0

        for msg in messages:
            parsed = parse_graph_message(msg)
            classified = await classify_email(parsed)
            classified_list.append(classified)

            # Tag importance in Outlook
            await set_message_importance(parsed.message_id, classified.urgency)

            # Move to urgency folder
            if req.move_to_folders:
                folder_name = f"{classified.urgency.capitalize()} Urgency"
                await move_message(parsed.message_id, folder_name)
                moved += 1

            # Auto-generate draft if confidence meets threshold
            if req.auto_draft and classified.confidence >= req.confidence_threshold:
                body = await llm_draft_reply(parsed)
                recipients = [r.address for r in parsed.recipients] or [parsed.sender.address]
                # Save to Outlook Drafts folder
                await save_draft_to_outlook(
                    to=recipients,
                    subject=f"Re: {parsed.subject}",
                    body=body,
                )
                # Also track in local store for API access
                draft = draft_service.create_draft(
                    DraftCreate(
                        original_message_id=parsed.message_id,
                        to=recipients,
                        subject=f"Re: {parsed.subject}",
                        body=body,
                    )
                )
                drafts.append(draft)

        return APIResponse(
            success=True,
            data=PipelineResult(
                classified=classified_list,
                drafts_generated=drafts,
                emails_moved=moved,
            ),
            message=f"Processed {len(classified_list)} emails, generated {len(drafts)} drafts, moved {moved} to folders",
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
