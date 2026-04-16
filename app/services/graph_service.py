from __future__ import annotations

import asyncio
from typing import Any

import msal
import httpx

from app.core.config import settings
from app.core.logging import get_logger

logger = get_logger(__name__)

_SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite", "https://graph.microsoft.com/Mail.Send"]
_token_cache = msal.SerializableTokenCache()
_cached_token: str | None = None


def _get_msal_app() -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=settings.azure_client_id,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=_token_cache,
    )


def get_access_token() -> str | None:
    """Return a cached token if available."""
    app = _get_msal_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    return None


def initiate_device_flow() -> dict[str, Any]:
    """Start device code flow — returns user_code and verification_uri."""
    app = _get_msal_app()
    flow = app.initiate_device_flow(scopes=_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow error: {flow.get('error_description')}")
    return flow


def complete_device_flow(flow: dict[str, Any]) -> str:
    """Poll until the user completes sign-in, return access token."""
    app = _get_msal_app()
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Auth error: {result.get('error_description')}")
    return result["access_token"]


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


_URGENCY_TO_IMPORTANCE = {"high": "high", "medium": "normal", "low": "low"}
_URGENCY_TO_CATEGORY = {"high": "High Urgency", "medium": "Medium Urgency", "low": "Low Urgency"}
_CATEGORY_COLOURS = {
    "High Urgency": "red",
    "Medium Urgency": "yellow",
    "Low Urgency": "green",
}


async def ensure_categories() -> None:
    """Create urgency colour categories in Outlook if supported (work/school accounts only)."""
    token = get_access_token()
    if not token:
        return
    url = f"{settings.graph_api_base}/me/outlook/masterCategories"
    try:
        async with httpx.AsyncClient() as client:
            resp = await client.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=15)
            if resp.status_code == 403:
                logger.info("masterCategories not supported for this account type — skipping colour categories")
                return
            resp.raise_for_status()
            existing = {c["displayName"] for c in resp.json().get("value", [])}

            for name, colour in _CATEGORY_COLOURS.items():
                if name not in existing:
                    await client.post(
                        url,
                        headers={
                            "Authorization": f"Bearer {token}",
                            "Content-Type": "application/json",
                        },
                        json={"displayName": name, "color": colour},
                        timeout=15,
                    )
                    logger.info("Created Outlook category: %s (%s)", name, colour)
    except Exception as exc:
        logger.warning("Could not create Outlook categories: %s", exc)


async def set_message_importance(message_id: str, urgency: str) -> None:
    """Tag a message with an importance flag and colour category based on AI urgency."""
    token = get_access_token()
    if not token:
        return  # silently skip if not authenticated
    importance = _URGENCY_TO_IMPORTANCE.get(urgency, "normal")
    category = _URGENCY_TO_CATEGORY.get(urgency, "Medium Urgency")
    url = f"{settings.graph_api_base}/me/messages/{message_id}"
    async with httpx.AsyncClient() as client:
        # Try with categories first, fall back to importance only if categories are unsupported
        resp = await client.patch(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"importance": importance, "categories": [category]},
            timeout=15,
        )
        if resp.status_code == 403:
            resp = await client.patch(
                url,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json={"importance": importance},
                timeout=15,
            )
        resp.raise_for_status()
    logger.info("Set importance=%s category=%s on message %s", importance, category, message_id)


_URGENCY_FOLDERS = ["High Urgency", "Medium Urgency", "Low Urgency", "AI Processed"]
_folder_id_cache: dict[str, str] = {}


async def ensure_folders() -> dict[str, str]:
    """Create urgency mail folders if they don't exist. Returns a name→id map."""
    token = get_access_token()
    if not token:
        return {}
    if _folder_id_cache:
        return _folder_id_cache

    url = f"{settings.graph_api_base}/me/mailFolders"
    async with httpx.AsyncClient() as client:
        resp = await client.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=15)
        resp.raise_for_status()
        existing = {f["displayName"]: f["id"] for f in resp.json().get("value", [])}

        for name in _URGENCY_FOLDERS:
            if name in existing:
                _folder_id_cache[name] = existing[name]
            else:
                r = await client.post(
                    url,
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json",
                    },
                    json={"displayName": name},
                    timeout=15,
                )
                r.raise_for_status()
                _folder_id_cache[name] = r.json()["id"]
                logger.info("Created mail folder: %s", name)

    return _folder_id_cache


async def move_message(message_id: str, folder_name: str) -> None:
    """Move a message into the named mail folder."""
    token = get_access_token()
    if not token:
        return
    folders = await ensure_folders()
    folder_id = folders.get(folder_name)
    if not folder_id:
        logger.warning("Folder not found: %s", folder_name)
        return
    url = f"{settings.graph_api_base}/me/messages/{message_id}/move"
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"destinationId": folder_id},
            timeout=15,
        )
        resp.raise_for_status()
    logger.info("Moved message %s to folder %s", message_id, folder_name)


async def save_draft_to_outlook(to: list[str], subject: str, body: str) -> str:
    """Save a draft to Outlook's Drafts folder. Returns the Outlook draft message ID."""
    token = get_access_token()
    if not token:
        raise RuntimeError("Not authenticated. Call GET /auth/login first.")
    url = f"{settings.graph_api_base}/me/messages"
    payload = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        "isDraft": True,
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
        outlook_id = resp.json()["id"]
    logger.info("Saved draft to Outlook drafts folder: %s", outlook_id)
    return outlook_id


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
