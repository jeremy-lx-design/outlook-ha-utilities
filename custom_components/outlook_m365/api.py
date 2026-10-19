"""Microsoft Graph API client for the Outlook M365 integration.

A thin async wrapper around the subset of Microsoft Graph endpoints we need
to read, send and organise mail, plus manage inbox (message) rules.

Authentication is delegated to Home Assistant's OAuth2 session, which handles
refreshing the access token transparently.
"""

from __future__ import annotations

import logging
from typing import Any

from aiohttp import ClientResponse, ClientSession

from homeassistant.helpers import config_entry_oauth2_flow

from .const import GRAPH_BASE_URL

_LOGGER = logging.getLogger(__name__)


class OutlookApiError(Exception):
    """Raised when the Microsoft Graph API returns an error."""


class OutlookAuthError(OutlookApiError):
    """Raised when authentication/authorization fails (401/403)."""


class OutlookApiClient:
    """Minimal async client for Microsoft Graph mail operations."""

    def __init__(
        self,
        websession: ClientSession,
        oauth_session: config_entry_oauth2_flow.OAuth2Session,
    ) -> None:
        """Initialise the client with a web session and OAuth2 session."""
        self._websession = websession
        self._oauth_session = oauth_session

    async def _async_token(self) -> str:
        """Return a valid access token, refreshing if required."""
        await self._oauth_session.async_ensure_token_valid()
        return self._oauth_session.token["access_token"]

    async def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        json: dict[str, Any] | None = None,
    ) -> Any:
        """Perform an authenticated request against Microsoft Graph.

        ``path`` may be an absolute URL (e.g. a Graph ``@odata.nextLink``) or a
        path relative to the Graph v1.0 base URL.
        """
        token = await self._async_token()
        url = path if path.startswith("http") else f"{GRAPH_BASE_URL}{path}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        _LOGGER.debug("Graph %s %s params=%s", method, url, params)
        resp: ClientResponse = await self._websession.request(
            method, url, headers=headers, params=params, json=json
        )

        if resp.status in (401, 403):
            text = await resp.text()
            raise OutlookAuthError(f"Authentication failed ({resp.status}): {text}")

        if resp.status == 429:
            retry_after = resp.headers.get("Retry-After", "unknown")
            raise OutlookApiError(
                f"Rate limited by Microsoft Graph (retry after {retry_after}s)"
            )

        if resp.status >= 400:
            text = await resp.text()
            raise OutlookApiError(f"Graph request failed ({resp.status}): {text}")

        if resp.status == 204 or resp.content_length == 0:
            return None

        # Some endpoints (e.g. sendMail) return 202 with an empty body.
        if resp.content_type != "application/json":
            return None

        return await resp.json()

    # ------------------------------------------------------------------
    # Profile
    # ------------------------------------------------------------------
    async def async_get_profile(self) -> dict[str, Any]:
        """Return the signed-in user's profile."""
        return await self._request("GET", "/me")

    # ------------------------------------------------------------------
    # Reading messages
    # ------------------------------------------------------------------
    async def async_list_messages(
        self,
        folder: str = "inbox",
        *,
        top: int = 25,
        select: list[str] | None = None,
        filter_query: str | None = None,
        orderby: str | None = None,
    ) -> list[dict[str, Any]]:
        """List messages in a well-known folder (default: inbox)."""
        params: dict[str, Any] = {"$top": top}
        if select:
            params["$select"] = ",".join(select)
        if filter_query:
            params["$filter"] = filter_query
        if orderby:
            params["$orderby"] = orderby

        data = await self._request(
            "GET", f"/me/mailFolders/{folder}/messages", params=params
        )
        return data.get("value", []) if data else []

    async def async_get_message(self, message_id: str) -> dict[str, Any]:
        """Return a single message by id."""
        return await self._request("GET", f"/me/messages/{message_id}")

    async def async_get_unread_count(self, folder: str = "inbox") -> int:
        """Return the unread item count for a well-known folder."""
        data = await self._request(
            "GET",
            f"/me/mailFolders/{folder}",
            params={"$select": "unreadItemCount"},
        )
        return int(data.get("unreadItemCount", 0)) if data else 0

    # ------------------------------------------------------------------
    # Sending mail
    # ------------------------------------------------------------------
    async def async_send_mail(
        self,
        *,
        subject: str,
        body: str,
        to: list[str],
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        body_type: str = "Text",
        save_to_sent_items: bool = True,
    ) -> None:
        """Send a new email."""
        message = {
            "subject": subject,
            "body": {"contentType": body_type, "content": body},
            "toRecipients": _recipients(to),
        }
        if cc:
            message["ccRecipients"] = _recipients(cc)
        if bcc:
            message["bccRecipients"] = _recipients(bcc)

        await self._request(
            "POST",
            "/me/sendMail",
            json={"message": message, "saveToSentItems": save_to_sent_items},
        )

    async def async_reply(
        self, message_id: str, comment: str, *, reply_all: bool = False
    ) -> None:
        """Reply (or reply-all) to a message."""
        action = "replyAll" if reply_all else "reply"
        await self._request(
            "POST", f"/me/messages/{message_id}/{action}", json={"comment": comment}
        )

    async def async_forward(
        self, message_id: str, to: list[str], comment: str = ""
    ) -> None:
        """Forward a message to one or more recipients."""
        await self._request(
            "POST",
            f"/me/messages/{message_id}/forward",
            json={"comment": comment, "toRecipients": _recipients(to)},
        )

    # ------------------------------------------------------------------
    # Organising mail
    # ------------------------------------------------------------------
    async def async_move_message(
        self, message_id: str, destination_id: str
    ) -> dict[str, Any]:
        """Move a message to another folder (id or well-known name)."""
        return await self._request(
            "POST",
            f"/me/messages/{message_id}/move",
            json={"destinationId": destination_id},
        )

    async def async_mark_read(self, message_id: str, is_read: bool = True) -> None:
        """Mark a message read or unread."""
        await self._request(
            "PATCH", f"/me/messages/{message_id}", json={"isRead": is_read}
        )

    async def async_set_categories(
        self, message_id: str, categories: list[str]
    ) -> None:
        """Set the categories on a message (replaces existing categories)."""
        await self._request(
            "PATCH", f"/me/messages/{message_id}", json={"categories": categories}
        )

    async def async_flag_message(
        self, message_id: str, flag_status: str = "flagged"
    ) -> None:
        """Set the follow-up flag on a message.

        ``flag_status`` is one of ``notFlagged``, ``flagged`` or ``complete``.
        """
        await self._request(
            "PATCH",
            f"/me/messages/{message_id}",
            json={"flag": {"flagStatus": flag_status}},
        )

    async def async_delete_message(self, message_id: str) -> None:
        """Delete a message (moves it to Deleted Items)."""
        await self._request("DELETE", f"/me/messages/{message_id}")

    # ------------------------------------------------------------------
    # Folders
    # ------------------------------------------------------------------
    async def async_list_folders(self) -> list[dict[str, Any]]:
        """List the user's mail folders."""
        data = await self._request(
            "GET", "/me/mailFolders", params={"$top": 100}
        )
        return data.get("value", []) if data else []

    async def async_create_folder(
        self, display_name: str, parent: str = "msgfolderroot"
    ) -> dict[str, Any]:
        """Create a mail folder under the given parent folder."""
        return await self._request(
            "POST",
            f"/me/mailFolders/{parent}/childFolders",
            json={"displayName": display_name},
        )

    # ------------------------------------------------------------------
    # Message (inbox) rules
    # ------------------------------------------------------------------
    async def async_list_rules(self) -> list[dict[str, Any]]:
        """List inbox message rules."""
        data = await self._request(
            "GET", "/me/mailFolders/inbox/messageRules"
        )
        return data.get("value", []) if data else []

    async def async_create_rule(
        self,
        *,
        display_name: str,
        sequence: int,
        conditions: dict[str, Any],
        actions: dict[str, Any],
        is_enabled: bool = True,
    ) -> dict[str, Any]:
        """Create an inbox message rule."""
        return await self._request(
            "POST",
            "/me/mailFolders/inbox/messageRules",
            json={
                "displayName": display_name,
                "sequence": sequence,
                "isEnabled": is_enabled,
                "conditions": conditions,
                "actions": actions,
            },
        )

    async def async_delete_rule(self, rule_id: str) -> None:
        """Delete an inbox message rule."""
        await self._request(
            "DELETE", f"/me/mailFolders/inbox/messageRules/{rule_id}"
        )


def _recipients(addresses: list[str]) -> list[dict[str, Any]]:
    """Convert a list of email addresses into Graph recipient objects."""
    return [{"emailAddress": {"address": addr}} for addr in addresses]
