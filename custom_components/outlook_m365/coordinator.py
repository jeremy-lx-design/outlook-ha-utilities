"""Data update coordinator for the Outlook M365 integration.

Polls the inbox on a fixed interval, fires a Home Assistant bus event for each
newly received message (so automations can trigger on them), and exposes the
latest unread count / most recent message for sensors.
"""

from __future__ import annotations

from datetime import datetime, timezone
import logging
from typing import Any

from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.exceptions import ConfigEntryAuthFailed
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed

from .api import OutlookApiClient, OutlookApiError, OutlookAuthError
from .const import (
    DEFAULT_SCAN_INTERVAL,
    DOMAIN,
    EVENT_NEW_EMAIL,
    MAX_MESSAGES_PER_POLL,
)

_LOGGER = logging.getLogger(__name__)

# Fields we ask Graph to return for each polled message.
_MESSAGE_SELECT = [
    "id",
    "subject",
    "bodyPreview",
    "from",
    "toRecipients",
    "receivedDateTime",
    "isRead",
    "hasAttachments",
    "importance",
    "webLink",
    "conversationId",
]


class OutlookDataUpdateCoordinator(DataUpdateCoordinator[dict[str, Any]]):
    """Coordinate polling of the Outlook mailbox."""

    def __init__(
        self,
        hass: HomeAssistant,
        entry: ConfigEntry,
        api: OutlookApiClient,
    ) -> None:
        """Initialise the coordinator."""
        super().__init__(
            hass,
            _LOGGER,
            name=DOMAIN,
            update_interval=DEFAULT_SCAN_INTERVAL,
        )
        self.entry = entry
        self.api = api
        # High-water mark for the most recent message we've already seen.
        self._last_received: datetime | None = None
        # Guard so the very first poll doesn't replay the whole inbox as "new".
        self._primed = False

    async def _async_update_data(self) -> dict[str, Any]:
        """Fetch new messages and the unread count."""
        try:
            messages = await self.api.async_list_messages(
                folder="inbox",
                top=MAX_MESSAGES_PER_POLL,
                select=_MESSAGE_SELECT,
                orderby="receivedDateTime desc",
            )
            unread_count = await self.api.async_get_unread_count("inbox")
        except OutlookAuthError as err:
            raise ConfigEntryAuthFailed(str(err)) from err
        except OutlookApiError as err:
            raise UpdateFailed(str(err)) from err

        new_messages = self._detect_new_messages(messages)

        if self._primed:
            for message in new_messages:
                self._fire_new_email_event(message)
        else:
            # First successful poll: record the baseline, don't fire events.
            self._primed = True

        latest = messages[0] if messages else None
        return {
            "unread_count": unread_count,
            "latest_message": latest,
            "new_messages": new_messages,
            "message_count": len(messages),
        }

    def _detect_new_messages(
        self, messages: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        """Return messages received after the last high-water mark."""
        new_messages: list[dict[str, Any]] = []
        highest = self._last_received

        for message in messages:
            received = _parse_dt(message.get("receivedDateTime"))
            if received is None:
                continue
            if self._last_received is None or received > self._last_received:
                new_messages.append(message)
            if highest is None or received > highest:
                highest = received

        self._last_received = highest
        # Oldest-first so automations fire in arrival order.
        new_messages.reverse()
        return new_messages

    def _fire_new_email_event(self, message: dict[str, Any]) -> None:
        """Fire a bus event describing a newly received email."""
        sender = (
            message.get("from", {})
            .get("emailAddress", {})
            if message.get("from")
            else {}
        )
        event_data = {
            "entry_id": self.entry.entry_id,
            "message_id": message.get("id"),
            "subject": message.get("subject"),
            "preview": message.get("bodyPreview"),
            "from_name": sender.get("name"),
            "from_address": sender.get("address"),
            "received": message.get("receivedDateTime"),
            "has_attachments": message.get("hasAttachments", False),
            "importance": message.get("importance"),
            "web_link": message.get("webLink"),
            "conversation_id": message.get("conversationId"),
        }
        _LOGGER.debug("Firing %s for message %s", EVENT_NEW_EMAIL, event_data["message_id"])
        self.hass.bus.async_fire(EVENT_NEW_EMAIL, event_data)


def _parse_dt(value: str | None) -> datetime | None:
    """Parse a Graph ISO-8601 timestamp into an aware datetime."""
    if not value:
        return None
    try:
        # Graph returns e.g. "2024-01-01T12:00:00Z".
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt
