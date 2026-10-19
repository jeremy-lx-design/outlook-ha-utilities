"""Service handlers for the Outlook M365 integration.

Registers Home Assistant services that let automations act on mail: send,
reply, forward, move, mark read/unread, flag, categorise, delete, create
folders and create/delete inbox rules.
"""

from __future__ import annotations

import logging
from typing import Any

import voluptuous as vol

from homeassistant.core import HomeAssistant, ServiceCall, ServiceResponse, SupportsResponse
from homeassistant.exceptions import HomeAssistantError, ServiceValidationError
from homeassistant.helpers import config_validation as cv

from .api import OutlookApiClient, OutlookApiError
from .const import DATA_API, DOMAIN

_LOGGER = logging.getLogger(__name__)

# Service names.
SERVICE_SEND_EMAIL = "send_email"
SERVICE_REPLY_EMAIL = "reply_email"
SERVICE_FORWARD_EMAIL = "forward_email"
SERVICE_MOVE_EMAIL = "move_email"
SERVICE_MARK_EMAIL = "mark_email"
SERVICE_FLAG_EMAIL = "flag_email"
SERVICE_CATEGORIZE_EMAIL = "categorize_email"
SERVICE_DELETE_EMAIL = "delete_email"
SERVICE_CREATE_FOLDER = "create_folder"
SERVICE_CREATE_RULE = "create_rule"
SERVICE_DELETE_RULE = "delete_rule"
SERVICE_LIST_FOLDERS = "list_folders"
SERVICE_LIST_RULES = "list_rules"

ATTR_CONFIG_ENTRY_ID = "config_entry_id"

# Shared validator for a comma-or-list of email addresses.
_ADDRESSES = vol.All(cv.ensure_list, [cv.string])

SEND_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("to"): _ADDRESSES,
        vol.Optional("cc"): _ADDRESSES,
        vol.Optional("bcc"): _ADDRESSES,
        vol.Required("subject"): cv.string,
        vol.Required("message"): cv.string,
        vol.Optional("body_type", default="Text"): vol.In(["Text", "HTML"]),
        vol.Optional("save_to_sent_items", default=True): cv.boolean,
    }
)

REPLY_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
        vol.Optional("comment", default=""): cv.string,
        vol.Optional("reply_all", default=False): cv.boolean,
    }
)

FORWARD_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
        vol.Required("to"): _ADDRESSES,
        vol.Optional("comment", default=""): cv.string,
    }
)

MOVE_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
        vol.Required("destination"): cv.string,
    }
)

MARK_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
        vol.Optional("read", default=True): cv.boolean,
    }
)

FLAG_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
        vol.Optional("flag_status", default="flagged"): vol.In(
            ["notFlagged", "flagged", "complete"]
        ),
    }
)

CATEGORIZE_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
        vol.Required("categories"): vol.All(cv.ensure_list, [cv.string]),
    }
)

DELETE_EMAIL_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("message_id"): cv.string,
    }
)

CREATE_FOLDER_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("name"): cv.string,
        vol.Optional("parent", default="msgfolderroot"): cv.string,
    }
)

CREATE_RULE_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("name"): cv.string,
        vol.Optional("sequence", default=1): cv.positive_int,
        vol.Optional("is_enabled", default=True): cv.boolean,
        # Conditions
        vol.Optional("from_addresses"): _ADDRESSES,
        vol.Optional("subject_contains"): vol.All(cv.ensure_list, [cv.string]),
        vol.Optional("body_contains"): vol.All(cv.ensure_list, [cv.string]),
        vol.Optional("importance"): vol.In(["low", "normal", "high"]),
        # Actions
        vol.Optional("move_to_folder"): cv.string,
        vol.Optional("copy_to_folder"): cv.string,
        vol.Optional("forward_to"): _ADDRESSES,
        vol.Optional("mark_as_read", default=False): cv.boolean,
        vol.Optional("mark_importance"): vol.In(["low", "normal", "high"]),
        vol.Optional("assign_categories"): vol.All(cv.ensure_list, [cv.string]),
        vol.Optional("delete", default=False): cv.boolean,
        vol.Optional("stop_processing_rules", default=False): cv.boolean,
    }
)

DELETE_RULE_SCHEMA = vol.Schema(
    {
        vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string,
        vol.Required("rule_id"): cv.string,
    }
)

LIST_SCHEMA = vol.Schema({vol.Optional(ATTR_CONFIG_ENTRY_ID): cv.string})


def _get_api(hass: HomeAssistant, call: ServiceCall) -> OutlookApiClient:
    """Resolve the API client for the targeted (or only) config entry."""
    entries: dict[str, dict[str, Any]] = hass.data.get(DOMAIN, {})
    if not entries:
        raise HomeAssistantError("No Outlook M365 account is configured")

    entry_id = call.data.get(ATTR_CONFIG_ENTRY_ID)
    if entry_id is None:
        if len(entries) > 1:
            raise ServiceValidationError(
                "Multiple Outlook M365 accounts configured; "
                "specify 'config_entry_id'"
            )
        entry_id = next(iter(entries))

    entry_data = entries.get(entry_id)
    if entry_data is None:
        raise ServiceValidationError(f"Unknown config_entry_id: {entry_id}")
    return entry_data[DATA_API]


def _build_rule_conditions(data: dict[str, Any]) -> dict[str, Any]:
    """Translate flat service data into Graph rule conditions."""
    conditions: dict[str, Any] = {}
    if data.get("from_addresses"):
        conditions["senderContains"] = data["from_addresses"]
    if data.get("subject_contains"):
        conditions["subjectContains"] = data["subject_contains"]
    if data.get("body_contains"):
        conditions["bodyContains"] = data["body_contains"]
    if data.get("importance"):
        conditions["importance"] = data["importance"]
    return conditions


def _build_rule_actions(data: dict[str, Any]) -> dict[str, Any]:
    """Translate flat service data into Graph rule actions."""
    actions: dict[str, Any] = {}
    if data.get("move_to_folder"):
        actions["moveToFolder"] = data["move_to_folder"]
    if data.get("copy_to_folder"):
        actions["copyToFolder"] = data["copy_to_folder"]
    if data.get("forward_to"):
        actions["forwardTo"] = [
            {"emailAddress": {"address": addr}} for addr in data["forward_to"]
        ]
    if data.get("mark_as_read"):
        actions["markAsRead"] = True
    if data.get("mark_importance"):
        actions["markImportance"] = data["mark_importance"]
    if data.get("assign_categories"):
        actions["assignCategories"] = data["assign_categories"]
    if data.get("delete"):
        actions["delete"] = True
    if data.get("stop_processing_rules"):
        actions["stopProcessingRules"] = True
    return actions


def async_setup_services(hass: HomeAssistant) -> None:
    """Register all Outlook M365 services (idempotent)."""
    if hass.services.has_service(DOMAIN, SERVICE_SEND_EMAIL):
        return

    # --- Mail actions -------------------------------------------------
    async def handle_send_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_send_mail(
                subject=call.data["subject"],
                body=call.data["message"],
                to=call.data["to"],
                cc=call.data.get("cc"),
                bcc=call.data.get("bcc"),
                body_type=call.data["body_type"],
                save_to_sent_items=call.data["save_to_sent_items"],
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to send email: {err}") from err

    async def handle_reply_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_reply(
                call.data["message_id"],
                call.data["comment"],
                reply_all=call.data["reply_all"],
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to reply: {err}") from err

    async def handle_forward_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_forward(
                call.data["message_id"], call.data["to"], call.data["comment"]
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to forward: {err}") from err

    async def handle_move_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_move_message(
                call.data["message_id"], call.data["destination"]
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to move email: {err}") from err

    async def handle_mark_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_mark_read(call.data["message_id"], call.data["read"])
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to mark email: {err}") from err

    async def handle_flag_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_flag_message(
                call.data["message_id"], call.data["flag_status"]
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to flag email: {err}") from err

    async def handle_categorize_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_set_categories(
                call.data["message_id"], call.data["categories"]
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to categorise email: {err}") from err

    async def handle_delete_email(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_delete_message(call.data["message_id"])
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to delete email: {err}") from err

    async def handle_create_folder(call: ServiceCall) -> ServiceResponse:
        api = _get_api(hass, call)
        try:
            folder = await api.async_create_folder(
                call.data["name"], call.data["parent"]
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to create folder: {err}") from err
        return {"id": folder.get("id"), "display_name": folder.get("displayName")}

    async def handle_create_rule(call: ServiceCall) -> ServiceResponse:
        api = _get_api(hass, call)
        conditions = _build_rule_conditions(call.data)
        actions = _build_rule_actions(call.data)
        if not actions:
            raise ServiceValidationError(
                "A rule must define at least one action"
            )
        try:
            rule = await api.async_create_rule(
                display_name=call.data["name"],
                sequence=call.data["sequence"],
                conditions=conditions,
                actions=actions,
                is_enabled=call.data["is_enabled"],
            )
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to create rule: {err}") from err
        return {"id": rule.get("id"), "display_name": rule.get("displayName")}

    async def handle_delete_rule(call: ServiceCall) -> None:
        api = _get_api(hass, call)
        try:
            await api.async_delete_rule(call.data["rule_id"])
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to delete rule: {err}") from err

    async def handle_list_folders(call: ServiceCall) -> ServiceResponse:
        api = _get_api(hass, call)
        try:
            folders = await api.async_list_folders()
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to list folders: {err}") from err
        return {
            "folders": [
                {
                    "id": f.get("id"),
                    "display_name": f.get("displayName"),
                    "unread": f.get("unreadItemCount"),
                    "total": f.get("totalItemCount"),
                }
                for f in folders
            ]
        }

    async def handle_list_rules(call: ServiceCall) -> ServiceResponse:
        api = _get_api(hass, call)
        try:
            rules = await api.async_list_rules()
        except OutlookApiError as err:
            raise HomeAssistantError(f"Failed to list rules: {err}") from err
        return {
            "rules": [
                {
                    "id": r.get("id"),
                    "display_name": r.get("displayName"),
                    "sequence": r.get("sequence"),
                    "is_enabled": r.get("isEnabled"),
                }
                for r in rules
            ]
        }

    register = hass.services.async_register

    register(DOMAIN, SERVICE_SEND_EMAIL, handle_send_email, schema=SEND_EMAIL_SCHEMA)
    register(DOMAIN, SERVICE_REPLY_EMAIL, handle_reply_email, schema=REPLY_EMAIL_SCHEMA)
    register(
        DOMAIN, SERVICE_FORWARD_EMAIL, handle_forward_email, schema=FORWARD_EMAIL_SCHEMA
    )
    register(DOMAIN, SERVICE_MOVE_EMAIL, handle_move_email, schema=MOVE_EMAIL_SCHEMA)
    register(DOMAIN, SERVICE_MARK_EMAIL, handle_mark_email, schema=MARK_EMAIL_SCHEMA)
    register(DOMAIN, SERVICE_FLAG_EMAIL, handle_flag_email, schema=FLAG_EMAIL_SCHEMA)
    register(
        DOMAIN,
        SERVICE_CATEGORIZE_EMAIL,
        handle_categorize_email,
        schema=CATEGORIZE_EMAIL_SCHEMA,
    )
    register(
        DOMAIN, SERVICE_DELETE_EMAIL, handle_delete_email, schema=DELETE_EMAIL_SCHEMA
    )
    register(
        DOMAIN,
        SERVICE_CREATE_FOLDER,
        handle_create_folder,
        schema=CREATE_FOLDER_SCHEMA,
        supports_response=SupportsResponse.OPTIONAL,
    )
    register(
        DOMAIN,
        SERVICE_CREATE_RULE,
        handle_create_rule,
        schema=CREATE_RULE_SCHEMA,
        supports_response=SupportsResponse.OPTIONAL,
    )
    register(DOMAIN, SERVICE_DELETE_RULE, handle_delete_rule, schema=DELETE_RULE_SCHEMA)
    register(
        DOMAIN,
        SERVICE_LIST_FOLDERS,
        handle_list_folders,
        schema=LIST_SCHEMA,
        supports_response=SupportsResponse.ONLY,
    )
    register(
        DOMAIN,
        SERVICE_LIST_RULES,
        handle_list_rules,
        schema=LIST_SCHEMA,
        supports_response=SupportsResponse.ONLY,
    )


def async_unload_services(hass: HomeAssistant) -> None:
    """Remove all Outlook M365 services."""
    for service in (
        SERVICE_SEND_EMAIL,
        SERVICE_REPLY_EMAIL,
        SERVICE_FORWARD_EMAIL,
        SERVICE_MOVE_EMAIL,
        SERVICE_MARK_EMAIL,
        SERVICE_FLAG_EMAIL,
        SERVICE_CATEGORIZE_EMAIL,
        SERVICE_DELETE_EMAIL,
        SERVICE_CREATE_FOLDER,
        SERVICE_CREATE_RULE,
        SERVICE_DELETE_RULE,
        SERVICE_LIST_FOLDERS,
        SERVICE_LIST_RULES,
    ):
        hass.services.async_remove(DOMAIN, service)
