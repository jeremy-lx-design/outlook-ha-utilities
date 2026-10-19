"""Constants for the Outlook M365 integration."""

from __future__ import annotations

from datetime import timedelta

DOMAIN = "outlook_m365"

# Microsoft identity platform OAuth2 endpoints (multi-tenant "common").
OAUTH2_AUTHORIZE = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
OAUTH2_TOKEN = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

# Microsoft Graph API.
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Delegated scopes requested during the OAuth2 flow.
#   User.Read                 - read the signed-in user's profile.
#   Mail.ReadWrite            - read/organise mail (move, flag, categorise, delete).
#   Mail.Send                 - send mail on behalf of the user.
#   MailboxSettings.ReadWrite - create and manage inbox (message) rules.
#   offline_access            - obtain a refresh token for long-lived access.
SCOPES = [
    "offline_access",
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "MailboxSettings.ReadWrite",
]

# How often the coordinator polls the inbox for new messages.
DEFAULT_SCAN_INTERVAL = timedelta(seconds=60)

# Bus event fired when a new email arrives. Use it in automation triggers:
#   trigger:
#     - platform: event
#       event_type: outlook_m365_new_email
EVENT_NEW_EMAIL = f"{DOMAIN}_new_email"

# Config entry runtime data keys.
DATA_COORDINATOR = "coordinator"
DATA_API = "api"

# Max messages to pull per poll.
MAX_MESSAGES_PER_POLL = 25
