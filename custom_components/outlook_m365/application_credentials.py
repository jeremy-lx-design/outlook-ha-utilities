"""Application credentials platform for Outlook M365.

This lets users add their Azure AD app registration (client id / secret)
through Home Assistant's *Application Credentials* UI, which Home Assistant
then uses to drive the OAuth2 authorization code flow.
"""

from __future__ import annotations

from homeassistant.components.application_credentials import AuthorizationServer
from homeassistant.core import HomeAssistant

from .const import OAUTH2_AUTHORIZE, OAUTH2_TOKEN


async def async_get_authorization_server(hass: HomeAssistant) -> AuthorizationServer:
    """Return the Microsoft identity platform authorization server."""
    return AuthorizationServer(
        authorize_url=OAUTH2_AUTHORIZE,
        token_url=OAUTH2_TOKEN,
    )


async def async_get_description_placeholders(hass: HomeAssistant) -> dict[str, str]:
    """Return description placeholders for the credentials dialog."""
    return {
        "oauth_consent_url": "https://portal.azure.com/",
        "redirect_url": "https://my.home-assistant.io/redirect/oauth",
        "more_info_url": "https://github.com/jeremy/outlook-ha-utilities",
    }
