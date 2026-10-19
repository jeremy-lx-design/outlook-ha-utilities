# Outlook M365 for Home Assistant

A custom [Home Assistant](https://www.home-assistant.io/) integration that connects your
**Microsoft 365 / Outlook** mailbox to Home Assistant via the
[Microsoft Graph API](https://learn.microsoft.com/graph/overview), so you can run
automations on your email.

Built-in Outlook rules are limited. This integration lets your always-on Home Assistant
box do the heavy lifting: trigger automations when mail arrives, and send, reply,
forward, move, flag, categorise, delete, and create server-side rules — all from HA
automations and scripts.

## Features

- 🔔 **Trigger automations on incoming email** via the `outlook_m365_new_email` event.
- ✉️ **Send / reply / forward** email from automations.
- 🗂️ **Organise mail**: move to folders, mark read/unread, flag, categorise, delete.
- 📏 **Manage server-side inbox rules**: create and delete message rules.
- 📁 **Folders & rules helpers**: create folders, list folders/rules (with ids).
- 📊 **Sensors**: unread count and latest email (with rich attributes).

## Requirements

- Home Assistant 2024.4 or newer.
- A Microsoft 365 / Outlook.com account.
- An **Azure app registration** (free) to obtain OAuth2 client credentials.
- Your HA instance reachable via its configured **External URL** (needed for the OAuth
  redirect).

## Installation

### HACS (recommended)

1. In HACS → Integrations → ⋮ → **Custom repositories**, add this repository as an
   *Integration*.
2. Install **Outlook M365** and restart Home Assistant.

### Manual

Copy `custom_components/outlook_m365` into your Home Assistant `config/custom_components`
directory and restart.

## Azure app registration

You need a client ID and secret from Azure. This is a one-time setup.

> **Using a personal account (@outlook.com / @hotmail.com / @live.com)?**
> Your mailbox account and the app registration are *separate things*. A personal
> Microsoft account has no client ID/secret and, by default, cannot open **App
> registrations** (you'll hit error `AADSTS50020`). To fix this, sign in once at
> [azure.microsoft.com/free](https://azure.microsoft.com/free) with that same personal
> account — this **creates a free Microsoft Entra (Azure AD) tenant** with you as admin at
> **no cost** (no credit-card charges, no paid Azure services required). You then create
> the app registration inside that free tenant. In step 3 below, make sure you select the
> option that **includes personal Microsoft accounts** so your Outlook.com mailbox can
> sign in and consent.

1. Go to the [Azure Portal](https://portal.azure.com/) → **Microsoft Entra ID** →
   **App registrations** → **New registration**.
2. **Name**: e.g. `Home Assistant Outlook`.
3. **Supported account types**: choose *Accounts in any organizational directory and
   personal Microsoft accounts* (this matches the `common` endpoint the integration uses).
4. **Redirect URI**: platform **Web**, value:
   ```
   https://my.home-assistant.io/redirect/oauth
   ```
5. Click **Register** and copy the **Application (client) ID**.
6. Go to **Certificates & secrets → New client secret**, create one, and copy its
   **Value** (not the Secret ID).
7. (Optional) Under **API permissions** you can review the delegated Microsoft Graph
   permissions; they are requested automatically during login:
   `offline_access`, `User.Read`, `Mail.ReadWrite`, `Mail.Send`,
   `MailboxSettings.ReadWrite`.

## Login / configuration

> **Does it pop open a Microsoft login page?** Not automatically — during setup you click
> a **Link account** button, which opens Microsoft's sign-in and consent page in your
> browser (standard OAuth2 Authorization Code flow). After the first login, a refresh
> token keeps you signed in; no repeated pop-ups.

1. In Home Assistant: **Settings → Devices & Services → Add Integration → Outlook M365**.
2. The first time, HA asks for **Application Credentials** — paste the **client ID** and
   **client secret** from Azure.
3. Click **Link account**. The Microsoft login page opens in your browser.
4. Sign in and **consent** to the requested permissions.
5. Microsoft redirects through `my.home-assistant.io` back to your HA instance, which
   exchanges the code for tokens. Done.

> **External URL required.** The `my.home-assistant.io` redirect bounces the login back to
> your HA instance using your **External URL** (Settings → System → Network). Azure
> requires HTTPS redirect URIs (except `localhost`), which is why the relay is used.

## Triggering automations on incoming email

Each new message fires an `outlook_m365_new_email` event on the HA bus:

```yaml
automation:
  - alias: "Notify on invoice emails"
    trigger:
      - platform: event
        event_type: outlook_m365_new_email
    condition:
      - condition: template
        value_template: "{{ 'invoice' in (trigger.event.data.subject | lower) }}"
    action:
      - service: notify.mobile_app_phone
        data:
          title: "New invoice from {{ trigger.event.data.from_name }}"
          message: "{{ trigger.event.data.subject }}"
      # Auto-file it
      - service: outlook_m365.move_email
        data:
          message_id: "{{ trigger.event.data.message_id }}"
          destination: archive
```

### Event data

| Field | Description |
|-------|-------------|
| `message_id` | Graph message id (use it in services) |
| `subject` | Email subject |
| `preview` | Short body preview |
| `from_name` / `from_address` | Sender |
| `received` | ISO-8601 received timestamp |
| `has_attachments` | Boolean |
| `importance` | `low` / `normal` / `high` |
| `web_link` | Open-in-Outlook URL |
| `conversation_id` | Thread id |
| `entry_id` | Which configured account received it |

## Services

| Service | Purpose |
|---------|---------|
| `outlook_m365.send_email` | Send a new email |
| `outlook_m365.reply_email` | Reply / reply-all to a message |
| `outlook_m365.forward_email` | Forward a message |
| `outlook_m365.move_email` | Move a message to a folder |
| `outlook_m365.mark_email` | Mark read / unread |
| `outlook_m365.flag_email` | Set follow-up flag |
| `outlook_m365.categorize_email` | Assign categories |
| `outlook_m365.delete_email` | Delete (to Deleted Items) |
| `outlook_m365.create_folder` | Create a mail folder |
| `outlook_m365.create_rule` | Create a server-side inbox rule |
| `outlook_m365.delete_rule` | Delete an inbox rule |
| `outlook_m365.list_folders` | List folders (returns ids) |
| `outlook_m365.list_rules` | List inbox rules (returns ids) |

### Examples

Send an email:

```yaml
service: outlook_m365.send_email
data:
  to: ["someone@example.com"]
  subject: "Garage door left open"
  message: "The garage has been open for 20 minutes."
```

Create a server-side rule (so it runs even when HA is offline):

```yaml
service: outlook_m365.create_rule
data:
  name: "File newsletters"
  from_addresses: ["newsletter@"]
  move_to_folder: archive
  mark_as_read: true
response_variable: rule
```

List folders to discover folder ids:

```yaml
service: outlook_m365.list_folders
response_variable: result
# result.folders -> [{ id, display_name, unread, total }, ...]
```

Well-known folder names you can use for `destination` / `move_to_folder`:
`inbox`, `archive`, `drafts`, `sentitems`, `deleteditems`, `junkemail`.

## How it works

- Uses Home Assistant's **Application Credentials** + **OAuth2** helpers for login and
  automatic token refresh.
- A **DataUpdateCoordinator** polls the inbox every 60 seconds, firing the new-email event
  for messages received since the last poll (the first poll primes the baseline and does
  not replay existing mail).
- Mail actions call Microsoft Graph `v1.0` endpoints directly.

## Troubleshooting

- **No URL available / redirect errors**: ensure your HA **External URL** is set and
  reachable over HTTPS, and that the Azure redirect URI is exactly
  `https://my.home-assistant.io/redirect/oauth`.
- **Login loops / no refresh token**: the flow requests `offline_access` and
  `prompt=consent` so a refresh token is always issued; re-link the account if it was set
  up before consenting.
- **403 from Graph**: confirm the account consented to all delegated scopes.

## Disclaimer

This is a community integration and is not affiliated with or endorsed by Microsoft.

## License

[MIT](LICENSE)
