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

## Setup — step by step

This is a one-time setup. You'll bounce between **two browser tabs**:

- 🟦 **Azure tab** — where you register the app and get a Client ID + Secret.
- 🟧 **Home Assistant tab** — where you paste those values and log in.

> 💡 **Keep both tabs open the whole time.** The guide tells you exactly when to switch.
> Have a notepad handy to paste two values into: a **Client ID** and a **Client Secret**.

### Before you start: do you use a personal account?

If your mailbox is a **personal** account (`@outlook.com`, `@hotmail.com`, `@live.com`),
do this **one extra thing first**; work/school M365 users can skip to Part A.

1. Open a **new tab** and go to **<https://azure.microsoft.com/free>**.
2. Sign in with your personal Microsoft account and complete the free sign-up.
   - ✅ This creates a **free** Microsoft Entra (Azure AD) tenant with you as admin.
   - ✅ No credit-card charges and no paid Azure services are required.
   - ℹ️ Why: a bare personal account can't open **App registrations** (you'd hit error
     `AADSTS50020`). This step gives you a tenant where you can.
3. ⏎ **Come back here** and continue with Part A.

---

### Part A — Create the app registration (🟦 Azure tab)

**A1.** Open a **new tab** and go to **<https://portal.azure.com/>**. Sign in with the
Microsoft account you'll manage this from.

**A2.** In the top search bar, type **App registrations** and open it.
→ *Direct link:* <https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade>

**A3.** Click **+ New registration** (top-left).

**A4.** Fill in the form:
- **Name:** `Home Assistant Outlook` (anything you like).
- **Supported account types:** select
  **“Accounts in any organizational directory (Any Microsoft Entra ID tenant –
  Multitenant) and personal Microsoft accounts.”**
  ⚠️ This exact option is required — it's what lets your mailbox sign in.
- **Redirect URI:** change the dropdown to **Web**, and paste exactly:
  ```
  https://my.home-assistant.io/redirect/oauth
  ```

**A5.** Click **Register**.

**A6.** You're now on the app's **Overview** page. Find **Application (client) ID**, click
the **copy** icon, and **paste it into your notepad** as *Client ID*. ✅ (value #1 of 2)

**A7.** In the left menu click **Certificates & secrets** → tab **Client secrets** →
**+ New client secret**.
- **Description:** `home-assistant`
- **Expires:** pick the longest option (e.g. 24 months).
- Click **Add**.

**A8.** ⚠️ **Immediately copy the secret’s `Value` column** (NOT the “Secret ID”) and
**paste it into your notepad** as *Client Secret*. ✅ (value #2 of 2)
> The `Value` is only shown **once**. If you navigate away before copying it, just delete
> the secret and create a new one.

You now have both values. **Leave the Azure tab open** (in case you need to redo a secret)
and switch to Home Assistant.

---

### Part B — Add the integration in Home Assistant (🟧 Home Assistant tab)

**B0.** *(One-time prerequisite)* Your HA must be reachable over HTTPS via its **External
URL**, because Microsoft sends the login result back through `my.home-assistant.io`.
Check **Settings → System → Network → Home Assistant URL** has a working external
`https://…` address. If using Nabu Casa Cloud, this is already set.

**B1.** Go to **Settings → Devices & Services**.
→ *Or open:* <https://my.home-assistant.io/redirect/integrations/>

**B2.** Click **+ Add Integration** (bottom-right), search **Outlook M365**, and select it.

**B3.** The first time only, HA shows an **Add Application Credentials** dialog. Now use
your notepad:
- **Client ID:** paste *value #1*.
- **Client Secret:** paste *value #2*.
- Click **Create** / **Submit**.

**B4.** HA shows **“… wants to link your account.”** Click **Link account / Open**.
→ A Microsoft **sign-in page opens in a new tab**. *(No, it doesn't pop open by itself —
you trigger it with this button.)*

**B5.** 🟦 **In that Microsoft tab:** sign in with the **mailbox** you want to automate,
then on the **permissions** screen click **Accept** to consent to:
`offline_access`, `User.Read`, `Mail.ReadWrite`, `Mail.Send`, `MailboxSettings.ReadWrite`.

**B6.** Microsoft redirects back through `my.home-assistant.io` to your HA instance. ⏎
**Return to the Home Assistant tab** — you should see **“Success!”** and can assign the
integration to an Area.

**✅ Done.** A refresh token now keeps you signed in indefinitely — you won't see the login
page again unless you remove the integration or revoke access.

---

### Setup troubleshooting

| Symptom | Fix |
|---------|-----|
| `AADSTS50020` when opening App registrations | You're on a bare personal account — do the free Azure sign-up in *“Before you start”*. |
| `redirect_uri` mismatch after login | The Azure **Redirect URI** must be exactly `https://my.home-assistant.io/redirect/oauth` under a **Web** platform (Part A, step A4). |
| “No URL available” / can't reach HA after login | Set a working **External URL** (step B0); Azure needs an HTTPS redirect. |
| Lost the client secret value | Azure tab → Certificates & secrets → delete it and create a new one (A7–A8), then re-add in HA. |
| Asks you to log in again later | Re-link the account; the flow requests `offline_access` + `prompt=consent` so a refresh token is issued. |

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
