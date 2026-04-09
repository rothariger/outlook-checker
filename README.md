# Checker Plus for Outlook

A Chrome extension (Manifest V3) inspired by **Checker Plus for Gmail**, but for **Outlook / Microsoft 365**.

![Chrome](https://img.shields.io/badge/Chrome-Extension-brightgreen?logo=googlechrome)
![Manifest V3](https://img.shields.io/badge/Manifest-V3-blue)
![Microsoft Graph](https://img.shields.io/badge/Microsoft-Graph%20API-0078d4?logo=microsoft)

---

## Features

- 🔔 **Badge counter** — total unread count across all accounts shown on the toolbar icon
- 📬 **Popup inbox** — emails grouped by account with subject, sender and relative time
- 📖 **Email preview** — read the full email body inline (HTML rendered in a sandboxed iframe)
- ✅ **Quick actions** — mark as read, archive, move to junk or delete without opening Outlook
- 👥 **Multi-account** — add as many Microsoft / Outlook accounts as you need
- 🔔 **Desktop notifications** — get notified when new emails arrive; click to open in Outlook
- 🔄 **Silent re-auth** — tokens refresh automatically; sign-in dialog only appears when needed
- ⚡ **Auto-polling** — checks for new email every minute via `chrome.alarms`

---

## Setup

### 1. Register an Azure AD Application (free, one-time)

> No credit card or subscription needed. Azure App Registration is a free feature available to any Microsoft account (Outlook, Hotmail, Live, or work/school).

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
   - **Name:** `Checker Plus for Outlook` (or any name you prefer)
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*
   - **Redirect URI:** leave blank for now
3. Click **Register** and copy the **Application (client) ID**
4. Go to **Authentication → Add a platform → Single-page application (SPA)**
   - Add the redirect URI: `https://<YOUR_EXTENSION_ID>.chromiumapp.org/`
   - *(You'll get the extension ID in step 3 below)*
5. Under **API permissions**, verify `User.Read` is present, then add:
   **Microsoft Graph → Delegated → Mail.ReadWrite**
6. For personal accounts, consent is automatic. For org accounts, an admin may need to click **Grant admin consent**.

---

### 2. Configure the extension

Copy the example config and add your client ID:

```bash
cp config.example.js config.js
```

Then edit `config.js`:

```js
export const CLIENT_ID = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'; // your Azure App client ID
```

---

### 3. Load the extension in Chrome

1. Open `chrome://extensions`
2. Enable **Developer mode** (top-right toggle)
3. Click **Load unpacked** and select the `outlook-checker` folder
4. Copy the **Extension ID** shown on the card (e.g. `abcdefghijklmnopabcdefghijklmnop`)
5. Go back to Azure Portal → Authentication and set the redirect URI:
   `https://<EXTENSION_ID>.chromiumapp.org/`

---

### 4. Add your first account

1. Click the extension icon in the toolbar
2. Click **+** or **Add your first account**
3. A Microsoft sign-in window opens — sign in and grant permissions
4. Your unread emails appear in the popup and the badge shows the count

Repeat for each additional account.

---

## File structure

```
outlook-checker/
├── manifest.json        # Extension manifest (Manifest V3)
├── config.example.js    # Config template — copy to config.js and fill in CLIENT_ID
├── config.js            # Your local config (git-ignored)
├── auth.js              # OAuth2 PKCE multi-account management
├── graph.js             # Microsoft Graph API calls
├── background.js        # Service worker: polling, badge updates, notifications
├── popup.html           # Popup UI markup
├── popup.css            # Popup styles
├── popup.js             # Popup logic (account list, email detail, actions)
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```

---

## Permissions

| Permission | Purpose |
|---|---|
| `storage` | Store account tokens and email cache |
| `alarms` | Poll for new emails every minute |
| `identity` | OAuth2 flow via `chrome.identity.launchWebAuthFlow` |
| `notifications` | Show desktop notifications for new emails |
| `https://graph.microsoft.com/*` | Call Microsoft Graph API |
| `https://login.microsoftonline.com/*` | Exchange auth codes for tokens |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `AADSTS50011: redirect URI mismatch` | The extension ID in Azure AD doesn't match `chrome://extensions` — update the redirect URI |
| Badge stays empty | Check `CLIENT_ID` in `config.js` is correct and permissions are granted |
| "Session expired" warning in popup | Click the warning link to re-authenticate |
| Extension reloaded and ID changed | Update the redirect URI in Azure Portal to the new extension ID |
| Notification click doesn't open email | Make sure the `tabs` permission is available (added automatically via `identity`) |

---

## License

MIT
