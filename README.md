# Checker Plus for Outlook

A Chrome extension (Manifest V3) similar to **Checker Plus for Gmail** but for **Outlook / Microsoft 365**.

- 🔔 **Badge** on the toolbar icon with the total unread count across all accounts
- 📬 **Popup** with emails grouped by account (subject, sender, time)
- 👥 **Multi-account** — add as many Microsoft/Outlook accounts as you need
- 🔄 **Silent re-auth** — if you're already signed into Outlook Web, tokens refresh without any dialog
- ⚡ **Auto-polling** every minute via `chrome.alarms`

---

## 1. Register an Azure AD Application (one-time setup — completely free ✅)

> **No credit card, no subscription, no cost.** The Azure App Registration is a free feature available to any Microsoft account — the same account you already use for Outlook/Hotmail/Live. You only need to do this once.


1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
   - Name: `Checker Plus for Outlook` (or any name)
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: **Single-page application (SPA)** — leave blank for now
3. Click **Register**
4. Copy the **Application (client) ID** — you'll need it in the next step
5. Go to **Authentication → Add a platform → Single-page application**
   - Redirect URI: `https://<YOUR_EXTENSION_ID>.chromiumapp.org/`
   - *(You'll find the extension ID after loading the extension in Chrome)*
6. Under **API permissions**, confirm `User.Read` is present, then **Add a permission → Microsoft Graph → Delegated → Mail.Read**
7. Click **Grant admin consent** if required by your organization (for personal accounts it's automatic)

---

## 2. Configure the extension

Open `config.js` and replace the placeholder with your client ID:

```js
export const CLIENT_ID = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx';
```

---

## 3. Load the extension in Chrome

1. Open Chrome and navigate to `chrome://extensions`
2. Enable **Developer mode** (toggle in the top-right)
3. Click **Load unpacked** and select the `outlook-checker` folder
4. Note the **Extension ID** shown on the card (e.g. `abcdefghijklmnopabcdefghijklmnop`)
5. Go back to Azure Portal and add the redirect URI:
   `https://<EXTENSION_ID>.chromiumapp.org/`

---

## 4. Add your first account

1. Click the extension icon in the toolbar
2. Click the **+** button (or **Add your first account**)
3. A Microsoft sign-in window opens — sign in and grant permissions
4. The popup will now show your emails and the badge will display your unread count

Repeat step 2 for each additional account.

---

## File structure

```
outlook-checker/
├── manifest.json      # Extension manifest (Manifest V3)
├── config.js          # Azure AD client_id and constants
├── auth.js            # OAuth2 multi-account management
├── graph.js           # Microsoft Graph API calls
├── background.js      # Service worker: polling + badge
├── popup.html         # Popup UI
├── popup.css          # Popup styles
├── popup.js           # Popup logic
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```

---

## Permissions used

| Permission | Reason |
|---|---|
| `storage` | Store account tokens and email cache |
| `alarms` | Poll emails every minute |
| `identity` | OAuth2 flow via `chrome.identity.launchWebAuthFlow` |
| `notifications` | (Reserved for future desktop notifications) |
| `https://graph.microsoft.com/*` | Call Microsoft Graph API |
| `https://login.microsoftonline.com/*` | Exchange auth codes for tokens |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "AADSTS50011: redirect URI mismatch" | Make sure the extension ID in Azure AD matches the one in `chrome://extensions` |
| Badge stays empty | Check that `CLIENT_ID` in `config.js` is set correctly |
| "Session expired" warning | Click the warning link to re-authenticate |
| Extension reloaded / ID changed | Update the redirect URI in Azure Portal and reload |
