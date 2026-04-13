# Privacy Policy — Outlook Checker

_Last updated: April 9, 2026_

## Overview

Outlook Checker is a Chrome browser extension that allows users to monitor their Microsoft 365 and Outlook email accounts from the Chrome toolbar. This policy explains what data is accessed, how it is used, and what is never done with it.

---

## Data collected and how it is used

### Authentication tokens
When you sign in to a Microsoft account, the extension receives an OAuth2 access token and refresh token issued by Microsoft. These tokens are stored locally in `chrome.storage.local` on your device and are used exclusively to make authenticated requests to the Microsoft Graph API on your behalf.

### Email data
The extension fetches a list of recent unread emails (subject, sender, received date, and message body) via the Microsoft Graph API. This data is:
- Displayed in the extension popup for your personal use
- Cached temporarily in `chrome.storage.local` to improve popup load time
- Never transmitted to any server other than Microsoft's own APIs

### Account email address and display name
Retrieved from Microsoft Graph (`User.Read` scope) to identify accounts in the popup UI.

---

## Data NOT collected

The extension does **not** collect, store, or transmit:

- Passwords or security credentials
- Web browsing history
- Location data
- Any analytics, telemetry, or usage statistics
- Any data to servers operated by this extension's developer

---

## Data sharing

User data is **never**:
- Sold or transferred to third parties
- Used for advertising or marketing purposes
- Used for purposes unrelated to reading and managing your email
- Used to determine creditworthiness or for lending purposes

All communication happens directly between your browser and Microsoft's servers (Graph API and login endpoints). The extension developer has no server and receives no data.

---

## Data storage and security

All tokens and cached email data are stored locally on your device using the Chrome extension storage API (`chrome.storage.local`). This data is accessible only to the extension itself and is not synced across devices.

You can remove all stored data at any time by removing the extension from Chrome (`chrome://extensions`).

---

## Third-party services

This extension communicates exclusively with:

- **Microsoft Graph API** (`https://graph.microsoft.com`) — to read and manage emails
- **Microsoft Identity Platform** (`https://login.microsoftonline.com`) — to authenticate users via OAuth2

Please refer to [Microsoft's Privacy Policy](https://privacy.microsoft.com/en-us/privacystatement) for details on how Microsoft handles your data.

---

## Changes to this policy

If this policy is updated, the _Last updated_ date at the top will be revised. Continued use of the extension after changes constitutes acceptance of the updated policy.

---

## Contact

For questions or concerns about this privacy policy, please open an issue at:
[https://github.com/rothariger/outlook-checker/issues](https://github.com/rothariger/outlook-checker/issues)
