import { getAccounts, getValidToken, markNeedsReauth } from './auth.js';
import { getUnreadCount, getRecentEmails } from './graph.js';
import { POLL_INTERVAL_MINUTES } from './config.js';

// ── Setup ─────────────────────────────────────────────────────────────────────

chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create('pollEmails', { periodInMinutes: POLL_INTERVAL_MINUTES });
  pollAllAccounts();
});

// Re-register alarm after browser restart (service worker doesn't persist)
chrome.runtime.onStartup.addListener(() => {
  chrome.alarms.get('pollEmails', alarm => {
    if (!alarm) chrome.alarms.create('pollEmails', { periodInMinutes: POLL_INTERVAL_MINUTES });
  });
  pollAllAccounts();
});

chrome.alarms.onAlarm.addListener(alarm => {
  if (alarm.name === 'pollEmails') pollAllAccounts();
});

// ── Message handling from popup ───────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type === 'forceRefresh') {
    pollAllAccounts().then(() => sendResponse({ ok: true })).catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }
  if (msg.type === 'accountAdded') {
    // Seed known IDs for the new account so first poll doesn't spam notifications
    pollAllAccounts({ seedOnly: true }).then(() => sendResponse({ ok: true })).catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }
});

// ── Notification click → open email in Outlook ────────────────────────────────

chrome.notifications.onClicked.addListener(notifId => {
  // notifId format: "outlook-<accountEmail>-<messageId>"
  chrome.storage.local.get('notifLinks', ({ notifLinks = {} }) => {
    const url = notifLinks[notifId];
    if (url) chrome.tabs.create({ url });
    chrome.notifications.clear(notifId);
  });
});

// ── Core polling ──────────────────────────────────────────────────────────────

async function pollAllAccounts({ seedOnly = false } = {}) {
  const accounts = await getAccounts();

  if (accounts.length === 0) {
    await chrome.action.setBadgeText({ text: '' });
    return;
  }

  const { knownIds = {}, notifLinks = {} } = await chrome.storage.local.get(['knownIds', 'notifLinks']);

  const cache = {};
  let totalUnread = 0;

  await Promise.allSettled(
    accounts.map(async account => {
      try {
        const token = await getValidToken(account);
        const [unreadCount, emails] = await Promise.all([
          getUnreadCount(token),
          getRecentEmails(token),
        ]);
        totalUnread += unreadCount;
        cache[account.email] = { unreadCount, emails, lastUpdated: Date.now(), error: null };

        const previousIds = new Set(knownIds[account.email] || []);
        const currentIds = emails.map(e => e.id);

        if (!seedOnly && previousIds.size > 0) {
          // Find emails we haven't seen before
          const newEmails = emails.filter(e => !previousIds.has(e.id));
          for (const email of newEmails) {
            await showNotification(email, account, notifLinks);
          }
        }

        // Update known IDs (keep last 100 to avoid unbounded growth)
        knownIds[account.email] = [...new Set([...previousIds, ...currentIds])].slice(-100);
      } catch (err) {
        await markNeedsReauth(account.email);
        cache[account.email] = { unreadCount: 0, emails: [], lastUpdated: Date.now(), error: err.message };
      }
    })
  );

  // Trim notifLinks to last 50 entries
  const notifEntries = Object.entries(notifLinks);
  const trimmedLinks = notifEntries.length > 50
    ? Object.fromEntries(notifEntries.slice(-50))
    : notifLinks;

  await chrome.storage.local.set({ emailsCache: cache, knownIds, notifLinks: trimmedLinks });
  await updateBadge(totalUnread);
}

async function showNotification(email, account, notifLinks) {
  const from = email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown sender';
  const subject = email.subject || '(no subject)';

  // Sanitize ID — Chrome notification IDs must not exceed 100 chars or contain odd chars
  const rawId = `outlook-${account.email}-${email.id}`;
  const notifId = rawId.replace(/[^a-zA-Z0-9._@-]/g, '_').slice(0, 100);

  // Build deep link with login_hint for correct account
  const isPersonal = /hotmail\.|outlook\.|live\.|msn\./i.test(account.email.split('@')[1] || '');
  const fallback = isPersonal ? 'https://outlook.live.com/mail/' : 'https://outlook.office.com/mail/';
  const webLink = email.webLink || fallback;
  const url = webLink + (webLink.includes('?') ? '&' : '?') + `login_hint=${encodeURIComponent(account.email)}`;

  notifLinks[notifId] = url;

  // Wrap in explicit Promise — more reliable in MV3 service workers
  await new Promise((resolve, reject) => {
    chrome.notifications.create(notifId, {
      type: 'basic',
      iconUrl: 'icons/icon128.png',
      title: from,
      message: subject,
      contextMessage: account.email,
      priority: 2,
    }, createdId => {
      if (chrome.runtime.lastError) {
        console.error('[Outlook Checker] Notification error:', chrome.runtime.lastError.message);
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        console.log('[Outlook Checker] Notification shown:', createdId);
        resolve(createdId);
      }
    });
  }).catch(err => console.error('[Outlook Checker] showNotification failed:', err.message));
}

async function updateBadge(count) {
  const text = count > 0 ? (count > 99 ? '99+' : String(count)) : '';
  await chrome.action.setBadgeText({ text });
  await chrome.action.setBadgeBackgroundColor({ color: '#0078d4' });
}
