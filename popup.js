import { addAccount, getAccounts, removeAccount, getValidToken } from './auth.js';
import { markAsRead, archiveEmail, deleteEmail, moveToJunk, getEmailBody } from './graph.js';

const ACCOUNT_COLORS = [
  '#0078d4', '#107c10', '#8764b8', '#c43e1c',
  '#038387', '#ca5010', '#004b1c', '#4f6bed',
];

const $ = id => document.getElementById(id);

// ── Utilities (defined first to avoid any reference issues) ───────────────────

function emailDeepLink(webLink, accountEmail) {
  if (!webLink) return inboxUrl(accountEmail);
  try {
    const url = new URL(webLink);
    url.searchParams.set('login_hint', accountEmail);
    return url.toString();
  } catch {
    return inboxUrl(accountEmail);
  }
}

function inboxUrl(email) {
  const hint = encodeURIComponent(email);
  const domain = email.split('@')[1]?.toLowerCase() || '';
  const isPersonal = ['hotmail.com', 'outlook.com', 'live.com', 'msn.com'].includes(domain);
  return isPersonal
    ? `https://outlook.live.com/mail/?login_hint=${hint}`
    : `https://outlook.office.com/mail/?login_hint=${hint}`;
}

function formatRelativeTime(isoString) {
  const diff = Date.now() - new Date(isoString).getTime();
  const m = Math.floor(diff / 60_000);
  if (m < 1) return 'just now';
  if (m < 60) return `${m}m`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h}h`;
  const d = Math.floor(h / 24);
  if (d < 7) return `${d}d`;
  return new Date(isoString).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
}

function escHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Init ──────────────────────────────────────────────────────────────────────

async function init() {
  $('loading').classList.remove('hidden');
  $('empty').classList.add('hidden');
  $('accounts-list').innerHTML = '';

  const [accounts, storage] = await Promise.all([
    getAccounts(),
    chrome.storage.local.get('emailsCache'),
  ]);

  $('loading').classList.add('hidden');

  if (accounts.length === 0) {
    $('empty').classList.remove('hidden');
    return;
  }

  const cache = storage.emailsCache || {};
  const list = $('accounts-list');
  accounts.forEach((account, i) => {
    const accountData = cache[account.email] || { unreadCount: 0, emails: [], error: null };
    list.appendChild(buildAccountSection(account, accountData, i));
  });
}

// ── Account section builder ───────────────────────────────────────────────────

function buildAccountSection(account, data, colorIndex) {
  const color = ACCOUNT_COLORS[colorIndex % ACCOUNT_COLORS.length];
  const initial = (account.displayName || account.email)[0].toUpperCase();
  const { unreadCount, emails, error } = data;

  const section = document.createElement('div');
  section.className = 'account-section';

  // Header row
  const header = document.createElement('div');
  header.className = 'account-header';
  header.innerHTML = `
    <div class="account-avatar" style="background:${color}">${initial}</div>
    <div class="account-info">
      <a class="account-email" title="${account.email}" href="${inboxUrl(account.email)}" target="_blank">${account.email}</a>
    </div>
    ${unreadCount > 0
      ? `<span class="account-badge has-unread">${unreadCount > 99 ? '99+' : unreadCount}</span>`
      : `<span class="account-badge">0 unread</span>`
    }
    <button class="btn-icon btn-refresh" data-email="${account.email}" title="Refresh">
      <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
        <path d="M10.5 6A4.5 4.5 0 1 1 6 1.5a4.5 4.5 0 0 1 3.18 1.32L10.5 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
        <path d="M10.5 1.5V4h-2.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    </button>
    <button class="btn-icon btn-remove" data-email="${account.email}" title="Remove account">
      <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
        <path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
      </svg>
    </button>
  `;

  section.appendChild(header);

  // Re-auth warning
  if (account.needsReauth || error) {
    const warn = document.createElement('div');
    warn.className = 'reauth-warning';
    if (account.needsReauth) {
      warn.innerHTML = `⚠ Session expired. <a href="#" class="reauth-link" data-email="${account.email}">Sign in again</a>`;
    } else {
      warn.textContent = `⚠ ${error}`;
    }
    section.appendChild(warn);
  }

  // Email list
  if (!account.needsReauth && !error) {
    const ul = document.createElement('ul');
    ul.className = 'email-list';

    if (emails.length === 0) {
      const li = document.createElement('div');
      li.className = 'no-emails';
      li.textContent = 'No unread emails.';
      section.appendChild(li);
    } else {
      emails.forEach(email => ul.appendChild(buildEmailItem(email, account)));
      section.appendChild(ul);
    }
  }

  // Open inbox link
  const openRow = document.createElement('div');
  openRow.className = 'open-inbox-row';
  openRow.innerHTML = `<a class="open-inbox-link" href="${inboxUrl(account.email)}" target="_blank">Open inbox ↗</a>`;
  section.appendChild(openRow);

  return section;
}

function buildEmailItem(email, account) {
  const accountEmail = account.email;
  const li = document.createElement('li');
  li.className = `email-item${email.isRead ? ' is-read' : ''}`;

  const from = email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown';
  const subject = email.subject || '(no subject)';
  const time = formatRelativeTime(email.receivedDateTime);

  li.innerHTML = `
    <div class="unread-dot${email.isRead ? ' read' : ''}"></div>
    <div class="email-content">
      <div class="email-subject">${escHtml(subject)}</div>
      <div class="email-meta">
        <span class="email-from">${escHtml(from)}</span>
        <span>·</span>
        <span class="email-time">${time}</span>
      </div>
    </div>
    <div class="email-actions">
      <button class="email-action-btn" data-action="read" title="Mark as read">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M1 7h12M1 4h12M1 10h7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
      </button>
      <button class="email-action-btn" data-action="archive" title="Archive">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><rect x="1" y="2" width="12" height="3" rx="1" stroke="currentColor" stroke-width="1.3"/><path d="M2 5v6a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5" stroke="currentColor" stroke-width="1.3"/><path d="M5 8h4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
      </button>
      <button class="email-action-btn" data-action="junk" title="Report as spam">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><circle cx="7" cy="7" r="5.5" stroke="currentColor" stroke-width="1.3"/><path d="M7 4v3M7 9.5v.5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
      </button>
      <button class="email-action-btn action-delete" data-action="delete" title="Delete">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M2 4h10M5 4V2.5a.5.5 0 0 1 .5-.5h3a.5.5 0 0 1 .5.5V4M3 4l.7 7.5a.5.5 0 0 0 .5.5h5.6a.5.5 0 0 0 .5-.5L11 4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
      </button>
    </div>
  `;

// ── Email detail view ─────────────────────────────────────────────────────────
  li.querySelector('.email-content').addEventListener('click', () => showEmailDetail(email, account, li));
  li.querySelector('.unread-dot').addEventListener('click', () => showEmailDetail(email, account, li));

  li.querySelectorAll('.email-action-btn').forEach(btn => {
    btn.addEventListener('click', async e => {
      e.stopPropagation();
      btn.disabled = true;
      try {
        const token = await getValidToken(account);
        const action = btn.dataset.action;
        if (action === 'read') await markAsRead(token, email.id);
        else if (action === 'archive') await archiveEmail(token, email.id);
        else if (action === 'junk') await moveToJunk(token, email.id);
        else if (action === 'delete') await deleteEmail(token, email.id);
        li.style.transition = 'opacity 0.2s';
        li.style.opacity = '0';
        setTimeout(() => li.remove(), 200);
      } catch (err) {
        btn.disabled = false;
        alert(`Action failed: ${err.message}`);
      }
    });
  });

  return li;
}

// ── Email detail view ─────────────────────────────────────────────────────────

async function showEmailDetail(email, account, listItem) {
  const detail = $('email-detail');
  const accountsList = $('accounts-list');
  const iframe = $('detail-iframe');
  const detailLoading = $('detail-loading');

  // Populate header
  $('detail-subject').textContent = email.subject || '(no subject)';
  $('detail-from').textContent = email.from?.emailAddress?.name
    ? `${email.from.emailAddress.name} <${email.from.emailAddress.address}>`
    : (email.from?.emailAddress?.address || '');
  $('detail-date').textContent = new Date(email.receivedDateTime).toLocaleString();

  // Build action buttons in detail header
  const actionsEl = $('detail-header-actions');
  actionsEl.innerHTML = `
    <button class="email-action-btn" data-action="read" title="Mark as read">
      <svg width="15" height="15" viewBox="0 0 14 14" fill="none"><path d="M1 7h12M1 4h12M1 10h7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
    </button>
    <button class="email-action-btn" data-action="archive" title="Archive">
      <svg width="15" height="15" viewBox="0 0 14 14" fill="none"><rect x="1" y="2" width="12" height="3" rx="1" stroke="currentColor" stroke-width="1.3"/><path d="M2 5v6a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V5" stroke="currentColor" stroke-width="1.3"/><path d="M5 8h4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
    </button>
    <button class="email-action-btn" data-action="junk" title="Spam">
      <svg width="15" height="15" viewBox="0 0 14 14" fill="none"><circle cx="7" cy="7" r="5.5" stroke="currentColor" stroke-width="1.3"/><path d="M7 4v3M7 9.5v.5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
    </button>
    <button class="email-action-btn action-delete" data-action="delete" title="Delete">
      <svg width="15" height="15" viewBox="0 0 14 14" fill="none"><path d="M2 4h10M5 4V2.5a.5.5 0 0 1 .5-.5h3a.5.5 0 0 1 .5.5V4M3 4l.7 7.5a.5.5 0 0 0 .5.5h5.6a.5.5 0 0 0 .5-.5L11 4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>
    </button>
    <a class="email-action-btn" href="${escHtml(emailDeepLink(email.webLink, account.email))}" target="_blank" title="Open in Outlook">
      <svg width="15" height="15" viewBox="0 0 14 14" fill="none"><path d="M6 2H2a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1V8M8 1h5v5M13 1L7 7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </a>
  `;

  actionsEl.querySelectorAll('.email-action-btn[data-action]').forEach(btn => {
    btn.addEventListener('click', async e => {
      e.stopPropagation();
      btn.disabled = true;
      try {
        const token = await getValidToken(account);
        const action = btn.dataset.action;
        if (action === 'read') await markAsRead(token, email.id);
        else if (action === 'archive') await archiveEmail(token, email.id);
        else if (action === 'junk') await moveToJunk(token, email.id);
        else if (action === 'delete') await deleteEmail(token, email.id);
        listItem?.remove();
        showList();
      } catch (err) {
        btn.disabled = false;
        alert(`Action failed: ${err.message}`);
      }
    });
  });

  // Show detail, hide list
  accountsList.classList.add('hidden');
  iframe.srcdoc = '';
  iframe.classList.add('hidden');
  detailLoading.classList.remove('hidden');
  detail.classList.remove('hidden');

  // Auto mark as read
  if (!email.isRead) {
    getValidToken(account).then(t => markAsRead(t, email.id)).catch(() => {});
    if (listItem) {
      listItem.classList.add('is-read');
      const dot = listItem.querySelector('.unread-dot');
      if (dot) dot.classList.add('read');
    }
  }

  // Fetch body
  try {
    const token = await getValidToken(account);
    const full = await getEmailBody(token, email.id);
    const isHtml = full.body?.contentType === 'html';
    const bodyContent = full.body?.content || '';

    const htmlDoc = isHtml ? bodyContent : `<pre style="font-family:sans-serif;white-space:pre-wrap">${escHtml(bodyContent)}</pre>`;
    const wrapped = `<!DOCTYPE html><html><head><base target="_blank"><style>body{margin:12px;font-family:'Segoe UI',sans-serif;font-size:13px;color:#201f1e}img{max-width:100%}a{color:#0078d4}</style></head><body>${htmlDoc}</body></html>`;

    iframe.srcdoc = wrapped;
    detailLoading.classList.add('hidden');
    iframe.classList.remove('hidden');

    // Populate footer link
    $('btn-open-outlook').href = emailDeepLink(email.webLink, account.email);
  } catch (err) {
    detailLoading.classList.add('hidden');
    iframe.srcdoc = `<p style="color:red;font-family:sans-serif;padding:12px">Could not load email: ${err.message}</p>`;
    iframe.classList.remove('hidden');
  }
}

function showList() {
  $('email-detail').classList.add('hidden');
  $('accounts-list').classList.remove('hidden');
}

async function handleAddAccount() {
  const btn = $('btn-add-account');
  btn.disabled = true;
  try {
    await addAccount();
    await chrome.runtime.sendMessage({ type: 'accountAdded' });
    await init();
  } catch (err) {
    if (!err.message?.includes('cancelled') && !err.message?.includes('closed')) {
      alert(`Could not add account: ${err.message}`);
    }
  } finally {
    btn.disabled = false;
  }
}

document.addEventListener('click', async e => {
  // Back button
  if (e.target.closest('#btn-back')) {
    showList();
    return;
  }

  // Add account
  if (e.target.closest('#btn-add-account') || e.target.closest('#btn-add-first')) {
    await handleAddAccount();
    return;
  }

  // Refresh all accounts
  if (e.target.closest('#btn-refresh-all')) {
    const btn = $('btn-refresh-all');
    btn.classList.add('spinning');
    try {
      await chrome.runtime.sendMessage({ type: 'forceRefresh' });
      await init();
    } finally {
      btn.classList.remove('spinning');
    }
    return;
  }

  // Refresh single account
  const refreshBtn = e.target.closest('.btn-refresh');
  if (refreshBtn) {
    refreshBtn.classList.add('spinning');
    try {
      await chrome.runtime.sendMessage({ type: 'forceRefresh' });
      await init();
    } finally {
      refreshBtn.classList.remove('spinning');
    }
    return;
  }

  // Remove account
  const removeBtn = e.target.closest('.btn-remove');
  if (removeBtn) {
    const email = removeBtn.dataset.email;
    if (confirm(`Remove account ${email}?`)) {
      await removeAccount(email);
      const cache = (await chrome.storage.local.get('emailsCache')).emailsCache || {};
      delete cache[email];
      await chrome.storage.local.set({ emailsCache: cache });
      await init();
    }
    return;
  }

  // Re-auth link
  const reauthLink = e.target.closest('.reauth-link');
  if (reauthLink) {
    e.preventDefault();
    await handleAddAccount();
  }
});

// Refresh when popup opens
chrome.runtime.sendMessage({ type: 'forceRefresh' }).catch(() => {});

init();
