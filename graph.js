import { GRAPH_BASE } from './config.js';

/**
 * Returns the number of unread messages in the inbox.
 */
export async function getUnreadCount(accessToken) {
  const url = `${GRAPH_BASE}/me/mailFolders/inbox/messages?$filter=isRead eq false&$count=true&$top=1&$select=id`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ConsistencyLevel: 'eventual',
    },
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data['@odata.count'] ?? 0;
}

/**
 * Returns up to 10 unread inbox messages with key metadata.
 */
export async function getRecentEmails(accessToken) {
  const select = 'id,subject,from,receivedDateTime,isRead,webLink';
  const url = `${GRAPH_BASE}/me/mailFolders/inbox/messages?$filter=isRead eq false&$top=10&$select=${select}&$orderby=receivedDateTime desc`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ConsistencyLevel: 'eventual',
    },
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.value || [];
}

export async function markAsRead(accessToken, messageId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${messageId}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ isRead: true }),
  });
  if (!res.ok) throw new Error(await extractError(res));
}

export async function archiveEmail(accessToken, messageId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${messageId}/move`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ destinationId: 'archive' }),
  });
  if (!res.ok) throw new Error(await extractError(res));
}

export async function deleteEmail(accessToken, messageId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${messageId}`, {
    method: 'DELETE',
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error(await extractError(res));
}

export async function moveToJunk(accessToken, messageId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${messageId}/move`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ destinationId: 'junkemail' }),
  });
  if (!res.ok) throw new Error(await extractError(res));
}

export async function getEmailBody(accessToken, messageId) {
  const url = `${GRAPH_BASE}/me/messages/${messageId}?$select=subject,from,toRecipients,receivedDateTime,body,isRead`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data;
}

async function extractError(res) {
  try {
    const d = await res.json();
    return d.error?.message || res.statusText;
  } catch {
    return res.statusText;
  }
}
