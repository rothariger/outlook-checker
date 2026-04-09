// Copy this file to config.js and replace with your Azure AD App Registration values.
// https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps
export const CLIENT_ID = 'YOUR_AZURE_AD_CLIENT_ID_HERE';

export const REDIRECT_URI = `https://${chrome.runtime.id}.chromiumapp.org/`;
export const SCOPES = 'Mail.ReadWrite User.Read openid profile offline_access';
export const AUTH_ENDPOINT = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
export const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
export const POLL_INTERVAL_MINUTES = 1;
