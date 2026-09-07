import { PublicClientApplication, InteractionRequiredAuthError, type AccountInfo } from '@azure/msal-browser';

const clientId = import.meta.env.VITE_M365_CLIENT_ID?.trim() || '';
const tenantId = import.meta.env.VITE_M365_TENANT_ID?.trim() || '';
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const microsoftConfigured = guid.test(clientId) && guid.test(tenantId);
const scopes = ['Mail.ReadWrite'];
let client: PublicClientApplication | undefined;
let ready: Promise<void> | undefined;

export function prepareMicrosoft(): Promise<void> {
  if (!microsoftConfigured) return Promise.reject(new Error('Microsoft 365 setup is required. Ask your administrator to configure this app.'));
  if (!client) client = new PublicClientApplication({
    auth: { clientId, authority: `https://login.microsoftonline.com/${tenantId}`, redirectUri: `${window.location.origin}/redirect.html`, postLogoutRedirectUri: `${window.location.origin}/redirect.html` },
    cache: { cacheLocation: 'sessionStorage' },
  });
  if (!ready) ready = client.initialize().catch(error => { ready = undefined; throw error; });
  return ready;
}

export function microsoftAccount(): AccountInfo | null { return client?.getActiveAccount() || null; }

// Initialization is completed when the export dialog mounts, before this click.
export async function connectMicrosoft(): Promise<AccountInfo> {
  if (!client || !ready) throw new Error('Microsoft sign-in is still preparing. Please try again.');
  const response = await client.loginPopup({ scopes, prompt: 'select_account' });
  if (!response.account) throw new Error('Microsoft did not return a signed-in account.');
  client.setActiveAccount(response.account);
  return response.account;
}

export async function microsoftToken(): Promise<string> {
  const account = microsoftAccount();
  if (!client || !account) throw new Error('Connect Microsoft 365 before creating a draft.');
  try {
    return (await client.acquireTokenSilent({ scopes, account })).accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) throw new Error('Your Microsoft 365 session needs attention. Click Change account / Reconnect, then try again.');
    throw error;
  }
}

export async function disconnectMicrosoft() {
  const account = microsoftAccount();
  if (client && account) { await client.clearCache({ account }); client.setActiveAccount(null); }
}
