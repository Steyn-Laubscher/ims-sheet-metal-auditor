import { broadcastResponseToMainFrame } from '@azure/msal-browser/redirect-bridge';
broadcastResponseToMainFrame().catch(() => {
  document.getElementById('status')!.textContent = 'Sign-in could not complete. Close this window and try connecting Microsoft 365 again from the auditor.';
});
