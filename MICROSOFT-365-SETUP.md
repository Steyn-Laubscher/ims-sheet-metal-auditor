# Microsoft 365 Outlook drafts setup

The auditor creates a draft with the selected CSV/PDF attached in the signed-in user's Exchange Online mailbox. Staff open Drafts in their installed Outlook on Windows or Mac, review the message, and send it. No automatic sending, local email-file download, or Outlook web compose window is used. This does not automatically open the desktop draft.

## Administrator registration

1. In Microsoft Entra admin center, open **App registrations → New registration**. Name it **IMS Sheet Metal Auditor**. Choose **Accounts in this organizational directory only**. Create a separate registration for this standalone app.
2. Under **Authentication → Add a platform → Single-page application**, register these exact redirect URIs:
   - `http://localhost:5173/redirect.html` for local testing (open the auditor at `http://localhost:5173/`).
   - `https://imssheetmauditor.netlify.app/redirect.html` for the existing public site when the update is deployed.
3. Under **API permissions**, add **Microsoft Graph → Delegated permissions → Mail.ReadWrite**. Do not add application permissions or Mail.Send. This permission can read/write mail in the signed-in mailbox; the implementation only creates drafts. Grant organization consent if required by your tenant policy.
4. Copy **Application (client) ID** and **Directory (tenant) ID** from Overview. These are identifiers, not passwords. Do not create or supply a client secret; this is a browser app using authorization code flow with PKCE through Microsoft MSAL.
5. For local testing, copy `.env.example` to `.env.local` and fill in:
   ```
   VITE_M365_CLIENT_ID=your-application-client-id
   VITE_M365_TENANT_ID=your-directory-tenant-id
   ```
   Restart the local server. For approved Netlify publication, set the same environment variables for the build and rebuild. Vite embeds these public IDs at build time. Never put secrets in VITE variables.

The app uses its current origin plus `/redirect.html`; the URL must match an Entra SPA registration exactly. The existing `127.0.0.1` local URL can still be used for auditing/downloads, but use `localhost` for this registered sign-in flow. Keep the bundled redirect page on the same origin, without Cross-Origin-Opener-Policy, and with `Cache-Control: no-store` (included in Netlify configuration).

## Staff workflow

1. Compare the documents and choose Generate CSV or Generate PDF.
2. Choose report categories, then Connect Microsoft 365. Select the same work account used in installed Outlook. Complete sign-in/consent in Microsoft's popup.
3. Confirm the Connected account shown in the dialog. Click Mail Selected.
4. Wait for **Draft created**. In installed Outlook, open that account's Drafts and find **IMS Sheet Metal Audit - CSV report** or **IMS Sheet Metal Audit - PDF report**. Allow a moment for synchronization.
5. Confirm the attachment, recipients (Werner and John; Nikita copied), and content before sending.

Disconnect clears this app's cached account/tokens for the session; it does not sign out other Microsoft applications. Original input documents remain in the browser. Only Mail Selected transmits the selected report, its summary and input filenames to Microsoft 365. Tokens are held by MSAL in session storage, not committed to source or logged.

## Verification and limits

- `npm test` checks recipients, attachment bytes, Graph request shape, failed requests and duplicate-prevention behavior with mocked HTTP. `npm run build` checks the app and the separate sign-in redirect bundle.
- Real authentication, tenant consent, draft creation and desktop synchronization need a configured test account. Do not claim those are verified until tested.
- Reports up to 2.5 MB are attached in the same request that creates the draft. Larger reports have a clear error and can be downloaded manually. Large-file upload sessions are not implemented.
- Failed or interrupted writes are not retried automatically. If the result is uncertain, check Outlook Drafts first to avoid duplicates. The app blocks another attempt until the user acknowledges checking.
- Signing in requires an online connection, allowed popups, and a supported browser. Conditional Access or admin restrictions may require assistance from IT.

Microsoft references: [Create draft message](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0), [MSAL initialization](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization), [Redirect bridge](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/redirect-bridge.md).
