import type { Comparison } from './types';
import type { ReportContext, ReportFormat } from './report-export';

// Stay below Graph's small-attachment/request limits after base64 expansion.
export const MAX_DRAFT_ATTACHMENT = 2_500_000;
export class DraftOutcomeUnknown extends Error {}

export async function draftPayload(report: Blob, format: ReportFormat, rows: Comparison[], context: ReportContext) {
  if (report.size > MAX_DRAFT_ATTACHMENT) throw new Error('This report is too large to attach through the app (2.5 MB limit). Select fewer categories or download and attach it in Outlook.');
  const bytes = new Uint8Array(await report.arrayBuffer());
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  const summary = ['critical', 'warning', 'revision', 'match'].map(severity => `${severity}: ${rows.filter(row => row.severity === severity).length}`).join(', ');
  return {
    subject: `IMS Sheet Metal Audit - ${format.toUpperCase()} report`,
    body: { contentType: 'Text', content: `Hi Werner and John,\n\nPlease review the attached sheet metal comparison report.\n\nBOM: ${context.bomName}\nJob card: ${context.jobName}\nBOM materials: Mild Steel and AISI 304 only.\nSelected rows: ${rows.length} (${summary})\n\nRegards` },
    toRecipients: ['werner@imssa.co.za', 'john@imssa.co.za'].map(address => ({ emailAddress: { address } })),
    ccRecipients: [{ emailAddress: { address: 'nikita@imssa.co.za' } }],
    attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: `ims-sheet-metal-audit.${format}`, contentType: format === 'pdf' ? 'application/pdf' : 'text/csv', contentBytes: btoa(binary) }],
  };
}

export async function createMicrosoftDraft(token: string, report: Blob, format: ReportFormat, rows: Comparison[], context: ReportContext, request: typeof fetch = fetch): Promise<string> {
  const payload = await draftPayload(report, format, rows, context);
  let response: Response;
  try {
    // One request creates both the draft and its attachment. Never auto-retry a write.
    response = await request('https://graph.microsoft.com/v1.0/me/messages', {
      method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload), signal: AbortSignal.timeout(60_000),
    });
  } catch {
    throw new DraftOutcomeUnknown('Microsoft did not confirm the result. Check Outlook Drafts before trying again; the draft may already exist.');
  }
  if (!response.ok) {
    if (response.status === 401) throw new Error('Your sign-in expired. Reconnect Microsoft 365 and try again.');
    if (response.status === 403) throw new Error('Microsoft denied mailbox access. Ask your administrator to approve delegated Mail.ReadWrite permission.');
    if (response.status === 413) throw new Error('Microsoft rejected the report size. Download it and attach it in Outlook.');
    if (response.status >= 500) throw new DraftOutcomeUnknown('Microsoft reported a service error. Check Outlook Drafts before trying again; the draft may already exist.');
    throw new Error(`Microsoft could not create the draft (HTTP ${response.status}). Please try again later.`);
  }
  let draft: { id?: string; isDraft?: boolean };
  try { draft = await response.json(); } catch { throw new DraftOutcomeUnknown('Microsoft accepted the request but its reply could not be read. Check Outlook Drafts before trying again.'); }
  if (!draft.id || draft.isDraft !== true) throw new DraftOutcomeUnknown('Microsoft did not confirm a draft. Check Outlook Drafts before trying again.');
  return draft.id;
}
