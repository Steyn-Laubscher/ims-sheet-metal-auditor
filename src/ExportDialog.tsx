import { useEffect, useState } from 'react';
import type { Comparison } from './types';
import { downloadBlob, makeReport, type ReportContext, type ReportFormat } from './report-export';
import { connectMicrosoft, disconnectMicrosoft, microsoftAccount, microsoftConfigured, microsoftToken, prepareMicrosoft } from './microsoft-auth';
import { createMicrosoftDraft, DraftOutcomeUnknown } from './microsoft-draft';

export function ExportDialog({ format, rows, context, onClose }: { format: ReportFormat; rows: Comparison[]; context: ReportContext; onClose: () => void }) {
  const [selection, setSelection] = useState({ all: true, critical: true, warning: true, revision: true });
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [account, setAccount] = useState('');
  const [authReady, setAuthReady] = useState(false);
  const [draftCreated, setDraftCreated] = useState(false);
  const [outcomeUnknown, setOutcomeUnknown] = useState(false);
  useEffect(() => {
    if (!microsoftConfigured) return;
    let active = true;
    prepareMicrosoft().then(() => { if (active) { setAuthReady(true); setAccount(microsoftAccount()?.username || ''); } }).catch(() => { if (active) setMessage('Microsoft sign-in could not initialize. Reload the page to try again.'); });
    return () => { active = false; };
  }, []);
  const selected = selection.all ? rows : rows.filter(row => row.severity !== 'match' && selection[row.severity]);
  const toggle = (key: keyof typeof selection) => { setDraftCreated(false); setSelection(s => key === 'all' ? { all: !s.all, critical: !s.all, warning: !s.all, revision: !s.all } : { ...s, all: false, [key]: !s[key] }); };
  const connect = async () => {
    setBusy(true); setMessage('');
    try { const signedIn = await connectMicrosoft(); setAccount(signedIn.username); setDraftCreated(false); }
    catch { setMessage('Microsoft sign-in was not completed. Allow the sign-in popup and try again. If Microsoft requests administrator approval, contact your Microsoft 365 administrator.'); }
    finally { setBusy(false); }
  };
  const disconnect = async () => {
    setBusy(true);
    try { await disconnectMicrosoft(); setAccount(''); setMessage('Microsoft 365 disconnected from this app.'); }
    catch { setMessage('Could not disconnect. Close this tab to clear this session.'); }
    finally { setBusy(false); }
  };
  const generate = async (mail: boolean) => {
    if (busy || !selected.length || (mail && (!account || draftCreated || outcomeUnknown))) return;
    setBusy(true);
    setMessage('');
    try {
      const report = await makeReport(format, selected, context);
      if (mail) {
        await createMicrosoftDraft(await microsoftToken(), report, format, selected, context);
        setDraftCreated(true);
        setMessage(`Draft created with the ${format.toUpperCase()} attached. In your installed Outlook, open Drafts for ${account} and look for “IMS Sheet Metal Audit - ${format.toUpperCase()} report”. Review and send it there. Allow a moment for Outlook to sync.`);
      } else {
        downloadBlob(report, `ims-sheet-metal-audit.${format}`);
        setMessage(`${format.toUpperCase()} report downloaded.`);
      }
    } catch (error) {
      if (error instanceof DraftOutcomeUnknown) setOutcomeUnknown(true);
      setMessage(`Could not generate report: ${error instanceof Error ? error.message : 'Please try again.'}`);
    } finally { setBusy(false); }
  };
  return <div className="export-screen"><div className="export-card" role="dialog" aria-modal="true" aria-labelledby="export-title">
    <div className="export-head"><div><label>REPORT EXPORT</label><h2 id="export-title">Generate {format.toUpperCase()}</h2><p>Select the audit categories to include in your report.</p></div><button className="close-export" aria-label="Close export" disabled={busy} onClick={onClose}>×</button></div>
    <div className="export-options">{([
      ['all', 'All', 'Include the complete comparison report, including exact matches'],
      ['critical', 'Critical', 'Missing parts, quantity/thickness mismatches and suspected wrong parts'],
      ['warning', 'Warnings', 'Name mismatches, unverified values and extra job-card parts'],
      ['revision', 'Revision Differences', 'Revision-only name changes requiring confirmation'],
    ] as const).map(([key, label, description]) => <label className={`export-option ${selection[key] ? 'checked' : ''}`} key={key}><input type="checkbox" disabled={busy} checked={selection[key]} onChange={() => toggle(key)} /><span><strong>{label}</strong><small>{description}</small></span></label>)}</div>
    <div className="mail-details"><strong>Email recipients</strong><div>To: werner@imssa.co.za; john@imssa.co.za</div><div>Cc: nikita@imssa.co.za</div><p>Mail Selected uploads only the selected report to your Microsoft 365 Drafts with the attachment included. Open Drafts in your installed Outlook to review and send. The original BOM and job card stay in your browser.</p>
      {!microsoftConfigured ? <p role="status">Microsoft 365 email setup is required. Ask your administrator to configure this app. Downloads are still available.</p> : <div className="microsoft-connection"><p>{account ? `Connected: ${account}` : 'Connect your Microsoft 365 work account to create a draft.'}</p><button disabled={!authReady || busy} onClick={connect}>{account ? 'Change account / Reconnect' : 'Connect Microsoft 365'}</button>{account && <button disabled={busy} onClick={disconnect}>Disconnect</button>}</div>}
    </div>
    <div className="export-footer"><span>{selected.length} rows selected</span><div className="export-actions"><button className="primary" disabled={!selected.length || busy} onClick={() => generate(false)}>Download Selected {format.toUpperCase()} ↓</button><button disabled={!selected.length || busy || !account || draftCreated || outcomeUnknown} onClick={() => generate(true)}>{draftCreated ? 'Draft created ✓' : `Mail Selected ${format.toUpperCase()} ✉`}</button></div></div>
    {busy && <p role="status">Preparing report…</p>}{message && <p role="status" className="export-message">{message}</p>}
    {outcomeUnknown && <button disabled={busy} onClick={() => { setOutcomeUnknown(false); setMessage('You can try creating the draft again.'); }}>I checked Outlook Drafts; allow another attempt</button>}
  </div></div>;
}
