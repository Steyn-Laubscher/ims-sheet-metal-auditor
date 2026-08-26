# IMS SHEET METAL AUDITOR

Standalone Vite + React + TypeScript application for local, pre-production BOM vs nesting job-card audits.

## Local use

```bash
npm install
npm run dev
```

Or double-click `start-local.command` on macOS. Open the local URL shown by Vite.

Files are parsed in the browser and are never uploaded. Excel/CSV uses SheetJS, DOCX uses Mammoth, and legacy DOC uses a local binary-text fallback for the plain-text job-card format used by the supplied validation file.

## Validation fixtures

- `/Users/steyn/Downloads/4T1875 (Remote) Solid Metal Sides_revA.xlsx`
- `/Users/steyn/Downloads/CM16-REMOTE-OK-SOLIDW-STD.DOC`

The current fixture run reads 74 BOM parts and 73 deduplicated job-card parts, and surfaces missing parts plus critical STD/RHS and 4T/3T changes. No deployment or external project integration is included.
