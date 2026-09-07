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
# September 2026 amendments

- BOM comparison now includes only Material / Material Type values `Mild Steel` and `AISI 304` (case and whitespace insensitive). `Plain Carbon Steel`, blank materials and other grades are excluded. A material column is required. Legacy DOC BOMs must be saved as XLSX, CSV or a DOCX table first; legacy nesting job cards remain supported.
- Material is no longer used as a thickness fallback. Missing quantity or thickness is flagged for review, not reported as an exact match.
- Generate CSV and Generate PDF use the same category selection. All selects every category and includes exact matches. Clearing All clears every category.
- Download Selected saves the report. Mail Selected now creates a Microsoft 365 draft with the selected report attached. Staff open Drafts in their installed Outlook and send it there. Administrator configuration is required; see [Microsoft 365 setup](MICROSOFT-365-SETUP.md).
- Run `npm test` for material filtering, column separation, quoted CSV and report export and Microsoft draft regression tests. PDF reports use browser-side jsPDF and AutoTable, loaded on demand.
