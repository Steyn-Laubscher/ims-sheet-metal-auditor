# IMS Sheet Metal Auditor — Integration Handoff

## Recommended integration model

Use the existing production application as the shell and mount this auditor as a feature route, for example:

```text
/production/sheet-metal-auditor
```

Keep the auditor’s parser and comparison logic as an isolated feature module. Do not share authentication, database, deployment, or runtime state with this standalone repository.

The hosted site is suitable for functional review only:

https://imssheetmauditor.netlify.app/

## What to hand over

- `src/parser.ts` — browser-side Excel, CSV, DOCX, and legacy DOC parsing
- `src/comparison.ts` — matching, revision, dangerous-name, quantity, thickness, missing-part, and extra-part rules
- `src/types.ts` — `Part`, `Upload`, `Comparison`, and severity types
- The two supplied fixtures for regression testing
- This document and the IMS design tokens below

The host application should own the page shell, navigation, user session, notifications, and deployment. The auditor feature should own file selection, parsing, comparison, filters, report display, and CSV export.

## Functional contract

Inputs:

```ts
type Part = {
  name: string;
  quantity: number | null;
  thickness: string;
  sourceRow: number;
};
```

Core comparison entry point:

```ts
compareParts(bomParts: Part[], jobCardParts: Part[]): Comparison[]
```

Each result contains:

```ts
type Comparison = {
  id: string;
  severity: 'match' | 'revision' | 'warning' | 'critical';
  bomPart: string;
  jobPart: string;
  bomQty: string;
  jobQty: string;
  bomThickness: string;
  jobThickness: string;
  explanation: string;
};
```

Required behavior:

- Never silently accept near matches.
- Flag revision-only differences separately.
- Treat `4T` versus `3T` as critical.
- Treat `STD` versus `RHS` or `LHS` as critical.
- Flag any changed word or number in a part name.
- Compare quantity and thickness when both values are available.
- Deduplicate identical job-card rows while summing their quantities.
- Preserve missing BOM parts and extra job-card parts as visible results.
- Sort the All report Critical → Warnings/Revision → Exact Match.
- Keep filters for All, Critical, Warnings, Revision Differences, and Matches.
- CSV export must allow All, Critical, Warnings, and Revision Differences selections.

## Browser-only privacy boundary

Files must be processed with `File.arrayBuffer()` / `File.text()` in the browser. No BOM or job-card file should be posted to an API, database, analytics event, server action, or third-party parser.

## Host-app adapter

The host can wrap the feature with a small adapter:

```ts
type AuditorHostAdapter = {
  onReportGenerated?: (summary: {
    totalBomParts: number;
    totalJobCardParts: number;
    exactMatches: number;
    warnings: number;
    criticalIssues: number;
  }) => void;
  onReset?: () => void;
};
```

The auditor should not assume the host’s router, auth provider, state manager, CSS framework, or component library. Pass host callbacks in; do not import host internals.

## IMS visual tokens

```css
--ims-black: #000000;
--ims-cream: #F4F5F0;
--ims-charcoal: #202624;
--ims-mint: #D4F3DA;
--ims-border: #DCE2DD;
--ims-success: #318048;
--ims-warning: #A56A11;
--ims-critical: #C24A40;
```

Use the host app’s typography and spacing primitives where available. Preserve the auditor’s visual meanings: green exact match, blue revision, amber warning, and red critical.

## Regression test expectation

With the supplied fixtures:

- BOM: 74 unique parts
- Job card: 113 placed parts across 73 unique rows
- Exact matches: 66
- Critical issues: 7
- General warnings: 1, plus revision differences shown separately

The expected critical examples include missing `CM13-12-2`, `STD/RHS` changes, and `4T/3T` changes. Revision examples include `CM13-3-2` versus `CM13-3-2_REVB`.

## Integration acceptance checklist

- Feature opens inside the colleague’s existing production route.
- Existing header, navigation, authentication, and deployment remain authoritative.
- No uploaded file leaves the browser.
- Both supplied fixtures produce the regression counts above.
- All report order and filters behave as specified.
- CSV export reflects the selected categories.
- Replacing/removing a file and starting a new comparison work without a full-app reload.
