export type SourceKind = 'bom' | 'jobcard';
export type Severity = 'match' | 'revision' | 'warning' | 'critical';
export type Part = { name: string; quantity: number | null; thickness: string; sourceRow: number };
export type Upload = { file: File; status: 'reading' | 'valid' | 'error'; message?: string; parts?: Part[] };
export type Comparison = { id: string; severity: Severity; bomPart: string; jobPart: string; bomQty: string; jobQty: string; bomThickness: string; jobThickness: string; explanation: string };
