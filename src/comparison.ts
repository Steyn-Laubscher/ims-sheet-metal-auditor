import type { Comparison, Part, Severity } from './types';
import { normal } from './parser';

const qty = (v: number | null) => v == null ? '—' : String(v);
const thickness = (v: string) => v || '—';
const dangerous = (a: string, b: string) => {
  const tokensA = new Set(a.toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean));
  const tokensB = new Set(b.toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean));
  return [...tokensA].some(t => (t === '4T' && tokensB.has('3T')) || (t === '3T' && tokensB.has('4T')) || (t === 'STD' && (tokensB.has('RHS') || tokensB.has('LHS'))) || (t === 'RHS' && tokensB.has('STD')) || (t === 'LHS' && tokensB.has('STD')));
};
const revisionOnly = (a: string, b: string) => normal(a) === normal(b) && a.toUpperCase() !== b.toUpperCase();

export function compareParts(bom: Part[], job: Part[]): Comparison[] {
  const result: Comparison[] = [];
  const used = new Set<number>();
  bom.forEach(b => {
    const exactIndex = job.findIndex((j, idx) => !used.has(idx) && j.name.trim().toUpperCase() === b.name.trim().toUpperCase());
    const rawIndex = exactIndex >= 0 ? exactIndex : job.findIndex((j, idx) => !used.has(idx) && (revisionOnly(b.name, j.name) || dangerous(b.name, j.name) || normal(j.name).includes(normal(b.name)) || normal(b.name).includes(normal(j.name))));
    if (rawIndex < 0) { result.push({ id: crypto.randomUUID(), severity: 'critical', bomPart: b.name, jobPart: 'Missing from job card', bomQty: qty(b.quantity), jobQty: '—', bomThickness: thickness(b.thickness), jobThickness: '—', explanation: 'Part is present in the BOM but was not found on the nesting job card.' }); return; }
    used.add(rawIndex); const j = job[rawIndex];
    const rev = revisionOnly(b.name, j.name); const nameChanged = normal(b.name) !== normal(j.name);
    const qtyMismatch = b.quantity != null && j.quantity != null && b.quantity !== j.quantity;
    const thickMismatch = b.thickness && j.thickness && b.thickness.toUpperCase() !== j.thickness.toUpperCase();
    let severity: Severity = 'match'; let explanation = 'Part name, quantity, and thickness match.';
    if (qtyMismatch || thickMismatch || dangerous(b.name, j.name)) { severity = 'critical'; explanation = [dangerous(b.name,j.name) ? 'Potentially dangerous part-name change (4T/3T or STD/LHS/RHS).' : '', qtyMismatch ? 'Quantity differs.' : '', thickMismatch ? 'Thickness differs.' : ''].filter(Boolean).join(' '); }
    else if (b.quantity == null || j.quantity == null || !b.thickness || !j.thickness) { severity = 'warning'; explanation = 'Quantity or thickness is missing; an exact match cannot be confirmed.' + (rev ? ' Revision-only name difference also requires confirmation.' : nameChanged ? ' Part names also differ.' : ''); }
    else if (rev || nameChanged) { severity = rev ? 'revision' : 'warning'; explanation = rev ? 'Revision-only difference; confirm the released revision.' : 'Part name differs; human confirmation required.'; }
    result.push({ id: crypto.randomUUID(), severity, bomPart: b.name, jobPart: j.name, bomQty: qty(b.quantity), jobQty: qty(j.quantity), bomThickness: thickness(b.thickness), jobThickness: thickness(j.thickness), explanation });
  });
  job.forEach((j, idx) => { if (!used.has(idx)) result.push({ id: crypto.randomUUID(), severity: 'warning', bomPart: 'Not in BOM', jobPart: j.name, bomQty: '—', jobQty: qty(j.quantity), bomThickness: '—', jobThickness: thickness(j.thickness), explanation: 'Part appears on the job card but is not present in the BOM.' }); });
  return result;
}
