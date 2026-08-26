import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import type { Part, SourceKind } from './types';

const clean = (value: unknown) => String(value ?? '').replace(/\s+/g, ' ').trim();
const numberValue = (value: unknown) => { const n = Number(String(value ?? '').replace(/,/g, '')); return Number.isFinite(n) && String(value).trim() !== '' ? n : null; };
const normal = (name: string) => clean(name).replace(/\.(dft|dwg|pdf)$/i, '').replace(/[_-]?rev(?:ision)?[ _-]?[a-z0-9]+$/i, '').replace(/\s+/g, ' ').trim().toUpperCase();
const findColumn = (headers: string[], words: string[]) => words.reduce((found, word) => found >= 0 ? found : headers.findIndex(h => h.includes(word)), -1);

function rowsToParts(rows: unknown[][], kind: SourceKind): Part[] {
  const headerAt = rows.findIndex(row => row.filter(Boolean).length >= 2 && row.some(v => /part|item|description|thickness|nested|cut qty/i.test(clean(v))));
  const headers = (rows[headerAt >= 0 ? headerAt : 0] ?? []).map(clean).map(v => v.toUpperCase());
  const nameIdx = findColumn(headers, ['PART NUMBER', 'PART NAME', 'DESCRIPTION', 'PART', 'ITEM']);
  const qtyIdx = findColumn(headers, [kind === 'bom' ? 'QTY' : 'CUT QTY', 'NESTED QTY', 'QUANTITY']);
  const thickIdx = findColumn(headers, ['THICKNESS', 'GAUGE', 'MATERIAL']);
  const result: Part[] = [];
  rows.slice(headerAt >= 0 ? headerAt + 1 : 0).forEach((row, i) => {
    const name = clean(row[nameIdx >= 0 ? nameIdx : 0]);
    if (!name || /part number|description|item no/i.test(name)) return;
    const rowText = row.map(clean).join(' ');
    if (kind === 'jobcard' && !/\.dft|\.dwg|\.pdf/i.test(rowText) && name.length < 4) return;
    result.push({ name: name.replace(/\.(dft|dwg|pdf)$/i, ''), quantity: numberValue(row[qtyIdx >= 0 ? qtyIdx : 1]), thickness: clean(row[thickIdx >= 0 ? thickIdx : -1]), sourceRow: i + 1 });
  });
  return result;
}

async function parseXlsx(file: File, kind: SourceKind) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const all: Part[] = [];
  workbook.SheetNames.forEach(sheet => all.push(...rowsToParts(XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { header: 1, defval: '' }) as unknown[][], kind)));
  return all;
}

function parseLegacyDoc(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const decoded = new TextDecoder('windows-1252').decode(bytes);
  return decoded.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '\n').replace(/[^\x09\x0A\x0D\x20-\x7EÀ-ÿ]/g, ' ');
}

function parseJobCardText(text: string): Part[] {
  const lines = text.split(/\r?\n/).map(clean).filter(Boolean);
  const parts: Part[] = [];
  lines.forEach((line, i) => {
    const match = line.match(/([A-Z0-9][A-Z0-9 _()\-./]+?)(?:\.dft|\.dwg|\.pdf)\b/i);
    if (match) {
      const nextNumbers = lines.slice(i + 1, i + 6).map(numberValue).filter((value): value is number => value != null);
      parts.push({ name: clean(match[1]), quantity: nextNumbers[1] ?? null, thickness: nextNumbers[0] == null ? '' : String(nextNumbers[0]), sourceRow: i + 1 });
    }
  });
  const unique = new Map<string, Part>();
  parts.forEach(part => {
    const key = part.name.toUpperCase();
    const existing = unique.get(key);
    if (!existing) unique.set(key, part);
    else if (existing.quantity != null && part.quantity != null) existing.quantity += part.quantity;
  });
  return [...unique.values()];
}

export async function parseDocument(file: File, kind: SourceKind): Promise<Part[]> {
  const extension = file.name.split('.').pop()?.toLowerCase();
  if (extension === 'xlsx' || extension === 'xls' || extension === 'csv') {
    if (extension === 'csv') return rowsToParts((await file.text()).split(/\r?\n/).map(line => line.split(',')), kind);
    return parseXlsx(file, kind);
  }
  if (extension === 'docx') return parseJobCardText((await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() })).value);
  if (extension === 'doc') return parseJobCardText(parseLegacyDoc(await file.arrayBuffer()));
  throw new Error('Unsupported file type. Please choose .xlsx, .xls, .csv, .doc, or .docx.');
}

export { normal };
