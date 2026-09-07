import type { Comparison } from './types';

export type ReportFormat = 'csv' | 'pdf';
export type ReportContext = { bomName: string; jobName: string };
const columns = ['Status', 'BOM part', 'Job-card part', 'BOM quantity', 'Job-card quantity', 'BOM thickness', 'Job-card thickness', 'Explanation'];
const values = (r: Comparison) => [r.severity, r.bomPart, r.jobPart, r.bomQty, r.jobQty, r.bomThickness, r.jobThickness, r.explanation];

export function csvReport(rows: Comparison[]): Blob {
  const cell = (value: string) => `"${(/^[\s]*[=+@-]/.test(value) ? "'" : '') + value.replaceAll('"', '""')}"`;
  return new Blob(['\uFEFF' + [columns, ...rows.map(values)].map(row => row.map(cell).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' });
}

export async function pdfReport(rows: Comparison[], context: ReportContext): Promise<Blob> {
  const { jsPDF } = await import('jspdf');
  const { autoTable } = await import('jspdf-autotable');
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  doc.setFillColor(32, 38, 36);
  doc.rect(0, 0, 297, 23, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(17);
  doc.text('IMS SHEET METAL AUDITOR', 12, 15);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(32, 38, 36);
  doc.setFontSize(9);
  const meta = doc.splitTextToSize(`BOM: ${context.bomName}\nJob card: ${context.jobName}\nBOM scope: Mild Steel and AISI 304 only\nSelected report: ${rows.length} rows | Generated: ${new Date().toLocaleString('en-ZA')}`, 273);
  doc.text(meta, 12, 31);
  autoTable(doc, {
    startY: 34 + meta.length * 4,
    margin: { left: 12, right: 12, top: 12, bottom: 17 },
    head: [['Status', 'BOM part', 'Job-card part', 'Qty\nBOM / Job', 'Thickness\nBOM / Job', 'Explanation']],
    body: rows.map(r => [r.severity.toUpperCase(), r.bomPart, r.jobPart, `${r.bomQty} / ${r.jobQty}`, `${r.bomThickness} / ${r.jobThickness}`, r.explanation]),
    styles: { font: 'helvetica', fontSize: 8, cellPadding: 2.5, overflow: 'linebreak', valign: 'top', lineColor: [220, 225, 220], lineWidth: 0.1 },
    headStyles: { fillColor: [32, 38, 36], textColor: [255, 255, 255] },
    alternateRowStyles: { fillColor: [244, 245, 240] },
    columnStyles: { 0: { cellWidth: 24 }, 1: { cellWidth: 60 }, 2: { cellWidth: 60 }, 3: { cellWidth: 22 }, 4: { cellWidth: 27 }, 5: { cellWidth: 80 } },
    rowPageBreak: 'avoid',
    didParseCell: data => {
      if (data.section === 'body' && data.column.index === 0) {
        const tones: Record<string, [number, number, number]> = { critical: [175, 35, 30], warning: [145, 85, 0], revision: [35, 95, 155], match: [30, 120, 65] };
        data.cell.styles.textColor = tones[rows[data.row.index].severity];
        data.cell.styles.fontStyle = 'bold';
      }
    },
  });
  for (let page = 1; page <= doc.getNumberOfPages(); page++) {
    doc.setPage(page);
    doc.setFontSize(8);
    doc.setTextColor(100, 110, 105);
    doc.text('IMS South Africa Management System', 12, 202);
    doc.text(`Page ${page} of ${doc.getNumberOfPages()}`, 285, 202, { align: 'right' });
  }
  return doc.output('blob');
}

export const makeReport = (format: ReportFormat, rows: Comparison[], context: ReportContext) => format === 'csv' ? Promise.resolve(csvReport(rows)) : pdfReport(rows, context);


export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
