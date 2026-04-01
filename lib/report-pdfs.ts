import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { ReportDetailRow } from './reports-data';

// VBA-matching margins (inches to mm: 1" = 25.4mm)
const INCH_MM = 25.4;
const MARGIN_LEFT_DETAILS = 0.4 * INCH_MM; // 10.16
const MARGIN_LEFT_SUMMARY = 0.5 * INCH_MM; // 12.7
const MARGIN_SMALL = 0.1 * INCH_MM; // 2.54
const A4_W = 210;
const A4_H = 297;
// 2 blocks per page: second block starts at this Y (mm) - same concept as VBA row 25
const SECOND_BLOCK_START_Y = 155;
const SMALL_TABLE_MAX_ROWS = 23; // VBA: ≤20 data rows + header + total ≈ 23
// Agar itni jagah nahi to block nayi page par start karo (half-split avoid)
const MIN_SPACE_FOR_BLOCK_MM = 55; // header + ~4–5 rows
const PAGE_BOTTOM_MARGIN_MM = 12;
// VBA RowHeight = 18.15 (points) → mm
const ROW_HEIGHT_MM = (18.15 / 72) * INCH_MM;
const LINE_NORMAL = 0.05;
const LINE_THICK = 0.2;

const MONTH_NAMES = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
const FOOTER_Y = A4_H - 8;

/** Block split across pages: add "1/2", "2/2" only on those pages; other pages get nothing */
function addBlockPageNumbers(
  doc: jsPDF,
  pageSpans: { startPage: number; endPage: number }[]
): void {
  if (pageSpans.length === 0) return;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  for (const { startPage, endPage } of pageSpans) {
    const total = endPage - startPage + 1;
    for (let p = startPage; p <= endPage; p++) {
      doc.setPage(p);
      doc.text(`${p - startPage + 1}/${total}`, A4_W / 2, FOOTER_Y, { align: 'center' });
    }
  }
}

function getDaysInMonth(month: number, year: number): number {
  if (month === 2) return (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)) ? 29 : 28;
  if ([4, 6, 9, 11].includes(month)) return 30;
  return 31;
}

function formatPeriodForDisplay(period: string): string {
  if (!period || period === 'All') return 'All Periods';
  const parts = period.trim().split('-');
  if (parts.length < 3) return period;
  const year = parseInt(parts[0], 10) || 0;
  const month = parseInt(parts[1], 10) || 0;
  const half = (parts[2] || '').replace(/\s/g, '').toLowerCase();
  if (month < 1 || month > 12) return period;
  const monthName = MONTH_NAMES[month - 1];
  const daysInMonth = getDaysInMonth(month, year);
  let range: string;
  if (half === '1&2' || half === '1and2') range = `1 to ${daysInMonth}`;
  else if (half === '1') range = '1 to 15';
  else if (half === '2') range = `16 to ${daysInMonth}`;
  else range = `1 to ${daysInMonth}`;
  return `${year}-${monthName}(${range})`;
}

function availToText(avail: number): string {
  if (avail === 1) return 'Yes';
  if (avail === 0) return 'NS';
  if (avail === 2) return 'B-Yes';
  return 'UNKNOWN';
}

function availToTextSummary(avail: number): string {
  if (avail === 1) return 'YES';
  if (avail === 0) return 'NS';
  if (avail === 2) return 'B-YES';
  return 'UNKNOWN';
}

function getLanguageShortCode(lang: string): string {
  const s = (lang || '').trim();
  if (!s) return '';
  const lower = s.toLowerCase();
  if (lower.startsWith('english')) return 'EN';
  if (lower.startsWith('hindi')) return 'HI';
  if (lower.startsWith('marathi')) return 'MR';
  if (lower.startsWith('gujarati')) return 'GJ';
  if (lower.startsWith('tamil')) return 'TM';
  if (lower.startsWith('telugu')) return 'TL';
  if (lower.startsWith('kannada')) return 'KN';
  if (lower.startsWith('bengali')) return 'BN';
  return s.length >= 2 ? s.slice(0, 2).toUpperCase() : s.toUpperCase();
}

function getContactNameOnly(contact: string): string {
  const s = (contact || '').trim();
  const i = s.indexOf('(');
  if (i >= 0) return s.slice(0, i).trim();
  const d = s.indexOf('-');
  if (d >= 0) return s.slice(0, d).trim();
  return s;
}

function getContactNumber(contact: string): string {
  const s = (contact || '').trim();
  let num = '';
  const start = s.indexOf('(');
  const end = s.indexOf(')');
  if (start >= 0 && end > start) num = s.slice(start + 1, end).trim();
  else {
    const d = s.indexOf('-');
    if (d >= 0) num = s.slice(d + 1).trim();
  }
  if (num.length === 10 && /^\d+$/.test(num))
    return num.slice(0, 5) + ' ' + num.slice(5);
  return num;
}

function toCamelCase(s: string): string {
  return (s || '')
    .split(/\s+/)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(' ');
}

// VBA order: group_sector, village, then Yes / Not In Stock / B-Yes, then contact, language, book
function sortedDetailRows(rows: ReportDetailRow[]): ReportDetailRow[] {
  return [...rows].sort((a, b) => {
    const g = (a.group_sector || '').localeCompare(b.group_sector || '');
    if (g !== 0) return g;
    const v = (a.village || '').localeCompare(b.village || '');
    if (v !== 0) return v;
    const availOrder = (x: number) => (x === 1 ? 0 : x === 0 ? 1 : 2);
    const av = availOrder(a.avail) - availOrder(b.avail);
    if (av !== 0) return av;
    const c = (a.contact_name || '').localeCompare(b.contact_name || '');
    if (c !== 0) return c;
    const lang = (a.language || '').localeCompare(b.language || '');
    if (lang !== 0) return lang;
    return (a.book_name || '').localeCompare(b.book_name || '');
  });
}

type DetailsCellDef =
  | string
  | number
  | { content: string; rowSpan?: number; styles?: { valign?: 'top' | 'middle' | 'bottom' } };

// VBA: Yes = full merge (Name, Mo.No., LG, AVL, Sign [total], Date); Not In Stock / B-Yes = limited merge (only Name, Mo.No., LG)
// Rows ordered by avail (Yes, NS, B-Yes), then contact, language, book. Merge per section (same avail).
function buildDetailsBodyWithMerges(blockRows: ReportDetailRow[]): DetailsCellDef[][] {
  const n = blockRows.length;
  if (n === 0) return [];

  const keyContact = (i: number) => (blockRows[i].contact_name || '').trim() + '|' + blockRows[i].avail;
  const keyLang = (i: number) =>
    keyContact(i) + '|' + getLanguageShortCode(blockRows[i].language);
  const keyAvail = (i: number) => keyContact(i) + '|' + availToText(blockRows[i].avail);

  const contactRun: number[] = [];
  for (let i = 0; i < n; i++) {
    if (i > 0 && keyContact(i) === keyContact(i - 1)) contactRun.push(0);
    else {
      let run = 1;
      while (i + run < n && keyContact(i + run) === keyContact(i)) run++;
      contactRun.push(run);
    }
  }

  const langRun: number[] = [];
  for (let i = 0; i < n; i++) {
    if (i > 0 && keyLang(i) === keyLang(i - 1)) langRun.push(0);
    else {
      let run = 1;
      while (i + run < n && keyLang(i + run) === keyLang(i)) run++;
      langRun.push(run);
    }
  }

  const availRun: number[] = [];
  for (let i = 0; i < n; i++) {
    if (blockRows[i].avail !== 1) {
      availRun.push(0);
      continue;
    }
    if (i > 0 && keyAvail(i) === keyAvail(i - 1)) availRun.push(0);
    else {
      let run = 1;
      while (i + run < n && keyAvail(i + run) === keyAvail(i)) run++;
      availRun.push(run);
    }
  }

  const contactTotalAtStart: number[] = [];
  for (let i = 0; i < n; i++) {
    if (blockRows[i].avail !== 1 || contactRun[i] === 0) {
      contactTotalAtStart.push(0);
      continue;
    }
    let sum = 0;
    for (let j = i; j < i + contactRun[i]; j++) sum += blockRows[j].qty;
    contactTotalAtStart.push(sum);
  }

  const valignMiddle = { valign: 'middle' as const };
  const body: DetailsCellDef[][] = [];

  for (let i = 0; i < n; i++) {
    const r = blockRows[i];
    const isYes = r.avail === 1;
    const row: DetailsCellDef[] = [];

    if (contactRun[i] > 0) {
      row.push({
        content: toCamelCase(getContactNameOnly(r.contact_name)),
        rowSpan: contactRun[i],
        styles: valignMiddle,
      });
      row.push({
        content: getContactNumber(r.contact_name),
        rowSpan: contactRun[i],
        styles: valignMiddle,
      });
    }
    if (langRun[i] > 0)
      row.push({
        content: getLanguageShortCode(r.language),
        rowSpan: langRun[i],
        styles: valignMiddle,
      });
    row.push(r.book_name);

    if (isYes && availRun[i] > 0)
      row.push({
        content: availToText(r.avail),
        rowSpan: availRun[i],
        styles: valignMiddle,
      });
    else if (!(isYes && availRun[i] === 0))
      row.push(availToText(r.avail));

    row.push(String(r.qty));

    if (isYes && contactRun[i] > 0) {
      const total = contactTotalAtStart[i];
      row.push({
        content: total ? `[${total}]` : '',
        rowSpan: contactRun[i],
        styles: valignMiddle,
      });
      row.push({ content: '', rowSpan: contactRun[i], styles: valignMiddle });
    } else {
      row.push('');
      row.push('');
    }
    body.push(row);
  }
  return body;
}

// --- Details Report (SectorVillageDetailsReport): Portrait, 8 cols, blocks per sector-village ---
export function generateDetailsReportPdf(
  rows: ReportDetailRow[],
  groupSector: string,
  period: string,
  org: string
): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const tableWidth = A4_W - MARGIN_LEFT_DETAILS - MARGIN_SMALL;
  const colRatios = [22, 10, 2, 22, 5, 3, 12, 8]; // Sign kam taaki Date PDF ke andar rahe
  const sumR = colRatios.reduce((a, b) => a + b, 0);
  const colWidths = colRatios.map((r, i) =>
    Math.max((tableWidth * r) / sumR, i === 2 ? 6 : i === 5 ? 8 : 0)
  );
  const detailsColStyles: Record<number, object> = {};
  colWidths.forEach((w, i) => {
    detailsColStyles[i] = {
      cellWidth: w,
      halign: i === 2 || i === 5 ? 'center' : i === 0 || i === 1 || i === 3 || i === 6 ? 'left' : 'center',
      overflow: i === 0 || i === 1 || i === 3 ? 'linebreak' : 'hidden',
    };
  });

  const sorted = sortedDetailRows(rows);
  const bySectorVillage = new Map<string, ReportDetailRow[]>();
  for (const r of sorted) {
    const key = `${r.group_sector}|||${r.village}`;
    if (!bySectorVillage.has(key)) bySectorVillage.set(key, []);
    bySectorVillage.get(key)!.push(r);
  }
  const blockList = Array.from(bySectorVillage.entries()).map(([k, blockRows]) => ({
    key: k,
    blockRows,
    isSmall: blockRows.length + 3 <= SMALL_TABLE_MAX_ROWS,
  }));

  let startY = MARGIN_SMALL;
  const blockMargin = 4;
  let sectorsOnPage = 0;
  const pageSpans: { startPage: number; endPage: number }[] = [];

  for (let idx = 0; idx < blockList.length; idx++) {
    const { key, blockRows, isSmall } = blockList[idx];
    const nextBlock = blockList[idx + 1];
    const nextIsSmall = nextBlock?.isSmall ?? false;

    if (sectorsOnPage >= 2 || (sectorsOnPage === 1 && !nextIsSmall)) {
      doc.addPage();
      startY = MARGIN_SMALL;
      sectorsOnPage = 0;
    }
    if (sectorsOnPage === 1 && nextIsSmall) {
      if (startY < SECOND_BLOCK_START_Y - 10) startY = SECOND_BLOCK_START_Y;
      else {
        doc.addPage();
        startY = MARGIN_SMALL;
        sectorsOnPage = 0;
      }
    }
    if (startY > MARGIN_SMALL) startY += blockMargin;

    const tableStartY = startY + 2;
    if (A4_H - tableStartY - PAGE_BOTTOM_MARGIN_MM < MIN_SPACE_FOR_BLOCK_MM) {
      doc.addPage();
      startY = MARGIN_SMALL;
      sectorsOnPage = 0;
    }

    const [gs, , , vill] = key.split('|');
    const groupKey = `${gs} - ${vill}`;

    const formattedPeriod = formatPeriodForDisplay(period);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(groupKey, MARGIN_LEFT_DETAILS, startY);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(`${formattedPeriod} | Summary: VK`, A4_W - MARGIN_SMALL, startY, { align: 'right' });
    startY += 2;

    const head = [['Name', 'Mo.No.', 'LG', 'Book Name', 'AVL', 'QTY', 'Sign', 'Date']];
    const body = buildDetailsBodyWithMerges(blockRows);

    const startPage = doc.getCurrentPageInfo().pageNumber;
    const lastColDetails = 7;
    autoTable(doc, {
      head,
      body,
      startY,
      margin: { left: MARGIN_LEFT_DETAILS, right: MARGIN_SMALL },
      tableWidth,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 1, minCellHeight: ROW_HEIGHT_MM, lineWidth: LINE_NORMAL },
      bodyStyles: { valign: 'middle' },
      headStyles: { fillColor: [220, 220, 220], fontStyle: 'bold', textColor: [0, 0, 0], halign: 'center' },
      columnStyles: detailsColStyles,
      didParseCell: (data) => {
        const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
        const isLeft = data.column.index === 0;
        const isRight = data.column.index + colSpan - 1 === lastColDetails;
        const isHead = data.section === 'head';
        const isLastRow = data.section === 'body' && data.row.index === body.length - 1;
        data.cell.styles.lineWidth = {
          top: isHead ? LINE_THICK : LINE_NORMAL,
          bottom: isHead || isLastRow ? LINE_THICK : LINE_NORMAL,
          left: isLeft ? LINE_THICK : LINE_NORMAL,
          right: isRight ? LINE_THICK : LINE_NORMAL,
        };
        const MIN_FONT = 7;
        if (data.section === 'body') {
          const fitCols = [0, 3];
          if (fitCols.includes(data.column.index)) {
            const text =
              Array.isArray((data.cell as any).text) && (data.cell as any).text.length
                ? (data.cell as any).text.join(' ')
                : String((data.cell as any).text ?? '');
            if (!text.trim()) return;
            const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
            const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
            const available = Math.max(2, cw - pad);
            let fs = (data.cell.styles.fontSize as number) ?? 8;
            doc.setFontSize(fs);
            let w = doc.getTextWidth(text);
            while (w > available && fs > MIN_FONT) {
              fs--;
              doc.setFontSize(fs);
              w = doc.getTextWidth(text);
            }
            data.cell.styles.fontSize = fs;
            data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
            doc.setFontSize(8);
          }
        }
        if (data.section === 'head' && [0, 3].includes(data.column.index)) {
          const text =
            Array.isArray((data.cell as any).text) && (data.cell as any).text.length
              ? (data.cell as any).text.join(' ')
              : String((data.cell as any).text ?? '');
          if (!text.trim()) return;
          const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
          const available = Math.max(2, cw - pad);
          let fs = (data.cell.styles.fontSize as number) ?? 8;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(text);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(text);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
          doc.setFontSize(8);
        }
      },
    });
    const endPage = doc.getCurrentPageInfo().pageNumber;
    if (endPage > startPage) pageSpans.push({ startPage, endPage });
    startY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 4;

    const totalQty = blockRows.reduce((s, r) => s + r.qty, 0);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.text(`Total Qty: ${totalQty}`, MARGIN_LEFT_DETAILS, startY);
    startY += 8;
    sectorsOnPage++;
  }

  addBlockPageNumbers(doc, pageSpans);
  return Buffer.from(doc.output('arraybuffer'));
}

// Natural sort for group names: E1, E2, ..., E10 (not E1, E10, E2)
function naturalCompare(a: string, b: string): number {
  const regex = /(\d+)|(\D+)/g;
  const aParts = a.match(regex) || [];
  const bParts = b.match(regex) || [];

  for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
    const aPart = aParts[i] || '';
    const bPart = bParts[i] || '';

    if (/^\d+$/.test(aPart) && /^\d+$/.test(bPart)) {
      const diff = parseInt(aPart, 10) - parseInt(bPart, 10);
      if (diff !== 0) return diff;
    } else {
      if (aPart < bPart) return -1;
      if (aPart > bPart) return 1;
    }
  }
  return 0;
}

// --- Summary Report (SectorVillageSummary): 6 cols LG, BOOK TITLE, AVAIL, QTY, SIGN, DATE ---
function buildSummaryRows(
  rows: ReportDetailRow[],
  groupBy: 'sector_village' | 'sector' | 'city'
): { key: string; rows: { lg: string; book: string; avail: string; qty: number }[] }[] {
  const map = new Map<string, { lg: string; book: string; avail: string; qty: number }[]>();

  for (const r of rows) {
    let key: string;
    if (groupBy === 'sector_village') key = `${r.group_sector}|||${r.village}`;
    else if (groupBy === 'sector') key = r.group_sector || 'All';
    else key = '__city__';

    const lg = getLanguageShortCode(r.language);
    const avail = availToTextSummary(r.avail);
    const book = r.book_name || '';

    if (!map.has(key)) map.set(key, []);
    const arr = map.get(key)!;
    const existing = arr.find((x) => x.lg === lg && x.book === book && x.avail === avail);
    if (existing) existing.qty += r.qty;
    else arr.push({ lg, book, avail, qty: r.qty });
  }

  const result: { key: string; rows: { lg: string; book: string; avail: string; qty: number }[] }[] = [];
  const availOrder = (a: string) => (a === 'YES' ? 0 : a === 'NS' ? 1 : 2);
  for (const [key, arr] of map) {
    arr.sort((a, b) => {
      const av = availOrder(a.avail) - availOrder(b.avail);
      if (av !== 0) return av;
      const lg = a.lg.localeCompare(b.lg);
      if (lg !== 0) return lg;
      return a.book.localeCompare(b.book);
    });
    result.push({ key, rows: arr });
  }
  result.sort((a, b) => a.key.localeCompare(b.key));
  return result;
}

// Build body rows with merge simulation: LG only on first row of language group; YES only on first row of YES block; SIGN shows [total] on first row of YES block
function buildSummaryBodyWithMerges(
  summaryRows: { lg: string; book: string; avail: string; qty: number }[]
): (string | number)[][] {
  const body: (string | number)[][] = [];
  let prevLg = '';
  let yesStartIdx = -1;
  let yesTotal = 0;
  for (let i = 0; i < summaryRows.length; i++) {
    const r = summaryRows[i];
    const lgShow = r.lg !== prevLg ? r.lg : '';
    prevLg = r.lg;
    let availShow = r.avail;
    let signShow = '';
    if (r.avail === 'YES') {
      if (yesStartIdx === -1) {
        yesStartIdx = body.length;
        yesTotal = r.qty;
        availShow = 'YES';
      } else {
        yesTotal += r.qty;
        availShow = '';
      }
    } else {
      if (yesStartIdx >= 0) {
        (body[yesStartIdx] as string[])[4] = `[${yesTotal}]`;
        yesStartIdx = -1;
      }
    }
    body.push([lgShow, r.book, availShow, String(r.qty), signShow, '']);
  }
  if (yesStartIdx >= 0) (body[yesStartIdx] as string[])[4] = `[${yesTotal}]`;
  return body;
}

// City: LG merge + YES block merge (AVAIL, SIGN with [total]) - same as VBA SummaryCityReport
function buildCityBodyWithMerges(
  rows: { lg: string; book: string; avail: string; qty: number }[]
): (string | number)[][] {
  const body: (string | number)[][] = [];
  let prevLg = '';
  let yesStartIdx = -1;
  let yesTotal = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const lgShow = r.lg !== prevLg ? r.lg : '';
    prevLg = r.lg;
    let availShow = r.avail;
    let signShow = '';
    if (r.avail === 'YES') {
      if (yesStartIdx === -1) {
        yesStartIdx = body.length;
        yesTotal = r.qty;
        availShow = 'YES';
      } else {
        yesTotal += r.qty;
        availShow = '';
      }
    } else {
      if (yesStartIdx >= 0) {
        (body[yesStartIdx] as string[])[4] = `[${yesTotal}]`;
        yesStartIdx = -1;
      }
    }
    body.push([lgShow, r.book, availShow, String(r.qty), signShow]);
  }
  if (yesStartIdx >= 0) (body[yesStartIdx] as string[])[4] = `[${yesTotal}]`;
  return body;
}

export function generateSummaryReportPdf(
  rows: ReportDetailRow[],
  groupSector: string,
  period: string,
  org: string
): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const tableWidth = A4_W - MARGIN_LEFT_SUMMARY - MARGIN_SMALL;
  const colRatios = [4, 35, 8, 5, 25, 12];
  const sumR = colRatios.reduce((a, b) => a + b, 0);
  const colWidths = colRatios.map((r, i) => Math.max((tableWidth * r) / sumR, i === 0 ? 6 : i === 3 ? 8 : 0));
  const colStyles: Record<number, object> = {};
  colWidths.forEach((w, i) => {
    colStyles[i] = {
      cellWidth: w,
      halign: i === 0 || i === 2 || i === 3 || i === 4 || i === 5 ? 'center' : 'left',
      overflow: i === 1 ? 'linebreak' : 'hidden',
    };
  });

  const groups = buildSummaryRows(rows, 'sector_village');
  const groupList = groups.map((g) => ({ ...g, dataRows: g.rows.length + 2, isSmall: g.rows.length + 2 <= SMALL_TABLE_MAX_ROWS }));
  let startY = MARGIN_SMALL;
  let sectorsOnPage = 0;
  const pageSpans: { startPage: number; endPage: number }[] = [];

  for (let idx = 0; idx < groupList.length; idx++) {
    const { key, rows: summaryRows, isSmall } = groupList[idx];
    const nextGroup = groupList[idx + 1];
    const nextIsSmall = nextGroup?.isSmall ?? false;

    if (sectorsOnPage >= 2 || (sectorsOnPage === 1 && !nextIsSmall)) {
      doc.addPage();
      startY = MARGIN_SMALL;
      sectorsOnPage = 0;
    }
    if (sectorsOnPage === 1 && nextIsSmall) {
      if (startY < SECOND_BLOCK_START_Y - 10) startY = SECOND_BLOCK_START_Y;
      else {
        doc.addPage();
        startY = MARGIN_SMALL;
        sectorsOnPage = 0;
      }
    }

    const tableStartY = startY + 2;
    if (A4_H - tableStartY - PAGE_BOTTOM_MARGIN_MM < MIN_SPACE_FOR_BLOCK_MM) {
      doc.addPage();
      startY = MARGIN_SMALL;
      sectorsOnPage = 0;
    }

    const [gs, , , vill] = key.split('|');
    const groupKey = `${gs} - ${vill}`;

    const formattedPeriod = formatPeriodForDisplay(period);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(groupKey, MARGIN_LEFT_SUMMARY, startY);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(`${formattedPeriod} | Summary: VK`, A4_W - MARGIN_SMALL, startY, { align: 'right' });
    startY += 2;

    const head = [['LG', 'BOOK TITLE', 'AVAIL', 'QTY', 'SIGN', 'DATE']];
    const body = buildSummaryBodyWithMerges(summaryRows);
    const totalQty = summaryRows.reduce((s, r) => s + r.qty, 0);
    body.push(['', 'Total Qty', '', String(totalQty), '', '']);

    const startPage = doc.getCurrentPageInfo().pageNumber;
    const lastColSummary = 5;
    autoTable(doc, {
      head,
      body,
      startY,
      margin: { left: MARGIN_LEFT_SUMMARY, right: MARGIN_SMALL },
      tableWidth,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 1, minCellHeight: ROW_HEIGHT_MM, lineWidth: LINE_NORMAL },
      bodyStyles: { valign: 'middle' },
      headStyles: { fillColor: [220, 220, 220], fontStyle: 'bold', textColor: [0, 0, 0], halign: 'center' },
      columnStyles: colStyles,
      didParseCell: (data) => {
        const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
        const isLeft = data.column.index === 0;
        const isRight = data.column.index + colSpan - 1 === lastColSummary;
        const isHead = data.section === 'head';
        const isLastRow = data.section === 'body' && data.row.index === body.length - 1;
        data.cell.styles.lineWidth = {
          top: isHead ? LINE_THICK : LINE_NORMAL,
          bottom: isHead || isLastRow ? LINE_THICK : LINE_NORMAL,
          left: isLeft ? LINE_THICK : LINE_NORMAL,
          right: isRight ? LINE_THICK : LINE_NORMAL,
        };
        const MIN_FONT = 7;
        if (data.section === 'body') {
          const fitCols = [1, 4];
          if (fitCols.includes(data.column.index)) {
            const text =
              Array.isArray((data.cell as any).text) && (data.cell as any).text.length
                ? (data.cell as any).text.join(' ')
                : String((data.cell as any).text ?? '');
            if (!text.trim()) return;
            const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
            const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
            const available = Math.max(2, cw - pad);
            let fs = (data.cell.styles.fontSize as number) ?? 8;
            doc.setFontSize(fs);
            let w = doc.getTextWidth(text);
            while (w > available && fs > MIN_FONT) {
              fs--;
              doc.setFontSize(fs);
              w = doc.getTextWidth(text);
            }
            data.cell.styles.fontSize = fs;
            data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
            doc.setFontSize(8);
          }
        }
        if (data.section === 'head' && [1, 4].includes(data.column.index)) {
          const text =
            Array.isArray((data.cell as any).text) && (data.cell as any).text.length
              ? (data.cell as any).text.join(' ')
              : String((data.cell as any).text ?? '');
          if (!text.trim()) return;
          const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
          const available = Math.max(2, cw - pad);
          let fs = (data.cell.styles.fontSize as number) ?? 8;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(text);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(text);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
          doc.setFontSize(8);
        }
      },
    });
    const endPage = doc.getCurrentPageInfo().pageNumber;
    if (endPage > startPage) pageSpans.push({ startPage, endPage });
    startY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 4;
    sectorsOnPage++;
  }

  addBlockPageNumbers(doc, pageSpans);
  return Buffer.from(doc.output('arraybuffer'));
}

// --- Sector Summary Report: 6 cols, group by sector ---
export function generateSectorSummaryReportPdf(
  rows: ReportDetailRow[],
  groupSector: string,
  period: string,
  org: string
): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const tableWidth = A4_W - MARGIN_LEFT_SUMMARY - MARGIN_SMALL;
  const colRatios = [4, 35, 8, 5, 25, 12];
  const sumR = colRatios.reduce((a, b) => a + b, 0);
  const colWidths = colRatios.map((r, i) => Math.max((tableWidth * r) / sumR, i === 0 ? 6 : i === 3 ? 8 : 0));
  const colStyles: Record<number, object> = {};
  colWidths.forEach((w, i) => {
    colStyles[i] = {
      cellWidth: w,
      halign: i === 0 || i === 2 || i === 3 || i === 4 || i === 5 ? 'center' : 'left',
      overflow: i === 1 ? 'linebreak' : 'hidden',
    };
  });

  const groups = buildSummaryRows(rows, 'sector');
  const groupList = groups.map((g) => ({ ...g, isSmall: g.rows.length + 2 <= SMALL_TABLE_MAX_ROWS }));
  let startY = MARGIN_SMALL;
  let sectorsOnPage = 0;
  const pageSpans: { startPage: number; endPage: number }[] = [];

  for (let idx = 0; idx < groupList.length; idx++) {
    const { key: sectorKey, rows: summaryRows, isSmall } = groupList[idx];
    const nextGroup = groupList[idx + 1];
    const nextIsSmall = nextGroup?.isSmall ?? false;

    if (sectorsOnPage >= 2 || (sectorsOnPage === 1 && !nextIsSmall)) {
      doc.addPage();
      startY = MARGIN_SMALL;
      sectorsOnPage = 0;
    }
    if (sectorsOnPage === 1 && nextIsSmall) {
      if (startY < SECOND_BLOCK_START_Y - 10) startY = SECOND_BLOCK_START_Y;
      else {
        doc.addPage();
        startY = MARGIN_SMALL;
        sectorsOnPage = 0;
      }
    }

    const tableStartY = startY + 2;
    if (A4_H - tableStartY - PAGE_BOTTOM_MARGIN_MM < MIN_SPACE_FOR_BLOCK_MM) {
      doc.addPage();
      startY = MARGIN_SMALL;
      sectorsOnPage = 0;
    }

    const formattedPeriod = formatPeriodForDisplay(period);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(sectorKey, MARGIN_LEFT_SUMMARY, startY);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(`${formattedPeriod} | Summary: Sector`, A4_W - MARGIN_SMALL, startY, { align: 'right' });
    startY += 2;

    const head = [['LG', 'BOOK TITLE', 'AVAIL', 'QTY', 'SIGN', 'DATE']];
    const body = buildSummaryBodyWithMerges(summaryRows);
    const totalQty = summaryRows.reduce((s, r) => s + r.qty, 0);
    body.push(['', 'Total Qty', '', String(totalQty), '', '']);

    const startPage = doc.getCurrentPageInfo().pageNumber;
    const lastColSector = 5;
    autoTable(doc, {
      head,
      body,
      startY,
      margin: { left: MARGIN_LEFT_SUMMARY, right: MARGIN_SMALL },
      tableWidth,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 1, minCellHeight: ROW_HEIGHT_MM, lineWidth: LINE_NORMAL },
      bodyStyles: { valign: 'middle' },
      headStyles: { fillColor: [220, 220, 220], fontStyle: 'bold', textColor: [0, 0, 0], halign: 'center' },
      columnStyles: colStyles,
      didParseCell: (data) => {
        const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
        const isLeft = data.column.index === 0;
        const isRight = data.column.index + colSpan - 1 === lastColSector;
        const isHead = data.section === 'head';
        const isLastRow = data.section === 'body' && data.row.index === body.length - 1;
        data.cell.styles.lineWidth = {
          top: isHead ? LINE_THICK : LINE_NORMAL,
          bottom: isHead || isLastRow ? LINE_THICK : LINE_NORMAL,
          left: isLeft ? LINE_THICK : LINE_NORMAL,
          right: isRight ? LINE_THICK : LINE_NORMAL,
        };
        const MIN_FONT = 7;
        if (data.section === 'body') {
          const fitCols = [1, 4];
          if (fitCols.includes(data.column.index)) {
            const text =
              Array.isArray((data.cell as any).text) && (data.cell as any).text.length
                ? (data.cell as any).text.join(' ')
                : String((data.cell as any).text ?? '');
            if (!text.trim()) return;
            const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
            const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
            const available = Math.max(2, cw - pad);
            let fs = (data.cell.styles.fontSize as number) ?? 8;
            doc.setFontSize(fs);
            let w = doc.getTextWidth(text);
            while (w > available && fs > MIN_FONT) {
              fs--;
              doc.setFontSize(fs);
              w = doc.getTextWidth(text);
            }
            data.cell.styles.fontSize = fs;
            data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
            doc.setFontSize(8);
          }
        }
        if (data.section === 'head' && [1, 4].includes(data.column.index)) {
          const text =
            Array.isArray((data.cell as any).text) && (data.cell as any).text.length
              ? (data.cell as any).text.join(' ')
              : String((data.cell as any).text ?? '');
          if (!text.trim()) return;
          const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
          const available = Math.max(2, cw - pad);
          let fs = (data.cell.styles.fontSize as number) ?? 8;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(text);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(text);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
          doc.setFontSize(8);
        }
      },
    });
    const endPage = doc.getCurrentPageInfo().pageNumber;
    if (endPage > startPage) pageSpans.push({ startPage, endPage });
    startY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 4;
    sectorsOnPage++;
  }

  addBlockPageNumbers(doc, pageSpans);
  return Buffer.from(doc.output('arraybuffer'));
}

// --- City Summary (All Summary): 5 cols LG, BOOK TITLE, AVAIL, QTY, SIGN ---
export function generateAllSummaryReportPdf(
  rows: ReportDetailRow[],
  period: string,
  org: string
): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const tableWidth = A4_W - MARGIN_LEFT_SUMMARY - MARGIN_SMALL;
  const colRatios = [8, 35, 14, 8, 25];
  const sumR = colRatios.reduce((a, b) => a + b, 0);
  const colWidths = colRatios.map((r, i) =>
    Math.max((tableWidth * r) / sumR, i === 0 ? 6 : i === 3 ? 8 : 0)
  );
  const cityColStyles: Record<number, object> = {};
  colWidths.forEach((w, i) => {
    cityColStyles[i] = {
      cellWidth: w,
      halign: i === 0 || i === 2 || i === 3 || i === 4 ? 'center' : 'left',
      overflow: i === 1 ? 'linebreak' : 'hidden',
    };
  });

  const groups = buildSummaryRows(rows, 'city');
  const cityRows = groups.length ? groups[0].rows : [];
  const totalQty = cityRows.reduce((s, r) => s + r.qty, 0);

  let startY = MARGIN_SMALL;
  const formattedPeriod = formatPeriodForDisplay(period);
  doc.setFontSize(12);
  doc.setFont('helvetica', 'bold');
  doc.text('City Summary', MARGIN_LEFT_SUMMARY, startY);
  const periodStr =
    period && period !== 'All'
      ? `${org}: ${formattedPeriod} - Summary`
      : `${org}: All Periods - Summary`;
  doc.text(periodStr, A4_W - MARGIN_SMALL, startY, { align: 'right' });
  startY += 2;

  const head = [['LG', 'BOOK TITLE', 'AVAIL', 'QTY', 'SIGN']];
  const body = buildCityBodyWithMerges(cityRows);
  body.push(['', 'Total Qty', '', String(totalQty), '']);

  const startPage = doc.getCurrentPageInfo().pageNumber;
  const lastColCity = 4;
  autoTable(doc, {
    head,
    body,
    startY,
    margin: { left: MARGIN_LEFT_SUMMARY, right: MARGIN_SMALL },
    tableWidth,
    theme: 'grid',
    styles: { fontSize: 8, cellPadding: 1, minCellHeight: ROW_HEIGHT_MM, lineWidth: LINE_NORMAL },
    bodyStyles: { valign: 'middle' },
    headStyles: { fillColor: [220, 220, 220], fontStyle: 'bold', textColor: [0, 0, 0], halign: 'center' },
    columnStyles: cityColStyles,
    didParseCell: (data) => {
      const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
      const isLeft = data.column.index === 0;
      const isRight = data.column.index + colSpan - 1 === lastColCity;
      const isHead = data.section === 'head';
      const isLastRow = data.section === 'body' && data.row.index === body.length - 1;
      data.cell.styles.lineWidth = {
        top: isHead ? LINE_THICK : LINE_NORMAL,
        bottom: isHead || isLastRow ? LINE_THICK : LINE_NORMAL,
        left: isLeft ? LINE_THICK : LINE_NORMAL,
        right: isRight ? LINE_THICK : LINE_NORMAL,
      };
      const MIN_FONT = 7;
      if (data.section === 'body') {
        const fitCols = [1, 4];
        if (fitCols.includes(data.column.index)) {
          const text =
            Array.isArray((data.cell as any).text) && (data.cell as any).text.length
              ? (data.cell as any).text.join(' ')
              : String((data.cell as any).text ?? '');
          if (!text.trim()) return;
          const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
          const available = Math.max(2, cw - pad);
          let fs = (data.cell.styles.fontSize as number) ?? 8;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(text);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(text);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
          doc.setFontSize(8);
        }
      }
      if (data.section === 'head' && [1, 4].includes(data.column.index)) {
        const text =
          Array.isArray((data.cell as any).text) && (data.cell as any).text.length
            ? (data.cell as any).text.join(' ')
            : String((data.cell as any).text ?? '');
        if (!text.trim()) return;
        const cw = (data.cell.styles.cellWidth as number) ?? colWidths[data.column.index] ?? 40;
        const pad = ((data.cell.styles.cellPadding as number) ?? 2) * 2;
        const available = Math.max(2, cw - pad);
        let fs = (data.cell.styles.fontSize as number) ?? 8;
        doc.setFontSize(fs);
        let w = doc.getTextWidth(text);
        while (w > available && fs > MIN_FONT) {
          fs--;
          doc.setFontSize(fs);
          w = doc.getTextWidth(text);
        }
        data.cell.styles.fontSize = fs;
        data.cell.styles.overflow = w <= available ? 'hidden' : 'linebreak';
        doc.setFontSize(8);
      }
    },
  });

  const endPage = doc.getCurrentPageInfo().pageNumber;
  const pageSpans =
    endPage > startPage ? [{ startPage, endPage }] : [];
  addBlockPageNumbers(doc, pageSpans);
  return Buffer.from(doc.output('arraybuffer'));
}

// --- Book Group Summary: Landscape, 6 cols Sr, Group, Total Qty, Name, Mobile No., Date ---
export interface BookGroupSummaryRow {
  group: string;
  total_qty: number;
}

export function generateBookGroupSummaryPdf(
  rows: BookGroupSummaryRow[],
  period: string,
  org: string
): Buffer {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  const marginLeft = 14;
  const marginRight = 14;
  const pageWidth = doc.internal.pageSize.getWidth();
  const tableWidth = pageWidth - marginLeft - marginRight;

  // Column width percentages: 5, 15, 10, 45, 12.5, 12.5
  const perc = [5, 15, 10, 45, 12.5, 12.5];
  const colWidths = perc.map((p) => (tableWidth * p) / 100);

  // Header: "{APP} - {Period}"
  const titlePeriod = formatPeriodForDisplay(period);
  const headerText = `${org} - ${titlePeriod}`;
  doc.setFontSize(14);
  doc.setFont('helvetica', 'bold');
  doc.text(headerText, pageWidth / 2, 12, { align: 'center' });

  const head = [['Sr.', 'Group', 'Total Qty', 'Name', 'Mobile No.', 'Date']];

  const sorted = [...rows].sort((a, b) =>
    naturalCompare(a.group || '', b.group || '')
  );

  const body: (string | number)[][] = [];
  let sr = 1;
  let grandTotal = 0;
  for (const r of sorted) {
    const qty = Number((r as any).total_qty) || 0;
    grandTotal += qty;
    body.push([sr, r.group || '', qty || '', '', '', '']);
    sr += 1;
  }
  body.push(['', 'Total', grandTotal || '', '', '', '']);

  autoTable(doc, {
    head,
    body,
    startY: 20,
    theme: 'grid',
    styles: {
      fontSize: 9,
      cellPadding: 1.5,
      halign: 'center',
      valign: 'middle',
      font: 'helvetica',
      minCellHeight: 5,
      lineWidth: 0.4,
      lineColor: [0, 0, 0],
    },
    headStyles: {
      fillColor: [217, 217, 217],
      textColor: [0, 0, 0],
      fontStyle: 'bold',
      minCellHeight: 5.5,
      lineWidth: 0.4,
      lineColor: [0, 0, 0],
    },
    columnStyles: {
      0: { cellWidth: colWidths[0] },
      1: { cellWidth: colWidths[1], halign: 'left' },
      2: { cellWidth: colWidths[2] },
      3: { cellWidth: colWidths[3], halign: 'left' },
      4: { cellWidth: colWidths[4] },
      5: { cellWidth: colWidths[5] },
    },
    margin: { left: marginLeft, right: marginRight },
    didParseCell: (data) => {
      const isLastRow = data.section === 'body' && data.row.index === body.length - 1;
      if (isLastRow) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.fillColor = [230, 230, 230];
      }
      if (data.section === 'body' && data.column.index === 1) {
        data.cell.styles.halign = 'left';
      }
      if (data.section === 'body' && data.column.index === 3) {
        data.cell.styles.halign = 'left';
      }
    },
  });

  return Buffer.from(doc.output('arraybuffer'));
}
