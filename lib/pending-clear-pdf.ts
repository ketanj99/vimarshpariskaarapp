import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { query, DB_SCHEMA } from './db';

const INCH_MM = 25.4;
const MARGIN_LEFT = 0.3 * INCH_MM;
const MARGIN_RIGHT = 0.1 * INCH_MM;
const A4_W = 210;
const A4_H = 297;
const TABLE_WIDTH = A4_W - MARGIN_LEFT - MARGIN_RIGHT;
const SECOND_BLOCK_START_Y = 155;
const SMALL_TABLE_MAX_ROWS = 23;
const ROW_HEIGHT = 17.5 / 7;
// Header thoda niche from top; table turant header ke niche (no gap)
const HEADER_TOP_MM = 10;
const GAP_HEADER_TO_TABLE_MM = 1.5;

export interface PendingClearRow {
  book_id: number;
  book_name: string;
  display_name: string;
  language: string;
  year: number;
  month: number;
  half: string;
  group_sector: string;
  app: string;
  total_qty: number;
  stock_number: string;
  village_breakdown: string;
}

/** Village-wise row for Pending Clear Village PDF (group_sector - village tables) */
export interface PendingClearVillageRow {
  group_sector: string;
  village: string;
  stock_number: string;
  book_name: string;
  display_name: string;
  language: string;
  year: number;
  month: number;
  half: string;
  qty: number;
  app: string;
}

function getLanguageShortCode(lang: string): string {
  const s = (lang || '').trim().toLowerCase();
  if (s.startsWith('gujarati')) return 'GJ';
  if (s.startsWith('hindi')) return 'HN';
  if (s.startsWith('marathi')) return 'MR';
  if (s.startsWith('telugu') || s.startsWith('telagu')) return 'TL';
  if (s.startsWith('tamil')) return 'TM';
  return (lang || '').slice(0, 2).toUpperCase();
}

/** LG order in village table: GJ, HN, MR, TL, TM, then others */
function getLanguageSortOrder(lang: string): number {
  const code = getLanguageShortCode(lang);
  const order: Record<string, number> = { GJ: 0, HN: 1, MR: 2, TL: 3, TM: 4 };
  return order[code] ?? 99;
}

/** Natural sort so E1, E2, ... E9, E10 (not E1, E10, E2) */
function naturalCompare(a: string, b: string): number {
  const parts = (s: string) => s.split(/(\d+)/).filter(Boolean);
  const aa = parts(a);
  const bb = parts(b);
  for (let i = 0; i < Math.min(aa.length, bb.length); i++) {
    const na = parseInt(aa[i], 10);
    const nb = parseInt(bb[i], 10);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      if (na !== nb) return na - nb;
    } else {
      const c = aa[i].localeCompare(bb[i]);
      if (c !== 0) return c;
    }
  }
  return aa.length - bb.length;
}

function formatPeriod(year: number, month: number, half: string): string {
  const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthName = MONTH_NAMES[month - 1] || '';
  const h = (half || '').replace(/\s/g, '').toLowerCase();
  if (h === '1&2' || h === '1and2') {
    const lastDay = new Date(year, month, 0).getDate();
    return `${year}-${monthName} (1 To ${lastDay})`;
  }
  if (half === '1') return `${year}-${monthName} (1 To 15)`;
  if (half === '2') {
    const lastDay = new Date(year, month, 0).getDate();
    return `${year}-${monthName} (16 To ${lastDay})`;
  }
  return `${year}-${monthName} (${year}-${String(month).padStart(2, '0')}-${half})`;
}

function toCamelCase(s: string): string {
  const t = (s || '').trim();
  if (!t) return '';
  return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
}

function shortVillageName(full: string, maxLen: number = 6): string {
  return toCamelCase(full).slice(0, maxLen);
}

export async function getPendingClearSectorData(
  pNos: string[],
  app?: string,
  groupSector?: string
): Promise<PendingClearRow[]> {
  if (!pNos.length) return [];
  const pNoList = pNos.map((p) => `'${String(p).replace(/'/g, "''")}'`).join(',');
  const appFilter = app && app !== 'All' ? `AND t.app = '${String(app).replace(/'/g, "''")}'` : '';
  const groupFilter =
    groupSector && groupSector !== 'All' ? `AND t.group_sector = '${String(groupSector).replace(/'/g, "''")}'` : '';

  const mainSql = `
    SELECT
      b.book_id,
      b.book_name,
      COALESCE(b.display_name, b.book_name) AS display_name,
      b.language,
      t.app,
      t.year,
      t.month,
      t.half,
      t.group_sector,
      SUM(t.qty)::int AS total_qty,
      t.stock_number
    FROM ${DB_SCHEMA}.transactions t
    INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
    WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE
      AND t.avail = 0
      AND COALESCE(t.stock_number, '') <> ''
      AND t.group_sector IS NOT NULL AND TRIM(t.group_sector) <> ''
      AND t.stock_number IN (${pNoList})
      ${appFilter}
      ${groupFilter}
    GROUP BY b.book_id, b.book_name, b.language, t.app, t.year, t.month, t.half, t.group_sector, t.stock_number
    ORDER BY t.group_sector ASC, t.app ASC, b.language ASC, b.book_name ASC, t.year DESC, t.month DESC, t.half DESC
  `;
  const mainResult = await query<{
    book_id: number;
    book_name: string;
    display_name: string | null;
    language: string;
    app: string;
    year: number;
    month: number;
    half: string;
    group_sector: string;
    total_qty: number;
    stock_number: string;
  }>(mainSql);

  const villageSql = `
    SELECT
      t.book_id,
      t.year,
      t.month,
      t.half,
      t.group_sector,
      t.stock_number,
      t.village,
      SUM(t.qty)::int AS village_qty
    FROM ${DB_SCHEMA}.transactions t
    INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
    WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE
      AND t.avail = 0
      AND COALESCE(t.stock_number, '') <> ''
      AND t.stock_number IN (${pNoList})
      AND t.village IS NOT NULL AND TRIM(t.village) <> ''
      ${appFilter}
      ${groupFilter}
    GROUP BY t.book_id, t.year, t.month, t.half, t.group_sector, t.stock_number, t.village
    ORDER BY t.book_id, t.year, t.month, t.half, t.group_sector, t.stock_number, t.village
  `;
  const villageResult = await query<{
    book_id: number;
    year: number;
    month: number;
    half: string;
    group_sector: string;
    stock_number: string;
    village: string;
    village_qty: number;
  }>(villageSql);

  const villageByKey = new Map<string, { village: string; qty: number }[]>();
  for (const v of villageResult.rows || []) {
    const key = `${v.book_id}|${v.year}|${v.month}|${v.half}|${v.group_sector}|${v.stock_number}`;
    if (!villageByKey.has(key)) villageByKey.set(key, []);
    villageByKey.get(key)!.push({ village: v.village, qty: v.village_qty });
  }

  const rows: PendingClearRow[] = (mainResult.rows || []).map((r) => {
    const key = `${r.book_id}|${r.year}|${r.month}|${r.half}|${r.group_sector}|${r.stock_number}`;
    const villages = villageByKey.get(key) || [];
    const shortNames = new Map<string, string>();
    const count = new Map<string, number>();
    for (const v of villages) {
      const short = shortVillageName(v.village);
      shortNames.set(v.village, short);
      count.set(short, (count.get(short) || 0) + 1);
    }
    const parts: string[] = [];
    for (const v of villages) {
      let name = shortNames.get(v.village)!;
      if ((count.get(name) || 0) > 1) name = shortVillageName(v.village, 8);
      parts.push(`${name}- ${v.qty}`);
    }
    const village_breakdown = parts.join(', ');
    return {
      book_id: r.book_id,
      book_name: r.book_name,
      display_name: (r.display_name ?? r.book_name) || '',
      language: r.language,
      app: r.app,
      year: r.year,
      month: r.month,
      half: r.half,
      group_sector: r.group_sector,
      total_qty: r.total_qty,
      stock_number: r.stock_number,
      village_breakdown,
    };
  });
  return rows;
}

export async function getPendingClearVillageData(
  pNos: string[],
  app?: string,
  groupSector?: string
): Promise<PendingClearVillageRow[]> {
  if (!pNos.length) return [];
  const pNoList = pNos.map((p) => `'${String(p).replace(/'/g, "''")}'`).join(',');
  const appFilter = app && app !== 'All' ? `AND t.app = '${String(app).replace(/'/g, "''")}'` : '';
  const groupFilter =
    groupSector && groupSector !== 'All' ? `AND t.group_sector = '${String(groupSector).replace(/'/g, "''")}'` : '';

  const sql = `
    SELECT
      t.group_sector,
      t.village,
      t.stock_number,
      b.book_name,
      COALESCE(b.display_name, b.book_name) AS display_name,
      b.language,
      t.year,
      t.month,
      t.half,
      SUM(t.qty)::int AS qty,
      t.app
    FROM ${DB_SCHEMA}.transactions t
    INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
    WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE
      AND t.avail = 0
      AND COALESCE(t.stock_number, '') <> ''
      AND t.group_sector IS NOT NULL AND TRIM(t.group_sector) <> ''
      AND t.village IS NOT NULL AND TRIM(t.village) <> ''
      AND t.stock_number IN (${pNoList})
      ${appFilter}
      ${groupFilter}
    GROUP BY t.group_sector, t.village, t.stock_number, b.book_name, b.display_name, b.language, t.year, t.month, t.half, t.app
    ORDER BY t.group_sector ASC, t.village ASC, t.stock_number ASC, b.language ASC, b.book_name ASC, t.year DESC, t.month DESC, t.half DESC
  `;
  const result = await query<PendingClearVillageRow>(sql);
  return result.rows || [];
}

function buildVillageTotalBreakdown(rows: PendingClearRow[]): string {
  const villageTotals = new Map<string, number>();
  for (const r of rows) {
    if (!r.village_breakdown) continue;
    const parts = r.village_breakdown.split(', ');
    for (const p of parts) {
      const dash = p.indexOf('-');
      if (dash > 0) {
        const name = p.slice(0, dash).trim();
        const qty = parseInt(p.slice(dash + 1).trim(), 10) || 0;
        if (name && qty) villageTotals.set(name, (villageTotals.get(name) || 0) + qty);
      }
    }
  }
  return Array.from(villageTotals.entries())
    .map(([name, qty]) => `${name}- ${qty}`)
    .join(', ');
}

type CellDef =
  | string
  | number
  | {
      content: string;
      rowSpan?: number;
      colSpan?: number;
      styles?: { valign?: 'top' | 'middle' | 'bottom' };
    };

// VBA-style merge with rowSpan so merged cell content is vertically middle-aligned
function buildBodyWithMerges(sectorRows: PendingClearRow[]): CellDef[][] {
  const n = sectorRows.length;
  if (n === 0) return [];

  const pNoRun: number[] = [];
  for (let i = 0; i < n; i++) {
    if (i > 0 && sectorRows[i].stock_number === sectorRows[i - 1].stock_number) {
      pNoRun.push(0);
    } else {
      let run = 1;
      while (i + run < n && sectorRows[i + run].stock_number === sectorRows[i].stock_number) run++;
      pNoRun.push(run);
    }
  }

  const bookLangRun: number[] = [];
  const signTotalForRun: number[] = [];
  for (let i = 0; i < n; i++) {
    const key = `${sectorRows[i].book_name}|${getLanguageShortCode(sectorRows[i].language)}`;
    const sameAsPrev = i > 0 && key === `${sectorRows[i - 1].book_name}|${getLanguageShortCode(sectorRows[i - 1].language)}`;
    if (sameAsPrev) {
      bookLangRun.push(0);
      signTotalForRun.push(0);
    } else {
      let run = 1;
      let total = sectorRows[i].total_qty;
      while (i + run < n && `${sectorRows[i + run].book_name}|${getLanguageShortCode(sectorRows[i + run].language)}` === key) {
        total += sectorRows[i + run].total_qty;
        run++;
      }
      bookLangRun.push(run);
      signTotalForRun.push(total);
    }
  }

  const body: CellDef[][] = [];
  const valignMiddle: { valign: 'top' | 'middle' | 'bottom' } = { valign: 'middle' };

  for (let i = 0; i < n; i++) {
    const r = sectorRows[i];
    const row: CellDef[] = [];

    if (pNoRun[i] > 0) {
      row.push({ content: r.stock_number, rowSpan: pNoRun[i], styles: valignMiddle });
    }
    if (bookLangRun[i] > 0) {
      const lg = getLanguageShortCode(r.language);
      row.push({
        // Always use original book name in PDFs (ignore display_name)
        content: (r.book_name || '').trim(),
        rowSpan: bookLangRun[i],
        styles: valignMiddle,
      });
      row.push({ content: lg, rowSpan: bookLangRun[i], styles: valignMiddle });
    }
    row.push(formatPeriod(r.year, r.month, r.half));
    row.push(String(r.total_qty));
    if (bookLangRun[i] > 0) {
      row.push({ content: `[${signTotalForRun[i]}]`, rowSpan: bookLangRun[i], styles: valignMiddle });
    }
    row.push(r.village_breakdown || '');

    body.push(row);
  }

  return body;
}

export interface PendingClearSectorOptions {
  includeSector?: boolean;
  includeSectorCity?: boolean;
}

export function generatePendingClearSectorPdf(
  rows: PendingClearRow[],
  options: PendingClearSectorOptions = {}
): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const colRatios = [5, 20, 4, 15, 5, 20, 26];
  const sumR = colRatios.reduce((a, b) => a + b, 0);
  const colWidths = colRatios.map((r) => (TABLE_WIDTH * r) / sumR);

  const bySector = new Map<string, PendingClearRow[]>();
  for (const r of rows) {
    const s = r.group_sector || '';
    if (!bySector.has(s)) bySector.set(s, []);
    bySector.get(s)!.push(r);
  }
  const sectors = Array.from(bySector.keys()).sort(naturalCompare);
  let sectorsOnPage = 0;
  let startY = HEADER_TOP_MM;
  const { includeSector = true, includeSectorCity = true } = options;
  let hasContent = false;

  if (includeSector) {
    for (let idx = 0; idx < sectors.length; idx++) {
    const sectorKey = sectors[idx];
    const sectorRows = bySector.get(sectorKey)!;
    const sectorRowsWithIdx = sectorRows.map((r, index) => ({ r, index }));
    sectorRowsWithIdx.sort((a, b) => {
      const appA = (a.r.app || '').toUpperCase();
      const appB = (b.r.app || '').toUpperCase();
      if (appA !== appB) return appA < appB ? -1 : 1;
      return a.index - b.index;
    });
    const orderedSectorRows = sectorRowsWithIdx.map((x) => x.r);

    const appSet = new Set<string>();
    for (const r of sectorRows) {
      if (r.app) appSet.add((r.app as string).toUpperCase());
    }
    const apps = Array.from(appSet).sort();
    const centerTitle = apps.length ? apps.join(' & ') : '';
    const nextSector = sectors[idx + 1];
    const nextRows = nextSector ? bySector.get(nextSector)!.length : 0;
    const estimatedRows = 3 + sectorRows.length + 1;
    const nextEstimated = 3 + nextRows + 1;

    if (sectorsOnPage >= 2 || (sectorsOnPage === 1 && (estimatedRows > 23 || nextEstimated > 23))) {
      doc.addPage();
      startY = HEADER_TOP_MM;
      sectorsOnPage = 0;
    }
    if (sectorsOnPage === 1 && estimatedRows <= 23 && nextEstimated <= 23) {
      if (startY < SECOND_BLOCK_START_Y - 15) startY = SECOND_BLOCK_START_Y;
      else {
        doc.addPage();
        startY = HEADER_TOP_MM;
        sectorsOnPage = 0;
      }
    }

    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 0);
    doc.text(sectorKey, MARGIN_LEFT, startY);
    doc.text(centerTitle, A4_W / 2, startY, { align: 'center' });
    doc.text('Pending -> Clear Summary : Sector', A4_W - MARGIN_RIGHT, startY, { align: 'right' });
    startY += GAP_HEADER_TO_TABLE_MM;

    const head = [['P No.', 'Book Name', 'LG', 'Period', 'Qty', 'Sign', 'VK Books']];

    const body: CellDef[][] = [];
    const appKey = (appName?: string) => (appName || '').toUpperCase();
    let currentGroupKey: string | null = null;
    let currentGroupRows: PendingClearRow[] = [];

    const flushGroup = () => {
      if (!currentGroupRows.length) return;
      const groupBody = buildBodyWithMerges(currentGroupRows);
      for (const row of groupBody) body.push(row);
    };

    for (const r of orderedSectorRows) {
      const key = appKey(r.app);
      if (currentGroupKey === null) {
        currentGroupKey = key;
        currentGroupRows = [r];
      } else if (key === currentGroupKey) {
        currentGroupRows.push(r);
      } else {
        flushGroup();
        body.push(['', '', '', '', '', '', '']);
        currentGroupKey = key;
        currentGroupRows = [r];
      }
    }
    flushGroup();

    const sectorTotal = sectorRows.reduce((s, r) => s + r.total_qty, 0);
    const villageTotalStr = buildVillageTotalBreakdown(sectorRows);
    body.push([
      '',
      '',
      '',
      'Total Qty',
      String(sectorTotal),
      { content: villageTotalStr, colSpan: 2 },
    ]);

    const LINE_NORMAL = 0.05;
    const LINE_THICK = 0.2;
    const lastColIndex = 6;

    autoTable(doc, {
      head,
      body,
      startY,
      margin: { left: MARGIN_LEFT, right: MARGIN_RIGHT + 5 }, // more right side white space
      tableWidth: TABLE_WIDTH - 10, // thoda chhota table width
      theme: 'grid',
      styles: {
        fontSize: 8,
        cellPadding: 1,
        minCellHeight: 3,
        lineWidth: LINE_NORMAL,
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0], // sab border dark black
      },
      bodyStyles: { valign: 'middle', textColor: [0, 0, 0], lineColor: [0, 0, 0] },
      headStyles: {
        fillColor: [220, 220, 220],
        fontStyle: 'bold',
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
      },
      columnStyles: {
        0: { halign: 'center', cellWidth: colWidths[0] },
        1: { halign: 'left', cellWidth: colWidths[1], overflow: 'hidden', fontSize: 7 },
        2: { halign: 'center', cellWidth: colWidths[2] },
        3: { halign: 'left', cellWidth: colWidths[3], overflow: 'hidden', fontSize: 7 },
        4: { halign: 'center', cellWidth: colWidths[4] },
        5: { halign: 'left', cellWidth: colWidths[5], overflow: 'hidden', fontSize: 7 },
        // VK Books: allow wrap so full text dikhe
        6: { halign: 'left', cellWidth: colWidths[6], overflow: 'linebreak', fontSize: 6 },
      },
      didParseCell: (data) => {
        const colIdx = Number(data.column.index);
        if (data.section === 'body' && data.row.index === body.length - 1) {
          data.cell.styles.fontStyle = 'bold';
        }
        // Borders: normal inner, slightly thick outer and header
        const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
        const isLeftEdge = colIdx === 0;
        const isRightEdge = colIdx + colSpan - 1 === lastColIndex;
        const isHead = data.section === 'head';
        const isLastBodyRow = data.section === 'body' && data.row.index === body.length - 1;
        data.cell.styles.lineWidth = {
          top: isHead ? LINE_THICK : LINE_NORMAL,
          bottom: isHead || isLastBodyRow ? LINE_THICK : LINE_NORMAL,
          left: isLeftEdge ? LINE_THICK : LINE_NORMAL,
          right: isRightEdge ? LINE_THICK : LINE_NORMAL,
        };
        // Fit in one line; font not smaller than 5 (zyada small na ho, par cut bhi na ho)
        const MIN_FONT = 5;
        // VK Books (col 6) ko wrap allow kar rahe hain; yaha shrink sirf 1,3,5 ke liye
        const fitColumns = [1, 3, 5];
        if (data.section === 'body' && fitColumns.includes(colIdx)) {
          const content = Array.isArray(data.cell.text) ? data.cell.text.join(' ') : String(data.cell.text ?? '');
          if (!content.trim()) return;
          const isTotalRowVk = data.row.index === body.length - 1 && colIdx === 5 && colSpan === 2;
          const cwRaw = isTotalRowVk
            ? (colWidths[5] ?? 0) + (colWidths[6] ?? 0)
            : (data.cell.styles.cellWidth ?? colWidths[colIdx] ?? 40);
          const cw = typeof cwRaw === 'number' ? cwRaw : 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? 1) * 2;
          const available = Math.max(2, cw - pad);
          const defaultFs = colIdx === 6 || isTotalRowVk ? 7 : 8;
          let fs = (data.cell.styles.fontSize as number) ?? defaultFs;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(content);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(content);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = 'hidden';
          doc.setFontSize(8);
        }
        if (data.section === 'head' && [1, 6].includes(colIdx)) {
          const content = Array.isArray(data.cell.text) ? data.cell.text.join(' ') : String(data.cell.text ?? '');
          if (!content.trim()) return;
          const cwRawHead = data.cell.styles.cellWidth ?? colWidths[colIdx] ?? 40;
          const cwHead = typeof cwRawHead === 'number' ? cwRawHead : 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? 1) * 2;
          const available = Math.max(2, cwHead - pad);
          let fs = (data.cell.styles.fontSize as number) ?? 7;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(content);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(content);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = 'hidden';
          doc.setFontSize(7);
        }
      },
      });
      startY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 4;
      sectorsOnPage++;
      hasContent = true;
    }
  }

  // Duplicate section: "Summary: Sector - City" style (continuous sectors, no half-page split)
  if (sectors.length > 0 && includeSectorCity) {
    let startYCity = HEADER_TOP_MM;
    if (hasContent) {
      doc.addPage();
    }
    const PAGE_BOTTOM_MM = A4_H - 15;

    for (let idx = 0; idx < sectors.length; idx++) {
      const sectorKey = sectors[idx];
      const sectorRows = bySector.get(sectorKey)!;

      const sectorRowsWithIdx = sectorRows.map((r, index) => ({ r, index }));
      sectorRowsWithIdx.sort((a, b) => {
        const appA = (a.r.app || '').toUpperCase();
        const appB = (b.r.app || '').toUpperCase();
        if (appA !== appB) return appA < appB ? -1 : 1;
        return a.index - b.index;
      });
      const orderedSectorRows = sectorRowsWithIdx.map((x) => x.r);

      const appSet = new Set<string>();
      for (const r of sectorRows) {
        if (r.app) appSet.add((r.app as string).toUpperCase());
      }
      const apps = Array.from(appSet).sort();
      const centerTitle = apps.length ? apps.join(' & ') : '';

      // Conservative estimate of required height: header + table rows.
      // Thoda zyada estimate rakhte hain taki koi bhi Sector - City table aadha next page par na jaye.
      const estimatedRows = 3 + sectorRows.length + 1; // header + data + total
      const approxHeight = 12 + estimatedRows * 7; // header ~12mm + 7mm per row (wrap ko dhyan me rakhkar)
      if (startYCity + approxHeight > PAGE_BOTTOM_MM) {
        doc.addPage();
        startYCity = HEADER_TOP_MM;
      }

      // 1 blank line between two sectors in Summary: Sector - City section
      if (idx > 0 && startYCity > HEADER_TOP_MM + 1) {
        startYCity += 4; // approx 1 row height gap
      }

      // Second section header (PARISHKAAR & VIMARSH / Sector - City) a little smaller to avoid overlap
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text(sectorKey, MARGIN_LEFT, startYCity);
      doc.text(centerTitle, A4_W / 2, startYCity, { align: 'center' });
      doc.text('Pending -> Clear Summary : Sector - City', A4_W - MARGIN_RIGHT, startYCity, {
        align: 'right',
      });
      startYCity += GAP_HEADER_TO_TABLE_MM;

      const headCity = [['P No.', 'Book Name', 'LG', 'Period', 'Qty', 'Sign', 'VK Books']];

      const bodyCity: CellDef[][] = [];
      const appKeyCity = (appName?: string) => (appName || '').toUpperCase();
      let currentGroupKeyCity: string | null = null;
      let currentGroupRowsCity: PendingClearRow[] = [];

      const flushGroupCity = () => {
        if (!currentGroupRowsCity.length) return;
        const groupBody = buildBodyWithMerges(currentGroupRowsCity);
        for (const row of groupBody) bodyCity.push(row);
      };

      for (const r of orderedSectorRows) {
        const key = appKeyCity(r.app);
        if (currentGroupKeyCity === null) {
          currentGroupKeyCity = key;
          currentGroupRowsCity = [r];
        } else if (key === currentGroupKeyCity) {
          currentGroupRowsCity.push(r);
        } else {
          flushGroupCity();
          bodyCity.push(['', '', '', '', '', '', '']);
          currentGroupKeyCity = key;
          currentGroupRowsCity = [r];
        }
      }
      flushGroupCity();

      const sectorTotalCity = sectorRows.reduce((s, r) => s + r.total_qty, 0);
      const villageTotalStrCity = buildVillageTotalBreakdown(sectorRows);
      bodyCity.push([
        '',
        '',
        '',
        'Total Qty',
        String(sectorTotalCity),
        { content: villageTotalStrCity, colSpan: 2 },
      ]);

      const LINE_NORMAL_CITY = 0.05;
      const LINE_THICK_CITY = 0.2;
      const lastColIndexCity = 6;

      autoTable(doc, {
        head: headCity,
        body: bodyCity,
        startY: startYCity,
        margin: { left: MARGIN_LEFT, right: MARGIN_RIGHT + 5 },
        tableWidth: TABLE_WIDTH - 10,
        theme: 'grid',
        styles: {
          fontSize: 8,
          cellPadding: 1,
          minCellHeight: 3,
          lineWidth: LINE_NORMAL_CITY,
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
        },
        bodyStyles: { valign: 'middle', textColor: [0, 0, 0], lineColor: [0, 0, 0] },
        headStyles: {
          fillColor: [220, 220, 220],
          fontStyle: 'bold',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
        },
        columnStyles: {
          0: { halign: 'center', cellWidth: colWidths[0] },
          1: { halign: 'left', cellWidth: colWidths[1], overflow: 'hidden', fontSize: 7 },
          2: { halign: 'center', cellWidth: colWidths[2] },
          3: { halign: 'left', cellWidth: colWidths[3], overflow: 'hidden', fontSize: 7 },
          4: { halign: 'center', cellWidth: colWidths[4] },
          5: { halign: 'left', cellWidth: colWidths[5], overflow: 'hidden', fontSize: 7 },
          6: { halign: 'left', cellWidth: colWidths[6], overflow: 'linebreak', fontSize: 6 },
        },
        didParseCell: (data) => {
          const colIdx = Number(data.column.index);
          if (data.section === 'body' && data.row.index === bodyCity.length - 1) {
            data.cell.styles.fontStyle = 'bold';
          }
          const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
          const isLeftEdge = colIdx === 0;
          const isRightEdge = colIdx + colSpan - 1 === lastColIndexCity;
          const isHead = data.section === 'head';
          const isLastBodyRow = data.section === 'body' && data.row.index === bodyCity.length - 1;
          data.cell.styles.lineWidth = {
            top: isHead ? LINE_THICK_CITY : LINE_NORMAL_CITY,
            bottom: isHead || isLastBodyRow ? LINE_THICK_CITY : LINE_NORMAL_CITY,
            left: isLeftEdge ? LINE_THICK_CITY : LINE_NORMAL_CITY,
            right: isRightEdge ? LINE_THICK_CITY : LINE_NORMAL_CITY,
          };
        },
      });
      startYCity = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 4;
    }
  }

  return Buffer.from(doc.output('arraybuffer'));
}

const GAP_TWO_COL_MM = 2;
const GAP_BETWEEN_VILLAGES_MM = 7; // 2 alag village ke table ke bich 1 row empty

type VillageCellDef =
  | string
  | number
  | { content: string; rowSpan?: number; styles?: { valign?: 'top' | 'middle' | 'bottom' } };

/** Build body with Period merged (rowSpan), center; and mark period boundaries for thick line */
function buildVillageBodyWithPeriodMerge(
  dataRows: (string | number)[][]
): { body: VillageCellDef[][]; periodBoundaryAfter: boolean[] } {
  const periodBoundaryAfter: boolean[] = [];
  const body: VillageCellDef[][] = [];
  const valignMiddle: { valign: 'top' | 'middle' | 'bottom' } = { valign: 'middle' };
  let i = 0;
  while (i < dataRows.length) {
    const periodStr = String(dataRows[i][0]);
    let run = 1;
    while (i + run < dataRows.length && String(dataRows[i + run][0]) === periodStr) run++;

    // Within same Period, merge LG column vertically when LG same (rowSpan), show centered once
    const lgRun: number[] = new Array(run).fill(0);
    let p = 0;
    while (p < run) {
      const lgVal = dataRows[i + p][2];
      let len = 1;
      while (p + len < run && dataRows[i + p + len][2] === lgVal) len++;
      lgRun[p] = len;
      p += len;
    }

    // First row for this Period (with Period cell)
    const firstRow: VillageCellDef[] = [];
    firstRow.push({ content: periodStr, rowSpan: run, styles: valignMiddle });
    firstRow.push(dataRows[i][1]); // Book Name
    if (lgRun[0] > 0) {
      firstRow.push({ content: String(dataRows[i][2]), rowSpan: lgRun[0], styles: valignMiddle });
    }
    firstRow.push(dataRows[i][3]); // Qty
    body.push(firstRow);

    // Remaining rows for this Period (no Period cell; LG only when new group starts)
    for (let j = 1; j < run; j++) {
      const row: VillageCellDef[] = [];
      row.push(dataRows[i + j][1]); // Book Name
      if (lgRun[j] > 0) {
        row.push({ content: String(dataRows[i + j][2]), rowSpan: lgRun[j], styles: valignMiddle });
      }
      row.push(dataRows[i + j][3]); // Qty
      body.push(row);
    }
    for (let j = 0; j < run - 1; j++) periodBoundaryAfter.push(false);
    periodBoundaryAfter.push(true);
    i += run;
  }
  return { body, periodBoundaryAfter };
}

/** Village-wise PDF: each village = left/right split table; 2 villages ke bich sirf 1 row gap, half-page nahi */
export function generatePendingClearVillagePdf(rows: PendingClearVillageRow[]): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const colRatios = [12, 24, 6, 4]; // Book Name thoda wide rakha hai, wrap kam ho
  const sumR = colRatios.reduce((a, b) => a + b, 0);
  const USABLE_TABLE_WIDTH = TABLE_WIDTH - 10; // overall thoda narrow; right side extra white space
  const halfTableWidth = (USABLE_TABLE_WIDTH - GAP_TWO_COL_MM) / 2;
  const colWidthsHalf = colRatios.map((r) => (halfTableWidth * r) / sumR);
  const CONTENT_RIGHT_X = MARGIN_LEFT + USABLE_TABLE_WIDTH; // header text (Total / P No.) isi width ke andar rahe
  const LINE_NORMAL = 0.05;
  const LINE_THICK = 0.2;
  const LINE_OUTER = 0.45; // village table outer border clearly visible
  const LINE_PERIOD = 0.15; // thoda thick between periods
  const lastColIndex = 3;
  const PAGE_BOTTOM_MM = A4_H - 15;
  const VILLAGE_BODY_FONT = 7;
  const VILLAGE_MIN_CELL_HEIGHT = 2;
  const VILLAGE_CELL_PADDING = 0.8;

  const byBlock = new Map<string, PendingClearVillageRow[]>();
  for (const r of rows) {
    const key = `${r.group_sector}|${r.village}|${r.stock_number}`;
    if (!byBlock.has(key)) byBlock.set(key, []);
    byBlock.get(key)!.push(r);
  }
  const bySector = new Map<string, [string, PendingClearVillageRow[]][]>();
  for (const [key, blockRows] of byBlock.entries()) {
    const sector = key.split('|')[0] || '';
    if (!bySector.has(sector)) bySector.set(sector, []);
    bySector.get(sector)!.push([key, blockRows]);
  }
  const sectorsOrdered = Array.from(bySector.keys()).sort(naturalCompare);
  const PAGE_CAPACITY_MM = PAGE_BOTTOM_MM - HEADER_TOP_MM;
  const MIN_SLOT_MM = 25;

  let simY = HEADER_TOP_MM;
  const blocks: [string, PendingClearVillageRow[]][] = [];
  for (const sector of sectorsOrdered) {
    const sectorBlocks = bySector.get(sector)!;
    const withHeight = sectorBlocks.map(([key, blockRows]) => {
      const tableRows = blockRows.length;
      const approxHeight = 6 + tableRows * 5;
      const totalWithGap = approxHeight + GAP_BETWEEN_VILLAGES_MM;
      return { key, blockRows, approxHeight, totalWithGap };
    });
    let firstSlotCapacity = PAGE_BOTTOM_MM - simY;
    if (firstSlotCapacity < MIN_SLOT_MM) firstSlotCapacity = PAGE_CAPACITY_MM;
    const pageSlots: { remaining: number; items: typeof withHeight }[] = [{ remaining: firstSlotCapacity, items: [] }];

    const bySizeAsc = [...withHeight].sort((a, b) => a.approxHeight - b.approxHeight);
    const placed = new Set<typeof withHeight[0]>();
    for (const w of bySizeAsc) {
      if (pageSlots[0].remaining >= w.totalWithGap) {
        pageSlots[0].items.push(w);
        pageSlots[0].remaining -= w.totalWithGap;
        placed.add(w);
      }
    }
    const remaining = withHeight.filter((w) => !placed.has(w));
    remaining.sort((a, b) => b.approxHeight - a.approxHeight);
    for (const w of remaining) {
      let put = false;
      for (let i = 1; i < pageSlots.length; i++) {
        if (pageSlots[i].remaining >= w.totalWithGap) {
          pageSlots[i].items.push(w);
          pageSlots[i].remaining -= w.totalWithGap;
          put = true;
          break;
        }
      }
      if (!put) {
        pageSlots.push({ remaining: PAGE_CAPACITY_MM - w.totalWithGap, items: [w] });
      }
    }
    for (const slot of pageSlots) {
      for (const w of slot.items) {
        blocks.push([w.key, w.blockRows]);
        if (simY + w.approxHeight > PAGE_BOTTOM_MM) simY = HEADER_TOP_MM;
        simY += w.totalWithGap;
      }
    }
  }
  let startY = HEADER_TOP_MM;

  const drawVillageTable = (
    tableBody: VillageCellDef[][],
    periodBoundaryAfter: boolean[],
    isLeftTable: boolean,
    startX: number,
    tableWidth: number,
    colWidths: number[]
  ) => {
    const body = tableBody.map((row) => row.map((c) => (typeof c === 'object' && c !== null ? { ...c } : c)));
    const head = [['Period', 'Book Name', 'LG', 'Qty']];
    autoTable(doc, {
      head,
      body,
      startY,
      margin: { left: startX, right: A4_W - startX - tableWidth },
      tableWidth,
      theme: 'grid',
      styles: {
        fontSize: VILLAGE_BODY_FONT,
        cellPadding: VILLAGE_CELL_PADDING,
        minCellHeight: VILLAGE_MIN_CELL_HEIGHT,
        lineWidth: LINE_NORMAL,
        lineColor: [0, 0, 0],
        textColor: [0, 0, 0],
      },
      bodyStyles: { valign: 'middle', fontSize: VILLAGE_BODY_FONT, lineColor: [0, 0, 0], textColor: [0, 0, 0] },
      headStyles: {
        fillColor: [220, 220, 220],
        fontStyle: 'bold',
        textColor: [0, 0, 0],
        fontSize: VILLAGE_BODY_FONT,
        lineColor: [0, 0, 0],
      },
      columnStyles: {
        0: { halign: 'center', cellWidth: colWidths[0], overflow: 'hidden', fontSize: VILLAGE_BODY_FONT },
        1: { halign: 'left', cellWidth: colWidths[1], overflow: 'hidden', fontSize: VILLAGE_BODY_FONT },
        2: { halign: 'center', cellWidth: colWidths[2], fontSize: VILLAGE_BODY_FONT },
        3: { halign: 'center', cellWidth: colWidths[3], fontSize: VILLAGE_BODY_FONT },
      },
      didParseCell: (data) => {
        const colIdx = Number(data.column.index);
        const rowIdx = data.row.index;
        const isLastBodyRow = data.section === 'body' && rowIdx === body.length - 1;
        const colSpan = (data.cell as { colSpan?: number }).colSpan ?? 1;
        const isLeftEdge = colIdx === 0;
        const isRightEdge = colIdx + colSpan - 1 === lastColIndex;
        const isHead = data.section === 'head';
        const isPeriodBoundary = data.section === 'body' && rowIdx < periodBoundaryAfter.length && periodBoundaryAfter[rowIdx];
        const thickBottom = isHead || isLastBodyRow || isPeriodBoundary;
        const bottomWidth = thickBottom ? (isPeriodBoundary && !isLastBodyRow ? LINE_PERIOD : LINE_THICK) : LINE_NORMAL;
        const outerLeft = isLeftTable && isLeftEdge;
        const outerRight = !isLeftTable && isRightEdge;
        data.cell.styles.lineColor = [0, 0, 0];
        data.cell.styles.lineWidth = {
          top: isHead ? LINE_OUTER : LINE_NORMAL,
          bottom: thickBottom ? (isPeriodBoundary && !isLastBodyRow ? LINE_PERIOD : LINE_OUTER) : LINE_NORMAL,
          left: outerLeft ? LINE_OUTER : LINE_NORMAL,
          right: outerRight ? LINE_OUTER : LINE_NORMAL,
        };

        // Book name / Period ko ek hi line me fit karne ke liye font size kam karna (wrap nahi)
        const FIT_COLS = [0, 1]; // 0: Period, 1: Book Name
        const MIN_FONT = VILLAGE_BODY_FONT - 2; // e.g. 5
        if (data.section === 'body' && FIT_COLS.includes(colIdx)) {
          const content = Array.isArray(data.cell.text) ? data.cell.text.join(' ') : String(data.cell.text ?? '');
          if (!content.trim()) return;
          const cwRaw = data.cell.styles.cellWidth ?? colWidths[colIdx] ?? 40;
          const cw = typeof cwRaw === 'number' ? cwRaw : 40;
          const pad = ((data.cell.styles.cellPadding as number) ?? VILLAGE_CELL_PADDING) * 2;
          const available = Math.max(2, cw - pad);
          let fs = (data.cell.styles.fontSize as number) ?? VILLAGE_BODY_FONT;
          doc.setFontSize(fs);
          let w = doc.getTextWidth(content);
          while (w > available && fs > MIN_FONT) {
            fs--;
            doc.setFontSize(fs);
            w = doc.getTextWidth(content);
          }
          data.cell.styles.fontSize = fs;
          data.cell.styles.overflow = 'hidden';
          doc.setFontSize(VILLAGE_BODY_FONT);
        }
      },
    });
  };

  for (let idx = 0; idx < blocks.length; idx++) {
    const [, blockRows] = blocks[idx];
    const first = blockRows[0];
    const sectorVillageLabel = `${first.group_sector} - ${first.village}`;
    const stock_number = first.stock_number;
    const totalQty = blockRows.reduce((s, r) => s + r.qty, 0);

    const halfOrder = (h: string) => {
      const x = (h || '').replace(/\s/g, '').toLowerCase();
      if (x === '1') return 0;
      if (x === '2') return 1;
      return 2;
    };
    // Order: Period old -> new, then LG wise (GJ, HN, ...), then Book name
    const sortedRows = [...blockRows].sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      if (a.month !== b.month) return a.month - b.month;
      if (halfOrder(a.half) !== halfOrder(b.half)) return halfOrder(a.half) - halfOrder(b.half);
      if (getLanguageSortOrder(a.language) !== getLanguageSortOrder(b.language))
        return getLanguageSortOrder(a.language) - getLanguageSortOrder(b.language);
      return a.book_name.localeCompare(b.book_name);
    });

    const dataRows = sortedRows.map((r) => [
      formatPeriod(r.year, r.month, r.half),
      (r.book_name || '').trim(),
      toCamelCase(r.language || ''),
      String(r.qty),
    ]);
    const n = dataRows.length;
    const leftCount = Math.ceil(n / 2);
    const leftData = dataRows.slice(0, leftCount);
    const rightData = dataRows.slice(leftCount);
    const { body: leftBody, periodBoundaryAfter: leftPeriodBoundary } = buildVillageBodyWithPeriodMerge(leftData);
    const { body: rightBody, periodBoundaryAfter: rightPeriodBoundary } = buildVillageBodyWithPeriodMerge(rightData);
    const tableRows = Math.max(leftBody.length, rightBody.length);
    const approxBlockHeight = 6 + tableRows * 5;

    if (startY + approxBlockHeight > PAGE_BOTTOM_MM) {
      doc.addPage();
      startY = HEADER_TOP_MM;
    }

    const headerFontSize = 9;
    doc.setFontSize(headerFontSize);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(0, 0, 0);
    doc.text(sectorVillageLabel, MARGIN_LEFT, startY);
    doc.setFont('helvetica', 'normal');
    doc.text(`Total : ${totalQty}`, MARGIN_LEFT + USABLE_TABLE_WIDTH / 2, startY, { align: 'center' });
    doc.setFont('helvetica', 'bold');
    doc.text(`P No. ${stock_number}`, CONTENT_RIGHT_X, startY, { align: 'right' });
    startY += GAP_HEADER_TO_TABLE_MM;

    const leftX = MARGIN_LEFT;
    const rightX = MARGIN_LEFT + halfTableWidth + GAP_TWO_COL_MM;
    drawVillageTable(leftBody, leftPeriodBoundary, true, leftX, halfTableWidth, colWidthsHalf);
    const leftFinalY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY;
    drawVillageTable(rightBody, rightPeriodBoundary, false, rightX, halfTableWidth, colWidthsHalf);
    const rightFinalY = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY;
    startY = Math.max(leftFinalY, rightFinalY) + GAP_BETWEEN_VILLAGES_MM;
  }

  return Buffer.from(doc.output('arraybuffer'));
}
