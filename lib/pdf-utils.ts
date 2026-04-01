import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { GroupVillageRow } from './tatvagnan-data';

export interface TatvagnanData {
  group: string;
  lang: string;
  old_qty: number;
  new_qty: number;
  rmv_qty: number;
  total_qty: number;
}

function cellVal(n: number): number | '' {
  return n === 0 ? '' : n;
}

export interface LanguageData {
  code: string;
  fullName: string;
}

const languageMap: { [key: string]: string } = {
  'GJ': 'Gujarati',
  'HN': 'Hindi',
  'EN': 'English',
  'MR': 'Marathi',
  'MH': 'Marathi',
  'SD': 'Sindhi',
  'UR': 'Urdu',
  'PJ': 'Punjabi',
  'BN': 'Bengali',
  'TA': 'Tamil',
  'TE': 'Telugu',
  'KN': 'Kannada',
  'ML': 'Malayalam',
};

export function getFullLanguageName(code: string): string {
  return languageMap[code.toUpperCase()] || code;
}

const langOrder = ['GJ', 'HN', 'EN', 'MR', 'MH', 'SD', 'UR', 'PJ', 'BN', 'TA', 'TE', 'KN', 'ML'];

function getLangSortKey(lang: string): number {
  const code = (lang || '').trim().toUpperCase();
  const two = code.length >= 2 ? code.slice(0, 2) : code;
  const idx = langOrder.indexOf(two);
  if (idx >= 0) return idx;
  if (lang.toUpperCase().startsWith('GUJ')) return 0;
  if (lang.toUpperCase().startsWith('HIN')) return 1;
  if (lang.toUpperCase().startsWith('MAR')) return 3;
  return 999;
}

function getLangCode(lang: string): string {
  const s = (lang || '').trim();
  if (!s) return '—';
  if (s.length <= 2) return s.toUpperCase();
  const u = s.toUpperCase();
  for (const [code, full] of Object.entries(languageMap)) {
    if (full.toUpperCase().startsWith(u) || u.startsWith(full.toUpperCase().slice(0, 2))) return code;
  }
  return s.slice(0, 2).toUpperCase();
}

function naturalSort(a: string, b: string): number {
  const regex = /(\d+)|(\D+)/g;
  const aParts = a.match(regex) || [];
  const bParts = b.match(regex) || [];

  for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
    const aPart = aParts[i] || '';
    const bPart = bParts[i] || '';

    if (/^\d+$/.test(aPart) && /^\d+$/.test(bPart)) {
      const diff = parseInt(aPart) - parseInt(bPart);
      if (diff !== 0) return diff;
    } else {
      if (aPart < bPart) return -1;
      if (aPart > bPart) return 1;
    }
  }
  return 0;
}


export function generateSummaryPdf(data: TatvagnanData[], pushpNo: string): Buffer {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  // Group data by group and language
  const grouped: { [group: string]: { [lang: string]: TatvagnanData } } = {};
  const languages = new Set<string>();

  data.forEach((row) => {
    if (!grouped[row.group]) {
      grouped[row.group] = {};
    }
    grouped[row.group][row.lang] = row;
    languages.add(row.lang);
  });

  const langArray = Array.from(languages).sort();
  const groups = Object.keys(grouped).sort(naturalSort);

  const prevPushpNo = parseInt(pushpNo) - 1;

  // Title (Excel row height 31pt ≈ 11mm)
  doc.setFontSize(14);
  doc.setFont('helvetica', 'bold');
  doc.text(`TATVAGNAN SUMMARY - Pushp No. ${pushpNo}`, doc.internal.pageSize.getWidth() / 2, 12, { align: 'center' });

  const dateTimeStr = new Date().toLocaleString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
  });
  doc.setFontSize(8);
  doc.setFont('helvetica', 'normal');
  doc.text(`Dt : ${dateTimeStr}`, doc.internal.pageSize.getWidth() - 10, 17, { align: 'right' });

  // Prepare table data
  const headers: any[] = [
    { content: 'Sr.', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
    { content: 'Group', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
  ];

  langArray.forEach((lang) => {
    const fullName = getFullLanguageName(lang);
    headers.push({ content: fullName, colSpan: 4, styles: { halign: 'center', valign: 'middle' } });
  });

  headers.push({ content: 'Total', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } });
  headers.push({ content: 'Sign', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } });
  headers.push({ content: 'Date', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } });

  const subHeaders: any[] = [];
  langArray.forEach(() => {
    subHeaders.push(`P-${prevPushpNo}`, 'New', 'Remove', 'Total');
  });

  // Data rows
  const tableData: any[] = [];
  const summaryBlankMarker = '__SUMMARY_BLANK_ROW__';
  groups.forEach((group, index) => {
    const row: any[] = [index + 1, group];
    let rowTotal = 0;

    langArray.forEach((lang) => {
      const langData = grouped[group][lang];
      if (langData) {
        row.push(
          cellVal(langData.old_qty),
          cellVal(langData.new_qty),
          cellVal(langData.rmv_qty),
          cellVal(langData.total_qty)
        );
        rowTotal += langData.total_qty || 0;
      } else {
        row.push('', '', '', '');
      }
    });

    row.push(cellVal(rowTotal), '', ''); // Total column, Sign, Date
    tableData.push(row);
    if (index < groups.length - 1) {
      const blankRow = row.map(() => '');
      blankRow[0] = summaryBlankMarker;
      tableData.push(blankRow);
    }
  });

  // Total row (same as VBA: sum each column from first data row to last)
  const totalRow: any[] = ['Total', ''];
  let grandTotal = 0;
  langArray.forEach((lang) => {
    let langOldTotal = 0;
    let langNewTotal = 0;
    let langRmvTotal = 0;
    let langTotalTotal = 0;
    groups.forEach((group) => {
      const langData = grouped[group][lang];
      if (langData) {
        langOldTotal += langData.old_qty || 0;
        langNewTotal += langData.new_qty || 0;
        langRmvTotal += langData.rmv_qty || 0;
        langTotalTotal += langData.total_qty || 0;
      }
    });
    totalRow.push(cellVal(langOldTotal), cellVal(langNewTotal), cellVal(langRmvTotal), cellVal(langTotalTotal));
    grandTotal += langTotalTotal;
  });
  totalRow.push(cellVal(grandTotal), '', ''); // Total column, Sign, Date
  tableData.push(totalRow);

  // Width tuning as requested:
  // - Group = 8
  // - Numeric columns = 6 each
  // - Remaining space to Sign
  const dataCols = 4 * langArray.length + 1; // 4 per lang + one Total column
  const totalUnits = 5 + 8 + dataCols * 6 + 26 + 10;
  const mmPerUnit = (doc.internal.pageSize.getWidth() - 20) / totalUnits;
  const col = (units: number) => Math.round(units * mmPerUnit * 10) / 10;

  const columnStyles: { [key: number]: any } = {
    0: { cellWidth: col(5) },
    1: { cellWidth: col(8) },
  };
  for (let c = 2; c < 2 + dataCols; c++) {
    columnStyles[c] = { cellWidth: col(6) };
  }
  langArray.forEach((_, i) => {
    columnStyles[2 + 4 * i + 3] = { cellWidth: col(6), fillColor: [230, 230, 230] };
  });
  columnStyles[2 + 4 * langArray.length] = { cellWidth: col(6), fillColor: [230, 230, 230] }; // Total column
  columnStyles[2 + dataCols] = { cellWidth: col(26) };
  columnStyles[2 + dataCols + 1] = { cellWidth: col(10) };

  autoTable(doc, {
    head: [headers, subHeaders],
    body: tableData,
    startY: 20,
    theme: 'grid',
    styles: {
      fontSize: 9,
      cellPadding: 1.5,
      halign: 'center',
      valign: 'middle',
      font: 'helvetica',
      minCellHeight: 5,
      overflow: 'linebreak',
      lineColor: [0, 0, 0],
      lineWidth: 0.2,
    },
    headStyles: {
      fillColor: [230, 234, 240],
      textColor: [0, 0, 0],
      fontStyle: 'bold',
      minCellHeight: 5.5,
      lineColor: [0, 0, 0],
      lineWidth: 0.2,
    },
    alternateRowStyles: {
      fillColor: [248, 250, 252],
    },
    bodyStyles: {
      fillColor: [255, 255, 255],
    },
    columnStyles,
    didParseCell: (data) => {
      if (data.section === 'head') {
        data.cell.styles.lineColor = [0, 0, 0];
        data.cell.styles.lineWidth = 0.2;
        if (data.row.index === 1) {
          data.cell.styles.fillColor = [217, 217, 217];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.fontSize = 9;
          data.cell.styles.textColor = [0, 0, 0];
        }
      } else if (data.section === 'body') {
        const raw = data.row.raw as any[];
        const isTotalRow = Array.isArray(raw) && String(raw[0] || '').toUpperCase() === 'TOTAL';
        if (isTotalRow) {
          data.cell.styles.fillColor = [236, 240, 245];
          data.cell.styles.fontStyle = 'bold';
        }
        if (data.column?.index === 1) {
          const txt = String(data.cell.raw ?? '');
          if (txt.length > 10) data.cell.styles.fontSize = 8;
          if (txt.length > 14) data.cell.styles.fontSize = 7;
        }
        if (Array.isArray(raw) && raw[0] === summaryBlankMarker) {
          data.cell.text = [''];
          data.cell.styles.minCellHeight = 1.2; // Doubled blank spacer row height
          data.cell.styles.cellPadding = 0;
          data.cell.styles.fontSize = 1;
          data.cell.styles.lineWidth = 0;
        }
      }
    },
  });

  return Buffer.from(doc.output('arraybuffer'));
}

export function generateGroupWisePdf(data: TatvagnanData[], pushpNo: string): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  // Group data
  const grouped: { [group: string]: TatvagnanData[] } = {};
  data.forEach((row) => {
    if (!grouped[row.group]) {
      grouped[row.group] = [];
    }
    grouped[row.group].push(row);
  });

  const groups = Object.keys(grouped).sort(naturalSort);
  const prevPushpNo = parseInt(pushpNo) - 1;

  const pageHeight = doc.internal.pageSize.getHeight();
  const topMargin = 12;
  const bottomMargin = 12;
  const usableHeight = pageHeight - topMargin - bottomMargin;
  const slotHeight = usableHeight / 4;

  groups.forEach((group, index) => {
    const pageIndex = Math.floor(index / 4);
    const slotIndex = index % 4;
    if (pageIndex > 0 && slotIndex === 0) {
      doc.addPage();
    }

    const slotStartY = topMargin + slotIndex * slotHeight;
    const tableStartY = slotStartY + 2;

    // Sort by total quantity (descending)
    const groupData = grouped[group].sort((a, b) => b.total_qty - a.total_qty);

    const tableData: (string | number | { content: string; colSpan: number })[][] = [
      [
        { content: `Tatvagnan Pushp No. ${pushpNo}`, colSpan: 3 },
        { content: `Group : ${group}`, colSpan: 3 },
      ],
      ['Sr.', 'Language', `Pushp - ${prevPushpNo}`, 'New', 'Remove', 'Total'],
      ...groupData.map((row, idx) => [
        idx + 1,
        row.lang || getFullLanguageName(row.lang),
        cellVal(row.old_qty),
        cellVal(row.new_qty),
        cellVal(row.rmv_qty),
        cellVal(row.total_qty),
      ]),
    ];

    const totalOld = groupData.reduce((sum, r) => sum + (r.old_qty || 0), 0);
    const totalNew = groupData.reduce((sum, r) => sum + (r.new_qty || 0), 0);
    const totalRmv = groupData.reduce((sum, r) => sum + (r.rmv_qty || 0), 0);
    const totalTotal = groupData.reduce((sum, r) => sum + (r.total_qty || 0), 0);

    tableData.push(['', 'Total', cellVal(totalOld), cellVal(totalNew), cellVal(totalRmv), cellVal(totalTotal)]);
    tableData.push([
      { content: 'Sign : _________________', colSpan: 3 },
      { content: 'Date : _________________', colSpan: 3 },
    ]);

    const gwCol = (chars: number) => chars * 2.2;
    autoTable(doc, {
      body: tableData,
      startY: tableStartY,
      theme: 'grid',
      styles: {
        fontSize: 9,
        cellPadding: 1,
        halign: 'center',
        valign: 'middle',
        font: 'helvetica',
        minCellHeight: 3.2,
        overflow: 'linebreak',
        lineColor: [0, 0, 0],
      },
      headStyles: {
        fillColor: [230, 234, 240],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        minCellHeight: 3.5,
        lineColor: [0, 0, 0],
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252],
      },
      bodyStyles: {
        fillColor: [255, 255, 255],
      },
      columnStyles: {
        0: { cellWidth: gwCol(8) },
        1: { cellWidth: gwCol(20) },
        2: { cellWidth: gwCol(10) },
        3: { cellWidth: gwCol(10) },
        4: { cellWidth: gwCol(10) },
        5: { cellWidth: gwCol(12) },
      },
      margin: { left: 20 },
      didParseCell: (data) => {
        const raw = data.row.raw as (string | number | { content: string; colSpan: number })[];
        const colIdx = data.column?.index ?? 0;
        const cell = raw[colIdx];
        if (data.row.index <= 1) {
          data.cell.styles.lineColor = [255, 255, 255]; // blank border – header rows have gray background
        } else {
          data.cell.styles.lineColor = [0, 0, 0];
        }
        if (typeof cell === 'object' && cell !== null && 'colSpan' in cell) {
          if (data.row.index === 0) {
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.fontSize = 10;
            data.cell.styles.halign = colIdx === 0 ? 'left' : 'right';
          }
          if (data.row.index === tableData.length - 1) {
            data.cell.styles.fontSize = 8;
            data.cell.styles.fontStyle = 'normal';
            data.cell.styles.halign = colIdx === 0 ? 'left' : 'right';
            data.cell.styles.minCellHeight = 20;
          }
        }
        if (data.row.index === 1) {
          data.cell.styles.fillColor = [230, 234, 240];
          data.cell.styles.fontStyle = 'bold';
        }
        if (data.row.index === tableData.length - 2) {
          data.cell.styles.fillColor = [236, 240, 245];
          data.cell.styles.fontStyle = 'bold';
        }
        if (data.section === 'body' && data.column?.index === 1 && data.row.index > 1 && data.row.index < tableData.length - 2) {
          const txt = String(data.cell.raw ?? '');
          if (txt.length > 12) data.cell.styles.fontSize = 8;
          if (txt.length > 18) data.cell.styles.fontSize = 7;
        }
      },
    });
  });

  return Buffer.from(doc.output('arraybuffer'));
}

/** Stock Update PDF - same layout as VBA StockUpdateForm (Pending : Surat City, book-wise totals, grand total) */
export interface StockUpdatePdfRow {
  book_name: string;
  language: string;
  period: string;
  qty: number;
  stock_number: string | null;
}

export function generateStockUpdatePdf(rows: StockUpdatePdfRow[]): Buffer {
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  // Row 1: "Pending : Surat City" (left), "Dt : dd/mm/yyyy hh:mm AM/PM" (right)
  doc.setFontSize(11);
  doc.setFont('helvetica', 'bold');
  doc.text('Pending : Surat City', 14, 10);
  const dateStr =
    'Dt : ' +
    new Date().toLocaleString('en-IN', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    });
  doc.text(dateStr, doc.internal.pageSize.getWidth() - 14, 10, { align: 'right' });

  // Build body rows with book-wise total rows and grand total (same as VBA)
  const bodyRows: (string | number)[][] = [];
  let currentBook = '';
  let bookTotal = 0;
  let grandTotal = 0;

  rows.forEach((row) => {
    const qty = Number(row.qty) || 0;
    if (row.book_name !== currentBook) {
      if (currentBook !== '') {
        bodyRows.push(['', '', `Total : ${bookTotal}`, '', '']);
      }
      currentBook = row.book_name;
      bookTotal = 0;
    }
    bookTotal += qty;
    grandTotal += qty;
    bodyRows.push([
      row.book_name,
      row.language,
      row.period,
      qty,
      row.stock_number || '',
    ]);
  });
  if (currentBook !== '') {
    bodyRows.push(['', '', `Total : ${bookTotal}`, '', '']);
  }
  bodyRows.push(['', '', `Total All Qty : ${grandTotal}`, '', '']);

  autoTable(doc, {
    head: [['Book Name', 'Language', 'Period', 'Qty', 'P No.']],
    body: bodyRows,
    startY: 16,
    theme: 'grid',
    styles: {
      fontSize: 9,
      cellPadding: 1.2,
      font: 'helvetica',
      minCellHeight: 4,
    },
    headStyles: {
      fillColor: [200, 200, 200],
      textColor: [0, 0, 0],
      fontStyle: 'bold',
      halign: 'left',
      valign: 'middle',
      minCellHeight: 4,
      cellPadding: 1.2,
    },
    columnStyles: {
      0: { halign: 'left', cellWidth: 55 },
      1: { halign: 'center', cellWidth: 28 },
      2: { halign: 'left', cellWidth: 45 },
      3: { halign: 'center', cellWidth: 18 },
      4: { halign: 'left', cellWidth: 38 },
    },
    didParseCell: (data) => {
      const raw = data.row.raw as (string | number)[];
      const periodCell = raw[2] ? String(raw[2]) : '';
      if (periodCell.startsWith('Total All Qty')) {
        data.cell.styles.fillColor = [220, 220, 220];
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.textColor = [0, 0, 0];
      } else if (periodCell.startsWith('Total')) {
        data.cell.styles.fillColor = [240, 240, 240];
        data.cell.styles.fontStyle = 'bold';
      }
    },
    margin: { left: 14, right: 14 },
  });

  return Buffer.from(doc.output('arraybuffer'));
}

/** Group + Village wise PDF: Group & Village on one line; table sorted by lang (GJ order); Name/Member double width; Lang short; Mobile 10 digit; Add/Remove last. */
export function generateGroupVillagePdf(rows: GroupVillageRow[], pushpNo: string): Buffer {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);

  const byGroup = new Map<string, Map<string, GroupVillageRow[]>>();
  for (const r of rows) {
    const g = r.group || '—';
    const v = r.village || '—';
    if (!byGroup.has(g)) byGroup.set(g, new Map());
    const byVillage = byGroup.get(g)!;
    if (!byVillage.has(v)) byVillage.set(v, []);
    byVillage.get(v)!.push(r);
  }

  const oneRowGap = 7;
  let startY = oneRowGap;
  const pageH = doc.internal.pageSize.getHeight();
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 14;
  const tableWidth = 10 + 68 + 78 + 10 + 22 + 40 + 40;
  const tableRight = margin + tableWidth;
  const head = [['+/-', 'Name', 'Member', 'Lang', 'Mobile', 'Add 1', 'Add 2']];

  const mobile10 = (s: string) => {
    const d = (s || '').replace(/\D/g, '').slice(-10);
    if (d.length !== 10) return d || '—';
    return `${d.slice(0, 5)} ${d.slice(5)}`;
  };

  const toTitleCase = (s: string) =>
    (s || '')
      .trim()
      .split(/\s+/)
      .map((w) => (w.length ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : w))
      .join(' ');

  const fitFontSize = (len: number, widthMm: number): number => {
    if (len <= 0) return 8;
    const at8 = Math.floor(widthMm / 1.5);
    const at7 = Math.floor(widthMm / 1.35);
    const at6 = Math.floor(widthMm / 1.2);
    const at5 = Math.floor(widthMm / 1.0);
    if (len <= at8) return 8;
    if (len <= at7) return 7;
    if (len <= at6) return 6;
    if (len <= at5) return 5;
    return 5;
  };

  const minSectionHeight = 45;
  const sortedGroups = Array.from(byGroup.keys()).sort(naturalSort);
  for (const group of sortedGroups) {
    const byVillage = byGroup.get(group)!;
    const sortedVillages = Array.from(byVillage.keys()).sort(naturalSort);
    for (const village of sortedVillages) {
      const list = byVillage.get(village)!;
      if (startY > pageH - minSectionHeight) {
        doc.addPage('landscape');
        startY = oneRowGap;
      }
      const addCount = list.filter((r) => r.action_type === 'ADD').length;
      const removeCount = list.filter((r) => r.action_type === 'REMOVE').length;
      const addStr = addCount > 0 ? `Add = ${addCount}` : '';
      const removeStr = removeCount > 0 ? `Rem = ${removeCount}` : '';
      const centerStr = [addStr, removeStr].filter(Boolean).join(' , ');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.text(`${group}- ${village}`, margin, startY);
      if (centerStr) {
        doc.text(centerStr, pageW / 2, startY, { align: 'center' });
      }
      doc.text(`Tatvagnan Pushp No. ${pushpNo}`, tableRight, startY, { align: 'right' });
      doc.setFont('helvetica', 'normal');
      startY += 3;

      const sorted = [...list].sort((a, b) => {
        const addFirst = (a.action_type === 'ADD' ? 0 : 1) - (b.action_type === 'ADD' ? 0 : 1);
        if (addFirst !== 0) return addFirst;
        return (a.name || '').localeCompare(b.name || '', 'en', { sensitivity: 'base' });
      });
      const body = sorted.map((r) => [
        r.action_type === 'ADD' ? 'Add' : r.action_type === 'REMOVE' ? 'Rem' : r.action_type || '—',
        toTitleCase(r.name || '—'),
        toTitleCase(r.member || '—'),
        getLangCode(r.language),
        mobile10(r.mobile_number),
        toTitleCase((r.address_line1 || '—').slice(0, 38)),
        toTitleCase((r.address_line2 || '—').slice(0, 38)),
      ]);
      autoTable(doc, {
        head,
        body: body.length ? body : [['—', '—', '—', '—', '—', '—', '—']],
        startY,
        theme: 'grid',
        margin: { bottom: oneRowGap },
        styles: { fontSize: 8, cellPadding: 1.5, lineColor: [0, 0, 0], overflow: 'linebreak' },
        headStyles: {
          fillColor: [230, 234, 240],
          fontStyle: 'bold',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          overflow: 'linebreak',
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
        },
        bodyStyles: {
          fillColor: [255, 255, 255],
        },
        showHead: 'everyPage',
        columnStyles: {
          0: { cellWidth: 10, overflow: 'linebreak' },
          1: { cellWidth: 68, overflow: 'linebreak' },
          2: { cellWidth: 78, overflow: 'linebreak' },
          3: { cellWidth: 10, overflow: 'linebreak' },
          4: { cellWidth: 22, overflow: 'linebreak' },
          5: { cellWidth: 40, overflow: 'linebreak' },
          6: { cellWidth: 40, overflow: 'linebreak' },
        },
        didParseCell: (data) => {
          if (data.section === 'body' && data.column?.index != null) {
            const raw = data.row.raw as string[];
            const col = data.column.index;
            const len = (i: number) => (raw[i] ? String(raw[i]).length : 0);
            if (col === 1) data.cell.styles.fontSize = fitFontSize(len(1), 68);
            else if (col === 2) data.cell.styles.fontSize = fitFontSize(len(2), 78);
            else if (col === 5) data.cell.styles.fontSize = fitFontSize(len(5), 40);
            else if (col === 6) data.cell.styles.fontSize = fitFontSize(len(6), 40);
          }
        },
      });
      startY = (doc as any).lastAutoTable.finalY + oneRowGap;
    }
  }

  if (byGroup.size === 0) {
    doc.text('No group + village data for this Pushp No.', margin, startY);
  }

  return Buffer.from(doc.output('arraybuffer'));
}
