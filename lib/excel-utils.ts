import * as XLSX from 'xlsx';

export interface ExcelRow {
  [key: string]: any;
}

export function parseExcelFile(buffer: Buffer): XLSX.WorkBook {
  return XLSX.read(buffer, { type: 'buffer' });
}

export function getFirstSheet(workbook: XLSX.WorkBook): XLSX.WorkSheet | null {
  const sheetName = workbook.SheetNames[0];
  return sheetName ? workbook.Sheets[sheetName] : null;
}

export function sheetToJSON(sheet: XLSX.WorkSheet): ExcelRow[] {
  return XLSX.utils.sheet_to_json(sheet, { raw: false, defval: '' });
}

function normalizeHeaderKey(key: string): string {
  return String(key || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .replace(/[\s._()]/g, '');
}

export function getCellByHeader(row: ExcelRow, candidates: string[]): unknown {
  if (!row) return undefined;
  const normalizedCandidates = candidates.map((c) => normalizeHeaderKey(c));

  for (const [key, value] of Object.entries(row)) {
    const normKey = normalizeHeaderKey(key);
    if (normalizedCandidates.includes(normKey)) {
      return value;
    }
  }

  return undefined;
}

export function extractPeriodFromCell(cellValue: string): {
  year: number;
  month: number;
  half: string;
} | null {
  if (!cellValue) return null;

  // Format: "PARISHKAAR DEC-2025(1-31)" or "VIMARSH NOV-2025(1-30)"
  const match = cellValue.match(/(\w+)-(\d{4})\((\d+)-(\d+)\)/i);
  
  if (match) {
    const monthStr = match[1].toUpperCase();
    const year = parseInt(match[2]);
    const startDay = parseInt(match[3]);
    const endDay = parseInt(match[4]);

    // Convert month name to number
    const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const month = monthNames.indexOf(monthStr) + 1;

    if (month === 0) return null;

    // Determine half
    let half = '';
    if (startDay === 1 && endDay === 15) {
      half = '1';
    } else if (startDay === 16 && endDay >= 28) {
      half = '2';
    } else if (startDay === 1 && endDay >= 28) {
      half = '1 & 2';  // Full month
    }

    return { year, month, half };
  }

  return null;
}

/**
 * Normalize HALF value to one of: "1", "2", "1&2" (for full month), or original string if unknown.
 * Mirrors VBA logic in CalculatePeriodFromTransactions for handling "1 & 2", "1_2", "1 and 2", etc.
 */
export function normalizeHalfValue(raw: unknown): string {
  const s = String(raw ?? '').trim();
  if (!s) return '';

  const lower = s.toLowerCase().replace(/&amp;/gi, '&');
  const compact = lower.replace(/\s/g, '').replace(/_/g, '');

  // Full month: contains 1 and 2 with a separator
  if (
    compact.includes('1') &&
    compact.includes('2') &&
    (lower.includes('&') || lower.includes('_') || lower.includes(' and '))
  ) {
    return '1&2';
  }

  if (compact === '1' || compact === 'first' || compact === '1st') return '1';
  if (compact === '2' || compact === 'second' || compact === '2nd') return '2';

  return s;
}

/**
 * Derive period (year, month, half) from transaction rows when there is no PERIOD sheet.
 * Follows the old VBA logic: take the first row that has valid YEAR, MONTH, HALF and a book title.
 */
export function extractPeriodFromDataRows(rows: ExcelRow[]): {
  year: number;
  month: number;
  half: string;
} | null {
  if (!rows?.length) return null;

  for (const row of rows) {
    const bookName = getCellByHeader(row, [
      'BOOK TITLE',
      'BOOK_NAME',
      'BOOK NAME',
      'TITLE',
      'BOOK',
    ]) as string;

    const yearRaw = getCellByHeader(row, ['YEAR']);
    const monthRaw = getCellByHeader(row, ['MONTH']);
    const halfRaw = getCellByHeader(row, ['HALF', 'HALF PERIOD', 'HALF_PERIOD']);

    if (!String(bookName || '').trim()) continue;

    const year = parseInt(String(yearRaw ?? '').trim(), 10);
    const month = parseInt(String(monthRaw ?? '').trim(), 10);

    if (!year || !month) continue;

    const half = normalizeHalfValue(halfRaw);
    if (!half) continue;

    return { year, month, half };
  }

  return null;
}

/**
 * Find the main data sheet ("VillageNameWise" / "village_namewise" / similar).
 * Mirrors the VBA logic that searches for a sheet whose name contains both "village" and "namewise".
 */
export function getVillageNamewiseSheet(workbook: XLSX.WorkBook): XLSX.WorkSheet | null {
  if (!workbook?.SheetNames?.length) return null;

  const targetName =
    workbook.SheetNames.find((name) => {
      const s = String(name || '').trim().toLowerCase();
      if (!s) return false;
      if (s === 'villagenamewise' || s === 'village_namewise' || s === 'village namewise') return true;
      return s.includes('village') && s.includes('namewise');
    }) || workbook.SheetNames[0];

  return workbook.Sheets[targetName] || null;
}

/**
 * Parse AVAIL value to numeric code.
 * 1 = YES, 0 = NS/NO, 2 = B-YES. Also accepts numeric 0/1/2 directly.
 * This keeps semantics in line with the original VBA import.
 */
export function parseAvailValue(value: unknown): number {
  if (value === null || value === undefined) return 0;

  if (typeof value === 'number' && !Number.isNaN(value)) {
    if (value === 0 || value === 1 || value === 2) return value;
  }

  const s = String(value).trim();
  if (!s) return 0;

  const upper = s.toUpperCase();
  if (upper === 'YES') return 1;
  if (upper === 'NO' || upper === 'NS' || upper === 'N') return 0;
  if (upper === 'B-YES' || upper === 'BYES' || upper === 'B YES') return 2;

  const n = parseInt(s, 10);
  if (!Number.isNaN(n) && (n === 0 || n === 1 || n === 2)) return n;

  return 0;
}

export function extractAppName(cellValue: string): string {
  if (!cellValue) return 'VIMARSH';

  const upperValue = cellValue.toUpperCase();
  if (upperValue.includes('PARISHKAAR')) {
    return 'PARISHKAAR';
  } else if (upperValue.includes('VIMARSH')) {
    return 'VIMARSH';
  }

  return 'VIMARSH';
}

export function formatPeriodDisplay(year: number, month: number, half: string): string {
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthName = monthNames[month - 1];
  const halfNormalized = half.replace(/\s/g, '').toLowerCase();

  if (halfNormalized === '1&2') {
    const lastDay = new Date(year, month, 0).getDate();
    return `${year}-${monthName} (1 To ${lastDay})`;
  } else if (half === '1') {
    return `${year}-${monthName} (1 To 15)`;
  } else if (half === '2') {
    const lastDay = new Date(year, month, 0).getDate();
    return `${year}-${monthName} (16 To ${lastDay})`;
  }

  return `${year}-${String(month).padStart(2, '0')}-${half}`;
}
