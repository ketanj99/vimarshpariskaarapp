import { query, DB_SCHEMA } from './db';

export interface ReportFilters {
  organizations: string[];
  groupSectors: string[];
  periods: string[];
}

export interface BookWithQty {
  book_id: number;
  book_name: string;
  language: string;
  display_label: string;
  total_qty: number;
}

export async function loadReportFilters(): Promise<Omit<ReportFilters, 'periods'>> {
  const [orgRes, sectorRes] = await Promise.all([
    query<{ app: string }>(
      `SELECT app FROM (SELECT DISTINCT app FROM ${DB_SCHEMA}.transactions WHERE app IS NOT NULL AND TRIM(app) <> '' AND is_deleted = FALSE) t ORDER BY CASE WHEN UPPER(TRIM(app)) = 'VIMARSH' THEN 0 ELSE 1 END, app`
    ),
    query<{ group_sector: string }>(
      `SELECT DISTINCT group_sector FROM ${DB_SCHEMA}.transactions WHERE group_sector IS NOT NULL AND TRIM(group_sector) <> '' AND is_deleted = FALSE ORDER BY group_sector`
    ),
  ]);

  const organizations = (orgRes.rows || []).map((r) => (r.app || '').trim()).filter(Boolean);
  const groupSectors = ['All', ...(sectorRes.rows || []).map((r) => (r.group_sector || '').trim()).filter(Boolean)];

  return { organizations, groupSectors };
}

/** Periods only for the selected organization (APP) – agar us app me period nahi hai to list me nahi aata */
export async function loadPeriodsForOrg(org: string): Promise<string[]> {
  if (!org?.trim()) return [];
  const result = await query<{ year: number; month: number; half: string }>(
    `SELECT DISTINCT year, month, half FROM ${DB_SCHEMA}.transactions WHERE app = $1 AND is_deleted = FALSE ORDER BY year ASC, month ASC, half ASC`,
    [org.trim()]
  );
  return (result.rows || []).map((r) => `${r.year}-${r.month}-${(r.half || '').replace(/\s/g, '')}`);
}

export async function loadBooksWithTotals(
  org: string,
  groupSector: string,
  period: string
): Promise<BookWithQty[]> {
  const [y, m, h] = period.split('-');
  const year = parseInt(y, 10);
  const month = parseInt(m, 10);
  const halfNorm = (h || '').replace(/\s/g, '').toLowerCase();
  const sectorVal = groupSector === 'All' ? 'All' : groupSector;

  const result = await query<{ book_id: number; book_name: string; language: string; total_qty: string }>(
    `SELECT b.book_id, b.book_name, COALESCE(b.language,'') AS language,
            SUM(t.qty) AS total_qty
     FROM ${DB_SCHEMA}.transactions t
     INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
     WHERE t.app = $1 AND t.year = $2 AND t.month = $3
       AND (REPLACE(LOWER(COALESCE(t.half,'')), ' ', '') IN ('1&2','1and2') OR t.half = $4)
       AND t.is_deleted = FALSE AND b.is_deleted = FALSE
       AND ($5::text = 'All' OR t.group_sector = $5)
     GROUP BY b.book_id, b.book_name, b.language
     HAVING SUM(t.qty) > 0
     ORDER BY b.book_name, b.language`,
    [org, year, month, halfNorm === '1&2' || halfNorm === '1and2' ? '1&2' : (h || ''), sectorVal]
  );

  return (result.rows || []).map((r) => ({
    book_id: r.book_id,
    book_name: r.book_name || '',
    language: r.language || '',
    display_label: r.language ? `${r.book_name} - ${r.language}` : r.book_name,
    total_qty: parseInt(String(r.total_qty), 10) || 0,
  }));
}

export interface ReportDetailRow {
  village: string;
  group_sector: string;
  book_name: string;
  language: string;
  contact_name: string;
  qty: number;
  avail: number; // 1=YES, 0=NS, 2=B-YES
}

export async function getFilteredReportData(
  org: string,
  groupSector: string,
  period: string,
  bookIds: number[]
): Promise<ReportDetailRow[]> {
  if (!bookIds.length) return [];
  const [y, m, h] = period.split('-');
  const year = parseInt(y, 10);
  const month = parseInt(m, 10);
  const halfNorm = (h || '').replace(/\s/g, '').toLowerCase();
  const sectorVal = groupSector === 'All' ? 'All' : groupSector;
  const placeholders = bookIds.map((_, i) => `$${i + 6}`).join(',');

  const result = await query<{
    village: string;
    group_sector: string;
    book_name: string;
    language: string;
    contact_name: string;
    qty: string;
    avail: number;
  }>(
    `SELECT t.village, t.group_sector, b.book_name, COALESCE(b.language,'') AS language,
            COALESCE(t.contact_name,'') AS contact_name, SUM(t.qty) AS qty, MAX(t.avail)::int AS avail
     FROM ${DB_SCHEMA}.transactions t
     INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
     WHERE t.app = $1 AND t.year = $2 AND t.month = $3
       AND (REPLACE(LOWER(COALESCE(t.half,'')), ' ', '') IN ('1&2','1and2') OR t.half = $4)
       AND ($5::text = 'All' OR t.group_sector = $5)
       AND t.book_id IN (${placeholders})
       AND t.is_deleted = FALSE AND b.is_deleted = FALSE
     GROUP BY t.village, t.group_sector, b.book_name, b.language, t.contact_name, t.avail
     ORDER BY t.group_sector, t.village, t.contact_name ASC, b.language ASC, b.book_name ASC`,
    [org, year, month, halfNorm === '1&2' || halfNorm === '1and2' ? '1&2' : (h || ''), sectorVal, ...bookIds]
  );

  return (result.rows || []).map((r) => ({
    village: r.village || '',
    group_sector: r.group_sector || '',
    book_name: r.book_name || '',
    language: r.language || '',
    contact_name: r.contact_name || '',
    qty: parseInt(String(r.qty), 10) || 0,
    avail: typeof r.avail === 'number' ? r.avail : parseInt(String(r.avail), 10) || 0,
  }));
}
