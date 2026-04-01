import { NextRequest, NextResponse } from 'next/server';
import { query, getClient, DB_SCHEMA } from '@/lib/db';

export const dynamic = 'force-dynamic';

interface Book {
  importid: number;
  book_id: number;
  book_name: string;
  language: string;
  village: string;
  group_sector: string;
  contact_name: string;
  qty: number;
  avail: number;
  year: number;
  month: number;
  half: string;
  stock_number: string | null;
  stock_date: string | null;
}

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const app = searchParams.get('app') || 'VIMARSH';
    const year = searchParams.get('year');
    const month = searchParams.get('month');
    const half = searchParams.get('half');

    const useAppFilter = app && app !== 'All';
    let sql = '';
    let params: any[] = [];

    if (!year || !month || !half) {
      if (useAppFilter) {
        sql = `
        SELECT 
          MIN(t.importid) as importid,
          b.book_id,
          b.book_name,
          b.language,
          SUM(t.qty) as qty,
          t.year,
          t.month,
          t.half,
          MAX(t.stock_number) as stock_number,
          MAX(t.stock_date) as stock_date
        FROM ${DB_SCHEMA}.transactions t
        INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
        WHERE t.app = $1
          AND t.avail = 0
          AND t.is_deleted = FALSE
          AND b.is_deleted = FALSE
        GROUP BY b.book_id, b.book_name, b.language, t.year, t.month, t.half
        ORDER BY b.language ASC, b.book_name ASC, t.year DESC, t.month DESC, t.half DESC
      `;
        params = [app];
      } else {
        sql = `
        SELECT 
          MIN(t.importid) as importid,
          b.book_id,
          b.book_name,
          b.language,
          SUM(t.qty) as qty,
          t.year,
          t.month,
          t.half,
          MAX(t.stock_number) as stock_number,
          MAX(t.stock_date) as stock_date,
          t.app
        FROM ${DB_SCHEMA}.transactions t
        INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
        WHERE t.avail = 0
          AND t.is_deleted = FALSE
          AND b.is_deleted = FALSE
        GROUP BY b.book_id, b.book_name, b.language, t.year, t.month, t.half, t.app
        ORDER BY b.language ASC, b.book_name ASC, t.app ASC, t.year DESC, t.month DESC, t.half DESC
      `;
        params = [];
      }
    } else {
      const halfNormalized = half.replace(/\s/g, '').toLowerCase();
      let halfCondition = '';
      let appWhere = '';
      if (useAppFilter) {
        appWhere = `t.app = $1 AND `;
        if (halfNormalized === '1&2') {
          halfCondition = `(REPLACE(LOWER(t.half), ' ', '') = '1&2' OR REPLACE(LOWER(t.half), ' ', '') = '1and2')`;
          params = [app, parseInt(year), parseInt(month)];
        } else {
          halfCondition = `t.half = $4`;
          params = [app, parseInt(year), parseInt(month), half];
        }
      } else {
        if (halfNormalized === '1&2') {
          halfCondition = `(REPLACE(LOWER(t.half), ' ', '') = '1&2' OR REPLACE(LOWER(t.half), ' ', '') = '1and2')`;
          params = [parseInt(year), parseInt(month)];
        } else {
          halfCondition = `t.half = $3`;
          params = [parseInt(year), parseInt(month), half];
        }
      }
      const paramYear = useAppFilter ? '$2' : '$1';
      const paramMonth = useAppFilter ? '$3' : '$2';
      const groupByApp = useAppFilter ? '' : ', t.app';
      const selectApp = useAppFilter ? '' : ', t.app';
      const orderByApp = useAppFilter ? '' : ', t.app ASC';
      sql = `
        SELECT 
          MIN(t.importid) as importid,
          b.book_id,
          b.book_name,
          b.language,
          SUM(t.qty) as qty,
          t.year,
          t.month,
          t.half,
          MAX(t.stock_number) as stock_number,
          MAX(t.stock_date) as stock_date${selectApp}
        FROM ${DB_SCHEMA}.transactions t
        INNER JOIN ${DB_SCHEMA}.books b ON t.book_id = b.book_id
        WHERE ${appWhere}t.year = ${paramYear}
          AND t.month = ${paramMonth}
          AND ${halfCondition}
          AND t.avail = 0
          AND t.is_deleted = FALSE
          AND b.is_deleted = FALSE
        GROUP BY b.book_id, b.book_name, b.language, t.year, t.month, t.half${groupByApp}
        ORDER BY b.language ASC, b.book_name ASC${orderByApp}, t.year DESC, t.month DESC, t.half DESC
      `;
    }

    const result = await query(sql, params);

    return NextResponse.json({ books: result.rows });
  } catch (error) {
    console.error('Stock update GET error:', error);
    return NextResponse.json(
      { error: 'Failed to load books: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  const client = await getClient();

  try {
    const body = await request.json();
    const { books, pNo, app } = body;
    const appValue = app || 'VIMARSH';

    if (!Array.isArray(books) || books.length === 0) {
      return NextResponse.json(
        { error: 'No books selected' },
        { status: 400 }
      );
    }

    await client.query('BEGIN');

    let updatedCount = 0;

    for (const book of books) {
      const bookApp = (book as { app?: string }).app || appValue;
      const halfNormalized = book.half.replace(/\s/g, '').toLowerCase();
      let halfCondition = '';
      let params: any[];

      if (halfNormalized === '1&2') {
        halfCondition = `(REPLACE(LOWER(t.half), ' ', '') = '1&2' OR REPLACE(LOWER(t.half), ' ', '') = '1and2')`;
        params = [pNo, bookApp, book.book_id, book.year, book.month];
      } else {
        halfCondition = `t.half = $6`;
        params = [pNo, bookApp, book.book_id, book.year, book.month, book.half];
      }

      const updateSql = `
        UPDATE ${DB_SCHEMA}.transactions t
        SET stock_number = $1::varchar,
            stock_date = CASE WHEN $1 IS NULL OR TRIM(COALESCE($1::varchar, '')) = '' THEN NULL ELSE CURRENT_DATE END,
            updated_date = CURRENT_TIMESTAMP
        FROM ${DB_SCHEMA}.books b
        WHERE t.book_id = b.book_id
          AND t.app = $2
          AND t.book_id = $3
          AND t.year = $4
          AND t.month = $5
          AND ${halfCondition}
          AND t.avail = 0
          AND t.is_deleted = FALSE
          AND b.is_deleted = FALSE
      `;

      const result = await client.query(updateSql, params);
      updatedCount += result.rowCount || 0;
    }

    await client.query('COMMIT');

    return NextResponse.json({
      message: `Successfully updated ${updatedCount} book(s).`,
      updatedCount,
    });
  } catch (error) {
    await client.query('ROLLBACK');
    console.error('Stock update POST error:', error);
    return NextResponse.json(
      { error: 'Update failed: ' + (error as Error).message },
      { status: 500 }
    );
  } finally {
    client.release();
  }
}

export async function DELETE(request: NextRequest) {
  const client = await getClient();

  try {
    const body = await request.json();
    const { books, app } = body as { books?: any[]; app?: string };
    const appValue = app || 'VIMARSH';

    if (!Array.isArray(books) || books.length === 0) {
      return NextResponse.json(
        { error: 'No books selected' },
        { status: 400 }
      );
    }

    await client.query('BEGIN');

    let deletedCount = 0;

    for (const book of books) {
      const bookApp = (book as { app?: string }).app || appValue;
      const halfNormalized = String(book.half || '').replace(/\s/g, '').toLowerCase();
      let halfCondition = '';
      let params: any[];

      if (halfNormalized === '1&2') {
        halfCondition = `(REPLACE(LOWER(t.half), ' ', '') = '1&2' OR REPLACE(LOWER(t.half), ' ', '') = '1and2')`;
        params = [bookApp, book.book_id, book.year, book.month];
      } else {
        halfCondition = `t.half = $5`;
        params = [bookApp, book.book_id, book.year, book.month, book.half];
      }

      const deleteSql = `
        UPDATE ${DB_SCHEMA}.transactions t
        SET is_deleted = TRUE,
            updated_date = CURRENT_TIMESTAMP
        FROM ${DB_SCHEMA}.books b
        WHERE t.book_id = b.book_id
          AND t.app = $1
          AND t.book_id = $2
          AND t.year = $3
          AND t.month = $4
          AND ${halfCondition}
          AND t.avail = 0
          AND t.is_deleted = FALSE
          AND b.is_deleted = FALSE
      `;

      const result = await client.query(deleteSql, params);
      deletedCount += result.rowCount || 0;
    }

    await client.query('COMMIT');

    return NextResponse.json({
      message: `Successfully deleted ${deletedCount} record(s).`,
      deletedCount,
    });
  } catch (error) {
    await client.query('ROLLBACK');
    console.error('Stock update DELETE error:', error);
    return NextResponse.json(
      { error: 'Delete failed: ' + (error as Error).message },
      { status: 500 }
    );
  } finally {
    client.release();
  }
}
