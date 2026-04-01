import { NextRequest, NextResponse } from 'next/server';
import { query, DB_SCHEMA } from '@/lib/db';

export async function GET(request: NextRequest) {
  try {
    const result = await query(
      `
      SELECT history_id,
             file_name,
             import_date,
             period,
             COALESCE(app, 'VIMARSH') AS app,
             total_book_qty,
             status
      FROM ${DB_SCHEMA}.import_history
      WHERE is_deleted = FALSE
      ORDER BY import_date DESC
      `
    );

    return NextResponse.json({ rows: result.rows });
  } catch (error) {
    console.error('Import history error:', error);
    return NextResponse.json(
      { error: 'Failed to load import history: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

export async function DELETE(request: NextRequest) {
  try {
    const body = await request.json().catch(() => ({}));
    const ids = Array.isArray(body.ids) ? body.ids : [];
    const numIds = ids.filter((id: unknown) => typeof id === 'number' && Number.isInteger(id));

    if (numIds.length === 0) {
      return NextResponse.json({ error: 'No valid history IDs provided' }, { status: 400 });
    }

    const placeholders = numIds.map((_, i) => `$${i + 1}`).join(',');
    await query(
      `UPDATE ${DB_SCHEMA}.import_history SET is_deleted = TRUE WHERE history_id IN (${placeholders})`,
      numIds
    );

    return NextResponse.json({ message: `${numIds.length} record(s) deleted.` });
  } catch (error) {
    console.error('Import history delete error:', error);
    return NextResponse.json(
      { error: 'Failed to delete: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

