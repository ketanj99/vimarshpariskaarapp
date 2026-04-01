import { NextRequest, NextResponse } from 'next/server';
import { query, DB_SCHEMA } from '@/lib/db';

export const dynamic = 'force-dynamic';

async function ensureDisplayNameColumn() {
  await query(
    `
      ALTER TABLE ${DB_SCHEMA}.books
      ADD COLUMN IF NOT EXISTS display_name TEXT
    `
  );
}

export async function GET(request: NextRequest) {
  try {
    await ensureDisplayNameColumn();

    const result = await query<{
      book_id: number;
      book_name: string;
      language: string;
      display_name: string | null;
    }>(
      `
        SELECT
          book_id,
          book_name,
          language,
          display_name
        FROM ${DB_SCHEMA}.books
        WHERE is_deleted = FALSE
        ORDER BY language ASC, book_name ASC
      `
    );

    return NextResponse.json({
      items: (result.rows || []).map((r) => ({
        book_id: r.book_id,
        book_name: r.book_name,
        language: r.language,
        display_name: r.display_name ?? '',
      })),
    });
  } catch (error) {
    console.error('Book aliases GET error:', error);
    return NextResponse.json(
      { error: 'Failed to load book aliases: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    await ensureDisplayNameColumn();
    const body = await request.json();
    const items = Array.isArray(body.items) ? body.items : [];

    if (!items.length) {
      return NextResponse.json(
        { error: 'No aliases provided' },
        { status: 400 }
      );
    }

    for (const raw of items) {
      const bookId = Number(raw.book_id);
      const displayName = String(raw.display_name || '').trim();
      if (!bookId) continue;

      await query(
        `
          UPDATE ${DB_SCHEMA}.books
          SET display_name = $2
          WHERE book_id = $1
        `,
        [bookId, displayName || null]
      );
    }

    return NextResponse.json({
      message: 'Book aliases saved successfully.',
      saved: items.length,
    });
  } catch (error) {
    console.error('Book aliases POST error:', error);
    return NextResponse.json(
      { error: 'Failed to save book aliases: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

