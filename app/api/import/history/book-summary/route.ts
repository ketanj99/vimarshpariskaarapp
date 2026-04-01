import { NextRequest, NextResponse } from 'next/server';
import { query, DB_SCHEMA } from '@/lib/db';
import { generateBookGroupSummaryPdf, BookGroupSummaryRow } from '@/lib/report-pdfs';

export const dynamic = 'force-dynamic';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const org = (searchParams.get('org') || '').trim();
    const period = (searchParams.get('period') || '').trim();

    if (!org || !period || period.toUpperCase() === 'ALL' || org.toUpperCase() === 'ALL') {
      return NextResponse.json(
        { error: 'Specific APP and Period are required to generate summary.' },
        { status: 400 }
      );
    }

    const [y, m, h] = period.split('-');
    const year = parseInt(y, 10);
    const month = parseInt(m, 10);
    if (!year || !month || !h) {
      return NextResponse.json(
        { error: 'Invalid period format.' },
        { status: 400 }
      );
    }

    const halfNorm = (h || '').replace(/\s/g, '').toLowerCase();
    const halfValue = halfNorm === '1&2' || halfNorm === '1and2' ? '1&2' : h;

    const sql = `
      SELECT
        COALESCE(group_sector, '') AS "group",
        SUM(qty)::bigint AS total_qty
      FROM ${DB_SCHEMA}.transactions
      WHERE app = $1
        AND year = $2
        AND month = $3
        AND (
          REPLACE(LOWER(COALESCE(half,'')), ' ', '') IN ('1&2','1and2')
          OR half = $4
        )
        AND is_deleted = FALSE
      GROUP BY group_sector
      ORDER BY group_sector;
    `;

    const result = await query<BookGroupSummaryRow>(sql, [org, year, month, halfValue]);
    const rows = result.rows || [];

    if (!rows.length) {
      return NextResponse.json(
        { error: 'No data found for selected APP and Period.' },
        { status: 404 }
      );
    }

    const pdfBuffer = generateBookGroupSummaryPdf(rows, period, org);

    return new NextResponse(pdfBuffer, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="Book_Summary_${org}_${period}.pdf"`,
      },
    });
  } catch (error) {
    console.error('Import history book-summary error:', error);
    return NextResponse.json(
      { error: 'Failed to generate summary: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

