import { NextRequest, NextResponse } from 'next/server';
import { getClient, DB_SCHEMA } from '@/lib/db';
import { parseExcelFile, sheetToJSON } from '@/lib/excel-utils';

export const dynamic = 'force-dynamic';

export async function POST(request: NextRequest) {
  const client = await getClient();

  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return NextResponse.json(
        { error: 'No file uploaded' },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = parseExcelFile(buffer);

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      return NextResponse.json(
        { error: 'No data sheet found' },
        { status: 400 }
      );
    }

    const rows = sheetToJSON(sheet);

    if (rows.length === 0) {
      return NextResponse.json(
        { error: 'Sheet is empty' },
        { status: 400 }
      );
    }

    const firstRow = rows[0];
    const firstCell = Object.values(firstRow)[0] as string;

    let pushpNo: string;
    if (firstCell && /pushp/i.test(firstCell)) {
      const match = firstCell.match(/\d+/);
      pushpNo = match ? match[0] : '';
    } else if (firstCell && /^\d+$/.test(firstCell)) {
      pushpNo = firstCell;
    } else {
      return NextResponse.json(
        { error: 'Could not extract Pushp No. from data' },
        { status: 400 }
      );
    }

    if (!pushpNo) {
      return NextResponse.json(
        { error: 'Invalid Pushp No.' },
        { status: 400 }
      );
    }

    await client.query('BEGIN');

    await client.query(
      `DELETE FROM ${DB_SCHEMA}.tatvagnan_data WHERE pushp_no = $1`,
      [parseInt(pushpNo)]
    );

    let recordsImported = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const group = row['GROUP'] || row['Group'] || row['group'];
      const lang = row['LANG'] || row['Lang'] || row['lang'] || row['Language'];

      if (!group || !lang) continue;

      const oldQty = parseInt(row['OLD'] || row['Old'] || row['old'] || '0') || 0;
      const newQty = parseInt(row['NEW'] || row['New'] || row['new'] || '0') || 0;
      const rmvQty = parseInt(row['RMV'] || row['Rmv'] || row['rmv'] || row['REMOVE'] || '0') || 0;

      await client.query(
        `INSERT INTO ${DB_SCHEMA}.tatvagnan_data (pushp_no, "group", lang, old_qty, new_qty, rmv_qty) VALUES ($1, $2, $3, $4, $5, $6)`,
        [parseInt(pushpNo), group, lang, oldQty, newQty, rmvQty]
      );
      recordsImported++;
    }

    await client.query('COMMIT');

    return NextResponse.json({
      message: 'Import successful',
      pushpNo: pushpNo,
      recordsImported,
    });
  } catch (error) {
    await client.query('ROLLBACK');
    console.error('Tatvagnan import error:', error);
    return NextResponse.json(
      { error: 'Import failed: ' + (error as Error).message },
      { status: 500 }
    );
  } finally {
    client.release();
  }
}
