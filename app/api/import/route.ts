import { NextRequest, NextResponse } from 'next/server';
import { query, getClient, DB_SCHEMA } from '@/lib/db';
import {
  parseExcelFile,
  getFirstSheet,
  sheetToJSON,
  extractPeriodFromDataRows,
  parseAvailValue,
  getCellByHeader,
  ExcelRow,
} from '@/lib/excel-utils';

export const dynamic = 'force-dynamic';

interface ImportResult {
  fileName: string;
  success: boolean;
  message: string;
  recordsImported?: number;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const appType = formData.get('appType') as string;
    const files = formData.getAll('files') as File[];

    if (!appType || !['1', '2'].includes(appType)) {
      return NextResponse.json(
        { error: 'Invalid APP type' },
        { status: 400 }
      );
    }

    if (!files || files.length === 0) {
      return NextResponse.json(
        { error: 'No files uploaded' },
        { status: 400 }
      );
    }

    const appName = appType === '1' ? 'VIMARSH' : 'PARISHKAAR';
    const results: ImportResult[] = [];

    for (const file of files) {
      const result = await importSingleFile(file, appName);
      results.push(result);
    }

    const successCount = results.filter(r => r.success).length;
    const totalCount = results.length;

    return NextResponse.json({
      message: `Import completed: ${successCount}/${totalCount} files successful`,
      results,
    });
  } catch (error) {
    console.error('Import error:', error);
    return NextResponse.json(
      { error: 'Import failed: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

async function importSingleFile(file: File, appName: string): Promise<ImportResult> {
  const client = await getClient();

  try {
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = parseExcelFile(buffer);

    // Always use FIRST sheet for import (match VBA behaviour)
    const sheetNames = workbook.SheetNames || [];
    console.log('Excel sheet names for import:', sheetNames);

    const dataSheet = getFirstSheet(workbook);
    if (!dataSheet) {
      return {
        fileName: file.name,
        success: false,
        message: 'No sheets found in Excel file',
      };
    }

    const dataRows = sheetToJSON(dataSheet);
    if (dataRows.length > 0) {
      console.log('Excel header keys (first data row):', Object.keys(dataRows[0] as ExcelRow));
    }

    // PERIOD always derived from YEAR, MONTH, HALF columns in data rows (VBA matching behaviour)
    const periodInfo = extractPeriodFromDataRows(dataRows as ExcelRow[]);

    if (!periodInfo) {
      return {
        fileName: file.name,
        success: false,
        message: 'Could not extract period information from YEAR/MONTH/HALF columns',
      };
    }

    await client.query('BEGIN');

    // Create import history first
    const historyQuery = `
      INSERT INTO ${DB_SCHEMA}.import_history (
        file_name, period, app, import_date, total_book_qty, status
      ) VALUES ($1, $2, $3, CURRENT_TIMESTAMP, $4, 'SUCCESS')
      RETURNING history_id
    `;

    const periodStr = `${periodInfo.year}-${periodInfo.month}-${periodInfo.half}`;
    // Total book quantity like old VBA (sum of QTY)
    const totalBookQty = (dataRows as ExcelRow[]).reduce((sum, row) => {
      const qtyRaw =
        (getCellByHeader(row, ['QTY', 'QUANTITY']) as string | number | undefined) ?? '0';
      const qtyNum = typeof qtyRaw === 'number' ? qtyRaw : parseInt(String(qtyRaw), 10);
      return sum + (Number.isNaN(qtyNum) ? 0 : qtyNum);
    }, 0);

    const historyResult = await client.query(historyQuery, [file.name, periodStr, appName, totalBookQty]);

    const historyId = historyResult.rows[0].history_id;
    let recordsImported = 0;

    for (const row of dataRows as ExcelRow[]) {
      // Skip header or empty rows
      const bookName = (getCellByHeader(row, [
        'BOOK NAME',
        'BOOK TITLE',
        'BOOK',
        'TITLE',
      ]) as string) || '';
      const language = (getCellByHeader(row, ['LANG CODE', 'LANGUAGE', 'LANG']) as string) || '';
      
      if (!bookName || !language) continue;

      // Check if book exists, insert if not
      const bookCheckQuery = `
        SELECT book_id FROM ${DB_SCHEMA}.books 
        WHERE book_name = $1 AND language = $2
      `;
      const bookCheck = await client.query(bookCheckQuery, [bookName, language]);

      let bookId: number;
      if (bookCheck.rows.length > 0) {
        bookId = bookCheck.rows[0].book_id;
      } else {
        const bookInsertQuery = `
          INSERT INTO ${DB_SCHEMA}.books (book_name, language)
          VALUES ($1, $2)
          RETURNING book_id
        `;
        const bookInsert = await client.query(bookInsertQuery, [bookName, language]);
        bookId = bookInsert.rows[0].book_id;
      }

      // Extract transaction data
      const qtyRaw =
        (getCellByHeader(row, ['QTY', 'QUANTITY']) as string | number | undefined) ?? '0';
      const qtyVal =
        typeof qtyRaw === 'number' ? qtyRaw : parseInt(String(qtyRaw), 10) || 0;
      const availRaw =
        getCellByHeader(row, ['AVAIL', 'AVAILABLE', 'AVAIL.']) ?? 0;
      const avail = parseAvailValue(availRaw);
      const village =
        (getCellByHeader(row, ['VILLAGE NAME', 'VILLAGE']) as string) || '';
      const groupSector =
        (getCellByHeader(row, ['GROUP(SECTOR)', 'GROUP SECTOR', 'GROUP', 'SECTOR']) as string) ||
        '';
      const contactName =
        (getCellByHeader(row, [
          'CONTACT NAME',
          'NAME(CONTACT NO.)',
          'NAME(CONTACT NO)',
          'NAME',
          'CONTACT',
        ]) as string) || '';

      // Year/month/half from row if present, otherwise from derived periodInfo
      const yearRaw = getCellByHeader(row, ['YEAR']);
      const monthRaw = getCellByHeader(row, ['MONTH']);
      const halfRaw =
        getCellByHeader(row, ['HALF', 'HALF PERIOD', 'HALF_PERIOD']) ?? periodInfo.half;

      const yearVal = parseInt(String(yearRaw ?? periodInfo.year), 10) || periodInfo.year;
      const monthVal = parseInt(String(monthRaw ?? periodInfo.month), 10) || periodInfo.month;
      const halfVal = halfRaw ? halfRaw.toString().trim() : periodInfo.half;

      // Insert transaction
      const transactionQuery = `
        INSERT INTO ${DB_SCHEMA}.transactions (
          history_id, book_id, year, month, half, qty, avail, 
          village, group_sector, contact_name, app
        ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
      `;

      await client.query(transactionQuery, [
        historyId,
        bookId,
        yearVal,
        monthVal,
        halfVal,
        qtyVal,
        avail,
        village,
        groupSector,
        contactName,
        appName,
      ]);

      recordsImported++;
    }

    await client.query('COMMIT');

    return {
      fileName: file.name,
      success: true,
      message: 'Import successful',
      recordsImported,
    };
  } catch (error) {
    await client.query('ROLLBACK');
    console.error('Import single file error:', error);
    return {
      fileName: file.name,
      success: false,
      message: 'Error: ' + (error as Error).message,
    };
  } finally {
    client.release();
  }
}
