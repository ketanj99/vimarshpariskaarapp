import { NextRequest, NextResponse } from 'next/server';
import { query, getClient, DB_SCHEMA } from '@/lib/db';

export const dynamic = 'force-dynamic';

interface BackupPayload {
  schema: string;
  createdAt: string;
  tables: Record<string, any[]>;
}

function quoteIdent(identifier: string): string {
  return `"${identifier.replace(/"/g, '""')}"`;
}

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const schema = searchParams.get('schema') || DB_SCHEMA;

    const tablesResult = await query<{ table_name: string }>(
      `
        SELECT table_name
        FROM information_schema.tables
        WHERE table_schema = $1
          AND table_type = 'BASE TABLE'
        ORDER BY table_name
      `,
      [schema]
    );

    const payload: BackupPayload = {
      schema,
      createdAt: new Date().toISOString(),
      tables: {},
    };

    await Promise.all(
      tablesResult.rows.map(async ({ table_name }) => {
        const tableData = await query<any>(
          `SELECT * FROM ${quoteIdent(schema)}.${quoteIdent(table_name)}`
        );
        payload.tables[table_name] = tableData.rows;
      })
    );

    return NextResponse.json(payload);
  } catch (error) {
    console.error('Backup GET error:', error);
    return NextResponse.json(
      { error: 'Failed to create backup: ' + (error as Error).message },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  const client = await getClient();

  try {
    const body = (await request.json()) as Partial<BackupPayload>;

    if (!body || !body.tables || typeof body.tables !== 'object') {
      return NextResponse.json(
        { error: 'Invalid backup payload: missing tables' },
        { status: 400 }
      );
    }

    const schema = body.schema || DB_SCHEMA;
    const tables = body.tables;

    const tableNames = Object.keys(tables);
    if (tableNames.length === 0) {
      return NextResponse.json(
        { error: 'No tables provided in backup payload' },
        { status: 400 }
      );
    }

    const preferredOrder = ['import_history', 'books', 'transactions'];
    const sortedTableNames = [...tableNames].sort((a, b) => {
      const ia = preferredOrder.indexOf(a);
      const ib = preferredOrder.indexOf(b);
      if (ia === -1 && ib === -1) return a.localeCompare(b);
      if (ia === -1) return 1;
      if (ib === -1) return -1;
      return ia - ib;
    });

    await client.query('BEGIN');

    for (const tableName of sortedTableNames) {
      const rows = tables[tableName];
      if (!Array.isArray(rows)) {
        continue;
      }

      const fullTableName = `${quoteIdent(schema)}.${quoteIdent(tableName)}`;

      await client.query(
        `TRUNCATE TABLE ${fullTableName} RESTART IDENTITY CASCADE`
      );

      if (rows.length === 0) {
        continue;
      }

      const columns = Object.keys(rows[0]);
      if (columns.length === 0) {
        continue;
      }

      const columnList = columns.map(quoteIdent).join(', ');

      const values: any[] = [];
      const valuePlaceholders: string[] = [];

      rows.forEach((row, rowIndex) => {
        const baseIndex = rowIndex * columns.length;
        const placeholdersForRow = columns.map(
          (_col, colIndex) => `$${baseIndex + colIndex + 1}`
        );
        valuePlaceholders.push(`(${placeholdersForRow.join(', ')})`);
        for (const col of columns) {
          values.push(
            // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
            (row as any)[col] ?? null
          );
        }
      });

      const insertSql = `INSERT INTO ${fullTableName} (${columnList}) VALUES ${valuePlaceholders.join(
        ', '
      )}`;

      await client.query(insertSql, values);
    }

    await client.query('COMMIT');

    return NextResponse.json({
      message: 'Restore completed successfully.',
      restoredTables: sortedTableNames.length,
    });
  } catch (error) {
    await client.query('ROLLBACK');
    console.error('Backup POST (restore) error:', error);
    return NextResponse.json(
      { error: 'Restore failed: ' + (error as Error).message },
      { status: 500 }
    );
  } finally {
    client.release();
  }
}

