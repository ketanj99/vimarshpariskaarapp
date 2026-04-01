import { NextResponse } from 'next/server';
import { query, DB_SCHEMA } from '@/lib/db';

export const dynamic = 'force-dynamic';

function naturalCompare(a: string, b: string): number {
  const parts = (s: string) => s.split(/(\d+)/).filter(Boolean);
  const aa = parts(a);
  const bb = parts(b);
  for (let i = 0; i < Math.min(aa.length, bb.length); i++) {
    const na = parseInt(aa[i], 10);
    const nb = parseInt(bb[i], 10);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      if (na !== nb) return na - nb;
    } else {
      const c = aa[i].localeCompare(bb[i]);
      if (c !== 0) return c;
    }
  }
  return aa.length - bb.length;
}

export async function GET() {
  try {
    const result = await query<{ group_sector: string }>(
      `SELECT DISTINCT group_sector FROM ${DB_SCHEMA}.transactions WHERE group_sector IS NOT NULL AND TRIM(group_sector) <> '' AND is_deleted = FALSE ORDER BY group_sector`
    );
    const list = (result.rows || []).map((r) => (r.group_sector || '').trim()).filter(Boolean);
    const groupSectors = ['All', ...list.sort(naturalCompare)];
    return NextResponse.json({ groupSectors });
  } catch (error) {
    console.error('Stock-update groups error:', error);
    return NextResponse.json(
      { error: 'Failed to load groups: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
