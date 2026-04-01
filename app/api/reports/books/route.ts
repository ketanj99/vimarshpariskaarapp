import { NextRequest, NextResponse } from 'next/server';
import { loadBooksWithTotals } from '@/lib/reports-data';

export const dynamic = 'force-dynamic';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const org = searchParams.get('org') || 'VIMARSH';
    const groupSector = searchParams.get('groupSector') || 'All';
    const period = searchParams.get('period') || '';

    if (!period) {
      return NextResponse.json({ books: [] });
    }

    const books = await loadBooksWithTotals(org, groupSector, period);
    return NextResponse.json({ books });
  } catch (error) {
    console.error('Reports books error:', error);
    return NextResponse.json(
      { error: 'Failed to load books: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
