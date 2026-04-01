import { NextResponse } from 'next/server';
import { loadReportFilters } from '@/lib/reports-data';

export const dynamic = 'force-dynamic';

export async function GET() {
  try {
    const filters = await loadReportFilters();
    return NextResponse.json(filters);
  } catch (error) {
    console.error('Reports filters error:', error);
    return NextResponse.json(
      { error: 'Failed to load filters: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
