import { NextRequest, NextResponse } from 'next/server';
import { loadPeriodsForOrg } from '@/lib/reports-data';

export const dynamic = 'force-dynamic';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const org = searchParams.get('org') || '';

    const periods = await loadPeriodsForOrg(org);
    return NextResponse.json({ periods });
  } catch (error) {
    console.error('Reports periods error:', error);
    return NextResponse.json(
      { error: 'Failed to load periods: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
