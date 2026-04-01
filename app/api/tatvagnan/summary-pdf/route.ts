import { NextRequest, NextResponse } from 'next/server';
import { fetchTatvagnanDataForPdf } from '@/lib/tatvagnan-data';
import { generateSummaryPdf } from '@/lib/pdf-utils';

export const dynamic = 'force-dynamic';

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const pushpNo = searchParams.get('pushpNo');

    if (!pushpNo) {
      return NextResponse.json(
        { error: 'Pushp No. is required' },
        { status: 400 }
      );
    }

    const rows = await fetchTatvagnanDataForPdf(pushpNo);

    if (rows.length === 0) {
      return NextResponse.json(
        { error: `No data found for Pushp No. ${pushpNo}` },
        { status: 404 }
      );
    }

    const pdfBuffer = generateSummaryPdf(rows, pushpNo);

    return new NextResponse(pdfBuffer, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="Tatvagnan_Summary_${pushpNo}.pdf"`,
      },
    });
  } catch (error) {
    console.error('Summary PDF generation error:', error);
    return NextResponse.json(
      { error: 'PDF generation failed: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
