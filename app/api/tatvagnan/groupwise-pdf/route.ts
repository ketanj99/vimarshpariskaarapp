import { NextRequest, NextResponse } from 'next/server';
import { fetchTatvagnanGroupWiseData } from '@/lib/tatvagnan-data';
import { generateGroupWisePdf } from '@/lib/pdf-utils';

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

    const rows = await fetchTatvagnanGroupWiseData(pushpNo);

    if (rows.length === 0) {
      return NextResponse.json(
        { error: `No data found for Pushp No. ${pushpNo}` },
        { status: 404 }
      );
    }

    const pdfBuffer = generateGroupWisePdf(rows, pushpNo);

    return new NextResponse(pdfBuffer, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="Tatvagnan_GroupWise_${pushpNo}.pdf"`,
      },
    });
  } catch (error) {
    console.error('Group-wise PDF generation error:', error);
    return NextResponse.json(
      { error: 'PDF generation failed: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
