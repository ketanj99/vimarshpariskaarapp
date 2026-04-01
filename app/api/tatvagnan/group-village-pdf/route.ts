import { NextRequest, NextResponse } from 'next/server';
import { fetchGroupVillageForPdf } from '@/lib/tatvagnan-data';
import { generateGroupVillagePdf } from '@/lib/pdf-utils';

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

    const rows = await fetchGroupVillageForPdf(pushpNo);

    const pdfBuffer = generateGroupVillagePdf(rows, pushpNo);

    return new NextResponse(pdfBuffer, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="Tatvagnan_GroupVillage_${pushpNo}.pdf"`,
      },
    });
  } catch (error) {
    console.error('Group Village PDF error:', error);
    return NextResponse.json(
      { error: 'PDF generation failed: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
