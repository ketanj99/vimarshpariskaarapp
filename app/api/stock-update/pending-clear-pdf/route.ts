import { NextRequest, NextResponse } from 'next/server';
import {
  getPendingClearSectorData,
  generatePendingClearSectorPdf,
} from '@/lib/pending-clear-pdf';

export const dynamic = 'force-dynamic';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    let pNos: string[] = [];
    if (Array.isArray(body.pNos)) {
      pNos = body.pNos.map((p: unknown) => String(p).trim()).filter(Boolean);
    } else if (typeof body.pNos === 'string') {
      pNos = body.pNos.split(/[\s,]+/).map((p: string) => p.trim()).filter(Boolean);
    }
    if (pNos.length === 0) {
      return NextResponse.json(
        { error: 'At least one P No. is required. Send { pNos: ["44","45"] } or { pNos: "44,45" }' },
        { status: 400 }
      );
    }
    const app = body.app == null || body.app === '' ? undefined : String(body.app);
    const groupSector = body.groupSector == null || body.groupSector === '' ? undefined : String(body.groupSector);
    const rows = await getPendingClearSectorData(pNos, app, groupSector);
    if (rows.length === 0) {
      return NextResponse.json(
        { error: 'No data found for the given P No. filter.' },
        { status: 404 }
      );
    }
    const includeSector =
      body.includeSector == null ? true : Boolean(body.includeSector);
    const includeSectorCity =
      body.includeSectorCity == null ? true : Boolean(body.includeSectorCity);

    const pdfBuffer = generatePendingClearSectorPdf(rows, {
      includeSector,
      includeSectorCity,
    });
    const filename = `PendingClear_SectorWise_${new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '')}.pdf`;
    return new NextResponse(pdfBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    });
  } catch (error) {
    console.error('Pending clear PDF error:', error);
    return NextResponse.json(
      { error: 'PDF generation failed: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
