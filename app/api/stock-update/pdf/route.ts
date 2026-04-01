import { NextRequest, NextResponse } from 'next/server';
import { generateStockUpdatePdf, StockUpdatePdfRow } from '@/lib/pdf-utils';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const rows = body.rows as StockUpdatePdfRow[] | undefined;
    if (!Array.isArray(rows) || rows.length === 0) {
      return NextResponse.json(
        { error: 'No data provided. Send { rows: [{ book_name, language, period, qty, stock_number }] }' },
        { status: 400 }
      );
    }

    const pdfBuffer = generateStockUpdatePdf(rows);
    const filename = `StockReport_${new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '')}.pdf`;

    return new NextResponse(pdfBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    });
  } catch (error) {
    console.error('Stock update PDF error:', error);
    return NextResponse.json(
      { error: 'PDF generation failed: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
