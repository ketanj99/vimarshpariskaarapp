import { NextRequest, NextResponse } from 'next/server';
import { getFilteredReportData } from '@/lib/reports-data';
import {
  generateDetailsReportPdf,
  generateSummaryReportPdf,
  generateSectorSummaryReportPdf,
  generateAllSummaryReportPdf,
} from '@/lib/report-pdfs';

export const dynamic = 'force-dynamic';

const safeName = (s: string) => s.replace(/[^a-zA-Z0-9_-]/g, '_');

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { org, groupSector, period, bookIds, reportTypes } = body as {
      org?: string;
      groupSector?: string;
      period?: string;
      bookIds?: number[];
      reportTypes?: { details?: boolean; summary?: boolean; sectorSummary?: boolean; allSummary?: boolean };
    };

    const organization = (org || 'VIMARSH').trim();
    const sector = (groupSector || 'All').trim();
    const periodStr = (period || '').trim();
    const ids = Array.isArray(bookIds) ? bookIds.filter((id) => Number.isInteger(id)) : [];
    const types = reportTypes || {};
    const wantDetails = types.details !== false;
    const wantSummary = types.summary !== false;
    const wantSectorSummary = types.sectorSummary !== false;
    const wantAllSummary = types.allSummary !== false;

    if (!periodStr || ids.length === 0) {
      return NextResponse.json(
        { error: 'Period and at least one book are required.' },
        { status: 400 }
      );
    }
    if (!wantDetails && !wantSummary && !wantSectorSummary && !wantAllSummary) {
      return NextResponse.json(
        { error: 'Please select at least one report type to generate.' },
        { status: 400 }
      );
    }

    const rows = await getFilteredReportData(organization, sector, periodStr, ids);
    if (rows.length === 0) {
      return NextResponse.json(
        { error: 'No data found for the selected filters and books.' },
        { status: 404 }
      );
    }

    const prefix = `Vimarsh_${safeName(sector)}_${safeName(periodStr)}`;
    const files: { name: string; content: string }[] = [];

    if (wantDetails) {
      const pdf = generateDetailsReportPdf(rows, sector, periodStr, organization);
      files.push({ name: `${prefix}_Details.pdf`, content: pdf.toString('base64') });
    }
    if (wantSummary) {
      const pdf = generateSummaryReportPdf(rows, sector, periodStr, organization);
      files.push({ name: `${prefix}_Summary.pdf`, content: pdf.toString('base64') });
    }
    if (wantSectorSummary) {
      const pdf = generateSectorSummaryReportPdf(rows, sector, periodStr, organization);
      files.push({ name: `${prefix}_SectorSummary.pdf`, content: pdf.toString('base64') });
    }
    if (wantAllSummary) {
      const pdf = generateAllSummaryReportPdf(rows, periodStr, organization);
      files.push({ name: `${prefix}_AllSummary.pdf`, content: pdf.toString('base64') });
    }

    return NextResponse.json({ files });
  } catch (error) {
    console.error('Reports generate error:', error);
    return NextResponse.json(
      { error: 'Failed to generate reports: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
