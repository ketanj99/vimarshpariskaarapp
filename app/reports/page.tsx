'use client';

import { useState, useEffect, useCallback } from 'react';

interface ReportFilters {
  organizations: string[];
  groupSectors: string[];
}

interface PeriodsState {
  list: string[];
  loading: boolean;
}

interface BookWithQty {
  book_id: number;
  book_name: string;
  language: string;
  display_label: string;
  total_qty: number;
}

export default function ReportsPage() {
  const [filters, setFilters] = useState<ReportFilters>({
    organizations: [],
    groupSectors: [],
  });
  const [periods, setPeriods] = useState<PeriodsState>({ list: [], loading: false });
  const [org, setOrg] = useState('VIMARSH');
  const [groupSector, setGroupSector] = useState('All');
  const [period, setPeriod] = useState('');
  const [books, setBooks] = useState<BookWithQty[]>([]);
  const [bookFilter, setBookFilter] = useState('');
  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [chkDetailsReport, setChkDetailsReport] = useState(true);
  const [chkSummaryReport, setChkSummaryReport] = useState(true);
  const [chkSecoreSummary, setChkSecoreSummary] = useState(true);
  const [chkAllSummary, setChkAllSummary] = useState(true);
  const [loadingFilters, setLoadingFilters] = useState(true);
  const [loadingBooks, setLoadingBooks] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [message, setMessage] = useState('');
  const [statusText, setStatusText] = useState('Select filters and books, then generate reports.');

  const loadFilters = useCallback(async () => {
    try {
      setLoadingFilters(true);
      const res = await fetch('/api/reports/filters');
      const data = await res.json();
      if (res.ok) {
        setFilters(data);
        if (data.organizations?.length && !data.organizations.includes(org)) {
          setOrg(data.organizations[0]);
        }
      }
    } catch (e) {
      setMessage('Error loading filters: ' + (e as Error).message);
    } finally {
      setLoadingFilters(false);
    }
  }, []);

  const loadPeriods = useCallback(async () => {
    if (!org?.trim()) {
      setPeriods({ list: [], loading: false });
      setPeriod('');
      return;
    }
    try {
      setPeriods((p) => ({ ...p, loading: true }));
      const res = await fetch(`/api/reports/periods?org=${encodeURIComponent(org)}`);
      const data = await res.json();
      if (res.ok) {
        const list = Array.isArray(data.periods) ? data.periods : [];
        setPeriods({ list, loading: false });
        setPeriod((prev) => (list.includes(prev) ? prev : list[0] || ''));
      } else {
        setPeriods({ list: [], loading: false });
        setPeriod('');
      }
    } catch (e) {
      setPeriods({ list: [], loading: false });
      setPeriod('');
    }
  }, [org]);

  const loadBooks = useCallback(async () => {
    if (!period) {
      setBooks([]);
      return;
    }
    try {
      setLoadingBooks(true);
      const res = await fetch(
        `/api/reports/books?org=${encodeURIComponent(org)}&groupSector=${encodeURIComponent(groupSector)}&period=${encodeURIComponent(period)}`
      );
      const data = await res.json();
      if (res.ok) {
        setBooks(Array.isArray(data.books) ? data.books : []);
        setSelectedIds(new Set());
      }
    } catch (e) {
      setMessage('Error loading books: ' + (e as Error).message);
    } finally {
      setLoadingBooks(false);
    }
  }, [org, groupSector, period]);

  useEffect(() => {
    loadFilters();
  }, [loadFilters]);

  useEffect(() => {
    loadPeriods();
  }, [loadPeriods]);

  useEffect(() => {
    loadBooks();
  }, [loadBooks]);

  const filteredBooks = books.filter(
    (b) =>
      !bookFilter.trim() ||
      b.display_label.toLowerCase().includes(bookFilter.toLowerCase()) ||
      String(b.total_qty).includes(bookFilter)
  );

  const toggleBook = (id: number) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectAll = () => {
    if (selectedIds.size === filteredBooks.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredBooks.map((b) => b.book_id)));
    }
  };

  const atLeastOneReportType = chkDetailsReport || chkSummaryReport || chkSecoreSummary || chkAllSummary;

  const handleGenerate = async () => {
    const ids = Array.from(selectedIds);
    if (ids.length === 0) {
      setMessage('Please keep at least one book selected to generate PDF reports.');
      setStatusText('Select at least one book.');
      return;
    }
    if (!atLeastOneReportType) {
      setMessage('Please select at least one report type to generate.');
      setStatusText('Select at least one report type.');
      return;
    }
    setGenerating(true);
    setMessage('');
    setStatusText('Generating PDF reports...');
    try {
      const res = await fetch('/api/reports/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          org,
          groupSector,
          period,
          bookIds: ids,
          reportTypes: {
            details: chkDetailsReport,
            summary: chkSummaryReport,
            sectorSummary: chkSecoreSummary,
            allSummary: chkAllSummary,
          },
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        setMessage(data.error || 'Generate failed.');
        setStatusText('Error.');
        return;
      }
      if (data.files?.length) {
        data.files.forEach((f: { name: string; content: string }) => {
          const a = document.createElement('a');
          a.href = `data:application/pdf;base64,${f.content}`;
          a.download = f.name;
          a.click();
        });
        setMessage(`Successfully generated ${data.files.length} PDF report(s).`);
        setStatusText(`Successfully generated ${data.files.length} PDF report(s)!`);
      } else {
        setMessage('No reports were generated. Please check your selections.');
        setStatusText('No reports generated.');
      }
    } catch (e) {
      setMessage('Error: ' + (e as Error).message);
      setStatusText('Error generating PDF.');
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="max-w-5xl mx-auto">
      <h1 className="text-2xl font-bold mb-2">Report Generation</h1>
      <p className="text-sm text-gray-600 mb-4">Same as ReportGenerationForm: filters, book selection, and report type checkboxes.</p>

      <div className="bg-white rounded-lg shadow-md p-4 border border-gray-200">
        {/* Dropdowns - same order as form: Organization, Group/Sector, Period */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
          <div>
            <label className="block text-xs font-medium text-gray-700 mb-1">Organization</label>
            <select
              value={org}
              onChange={(e) => setOrg(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded text-sm"
            >
              {loadingFilters ? (
                <option>Loading...</option>
              ) : (
                filters.organizations.map((o) => (
                  <option key={o} value={o}>{o}</option>
                ))
              )}
            </select>
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-700 mb-1">Group / Sector</label>
            <select
              value={groupSector}
              onChange={(e) => setGroupSector(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded text-sm"
            >
              {filters.groupSectors.map((s) => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-700 mb-1">Period (for selected Organization)</label>
            <select
              value={period}
              onChange={(e) => setPeriod(e.target.value)}
              disabled={periods.loading}
              className="w-full px-3 py-2 border border-gray-300 rounded text-sm"
            >
              {periods.loading ? (
                <option>Loading...</option>
              ) : periods.list.length === 0 ? (
                <option value="">No periods for this app</option>
              ) : (
                periods.list.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))
              )}
            </select>
          </div>
        </div>

        {/* Book filter + list - same as form */}
        <div className="mb-2">
          <label className="block text-xs font-medium text-gray-700 mb-1">Filter books</label>
          <input
            type="text"
            value={bookFilter}
            onChange={(e) => setBookFilter(e.target.value)}
            placeholder="Search book name, language, qty..."
            className="w-full px-3 py-2 border border-gray-300 rounded text-sm"
          />
        </div>

        <div className="border border-gray-200 rounded overflow-hidden mb-4" style={{ maxHeight: 280 }}>
          <div className="bg-gray-50 px-3 py-2 flex items-center justify-between">
            <span className="text-sm font-medium text-gray-700">Books (keep at least one selected to generate reports)</span>
            <button
              type="button"
              onClick={selectAll}
              className="text-xs text-blue-600 hover:underline"
            >
              {selectedIds.size === filteredBooks.length && filteredBooks.length > 0 ? 'Deselect all' : 'Select all'}
            </button>
          </div>
          <div className="overflow-y-auto p-2" style={{ maxHeight: 240 }}>
            {loadingBooks ? (
              <p className="text-sm text-gray-500 py-4">Loading books...</p>
            ) : filteredBooks.length === 0 ? (
              <p className="text-sm text-gray-500 py-4">No books for this period/sector. Try another period or organization.</p>
            ) : (
              <ul className="space-y-1">
                {filteredBooks.map((b) => (
                  <li key={b.book_id} className="flex items-center gap-2 text-sm">
                    <input
                      type="checkbox"
                      checked={selectedIds.has(b.book_id)}
                      onChange={() => toggleBook(b.book_id)}
                      className="rounded border-gray-300"
                    />
                    <span className="flex-1 text-gray-800">{b.display_label}</span>
                    <span className="text-gray-500 tabular-nums">{b.total_qty.toLocaleString()}</span>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>

        {/* 4 report type checkboxes - same as form (chkDetailsReport, chkSummaryReport, chkSecoreSummary, chkAllSummary) */}
        <div className="mb-4 p-3 bg-gray-50 rounded border border-gray-200">
          <span className="block text-xs font-medium text-gray-700 mb-2">Report types to generate</span>
          <div className="flex flex-wrap gap-6">
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={chkDetailsReport}
                onChange={(e) => setChkDetailsReport(e.target.checked)}
                className="rounded border-gray-300"
              />
              <span className="text-sm">Sector-Village Details</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={chkSummaryReport}
                onChange={(e) => setChkSummaryReport(e.target.checked)}
                className="rounded border-gray-300"
              />
              <span className="text-sm">Sector-Village Summary</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={chkSecoreSummary}
                onChange={(e) => setChkSecoreSummary(e.target.checked)}
                className="rounded border-gray-300"
              />
              <span className="text-sm">Sector Summary</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={chkAllSummary}
                onChange={(e) => setChkAllSummary(e.target.checked)}
                className="rounded border-gray-300"
              />
              <span className="text-sm">All Summary Selected Period</span>
            </label>
          </div>
        </div>

        <button
          type="button"
          onClick={handleGenerate}
          disabled={generating || selectedIds.size === 0 || !atLeastOneReportType}
          className="w-full py-3 bg-blue-600 text-white font-medium rounded-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
        >
          {generating ? 'Generating reports...' : 'Generate PDF reports'}
        </button>

        <p className="mt-3 text-sm text-gray-600">{statusText}</p>
        {message && (
          <p className={`mt-1 text-sm ${message.includes('Error') || message.includes('failed') || message.includes('Please') ? 'text-red-600' : 'text-green-600'}`}>
            {message}
          </p>
        )}
      </div>
    </div>
  );
}
