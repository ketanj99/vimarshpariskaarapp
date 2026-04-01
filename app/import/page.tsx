'use client';

import { useState, useEffect } from 'react';

interface ImportHistoryRow {
  history_id: number;
  file_name: string;
  import_date: string;
  period: string;
  app: string;
  total_book_qty: number;
  status: string;
}

export default function ImportPage() {
  const [files, setFiles] = useState<FileList | null>(null);
  const [appType, setAppType] = useState<'1' | '2'>('1');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [results, setResults] = useState<any[]>([]);

  const [history, setHistory] = useState<ImportHistoryRow[]>([]);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState('');
  const [filterPeriod, setFilterPeriod] = useState<'ALL' | string>('ALL');
  const [filterApp, setFilterApp] = useState<'ALL' | string>('ALL');
  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [deleteLoading, setDeleteLoading] = useState(false);
  const [deleteMessage, setDeleteMessage] = useState('');
  const [summaryLoading, setSummaryLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!files || files.length === 0) {
      setMessage('Please select at least one file');
      return;
    }

    setLoading(true);
    setMessage('');
    setResults([]);

    const formData = new FormData();
    formData.append('appType', appType);

    for (let i = 0; i < files.length; i++) {
      formData.append('files', files[i]);
    }

    try {
      const response = await fetch('/api/import', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (response.ok) {
        setMessage(data.message || 'Import successful');
        setResults(data.results || []);
      } else {
        setMessage(data.error || 'Import failed');
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const loadHistory = async () => {
    try {
      setHistoryLoading(true);
      setHistoryError('');
      const res = await fetch('/api/import/history');
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || res.statusText);
      }
      setHistory(Array.isArray(data.rows) ? data.rows : []);
    } catch (error) {
      setHistoryError('Error loading import history: ' + (error as Error).message);
    } finally {
      setHistoryLoading(false);
    }
  };

  useEffect(() => {
    loadHistory();
  }, []);

  const formatDate = (value: string) => {
    const d = new Date(value);
    if (isNaN(d.getTime())) return value;
    return d.toLocaleDateString('en-IN', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
  };

  const historyList = Array.isArray(history) ? history : [];
  const uniquePeriods = Array.from(new Set(historyList.map(h => h.period))).sort();
  const uniqueApps = Array.from(new Set(historyList.map(h => h.app))).sort();

  const filteredHistory = historyList.filter((h) => {
    if (filterPeriod !== 'ALL' && h.period !== filterPeriod) return false;
    if (filterApp !== 'ALL' && h.app !== filterApp) return false;
    return true;
  });

  const allFilteredIds = filteredHistory.map((h) => h.history_id);
  const allSelected = allFilteredIds.length > 0 && allFilteredIds.every((id) => selectedIds.has(id));

  const toggleSelectAll = () => {
    if (allSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        allFilteredIds.forEach((id) => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        allFilteredIds.forEach((id) => next.add(id));
        return next;
      });
    }
  };

  const toggleSelect = (id: number) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const handleDeleteSelected = async () => {
    const ids = Array.from(selectedIds);
    if (ids.length === 0) {
      setDeleteMessage('Please select at least one record to delete.');
      return;
    }
    setDeleteLoading(true);
    setDeleteMessage('');
    try {
      const res = await fetch('/api/import/history', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids }),
      });
      const data = await res.json();
      if (res.ok) {
        setDeleteMessage(data.message || 'Deleted successfully.');
        setSelectedIds(new Set());
        loadHistory();
      } else {
        setDeleteMessage(data.error || 'Delete failed.');
      }
    } catch (error) {
      setDeleteMessage('Error: ' + (error as Error).message);
    } finally {
      setDeleteLoading(false);
    }
  };

  const handleGenerateSummary = async () => {
    if (filterPeriod === 'ALL' || filterApp === 'ALL') {
      setDeleteMessage('Please select specific Period and APP to generate summary PDF.');
      return;
    }
    setSummaryLoading(true);
    setDeleteMessage('');
    try {
      const res = await fetch(
        `/api/import/history/book-summary?org=${encodeURIComponent(filterApp)}&period=${encodeURIComponent(filterPeriod)}`
      );
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        setDeleteMessage(data.error || 'Failed to generate summary PDF.');
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Book_Summary_${filterApp}_${filterPeriod}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setDeleteMessage('Error generating summary PDF: ' + (e as Error).message);
    } finally {
      setSummaryLoading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto">
      <h1 className="text-2xl font-bold mb-4">Import Excel to PostgreSQL</h1>

      <div className="bg-white rounded-lg shadow-md p-4 border border-gray-200">
        <form onSubmit={handleSubmit} className="flex flex-wrap items-end gap-3">
          <div className="min-w-[140px]">
            <label className="block text-xs font-medium text-gray-600 mb-1">APP Type</label>
            <select
              value={appType}
              onChange={(e) => setAppType(e.target.value as '1' | '2')}
              className="w-full px-3 py-1.5 text-sm border border-gray-300 rounded focus:ring-1 focus:ring-blue-500"
            >
              <option value="1">VIMARSH</option>
              <option value="2">PARISHKAAR</option>
            </select>
          </div>
          <div className="flex-1 min-w-[180px]">
            <label className="block text-xs font-medium text-gray-600 mb-1">Excel files</label>
            <input
              type="file"
              accept=".xlsx,.xls"
              multiple
              onChange={(e) => setFiles(e.target.files)}
              className="w-full px-3 py-1.5 text-sm border border-gray-300 rounded file:mr-2 file:py-1 file:px-2 file:rounded file:border-0 file:text-sm file:bg-blue-50 file:text-blue-700"
            />
            {files && <span className="text-xs text-gray-500">{files.length} selected</span>}
          </div>
          <button
            type="submit"
            disabled={loading}
            className="px-4 py-1.5 bg-blue-600 text-white text-sm font-medium rounded hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            {loading ? 'Importing...' : 'Import'}
          </button>
        </form>

        {message && (
          <div
            className={`mt-3 p-3 text-sm rounded ${
              message.includes('Error') || message.includes('failed')
                ? 'bg-red-50 text-red-700 border border-red-200'
                : 'bg-green-50 text-green-700 border border-green-200'
            }`}
          >
            {message}
          </div>
        )}

        {results.length > 0 && (
          <div className="mt-3">
            <h3 className="font-semibold text-sm mb-2">Results</h3>
            <div className="space-y-2">
              {results.map((result, index) => (
                <div
                  key={index}
                  className={`p-2 text-sm rounded border ${
                    result.success
                      ? 'bg-green-50 border-green-200'
                      : 'bg-red-50 border-red-200'
                  }`}
                >
                  <p className="font-medium">{result.fileName}</p>
                  <p className="text-xs mt-0.5">
                    {result.message}
                    {result.recordsImported && ` (${result.recordsImported} records)`}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      <div className="mt-6 bg-white rounded-lg shadow-md p-6 border border-gray-200">
        <div className="flex flex-col md:flex-row md:items-end md:justify-between gap-4 mb-4">
          <div>
            <h2 className="text-lg font-semibold">Import History</h2>
            <p className="text-xs text-gray-500">
              Excel ke Import History form jaisa, latest imports ka list.
            </p>
          </div>
          <div className="flex flex-wrap gap-4 items-end">
            {filteredHistory.length > 0 && (
              <button
                type="button"
                onClick={handleDeleteSelected}
                disabled={deleteLoading || selectedIds.size === 0}
                className="px-4 py-2 bg-red-600 text-white text-sm font-medium rounded-md hover:bg-red-700 disabled:bg-gray-300 disabled:cursor-not-allowed"
              >
                {deleteLoading ? 'Deleting...' : `Delete selected (${selectedIds.size})`}
              </button>
            )}
            <button
              type="button"
              onClick={handleGenerateSummary}
              disabled={summaryLoading || filterPeriod === 'ALL' || filterApp === 'ALL'}
              className="px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded-md hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed"
            >
              {summaryLoading ? 'Generating summary...' : 'Book Summary PDF'}
            </button>
            <div className="flex flex-wrap gap-4 text-sm">
            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">Period</label>
              <select
                value={filterPeriod}
                onChange={(e) => setFilterPeriod(e.target.value)}
                className="px-2 py-1.5 border border-gray-300 rounded text-sm"
              >
                <option value="ALL">All Periods</option>
                {uniquePeriods.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">APP</label>
              <select
                value={filterApp}
                onChange={(e) => setFilterApp(e.target.value)}
                className="px-2 py-1.5 border border-gray-300 rounded text-sm"
              >
                <option value="ALL">All</option>
                {uniqueApps.map((a) => (
                  <option key={a} value={a}>{a}</option>
                ))}
              </select>
            </div>
            </div>
          </div>
        </div>

        {deleteMessage && (
          <p className={`text-sm mb-2 ${deleteMessage.includes('Error') || deleteMessage.includes('failed') || deleteMessage.includes('select') ? 'text-red-600' : 'text-green-600'}`}>
            {deleteMessage}
          </p>
        )}

        {historyLoading && (
          <p className="text-sm text-gray-500">Loading import history...</p>
        )}

        {historyError && (
          <p className="text-sm text-red-600">{historyError}</p>
        )}

        {!historyLoading && !historyError && filteredHistory.length === 0 && (
          <p className="text-sm text-gray-500">No import history found.</p>
        )}

        {!historyLoading && !historyError && filteredHistory.length > 0 && (
          <div className="mt-2 overflow-x-auto">
            <table className="w-full border-collapse text-sm">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-2 py-2 w-10">
                    <input
                      type="checkbox"
                      checked={allSelected}
                      onChange={toggleSelectAll}
                      title="Select all"
                      className="rounded border-gray-300"
                    />
                  </th>
                  <th className="px-3 py-2 text-left font-medium text-gray-700">#</th>
                  <th className="px-3 py-2 text-left font-medium text-gray-700">File Name</th>
                  <th className="px-3 py-2 text-left font-medium text-gray-700">Import Date</th>
                  <th className="px-3 py-2 text-left font-medium text-gray-700">Period</th>
                  <th className="px-3 py-2 text-left font-medium text-gray-700">APP</th>
                  <th className="px-3 py-2 text-right font-medium text-gray-700">Total Qty</th>
                  <th className="px-3 py-2 text-left font-medium text-gray-700">Status</th>
                </tr>
              </thead>
              <tbody>
                {filteredHistory.map((row, index) => (
                  <tr key={row.history_id} className={index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-2 py-2">
                      <input
                        type="checkbox"
                        checked={selectedIds.has(row.history_id)}
                        onChange={() => toggleSelect(row.history_id)}
                        className="rounded border-gray-300"
                      />
                    </td>
                    <td className="px-3 py-2 text-gray-600">{index + 1}</td>
                    <td className="px-3 py-2 text-gray-800">{row.file_name}</td>
                    <td className="px-3 py-2 text-gray-700">{formatDate(row.import_date)}</td>
                    <td className="px-3 py-2 text-gray-700">{row.period}</td>
                    <td className="px-3 py-2 text-gray-700">{row.app}</td>
                    <td className="px-3 py-2 text-right text-gray-800">{row.total_book_qty}</td>
                    <td className="px-3 py-2">
                      <span
                        className={
                          row.status === 'SUCCESS'
                            ? 'px-2 py-1 rounded-full text-xs bg-green-50 text-green-700 border border-green-200'
                            : 'px-2 py-1 rounded-full text-xs bg-red-50 text-red-700 border border-red-200'
                        }
                      >
                        {row.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
