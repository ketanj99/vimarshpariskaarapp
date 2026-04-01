'use client';

import { useState, useEffect, Fragment } from 'react';

interface Book {
  importid: number;
  book_id: number;
  book_name: string;
  language: string;
  qty: number;
  year: number;
  month: number;
  half: string;
  stock_number: string | null;
  stock_date: string | null;
  app?: string;
}

export default function StockUpdatePage() {
  const [books, setBooks] = useState<Book[]>([]);
  const [selectedBooks, setSelectedBooks] = useState<Set<string>>(new Set());
  const [pNo, setPNo] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');

  const [filterEnabled, setFilterEnabled] = useState(false);
  const [pNoFilter, setPNoFilter] = useState<'all' | 'blank' | 'assigned'>('blank');
  const [pdfLoading, setPdfLoading] = useState(false);
  const [pendingClearPNos, setPendingClearPNos] = useState('');
  const [pendingClearPdfLoading, setPendingClearPdfLoading] = useState(false);
  const [pendingClearVillagePdfLoading, setPendingClearVillagePdfLoading] = useState(false);
  const [includeSectorSummary, setIncludeSectorSummary] = useState(true);
  const [includeSectorCitySummary, setIncludeSectorCitySummary] = useState(true);

  const [selectedApp, setSelectedApp] = useState('VIMARSH');
  const [organizations, setOrganizations] = useState<string[]>(['VIMARSH']);
  const [groupSectors, setGroupSectors] = useState<string[]>(['All']);
  const [selectedGroup, setSelectedGroup] = useState('All');
  const [pendingClearFormOpen, setPendingClearFormOpen] = useState(false);
  const [searchLanguage, setSearchLanguage] = useState('');
  const [searchBookName, setSearchBookName] = useState('');
  const [searchPeriod, setSearchPeriod] = useState('');
  const [searchPNo, setSearchPNo] = useState('');

  const cyclePNoFilter = () => {
    if (pNoFilter === 'all') {
      setPNoFilter('blank');
    } else if (pNoFilter === 'blank') {
      setPNoFilter('assigned');
    } else {
      setPNoFilter('all');
    }
  };

  const getPNoFilterLabel = () => {
    if (pNoFilter === 'all') return { icon: '🌐', text: 'All', color: 'bg-blue-600' };
    if (pNoFilter === 'blank') return { icon: '📋', text: 'Pending', color: 'bg-orange-600' };
    return { icon: '✅', text: 'Clear', color: 'bg-green-600' };
  };

  const handleGeneratePDF = async () => {
    if (sortedForGroup.length === 0) {
      setMessage('No data to generate PDF.');
      return;
    }
    setPdfLoading(true);
    setMessage('');
    try {
      const rows = sortedForGroup.map((b) => ({
        book_name: b.book_name,
        language: b.language,
        period: formatPeriod(b.year, b.month, b.half),
        qty: b.qty,
        stock_number: b.stock_number,
      }));
      const res = await fetch('/api/stock-update/pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || res.statusText);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `StockReport_${new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '')}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
      setMessage('PDF generated successfully.');
    } catch (e) {
      setMessage('Error generating PDF: ' + (e as Error).message);
    } finally {
      setPdfLoading(false);
    }
  };

  const handleGeneratePendingClearSectorPDF = async () => {
    const pNos = pendingClearPNos.split(/[\s,]+/).map((p) => p.trim()).filter(Boolean);
    if (pNos.length === 0) {
      setMessage('Enter at least one P No. (comma-separated) for Pending Clear Sector PDF.');
      return;
    }
    setPendingClearPdfLoading(true);
    setMessage('');
    try {
      const res = await fetch('/api/stock-update/pending-clear-pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          pNos,
          app: selectedApp === 'All' ? undefined : selectedApp,
          groupSector: selectedGroup === 'All' ? undefined : selectedGroup,
          includeSector: includeSectorSummary,
          includeSectorCity: includeSectorCitySummary,
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || res.statusText);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `PendingClear_SectorWise_${new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '')}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
      setMessage('Pending Clear Sector PDF generated (same as VBA).');
    } catch (e) {
      setMessage('Error: ' + (e as Error).message);
    } finally {
      setPendingClearPdfLoading(false);
    }
  };

  const handleGeneratePendingClearVillagePDF = async () => {
    const pNos = pendingClearPNos.split(/[\s,]+/).map((p) => p.trim()).filter(Boolean);
    if (pNos.length === 0) {
      setMessage('Enter at least one P No. (comma-separated) for Pending Clear Village PDF.');
      return;
    }
    setPendingClearVillagePdfLoading(true);
    setMessage('');
    try {
      const res = await fetch('/api/stock-update/pending-clear-village-pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          pNos,
          app: selectedApp === 'All' ? undefined : selectedApp,
          groupSector: selectedGroup === 'All' ? undefined : selectedGroup,
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || res.statusText);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `PendingClear_VillageWise_${new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '')}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
      setMessage('Pending Clear Village PDF generated.');
    } catch (e) {
      setMessage('Error: ' + (e as Error).message);
    } finally {
      setPendingClearVillagePdfLoading(false);
    }
  };

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch('/api/reports/filters');
        const data = await res.json();
        if (res.ok && data.organizations?.length) {
          setOrganizations(data.organizations);
          if (!data.organizations.includes(selectedApp)) {
            setSelectedApp(data.organizations[0]);
          }
        }
        if (res.ok && Array.isArray(data.groupSectors)) {
          setGroupSectors(data.groupSectors.length ? data.groupSectors : ['All']);
        }
      } catch {
        // keep default VIMARSH
      }
    })();
  }, []);

  const loadGroupSectors = async () => {
    try {
      const res = await fetch('/api/stock-update/groups');
      const data = await res.json();
      if (res.ok && Array.isArray(data.groupSectors)) {
        setGroupSectors(data.groupSectors.length ? data.groupSectors : ['All']);
      }
    } catch {
      setGroupSectors(['All']);
    }
  };

  useEffect(() => {
    loadBooks();
  }, [selectedApp]);

  // Format period like Excel: 2026-Jan(16 To 31)
  const formatStockDate = (stockDate: string | null | undefined): string => {
    if (!stockDate) return '-';
    try {
      const d = new Date(stockDate);
      if (Number.isNaN(d.getTime())) return String(stockDate);
      return d.toLocaleDateString('en-IN', { dateStyle: 'short' });
    } catch {
      return String(stockDate);
    }
  };

  const formatPeriod = (year: number, month: number, half: string): string => {
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthName = monthNames[month - 1];
    
    // Normalize half (remove spaces)
    const halfNormalized = half.replace(/\s/g, '').toLowerCase();
    
    if (halfNormalized === '1&2' || halfNormalized === '1and2') {
      // Full month - get last day
      const lastDay = new Date(year, month, 0).getDate();
      return `${year}-${monthName}(1 To ${lastDay})`;
    } else if (half === '1') {
      return `${year}-${monthName}(1 To 15)`;
    } else if (half === '2') {
      // Second half - get last day
      const lastDay = new Date(year, month, 0).getDate();
      return `${year}-${monthName}(16 To ${lastDay})`;
    }
    
    return `${year}-${monthName}-${half}`;
  };

  const uniqueLanguages = Array.from(new Set(books.map(b => b.language))).sort();
  const uniqueBookNames = Array.from(new Set(books.map(b => b.book_name))).sort();

  // Get unique periods for dropdown
  const uniquePeriods = Array.from(
    new Set(books.map(b => formatPeriod(b.year, b.month, b.half)))
  ).sort();

  // Filter books based on P No. filter and search criteria
  const filteredBooks = books.filter((book) => {
    // P No. status filter (All/Pending/Clear)
    if (pNoFilter === 'blank' && book.stock_number) return false;
    if (pNoFilter === 'assigned' && !book.stock_number) return false;

    if (searchLanguage && !book.language.toLowerCase().includes(searchLanguage.toLowerCase())) {
      return false;
    }
    if (searchBookName && !book.book_name.toLowerCase().includes(searchBookName.toLowerCase())) {
      return false;
    }

    // Period search (contains match)
    if (searchPeriod) {
      const bookPeriod = formatPeriod(book.year, book.month, book.half);
      if (!bookPeriod.toLowerCase().includes(searchPeriod.toLowerCase())) {
        return false;
      }
    }

    // P No. search (contains match)
    if (searchPNo) {
      if (!book.stock_number || !book.stock_number.toLowerCase().includes(searchPNo.toLowerCase())) {
        return false;
      }
    }

    return true;
  });

  // Group filtered books by book_id for visual separation
  const bookGroupColors = [
    'bg-blue-50 border-l-4 border-blue-300',
    'bg-amber-50 border-l-4 border-amber-300',
    'bg-emerald-50 border-l-4 border-emerald-300',
    'bg-sky-50 border-l-4 border-sky-300',
    'bg-rose-50 border-l-4 border-rose-300',
    'bg-violet-50 border-l-4 border-violet-300',
    'bg-teal-50 border-l-4 border-teal-300',
    'bg-orange-50 border-l-4 border-orange-300',
  ];

  // Same ordering as VBA StockUpdateForm: language ASC, book_name ASC, year DESC, month DESC, half DESC
  const halfOrder = (h: string) => {
    const n = (h || '').replace(/\s/g, '').toLowerCase();
    if (n === '2') return 2;
    if (n === '1') return 1;
    return 0; // 1&2 / 1and2
  };
  const sortedForGroup = [...filteredBooks].sort((a, b) => {
    if ((a.language || '') !== (b.language || '')) return (a.language || '').localeCompare(b.language || '');
    if ((a.book_name || '') !== (b.book_name || '')) return (a.book_name || '').localeCompare(b.book_name || '');
    if (a.year !== b.year) return b.year - a.year;
    if (a.month !== b.month) return b.month - a.month;
    return halfOrder(b.half) - halfOrder(a.half);
  });

  const bookGroups = sortedForGroup.reduce<{ book_id: number; book_name: string; language: string; books: typeof filteredBooks }[]>(
    (acc, book) => {
      const last = acc[acc.length - 1];
      if (last && last.book_id === book.book_id) {
        last.books.push(book);
      } else {
        acc.push({
          book_id: book.book_id,
          book_name: book.book_name,
          language: book.language,
          books: [book],
        });
      }
      return acc;
    },
    []
  );

  const loadBooks = async () => {
    setLoading(true);
    try {
      const response = await fetch(`/api/stock-update?app=${encodeURIComponent(selectedApp)}`);
      const data = await response.json();
      if (response.ok) {
        setBooks(data.books || []);
      } else {
        setMessage('Error loading books: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const handleSelectAll = () => {
    // Get indices of filtered books from original books array
    const filteredIndices = filteredBooks.map((filteredBook) =>
      String(books.findIndex((b) => b.importid === filteredBook.importid))
    );

    if (filteredIndices.every((idx) => selectedBooks.has(idx))) {
      // Deselect all filtered books
      const newSet = new Set(selectedBooks);
      filteredIndices.forEach((idx) => newSet.delete(idx));
      setSelectedBooks(newSet);
    } else {
      // Select all filtered books
      const newSet = new Set(selectedBooks);
      filteredIndices.forEach((idx) => newSet.add(idx));
      setSelectedBooks(newSet);
    }
  };

  const handleSelectBook = (index: number) => {
    const newSet = new Set(selectedBooks);
    if (newSet.has(String(index))) {
      newSet.delete(String(index));
    } else {
      newSet.add(String(index));
    }
    setSelectedBooks(newSet);
  };

  const handleUpdate = async () => {
    if (selectedBooks.size === 0) {
      setMessage('Please select at least one book');
      return;
    }

    const selectedBooksData = Array.from(selectedBooks).map((idx) => books[parseInt(idx)]);

    setLoading(true);
    setMessage('');

    try {
      const response = await fetch('/api/stock-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          books: selectedBooksData.map((b) => ({ ...b, app: selectedApp === 'All' ? b.app : selectedApp })),
          pNo: pNo === '' ? null : pNo.trim(),
          app: selectedApp === 'All' ? undefined : selectedApp,
        }),
      });

      const data = await response.json();

      if (response.ok) {
        setMessage(data.message || 'Update successful');
        setPNo('');
        setSelectedBooks(new Set());
        await loadBooks();
      } else {
        setMessage('Error: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const handleDelete = async () => {
    if (selectedBooks.size === 0) {
      setMessage('Please select at least one book');
      return;
    }

    const confirmed = window.confirm(
      `Are you sure you want to delete ${selectedBooks.size} selected record(s) from data? This cannot be undone.`
    );
    if (!confirmed) return;

    const selectedBooksData = Array.from(selectedBooks).map((idx) => books[parseInt(idx)]);

    setLoading(true);
    setMessage('');

    try {
      const response = await fetch('/api/stock-update', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          books: selectedBooksData.map((b) => ({
            ...b,
            app: selectedApp === 'All' ? b.app : selectedApp,
          })),
          app: selectedApp === 'All' ? undefined : selectedApp,
        }),
      });

      const data = await response.json();

      if (response.ok) {
        setMessage(data.message || `Deleted ${data.deletedCount || 0} record(s).`);
        setSelectedBooks(new Set());
        await loadBooks();
      } else {
        setMessage('Error: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const handleClearPNo = () => {
    setPNo('');
    setMessage('💡 P No. field cleared. Update selected books with blank P No. to remove from database.');
  };

  return (
    <div className="max-w-7xl mx-auto">

      {books.length > 0 && (
        <div className="bg-white rounded-lg shadow-md p-4 border border-gray-200 mb-6">
          <div className="flex items-start justify-between gap-4 mb-4">
            <div className="flex-1">
              <h2 className="text-base font-semibold mb-2 text-gray-800">Filter (Table header jaisa)</h2>
              <p className="text-xs text-gray-500 mb-3">Language, Book Name, Period, P No. — table ke column order me</p>
              
              <div className="grid grid-cols-1 md:grid-cols-6 gap-3 mb-3">
                <div>
                  <label className="block text-xs font-medium text-gray-700 mb-1">Organization (APP)</label>
                  <select
                    value={selectedApp}
                    onChange={(e) => setSelectedApp(e.target.value)}
                    className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm"
                  >
                    <option value="All">All</option>
                    {organizations.map((org) => (
                      <option key={org} value={org}>{org}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-700 mb-1">Language</label>
                  <select
                    value={searchLanguage}
                    onChange={(e) => setSearchLanguage(e.target.value)}
                    className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm"
                  >
                    <option value="">All</option>
                    {uniqueLanguages.map((lang) => (
                      <option key={lang} value={lang}>{lang}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-700 mb-1">Book Name</label>
                  <select
                    value={searchBookName}
                    onChange={(e) => setSearchBookName(e.target.value)}
                    className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm"
                  >
                    <option value="">All Books</option>
                    {uniqueBookNames.map((name) => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-gray-700 mb-1">Period</label>
                  <select
                    value={searchPeriod}
                    onChange={(e) => setSearchPeriod(e.target.value)}
                    className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm"
                  >
                    <option value="">All Periods</option>
                    {uniquePeriods.map((period) => (
                      <option key={period} value={period}>{period}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-gray-700 mb-1">P No. (Filter)</label>
                  <input
                    type="text"
                    value={searchPNo}
                    onChange={(e) => setSearchPNo(e.target.value)}
                    placeholder="P No. se search"
                    className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-gray-700 mb-1">P No. (Assign)</label>
                  <input
                    type="text"
                    value={pNo}
                    onChange={(e) => setPNo(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Escape') {
                        handleClearPNo();
                      }
                    }}
                    placeholder="P No. enter karein"
                    className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm"
                  />
                </div>
              </div>

              <div className="flex gap-2">
                {(searchLanguage || searchBookName || searchPeriod || searchPNo) && (
                  <button
                    onClick={() => {
                      setSearchLanguage('');
                      setSearchBookName('');
                      setSearchPeriod('');
                      setSearchPNo('');
                    }}
                    className="text-xs text-red-600 hover:text-red-700 font-medium"
                  >
                    Clear Filters
                  </button>
                )}
                {pNo && (
                  <button
                    onClick={handleClearPNo}
                    className="text-xs px-3 py-1 bg-gray-600 text-white rounded hover:bg-gray-700"
                  >
                    Clear P No.
                  </button>
                )}
                <button
                  onClick={handleDelete}
                  disabled={loading || selectedBooks.size === 0}
                  className="px-4 py-1.5 bg-red-600 text-white rounded hover:bg-red-700 disabled:bg-gray-400 text-sm font-medium"
                >
                  Delete ({selectedBooks.size})
                </button>
                <button
                  onClick={handleUpdate}
                  disabled={loading || selectedBooks.size === 0}
                  className="ml-auto px-4 py-1.5 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-gray-400 text-sm font-medium"
                >
                  {loading ? 'Updating...' : `Update (${selectedBooks.size})`}
                </button>
              </div>
            </div>

            <button
              onClick={cyclePNoFilter}
              className={`${getPNoFilterLabel().color} text-white px-3 py-2 rounded-md font-medium hover:opacity-90 transition-all shadow-sm text-sm whitespace-nowrap`}
              title="Click to cycle: All → Pending → Clear"
            >
              <span>{getPNoFilterLabel().icon}</span>
              <span className="ml-1">{getPNoFilterLabel().text}</span>
              <span className="ml-1 text-xs opacity-75">↻</span>
            </button>
            <button
              onClick={handleGeneratePDF}
              disabled={pdfLoading || sortedForGroup.length === 0}
              className="px-3 py-2 rounded-md font-medium shadow-sm text-sm whitespace-nowrap bg-red-600 text-white hover:bg-red-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
              title="Same PDF as Excel form: Pending : Surat City, book-wise totals"
            >
              {pdfLoading ? 'Generating…' : 'Generate PDF'}
            </button>
            <div className="flex items-center gap-2 border-l border-gray-300 pl-3">
              <button
                onClick={() => {
                  loadGroupSectors();
                  setPendingClearFormOpen(true);
                }}
                className="px-3 py-2 rounded-md font-medium shadow-sm text-sm whitespace-nowrap bg-violet-600 text-white hover:bg-violet-700"
                title="Open Pending Clear PDF form"
              >
                Pending Clear PDFs
              </button>
            </div>

            {pendingClearFormOpen && (
              <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={() => setPendingClearFormOpen(false)}>
                <div
                  className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md border border-gray-200"
                  onClick={(e) => e.stopPropagation()}
                >
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="text-lg font-semibold text-gray-800">Pending Clear PDF</h3>
                    <button
                      type="button"
                      onClick={() => setPendingClearFormOpen(false)}
                      className="text-gray-500 hover:text-gray-700 text-xl leading-none"
                    >
                      ×
                    </button>
                  </div>
                  <div className="space-y-4">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">P No.</label>
                      <input
                        type="text"
                        value={pendingClearPNos}
                        onChange={(e) => setPendingClearPNos(e.target.value)}
                        placeholder="e.g. 44,45,51"
                        className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">Group</label>
                      <select
                        value={selectedGroup}
                        onChange={(e) => setSelectedGroup(e.target.value)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm bg-white"
                      >
                        {groupSectors.map((g) => (
                          <option key={g} value={g}>
                            {g}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-3 pt-2">
                      <div className="flex flex-col gap-1 text-xs text-gray-700">
                        <span className="font-medium">Include sections:</span>
                        <label className="inline-flex items-center gap-2">
                          <input
                            type="checkbox"
                            checked={includeSectorSummary}
                            onChange={(e) => setIncludeSectorSummary(e.target.checked)}
                            className="w-4 h-4"
                          />
                          <span>Summary : Sector (2 per page layout)</span>
                        </label>
                        <label className="inline-flex items-center gap-2">
                          <input
                            type="checkbox"
                            checked={includeSectorCitySummary}
                            onChange={(e) => setIncludeSectorCitySummary(e.target.checked)}
                            className="w-4 h-4"
                          />
                          <span>Summary : Sector - City (continuous layout)</span>
                        </label>
                      </div>

                      <div className="flex gap-2">
                      <button
                        onClick={handleGeneratePendingClearSectorPDF}
                        disabled={pendingClearPdfLoading || pendingClearVillagePdfLoading || !pendingClearPNos.trim()}
                        className="flex-1 px-3 py-2 rounded-md font-medium shadow-sm text-sm bg-violet-600 text-white hover:bg-violet-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
                        title="Sector-wise Pending Clear report"
                      >
                        {pendingClearPdfLoading ? 'Generating…' : 'Pending Group'}
                      </button>
                      <button
                        onClick={handleGeneratePendingClearVillagePDF}
                        disabled={pendingClearPdfLoading || pendingClearVillagePdfLoading || !pendingClearPNos.trim()}
                        className="flex-1 px-3 py-2 rounded-md font-medium shadow-sm text-sm bg-indigo-600 text-white hover:bg-indigo-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
                        title="Village-wise Pending Clear report"
                      >
                        {pendingClearVillagePdfLoading ? 'Generating…' : 'Pending Village'}
                      </button>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {message && (
        <div
          className={`mb-6 p-4 rounded-md ${
            message.includes('Error')
              ? 'bg-red-50 text-red-700 border border-red-200'
              : 'bg-green-50 text-green-700 border border-green-200'
          }`}
        >
          {message}
        </div>
      )}


      <div className="bg-white rounded-lg shadow-md border border-gray-200">
        <div className="p-4 border-b border-gray-200">
          <div className="flex justify-between items-center mb-2">
            <div>
              <h2 className="text-xl font-semibold">
                Not In Stock Books 
                <span className="text-blue-600"> ({filteredBooks.length}</span>
                {filteredBooks.length !== books.length && (
                  <span className="text-gray-500 text-sm"> of {books.length}</span>
                )}
                <span className="text-blue-600">)</span>
              </h2>
              <div className="flex gap-2 mt-1 text-xs">
                {pNoFilter === 'blank' && (
                  <span className="text-orange-600 bg-orange-50 px-2 py-1 rounded">
                    📋 Pending Only
                  </span>
                )}
                {pNoFilter === 'assigned' && (
                  <span className="text-green-600 bg-green-50 px-2 py-1 rounded">
                    ✅ Clear Only
                  </span>
                )}
                {(searchLanguage || searchBookName || searchPeriod || searchPNo) && (
                  <span className="text-blue-600 bg-blue-50 px-2 py-1 rounded">
                    🔍 Filter Active
                  </span>
                )}
              </div>
            </div>
            <button
              onClick={handleSelectAll}
              className="px-4 py-2 text-sm bg-gray-200 hover:bg-gray-300 rounded-md transition-colors"
            >
              {filteredBooks.every((book) => selectedBooks.has(String(books.findIndex(b => b.importid === book.importid))))
                ? 'Deselect All'
                : 'Select All Visible'}
            </button>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full border-collapse">
            <thead className="bg-gray-50 sticky top-0">
              <tr>
                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">Select</th>
                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">Language</th>
                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">Book Name</th>
                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">Period</th>
                <th className="px-4 py-3 text-right text-sm font-medium text-gray-700">Total Qty</th>
                <th className="px-4 py-3 text-center text-sm font-medium text-gray-700">P No.</th>
                <th className="px-4 py-3 text-center text-sm font-medium text-gray-700">Updated Date</th>
              </tr>
            </thead>
            <tbody>
              {bookGroups.map((group, groupIndex) => (
                <Fragment key={group.book_id}>
                  {groupIndex > 0 && (
                    <tr>
                      <td colSpan={7} className="h-4 bg-gray-200 py-0" />
                    </tr>
                  )}
                  {group.books.map((book) => {
                    const originalIndex = books.findIndex(b => b.importid === book.importid);
                    const rowColor = bookGroupColors[groupIndex % bookGroupColors.length];
                    return (
                      <tr
                        key={book.importid}
                        role="button"
                        tabIndex={0}
                        onClick={() => handleSelectBook(originalIndex)}
                        onKeyDown={(e) => e.key === 'Enter' && handleSelectBook(originalIndex)}
                        className={`border-t border-gray-200 hover:opacity-90 cursor-pointer select-none ${rowColor} ${selectedBooks.has(String(originalIndex)) ? 'ring-2 ring-inset ring-blue-500' : ''}`}
                      >
                        <td className="px-4 py-3" onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            checked={selectedBooks.has(String(originalIndex))}
                            onChange={() => handleSelectBook(originalIndex)}
                            className="w-4 h-4 cursor-pointer"
                          />
                        </td>
                        <td className="px-4 py-3 text-sm">{book.language}</td>
                        <td className="px-4 py-3 text-sm">{book.book_name}</td>
                        <td className="px-4 py-3 text-sm text-gray-600">
                          {formatPeriod(book.year, book.month, book.half)}
                        </td>
                        <td className="px-4 py-3 text-sm text-right font-medium">{book.qty}</td>
                        <td className="px-4 py-3 text-sm text-center">
                          <span className={book.stock_number ? 'px-2 py-1 bg-blue-100 text-blue-800 rounded' : 'text-gray-400'}>
                            {book.stock_number || '-'}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-sm text-center text-gray-600 whitespace-nowrap">
                          {formatStockDate(book.stock_date)}
                        </td>
                      </tr>
                    );
                  })}
                </Fragment>
              ))}
            </tbody>
          </table>
        </div>

        {loading && filteredBooks.length === 0 && (
          <div className="p-8 text-center">
            <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600 mb-2"></div>
            <p className="text-gray-500">Loading books...</p>
          </div>
        )}

        {filteredBooks.length === 0 && !loading && books.length > 0 && (
          <div className="p-8 text-center text-gray-500">
            <p className="text-lg mb-2">🔍 No books match the current P No. filter</p>
            <p className="text-sm">
              {pNoFilter === 'blank'
                ? 'All books have P No. assigned. Try "All Books" filter.'
                : 'No books have P No. assigned yet. Try "P No. Blank" filter.'}
            </p>
          </div>
        )}

        {books.length === 0 && !loading && (
          <div className="p-8 text-center text-gray-500">
            <p className="text-lg mb-2">📦 No "Not In Stock" books found</p>
            <p className="text-sm">
              {searchLanguage || searchBookName || searchPeriod || searchPNo
                ? 'Try adjusting the filter or clear filters to see all books.'
                : 'No books are currently marked as "Not In Stock" in the database.'}
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
