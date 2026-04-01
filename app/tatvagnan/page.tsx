'use client';

import { useState } from 'react';

export default function TatvagnanPage() {
  const [file, setFile] = useState<File | null>(null);
  const [pushpNo, setPushpNo] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [importResult, setImportResult] = useState<any>(null);

  const handleImport = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!file) {
      setMessage('Please select a file');
      return;
    }

    setLoading(true);
    setMessage('');
    setImportResult(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('/api/tatvagnan/import', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (response.ok) {
        setMessage(data.message || 'Import successful');
        setImportResult(data);
        setFile(null);
      } else {
        setMessage('Error: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const handleGenerateSummaryPdf = async () => {
    if (!pushpNo) {
      setMessage('Please enter Pushp No.');
      return;
    }

    setLoading(true);
    setMessage('Generating Summary PDF...');

    try {
      const response = await fetch(`/api/tatvagnan/summary-pdf?pushpNo=${pushpNo}`);

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Tatvagnan_Summary_${pushpNo}.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        setMessage('Summary PDF generated successfully');
      } else {
        const data = await response.json();
        setMessage('Error: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const handleGenerateGroupWisePdf = async () => {
    if (!pushpNo) {
      setMessage('Please enter Pushp No.');
      return;
    }

    setLoading(true);
    setMessage('Generating Group-Wise PDF...');

    try {
      const response = await fetch(`/api/tatvagnan/groupwise-pdf?pushpNo=${pushpNo}`);

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Tatvagnan_GroupWise_${pushpNo}.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        setMessage('Group-Wise PDF generated successfully');
      } else {
        const data = await response.json();
        setMessage('Error: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  const handleGenerateGroupVillagePdf = async () => {
    if (!pushpNo) {
      setMessage('Please enter Pushp No.');
      return;
    }

    setLoading(true);
    setMessage('Generating Group + Village PDF...');

    try {
      const response = await fetch(`/api/tatvagnan/group-village-pdf?pushpNo=${pushpNo}`);

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Tatvagnan_GroupVillage_${pushpNo}.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        setMessage('Group + Village PDF generated successfully');
      } else {
        const data = await response.json();
        setMessage('Error: ' + data.error);
      }
    } catch (error) {
      setMessage('Error: ' + (error as Error).message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto">
      <h1 className="text-3xl font-bold mb-6">Tatvagnan Management</h1>

      {/* Import Section */}
      <div className="bg-white rounded-lg shadow-md p-6 border border-gray-200 mb-6">
        <h2 className="text-xl font-semibold mb-4">Import Tatvagnan Data</h2>
        <form onSubmit={handleImport}>
          <div className="mb-4">
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Select Excel File
            </label>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => setFile(e.target.files?.[0] || null)}
              className="w-full px-4 py-2 border border-gray-300 rounded-md"
            />
          </div>
          <button
            type="submit"
            disabled={loading}
            className="w-full bg-purple-600 text-white py-3 px-6 rounded-md hover:bg-purple-700 disabled:bg-gray-400 font-medium"
          >
            {loading ? 'Importing...' : 'Import Tatvagnan Data'}
          </button>
        </form>

        {importResult && (
          <div className="mt-4 p-4 bg-green-50 border border-green-200 rounded-md">
            <p className="text-green-700">
              <strong>Pushp No:</strong> {importResult.pushpNo}<br />
              <strong>Records Imported:</strong> {importResult.recordsImported}
            </p>
          </div>
        )}
      </div>

      {/* PDF Generation Section */}
      <div className="bg-white rounded-lg shadow-md p-6 border border-gray-200">
        <h2 className="text-xl font-semibold mb-4">Generate PDF Reports</h2>
        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Pushp No.
          </label>
          <input
            type="number"
            value={pushpNo}
            onChange={(e) => setPushpNo(e.target.value)}
            placeholder="Enter Pushp No."
            className="w-full px-4 py-2 border border-gray-300 rounded-md"
          />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <button
            onClick={handleGenerateSummaryPdf}
            disabled={loading || !pushpNo}
            className="bg-blue-600 text-white py-3 px-6 rounded-md hover:bg-blue-700 disabled:bg-gray-400 font-medium"
          >
            Generate Summary PDF
          </button>
          <button
            onClick={handleGenerateGroupWisePdf}
            disabled={loading || !pushpNo}
            className="bg-green-600 text-white py-3 px-6 rounded-md hover:bg-green-700 disabled:bg-gray-400 font-medium"
          >
            Generate Group-Wise PDF
          </button>
          <button
            onClick={handleGenerateGroupVillagePdf}
            disabled={loading || !pushpNo}
            className="bg-amber-600 text-white py-3 px-6 rounded-md hover:bg-amber-700 disabled:bg-gray-400 font-medium"
          >
            Group + Village PDF
          </button>
        </div>
      </div>

      {message && (
        <div
          className={`mt-6 p-4 rounded-md ${
            message.includes('Error')
              ? 'bg-red-50 text-red-700 border border-red-200'
              : 'bg-green-50 text-green-700 border border-green-200'
          }`}
        >
          {message}
        </div>
      )}

      <div className="mt-6 p-4 bg-purple-50 rounded-lg border border-purple-200">
        <h3 className="font-semibold text-purple-800 mb-2">Instructions:</h3>
        <ul className="text-sm text-purple-700 space-y-1">
          <li>1. Import Tatvagnan data from Excel file (contains language-wise group data)</li>
          <li>2. Enter Pushp No. to generate reports</li>
          <li>3. Summary PDF: All groups with language totals</li>
          <li>4. Group-Wise PDF: 4 groups per page with detailed language breakdown</li>
          <li>5. Group + Village PDF: From tatvagnan_simple_data — add/remove by group and village with names</li>
        </ul>
      </div>
    </div>
  );
}
