'use client';

import { useState } from 'react';

interface BackupResponse {
  schema: string;
  createdAt: string;
  tables: Record<string, any[]>;
}

export default function BackupPage() {
  const [backupLoading, setBackupLoading] = useState(false);
  const [restoreLoading, setRestoreLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const handleDownloadBackup = async () => {
    setBackupLoading(true);
    setMessage('');

    try {
      const res = await fetch('/api/backup');
      const data: BackupResponse | { error?: string } = await res.json();

      if (!res.ok || (data as any).error) {
        throw new Error((data as any).error || res.statusText);
      }

      const json = JSON.stringify(data, null, 2);
      const blob = new Blob([json], { type: 'application/json' });
      const url = URL.createObjectURL(blob);

      const createdAt =
        'createdAt' in data
          ? new Date(data.createdAt).toISOString().slice(0, 19).replace(/[-:T]/g, '')
          : new Date().toISOString().slice(0, 19).replace(/[-:T]/g, '');

      const a = document.createElement('a');
      a.href = url;
      a.download = `vimars_backup_${createdAt}.json`;
      a.click();
      URL.revokeObjectURL(url);

      setMessage('Backup JSON file downloaded successfully.');
    } catch (error) {
      setMessage('Error while creating backup: ' + (error as Error).message);
    } finally {
      setBackupLoading(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    setSelectedFile(files && files.length > 0 ? files[0] : null);
  };

  const handleRestore = async () => {
    if (!selectedFile) {
      setMessage('Please select a backup JSON file to restore.');
      return;
    }

    const confirmRestore = window.confirm(
      'This will clear existing data in the backup tables (TRUNCATE + RESTART IDENTITY) and restore from selected file. Continue?'
    );
    if (!confirmRestore) return;

    setRestoreLoading(true);
    setMessage('');

    try {
      const text = await selectedFile.text();
      let json: unknown;
      try {
        json = JSON.parse(text);
      } catch (e) {
        throw new Error('Selected file is not valid JSON: ' + (e as Error).message);
      }

      const res = await fetch('/api/backup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(json),
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || res.statusText);
      }

      setMessage(data.message || 'Restore completed successfully.');
    } catch (error) {
      setMessage('Error while restoring backup: ' + (error as Error).message);
    } finally {
      setRestoreLoading(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto">
      <h1 className="text-2xl font-bold mb-4">Full Database Backup &amp; Restore</h1>

      {message && (
        <div
          className={`mb-4 p-3 rounded text-sm border ${
            message.toLowerCase().includes('error')
              ? 'bg-red-50 text-red-700 border-red-200'
              : 'bg-green-50 text-green-700 border-green-200'
          }`}
        >
          {message}
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="bg-white rounded-lg shadow-md border border-gray-200 p-5">
          <h2 className="text-lg font-semibold mb-2 text-blue-700">Manual Backup</h2>
          <p className="text-sm text-gray-600 mb-4">
            Complete PostgreSQL data ka JSON backup file download kare. Is file ko safe jagah par
            rakh sakte hain (external drive / cloud).
          </p>
          <button
            type="button"
            onClick={handleDownloadBackup}
            disabled={backupLoading}
            className="px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
          >
            {backupLoading ? 'Creating backup…' : 'Download Full Backup (.json)'}
          </button>
        </div>

        <div className="bg-white rounded-lg shadow-md border border-gray-200 p-5">
          <h2 className="text-lg font-semibold mb-2 text-green-700">Manual Restore</h2>
          <p className="text-sm text-gray-600 mb-3">
            Pehle liya hua JSON backup file select kare, phir Restore button dabaye. Selected
            tables pehle TRUNCATE honge, phir file se data load hoga.
          </p>
          <div className="space-y-3">
            <div>
              <input
                type="file"
                accept="application/json,.json"
                onChange={handleFileChange}
                className="block w-full text-sm text-gray-700
                  file:mr-3 file:py-1.5 file:px-3
                  file:rounded file:border-0
                  file:text-sm file:font-medium
                  file:bg-green-50 file:text-green-700
                  hover:file:bg-green-100"
              />
              {selectedFile && (
                <p className="mt-1 text-xs text-gray-500">
                  Selected: <span className="font-medium">{selectedFile.name}</span>
                </p>
              )}
            </div>

            <button
              type="button"
              onClick={handleRestore}
              disabled={restoreLoading || !selectedFile}
              className="px-4 py-2 bg-green-600 text-white text-sm font-medium rounded hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              {restoreLoading ? 'Restoring…' : 'Restore from Backup'}
            </button>

            <p className="text-xs text-red-600 mt-1">
              Warning: Production database par restore karne se pehle hamesha latest manual backup
              jarur le.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

