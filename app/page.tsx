export default function Home() {
  return (
    <div className="max-w-4xl mx-auto">
      <h1 className="text-3xl font-bold mb-6">Welcome to VIMARS Management System</h1>
      
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <a href="/import" className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border border-gray-200">
          <h2 className="text-xl font-semibold mb-2 text-blue-600">📥 Import Data</h2>
          <p className="text-gray-600">Import Excel files to PostgreSQL database. Supports multiple file uploads.</p>
        </a>

        <a href="/backup" className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border border-gray-200">
          <h2 className="text-xl font-semibold mb-2 text-red-600">💾 Backup &amp; Restore</h2>
          <p className="text-gray-600">Pure database ka full JSON backup download karein aur kabhi bhi restore karein.</p>
        </a>

        <a href="/stock-update" className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border border-gray-200">
          <h2 className="text-xl font-semibold mb-2 text-green-600">📦 Stock Update</h2>
          <p className="text-gray-600">Update P No. for Not In Stock books. Clear or assign P numbers easily.</p>
        </a>

        <a href="/tatvagnan" className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border border-gray-200">
          <h2 className="text-xl font-semibold mb-2 text-purple-600">📊 Tatvagnan</h2>
          <p className="text-gray-600">Import Tatvagnan data and generate group-wise reports in PDF format.</p>
        </a>

        <a href="/reports" className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border border-gray-200">
          <h2 className="text-xl font-semibold mb-2 text-orange-600">📄 Reports</h2>
          <p className="text-gray-600">Generate various reports including Sector Summary, Village Details, and more.</p>
        </a>

      </div>

      <div className="mt-8 p-4 bg-blue-50 rounded-lg border border-blue-200">
        <h3 className="font-semibold text-blue-800 mb-2">System Information</h3>
        <ul className="text-sm text-blue-700 space-y-1">
          <li>• Database: PostgreSQL (vimarshbooks)</li>
          <li>• Schema: vimars</li>
          <li>• Features: Import, Stock Management, Tatvagnan, PDF Reports</li>
        </ul>
      </div>
    </div>
  );
}
