import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "VIMARS - Data Management System",
  description: "VIMARS Book Management and Reporting System",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">
        <nav className="bg-blue-600 text-white p-4 shadow-lg">
          <div className="container mx-auto flex justify-between items-center">
            <h1 className="text-2xl font-bold">VIMARS Management System</h1>
            <div className="space-x-4">
              <a href="/" className="hover:underline">Home</a>
              <a href="/import" className="hover:underline">Import</a>
              <a href="/stock-update" className="hover:underline">Stock Update</a>
              <a href="/tatvagnan" className="hover:underline">Tatvagnan</a>
            </div>
          </div>
        </nav>
        <main className="container mx-auto p-6">
          {children}
        </main>
      </body>
    </html>
  );
}
