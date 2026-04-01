-- VIMARS Database Schema Initialization
-- This matches the schema created by the Excel VBA code

-- Create schema if not exists
CREATE SCHEMA IF NOT EXISTS vimars;

-- Books table
CREATE TABLE IF NOT EXISTS vimars.books (
  book_id SERIAL PRIMARY KEY,
  book_name TEXT NOT NULL,
  language TEXT NOT NULL,
  created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  created_by TEXT DEFAULT 'Excel Import',
  updated_by TEXT DEFAULT 'Excel Import',
  is_deleted BOOLEAN DEFAULT FALSE,
  UNIQUE(book_name, language)
);

-- Import history table
CREATE TABLE IF NOT EXISTS vimars.import_history (
  history_id SERIAL PRIMARY KEY,
  file_name TEXT NOT NULL,
  import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  period TEXT,
  app TEXT DEFAULT 'VIMARSH',
  year INTEGER,
  month INTEGER,
  half TEXT,
  total_book_qty INTEGER DEFAULT 0,
  status TEXT DEFAULT 'SUCCESS',
  created_by TEXT DEFAULT 'Excel Import',
  is_deleted BOOLEAN DEFAULT FALSE
);

-- Transactions table
CREATE TABLE IF NOT EXISTS vimars.transactions (
  importid SERIAL PRIMARY KEY,
  history_id INTEGER REFERENCES vimars.import_history(history_id) ON DELETE CASCADE,
  book_id INTEGER REFERENCES vimars.books(book_id),
  year INTEGER,
  month INTEGER,
  half TEXT,
  qty INTEGER,
  avail INTEGER,
  village TEXT,
  group_sector TEXT,
  contact_name TEXT,
  app TEXT DEFAULT 'VIMARSH',
  stock_number TEXT,
  stock_date DATE,
  created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  created_by TEXT DEFAULT 'Excel Import',
  updated_by TEXT DEFAULT 'Excel Import',
  is_deleted BOOLEAN DEFAULT FALSE
);

-- Tatvagnan data table (for group-wise reports)
CREATE TABLE IF NOT EXISTS vimars.tatvagnan_data (
  id SERIAL PRIMARY KEY,
  pushp_no INTEGER NOT NULL,
  "group" TEXT NOT NULL,
  lang TEXT NOT NULL,
  old_qty INTEGER DEFAULT 0,
  new_qty INTEGER DEFAULT 0,
  rmv_qty INTEGER DEFAULT 0,
  created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Indexes for better performance
CREATE INDEX IF NOT EXISTS idx_books_language ON vimars.books(language);
CREATE INDEX IF NOT EXISTS idx_books_is_deleted ON vimars.books(is_deleted);

CREATE INDEX IF NOT EXISTS idx_transactions_book_id ON vimars.transactions(book_id);
CREATE INDEX IF NOT EXISTS idx_transactions_period ON vimars.transactions(year, month, half);
CREATE INDEX IF NOT EXISTS idx_transactions_avail ON vimars.transactions(avail);
CREATE INDEX IF NOT EXISTS idx_transactions_app ON vimars.transactions(app);
CREATE INDEX IF NOT EXISTS idx_transactions_is_deleted ON vimars.transactions(is_deleted);

CREATE INDEX IF NOT EXISTS idx_import_history_app ON vimars.import_history(app);
CREATE INDEX IF NOT EXISTS idx_import_history_period ON vimars.import_history(year, month, half);

CREATE INDEX IF NOT EXISTS idx_tatvagnan_pushp_no ON vimars.tatvagnan_data(pushp_no);
CREATE INDEX IF NOT EXISTS idx_tatvagnan_group ON vimars.tatvagnan_data("group");

-- Comments for documentation
COMMENT ON TABLE vimars.books IS 'Master table of books with unique book_name and language combination';
COMMENT ON TABLE vimars.import_history IS 'Records of all Excel file imports';
COMMENT ON TABLE vimars.transactions IS 'Transaction records from Excel imports with period and availability info';
COMMENT ON TABLE vimars.tatvagnan_data IS 'Tatvagnan group-wise data for PDF reports';

COMMENT ON COLUMN vimars.transactions.avail IS 'Availability: 0 = Not In Stock, >0 = Available quantity';
COMMENT ON COLUMN vimars.transactions.stock_number IS 'P No. for stock management (set via Stock Update module)';
COMMENT ON COLUMN vimars.transactions.stock_date IS 'Stock date for management';
