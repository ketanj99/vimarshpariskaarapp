-- Add app column to transactions table if it doesn't exist
-- Run this if you get "column app does not exist" error

ALTER TABLE vimars.transactions 
ADD COLUMN IF NOT EXISTS app TEXT DEFAULT 'VIMARSH';

-- Add stock_number and stock_date columns if they don't exist
ALTER TABLE vimars.transactions 
ADD COLUMN IF NOT EXISTS stock_number TEXT;

ALTER TABLE vimars.transactions 
ADD COLUMN IF NOT EXISTS stock_date DATE;

-- Update import_history to have year, month, half columns
ALTER TABLE vimars.import_history 
ADD COLUMN IF NOT EXISTS year INTEGER;

ALTER TABLE vimars.import_history 
ADD COLUMN IF NOT EXISTS month INTEGER;

ALTER TABLE vimars.import_history 
ADD COLUMN IF NOT EXISTS half TEXT;

COMMENT ON COLUMN vimars.transactions.app IS 'Application name: VIMARSH or PARISHKAAR';
COMMENT ON COLUMN vimars.transactions.stock_number IS 'P No. for stock management';
COMMENT ON COLUMN vimars.transactions.stock_date IS 'Date for stock management';
