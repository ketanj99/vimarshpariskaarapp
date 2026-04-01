-- VIMARS Database Backup
-- Created: 09-01-2026 11:36:30
-- Schema: vimars

BEGIN;

-- Delete existing data in correct order (considering foreign keys)
-- Delete from child tables first (transactions), then parent tables
DELETE FROM vimars.transactions WHERE is_deleted = FALSE OR is_deleted IS NULL;
DELETE FROM vimars.books WHERE is_deleted = FALSE OR is_deleted IS NULL;
DELETE FROM vimars.import_history WHERE is_deleted = FALSE OR is_deleted IS NULL;

-- Backup Books Table
