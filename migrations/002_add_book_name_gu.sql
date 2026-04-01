-- Optional Gujarati book name for PDF display
ALTER TABLE vimars.books ADD COLUMN IF NOT EXISTS book_name_gu TEXT;
COMMENT ON COLUMN vimars.books.book_name_gu IS 'Book name in Gujarati script for PDF; when set, used in place of book_name in reports';
