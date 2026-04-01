# VIMARS Next.js Application

A comprehensive web-based management system for VIMARS book data import, stock management, and report generation.

## Features

### 1. Excel Import
- Import multiple Excel files to PostgreSQL database
- Support for VIMARSH and PARISHKAAR apps
- Automatic period extraction from Excel files
- Bulk import with transaction safety

### 2. Stock Update
- Manage P No. for "Not In Stock" books
- Filter by year, month, and half period
- Bulk update with multi-select
- Clear P No. functionality (ESC key shortcut)
- Full month period support (1 & 2)

### 3. Tatvagnan Management
- Import Tatvagnan data from Excel
- Generate Summary PDF reports (landscape)
- Generate Group-Wise PDF reports (portrait, 4 groups per page)
- Automatic natural sorting (E1, E2... E9, E10)
- Full language name display

### 4. PDF Reports
- Summary reports with language-wise totals
- Group-wise detailed reports
- Professional formatting with tables and borders
- Download as PDF directly from browser

## Tech Stack

- **Frontend**: Next.js 15, React 19, TypeScript, Tailwind CSS
- **Backend**: Next.js API Routes
- **Database**: PostgreSQL with pg connection pooling
- **Excel Processing**: xlsx library
- **PDF Generation**: jspdf with jspdf-autotable
- **Styling**: Tailwind CSS

## Prerequisites

- Node.js 18+ installed
- PostgreSQL database running
- Database schema: `vimars`
- Required tables: `books`, `import_history`, `tatvagnan_data`

## Installation

1. Install dependencies:
```bash
npm install
```

2. Configure environment variables:
Create `.env.local` file with your database credentials:
```env
DB_HOST=localhost
DB_PORT=5432
DB_NAME=vimarshbooks
DB_USER=postgres
DB_PASSWORD=your_password
DB_SCHEMA=vimars
JWT_SECRET=your-secret-key
```

3. **Setup Database** (First time only):
```bash
npm run setup-db
```
This will create all required tables and indexes matching the Excel VBA schema.

4. Run development server:
```bash
npm run dev
```

5. Open [http://localhost:7001](http://localhost:7001)

## Database Schema

This application uses the **exact same database schema** as the Excel VBA code.

### books
- `book_id`: Primary key
- `book_name`: Book title
- `language`: Language code (GJ, HN, EN, etc.)
- `created_date`, `updated_date`: Timestamps
- `is_deleted`: Soft delete flag
- **Unique constraint**: (book_name, language)

### transactions
- `importid`: Primary key
- `history_id`: Foreign key to import_history
- `book_id`: Foreign key to books
- `year`, `month`, `half`: Period information
- `qty`: Quantity
- `avail`: Availability (0 = Not In Stock)
- `village`, `group_sector`, `contact_name`: Location details
- `app`: Application name (VIMARSH or PARISHKAAR)
- `stock_number`: P No. for stock management
- `stock_date`: Stock date
- `is_deleted`: Soft delete flag

### import_history
- `history_id`: Primary key
- `file_name`: Imported Excel file name
- `import_date`: Import timestamp
- `period`: Period string
- `app`: Application name
- `year`, `month`, `half`: Period components
- `total_book_qty`: Number of records imported
- `status`: Import status

### tatvagnan_data
- `id`: Primary key
- `pushp_no`: Pushp number
- `group`: Group name (E1, E2, etc.)
- `lang`: Language code
- `old_qty`, `new_qty`, `rmv_qty`: Quantities

## API Endpoints

### Import
- `POST /api/import` - Import Excel files to database

### Stock Update
- `GET /api/stock-update` - Get not-in-stock books by period
- `POST /api/stock-update` - Update P No. for selected books

### Tatvagnan
- `POST /api/tatvagnan/import` - Import Tatvagnan data
- `GET /api/tatvagnan/summary-pdf` - Generate summary PDF
- `GET /api/tatvagnan/groupwise-pdf` - Generate group-wise PDF

## Production Build

```bash
npm run build
npm start
```

## Key Features Implementation

### Period Matching
- Normalizes "1 & 2", "1&2", "1 and 2" for full month periods
- Accurate matching across different formats

### Natural Sorting
- Sorts groups naturally: E1, E2, E9, E10 (not E1, E10, E2, E9)
- Custom comparison function for alphanumeric strings

### Connection Pooling
- PostgreSQL connection pool with 20 max connections
- Automatic connection management and error handling

### Transaction Safety
- All import/update operations use transactions
- Automatic rollback on errors

### File Upload
- Supports multiple file uploads
- File size limit: 50MB
- Excel format validation

## Development

- Run linter: `npm run lint`
- Development mode: `npm run dev` (auto-reload on changes)

## License

Private - VIMARS Internal Use Only
