# VIMARS Books PostgreSQL Database Setup Script
# PowerShell version with enhanced error handling

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "VIMARS Books PostgreSQL Database Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if PostgreSQL is installed
Write-Host "Checking PostgreSQL installation..." -ForegroundColor Yellow
try {
    $psqlVersion = & psql --version 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "PostgreSQL found. Version:" -ForegroundColor Green
        Write-Host $psqlVersion -ForegroundColor White
    } else {
        throw "PostgreSQL not found"
    }
} catch {
    Write-Host "ERROR: PostgreSQL is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install PostgreSQL from: https://www.postgresql.org/download/" -ForegroundColor Yellow
    Write-Host "Make sure to add PostgreSQL bin directory to your PATH" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Check if PostgreSQL service is running
Write-Host "Checking PostgreSQL service..." -ForegroundColor Yellow
try {
    $service = Get-Service -Name "postgresql*" -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        Write-Host "PostgreSQL service is running." -ForegroundColor Green
    } else {
        Write-Host "WARNING: PostgreSQL service may not be running" -ForegroundColor Yellow
        Write-Host "Please start PostgreSQL service manually" -ForegroundColor Yellow
    }
} catch {
    Write-Host "WARNING: Could not check PostgreSQL service status" -ForegroundColor Yellow
}

Write-Host ""

# Test database connection
Write-Host "Testing database connection..." -ForegroundColor Yellow
try {
    $testConnection = & psql -U postgres -c "SELECT version();" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Database connection successful." -ForegroundColor Green
    } else {
        throw "Connection failed"
    }
} catch {
    Write-Host "ERROR: Cannot connect to PostgreSQL" -ForegroundColor Red
    Write-Host "Please check:" -ForegroundColor Yellow
    Write-Host "1. PostgreSQL service is running" -ForegroundColor Yellow
    Write-Host "2. Password for postgres user is correct" -ForegroundColor Yellow
    Write-Host "3. PostgreSQL is accepting connections" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Create database
Write-Host "Creating VIMARS Books database..." -ForegroundColor Yellow
try {
    $createDB = & psql -U postgres -c "CREATE DATABASE vimarshbooks;" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Database created successfully." -ForegroundColor Green
    } else {
        Write-Host "Database may already exist or connection failed" -ForegroundColor Yellow
        Write-Host "Continuing with existing database..." -ForegroundColor Yellow
    }
} catch {
    Write-Host "WARNING: Could not create database" -ForegroundColor Yellow
}

Write-Host ""

# Run the main database creation script
Write-Host "Running database creation script..." -ForegroundColor Yellow
try {
    $scriptPath = Join-Path $PSScriptRoot "create_vimarshbooks_db.sql"
    if (Test-Path $scriptPath) {
        $createStructure = & psql -U postgres -d vimarshbooks -f $scriptPath
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Database structure created successfully!" -ForegroundColor Green
        } else {
            throw "Failed to create database structure"
        }
    } else {
        throw "Database creation script not found: $scriptPath"
    }
} catch {
    Write-Host "ERROR: Failed to create database structure" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Ask if user wants to insert sample data
$insertSample = Read-Host "Do you want to insert sample data? (y/n)"
if ($insertSample -eq "y" -or $insertSample -eq "Y") {
    Write-Host ""
    Write-Host "Inserting sample data..." -ForegroundColor Yellow
    try {
        $sampleDataPath = Join-Path $PSScriptRoot "sample_data.sql"
        if (Test-Path $sampleDataPath) {
            $insertData = & psql -U postgres -d vimarshbooks -f $sampleDataPath
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Sample data inserted successfully!" -ForegroundColor Green
            } else {
                throw "Failed to insert sample data"
            }
        } else {
            throw "Sample data script not found: $sampleDataPath"
        }
    } catch {
        Write-Host "ERROR: Failed to insert sample data" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Database Details:" -ForegroundColor White
Write-Host "- Database Name: vimarshbooks" -ForegroundColor Yellow
Write-Host "- Schema: vimars" -ForegroundColor Yellow
Write-Host "- Tables: books, transactions, users, audit_log, reports, settings" -ForegroundColor Yellow
Write-Host "- Views: book_summary, monthly_summary, village_summary" -ForegroundColor Yellow
Write-Host ""
Write-Host "Default Users:" -ForegroundColor White
Write-Host "- admin/admin123 (Administrator)" -ForegroundColor Yellow
Write-Host "- manager/manager123 (Manager)" -ForegroundColor Yellow
Write-Host "- user1/user123 (Regular User)" -ForegroundColor Yellow
Write-Host "- user2/user123 (Regular User)" -ForegroundColor Yellow
Write-Host "- reports/reports123 (Reports User)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Connection Details:" -ForegroundColor White
Write-Host "- Host: localhost" -ForegroundColor Yellow
Write-Host "- Port: 5432" -ForegroundColor Yellow
Write-Host "- Database: vimarshbooks" -ForegroundColor Yellow
Write-Host "- Username: postgres (or create specific users)" -ForegroundColor Yellow
Write-Host ""
Write-Host "IMPORTANT: Change default passwords in production!" -ForegroundColor Red
Write-Host ""

# Test the database connection
Write-Host "Testing final database connection..." -ForegroundColor Yellow
try {
    $testFinal = & psql -U postgres -d vimarshbooks -c "SELECT COUNT(*) as book_count FROM vimars.books;" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Database connection test successful!" -ForegroundColor Green
    } else {
        Write-Host "WARNING: Could not test final database connection" -ForegroundColor Yellow
    }
} catch {
    Write-Host "WARNING: Could not test final database connection" -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Press Enter to exit"
