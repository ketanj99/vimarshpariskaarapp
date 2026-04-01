Option Explicit

' Custom type for column mappings
Private Type ColumnMappings
    bookTitleCol As Long
    languageCol As Long
    yearCol As Long
    monthCol As Long
    halfCol As Long
    qtyCol As Long
    availCol As Long
    villageCol As Long
    groupSectorCol As Long
    contactNameCol As Long
End Type

' PostgreSQL Import Module for VIMARS
' This module handles importing Excel data to PostgreSQL database

Private Const DB_HOST As String = "localhost"
Private Const DB_PORT As String = "5432"
Private Const DB_NAME As String = "vimarshbooks"
Private Const DB_USER As String = "postgres"
Private Const DB_PASSWORD As String = "Ketan@757399"
Private Const DB_SCHEMA As String = "vimars"

' Initialize System function
Public Sub InitializeSystem()
    On Error GoTo ErrorHandler
    
    ' Try to connect to PostgreSQL first
    Dim testConn As Object
    Set testConn = CreateObject("ADODB.Connection")
    
    Debug.Print "=== InitializeSystem Started ==="
    Debug.Print "Attempting PostgreSQL connection and database setup..."
    
    If ConnectToPostgreSQL(testConn) Then
        ' Connection successful - now create tables if they don't exist
        If CreatePostgreSQLTables(testConn) Then
            testConn.Close
            Set testConn = Nothing
            MsgBox "VIMARS PostgreSQL Import System Initialized Successfully!" & vbCrLf & vbCrLf & _
                   "✓ PostgreSQL connection established" & vbCrLf & _
                   "✓ Database '" & DB_NAME & "' is ready" & vbCrLf & _
                   "✓ All required tables created/verified" & vbCrLf & _
                   "✓ Schema '" & DB_SCHEMA & "' is ready" & vbCrLf & vbCrLf & _
                   "System is ready for Excel to PostgreSQL import.", vbInformation, "System Initialized"
        Else
            testConn.Close
            Set testConn = Nothing
            MsgBox "PostgreSQL connection successful but table creation failed." & vbCrLf & _
                   "Please check the Immediate Window (Ctrl+G) for details.", vbCritical, "Table Creation Failed"
        End If
    Else
        ' PostgreSQL not available - offer fallback
        Set testConn = Nothing
        Dim response As VbMsgBoxResult
        response = MsgBox("PostgreSQL connection or database creation failed." & vbCrLf & vbCrLf & _
                         "Possible reasons:" & vbCrLf & _
                         "• PostgreSQL server is not running" & vbCrLf & _
                         "• ODBC driver not installed" & vbCrLf & _
                         "• Incorrect credentials" & vbCrLf & _
                         "• Permission issues" & vbCrLf & vbCrLf & _
                         "Would you like to:" & vbCrLf & _
                         "• YES = Use Excel-based database (Recommended)" & vbCrLf & _
                         "• NO = Show PostgreSQL setup instructions", vbYesNo + vbQuestion, "Database Setup Failed")
        
        If response = vbYes Then
            ' Create Excel-based database
            CreateExcelBasedDatabase
        Else
            ' Show PostgreSQL setup instructions
            ShowPostgreSQLSetupInstructions
        End If
    End If
    
    Debug.Print "=== InitializeSystem Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in InitializeSystem ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    MsgBox "Error initializing system: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Check Immediate Window (Ctrl+G) for details", vbCritical, "Initialization Error"
End Sub

' Create Excel-based database as fallback
Private Sub CreateExcelBasedDatabase()
    On Error GoTo ErrorHandler
    
    Dim dataWb As Workbook
    Dim booksWs As Worksheet
    Dim transactionsWs As Worksheet
    Dim importHistoryWs As Worksheet
    
    ' Create new workbook
    Set dataWb = Workbooks.Add
    
    ' Create Books worksheet
    Set booksWs = dataWb.Sheets(1)
    booksWs.Name = "Books"
    
    ' Add headers to Books worksheet
    booksWs.Cells(1, 1).Value = "BookID"
    booksWs.Cells(1, 2).Value = "BookName"
    booksWs.Cells(1, 3).Value = "Language"
    booksWs.Cells(1, 4).Value = "Year"
    booksWs.Cells(1, 5).Value = "Month"
    booksWs.Cells(1, 6).Value = "Half"
    booksWs.Cells(1, 7).Value = "Quantity"
    booksWs.Cells(1, 8).Value = "Available"
    booksWs.Cells(1, 9).Value = "Village"
    booksWs.Cells(1, 10).Value = "GroupSector"
    booksWs.Cells(1, 11).Value = "ContactName"
    booksWs.Cells(1, 12).Value = "CreatedDate"
    
    ' Format headers
    booksWs.Range("A1:L1").Font.Bold = True
    booksWs.Range("A1:L1").Interior.Color = RGB(220, 220, 220)
    booksWs.Range("A1:L1").HorizontalAlignment = xlCenter
    
    ' Create Transactions worksheet
    Set transactionsWs = dataWb.Sheets.Add
    transactionsWs.Name = "Transactions"
    
    ' Add headers to Transactions worksheet
    transactionsWs.Cells(1, 1).Value = "TransactionID"
    transactionsWs.Cells(1, 2).Value = "BookID"
    transactionsWs.Cells(1, 3).Value = "TransactionDate"
    transactionsWs.Cells(1, 4).Value = "Amount"
    transactionsWs.Cells(1, 5).Value = "TransactionType"
    transactionsWs.Cells(1, 6).Value = "Description"
    transactionsWs.Cells(1, 7).Value = "CreatedDate"
    
    ' Format headers
    transactionsWs.Range("A1:G1").Font.Bold = True
    transactionsWs.Range("A1:G1").Interior.Color = RGB(220, 220, 220)
    transactionsWs.Range("A1:G1").HorizontalAlignment = xlCenter
    
    ' Create Import History worksheet
    Set importHistoryWs = dataWb.Sheets.Add
    importHistoryWs.Name = "ImportHistory"
    
    ' Add headers to Import History worksheet
    importHistoryWs.Cells(1, 1).Value = "HistoryID"
    importHistoryWs.Cells(1, 2).Value = "FileName"
    importHistoryWs.Cells(1, 3).Value = "ImportDate"
    importHistoryWs.Cells(1, 4).Value = "Period"
    importHistoryWs.Cells(1, 5).Value = "TotalQty"
    importHistoryWs.Cells(1, 6).Value = "Status"
    importHistoryWs.Cells(1, 7).Value = "CreatedDate"
    
    ' Format headers
    importHistoryWs.Range("A1:G1").Font.Bold = True
    importHistoryWs.Range("A1:G1").Interior.Color = RGB(220, 220, 220)
    importHistoryWs.Range("A1:G1").HorizontalAlignment = xlCenter
    
    ' Auto-fit columns
    booksWs.Columns.AutoFit
    transactionsWs.Columns.AutoFit
    importHistoryWs.Columns.AutoFit
    
    ' Save the data workbook
    Dim dataPath As String
    dataPath = ThisWorkbook.Path & "\VIMARS_Excel_Database.xlsx"
    dataWb.SaveAs dataPath
    
    ' Activate the data workbook
    dataWb.Activate
    
    MsgBox "Excel-based database created successfully!" & vbCrLf & vbCrLf & _
           "File: " & dataPath & vbCrLf & vbCrLf & _
           "Worksheets created:" & vbCrLf & _
           "• Books - for book inventory" & vbCrLf & _
           "• Transactions - for financial records" & vbCrLf & _
           "• ImportHistory - for import tracking" & vbCrLf & vbCrLf & _
           "You can now import Excel data to this database.", vbInformation, "Excel Database Created"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Excel database: " & Err.Description, vbCritical
End Sub

' Show PostgreSQL setup instructions
Private Sub ShowPostgreSQLSetupInstructions()
    On Error GoTo ErrorHandler
    
    Dim instructions As String
    instructions = "PostgreSQL Setup Instructions:" & vbCrLf & vbCrLf & _
                   "1. Install PostgreSQL Server:" & vbCrLf & _
                   "   - Download from: https://www.postgresql.org/download/" & vbCrLf & _
                   "   - Install with default settings" & vbCrLf & vbCrLf & _
                   "2. Install PostgreSQL ODBC Driver:" & vbCrLf & _
                   "   - Download from: https://odbc.postgresql.org/" & vbCrLf & _
                   "   - Choose 64-bit or 32-bit based on your Excel version" & vbCrLf & vbCrLf & _
                   "3. Create Database:" & vbCrLf & _
                   "   - Open pgAdmin (comes with PostgreSQL)" & vbCrLf & _
                   "   - Create database named 'vimarshbooks'" & vbCrLf & vbCrLf & _
                   "4. Alternative - Use PowerShell Script:" & vbCrLf & _
                   "   - Run setup_database.ps1 as Administrator" & vbCrLf & vbCrLf & _
                   "5. Restart Excel after installation" & vbCrLf & vbCrLf & _
                   "Current Settings:" & vbCrLf & _
                   "• Server: localhost" & vbCrLf & _
                   "• Port: 5432" & vbCrLf & _
                   "• Database: vimarshbooks" & vbCrLf & _
                   "• Username: postgres" & vbCrLf & _
                   "• Password: Ketan@757399"
    
    MsgBox instructions, vbInformation, "PostgreSQL Setup Guide"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing instructions: " & Err.Description, vbCritical
End Sub

' Main import function - called from menu
Public Sub ImportToPostgreSQL()
    On Error GoTo ErrorHandler
    
    Dim filePaths As String
    Dim fileCount As Integer
    Dim response As VbMsgBoxResult
    Dim appSelection As String
    Dim appName As String
    
    ' Optimize Excel for import process
    OptimizeExcelForImport
    
    ' Ask for APP selection first (before file selection)
    appSelection = InputBox("Select APP for import:" & vbCrLf & vbCrLf & _
                           "1 = VIMARSH" & vbCrLf & _
                           "2 = PARISHKAAR" & vbCrLf & vbCrLf & _
                           "Enter 1 or 2:", "APP Selection", "1")
    
    ' Validate and set APP name
    If Trim(appSelection) = "1" Then
        appName = "VIMARSH"
    ElseIf Trim(appSelection) = "2" Then
        appName = "PARISHKAAR"
    Else
        RestoreExcelSettings
        MsgBox "Invalid selection. Please enter 1 for VIMARSH or 2 for PARISHKAAR.", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    Debug.Print "APP Selected: " & appName
    
    ' Show file selection dialog
    response = MsgBox("Do you want to import multiple files?" & vbCrLf & vbCrLf & _
                     "APP: " & appName & vbCrLf & vbCrLf & _
                     "Note: Excel will be optimized during import to prevent blank screen issues.", vbYesNoCancel + vbQuestion, "Import Options")
    
    Select Case response
        Case vbYes
            ' Multiple file import
            filePaths = SelectMultipleFiles()
            If filePaths = "" Then
                RestoreExcelSettings
                MsgBox "No files selected. Import cancelled.", vbInformation
                Exit Sub
            End If
            fileCount = CountFiles(filePaths)
            ImportMultipleFilesToPostgreSQL filePaths, fileCount, appName
            
        Case vbNo
            ' Single file import
            filePaths = SelectSingleFile()
            If filePaths = "" Then
                RestoreExcelSettings
                MsgBox "No file selected. Import cancelled.", vbInformation
                Exit Sub
            End If
            ImportSingleFileToPostgreSQL filePaths, appName
            
        Case vbCancel
            RestoreExcelSettings
            MsgBox "Import cancelled by user.", vbInformation
            Exit Sub
    End Select
    
    ' Restore Excel settings
    RestoreExcelSettings
    Exit Sub
    
ErrorHandler:
    ' Restore Excel settings even if error occurs
    RestoreExcelSettings
    MsgBox "Error in ImportToPostgreSQL: " & Err.Description, vbCritical
End Sub

' Select multiple files using file dialog
Private Function SelectMultipleFiles() As String
    On Error GoTo ErrorHandler
    
    Dim fd As Object
    Dim filePaths As String
    Dim i As Integer
    Dim selectedFile As Variant
    
    ' Create file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Configure dialog for multiple file selection
    With fd
        .Title = "Select Excel Files to Import"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
    End With
    
    ' Try to enable multi-select
    On Error Resume Next
    fd.AllowMultiSelect = True
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "File dialog configured"
    
    ' Show dialog
    If fd.Show = True Then
        Debug.Print "User selected files in dialog"
        
        ' Build file paths string using For Each loop
        filePaths = ""
        i = 0
        For Each selectedFile In fd.SelectedItems
            i = i + 1
            If filePaths <> "" Then
                filePaths = filePaths & "|"
            End If
            filePaths = filePaths & CStr(selectedFile)
            Debug.Print "Selected file " & i & ": " & CStr(selectedFile)
        Next selectedFile
        
        Debug.Print "Total files selected: " & i
        Debug.Print "File paths string: " & filePaths
    Else
        Debug.Print "User cancelled file selection"
        filePaths = ""
    End If
    
    SelectMultipleFiles = filePaths
    Exit Function
    
ErrorHandler:
    SelectMultipleFiles = ""
End Function

' Select single file using file dialog
Private Function SelectSingleFile() As String
    On Error GoTo ErrorHandler
    
    Dim fd As Object
    Dim selectedItems As Variant
    
    ' Create file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Configure dialog for single file selection
    With fd
        .Title = "Select Excel File to Import"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = False
    End With
    
    ' Show dialog
    If fd.Show = True Then
        selectedItems = fd.SelectedItems
        SelectSingleFile = selectedItems(1)
    Else
        SelectSingleFile = ""
    End If
    
    Exit Function
    
ErrorHandler:
    SelectSingleFile = ""
End Function

' Count files in file paths string
Private Function CountFiles(filePaths As String) As Integer
    On Error GoTo ErrorHandler
    
    Dim fileArray() As String
    
    If filePaths = "" Then
        CountFiles = 0
        Exit Function
    End If
    
    fileArray = Split(filePaths, "|")
    CountFiles = UBound(fileArray) + 1
    
    Exit Function
    
ErrorHandler:
    CountFiles = 0
End Function

' Import multiple files to PostgreSQL
Private Sub ImportMultipleFilesToPostgreSQL(filePaths As String, fileCount As Integer, appName As String)
    On Error GoTo ErrorHandler
    
    Dim fileArray() As String
    Dim i As Integer
    Dim currentFile As String
    Dim successCount As Integer
    Dim failedFiles As String
    Dim progressForm As UserForm
    
    ' Split file paths
    fileArray = Split(filePaths, "|")
    
    ' Disable Excel features to prevent blank screen
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    ' Show progress
    MsgBox "Starting import of " & fileCount & " files to PostgreSQL..." & vbCrLf & vbCrLf & _
           "APP: " & appName & vbCrLf & vbCrLf & _
           "Excel will be temporarily disabled during import to prevent blank screen.", vbInformation
    
    successCount = 0
    failedFiles = ""
    
    ' Process each file
    For i = 0 To UBound(fileArray)
        currentFile = fileArray(i)
        
        ' Show progress
        Application.StatusBar = "Importing file " & (i + 1) & " of " & fileCount & ": " & GetFileName(currentFile)
        DoEvents
        
        ' Force Excel to process events and prevent blank screen
        ForceExcelRefresh
        
        ' Import current file (suppress individual success messages) with specified APP
        If ImportSingleFileToPostgreSQLSilent(currentFile, appName) Then
            successCount = successCount + 1
        Else
            ' Track failed files
            If failedFiles <> "" Then failedFiles = failedFiles & vbCrLf
            failedFiles = failedFiles & GetFileName(currentFile)
        End If
        
        ' Force Excel to refresh and prevent blank screen
        ForceExcelRefresh
    Next i
    
    ' Re-enable Excel features
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Force Excel to refresh
    DoEvents
    Application.Wait Now + TimeValue("00:00:02")
    
    Dim summaryMessage As String
    summaryMessage = "Bulk Import Summary:" & vbCrLf & vbCrLf & _
                     "Total Files: " & fileCount & vbCrLf & _
                     "Successfully Imported: " & successCount & vbCrLf & _
                     "Failed: " & (fileCount - successCount) & " files"
    
    If failedFiles <> "" Then
        summaryMessage = summaryMessage & vbCrLf & vbCrLf & "Failed Files:" & vbCrLf & failedFiles
    End If
    
    MsgBox summaryMessage, vbInformation, "Import Summary"
    
    Exit Sub
    
ErrorHandler:
    ' Re-enable Excel features even if error occurs
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    DoEvents
    MsgBox "Error in ImportMultipleFilesToPostgreSQL: " & Err.Description, vbCritical
End Sub

' Import single file to PostgreSQL (with success message)
Private Function ImportSingleFileToPostgreSQL(filePath As String, appName As String) As Boolean
    Dim success As Boolean
    success = ImportSingleFileToPostgreSQLCore(filePath, appName)
    
    If success Then
        ' Show success message for single file import
        MsgBox "Import completed successfully!" & vbCrLf & _
               "File: " & GetFileName(filePath) & vbCrLf & _
               "APP: " & appName, vbInformation, "Import Success"
    End If
    
    ImportSingleFileToPostgreSQL = success
End Function

' Import single file to PostgreSQL (silent version - no success message)
Private Function ImportSingleFileToPostgreSQLSilent(filePath As String, appName As String) As Boolean
    Dim result As Boolean
    result = ImportSingleFileToPostgreSQLCore(filePath, appName)
    ImportSingleFileToPostgreSQLSilent = result
End Function

' Core import function (no messages)
Private Function ImportSingleFileToPostgreSQLCore(filePath As String, appName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim bookData As Collection
    Dim transactionData As Collection
    Dim success As Boolean
    
    Debug.Print "=== ImportSingleFileToPostgreSQL Started ==="
    Debug.Print "File path: " & filePath
    
    ' Temporarily disable Excel features to prevent blank screen
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Open workbook
    Debug.Print "Opening workbook..."
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Debug.Print "Workbook opened successfully"
    
    ' Get first worksheet
    Debug.Print "Getting first worksheet..."
    
    ' Validate workbook is open
    If wb Is Nothing Then
        Debug.Print "ERROR: Workbook is Nothing"
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    
    ' Check if workbook has worksheets
    On Error Resume Next
    Dim sheetCount As Long
    sheetCount = wb.Worksheets.Count
    If Err.Number <> 0 Or sheetCount = 0 Then
        Debug.Print "ERROR: Workbook has no worksheets - " & Err.Description
        wb.Close False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "Total worksheets: " & sheetCount
    
    ' Get first worksheet for data extraction
    On Error Resume Next
    Set ws = wb.Worksheets(1)
    If Err.Number <> 0 Then
        Debug.Print "ERROR: Failed to get first worksheet - " & Err.Description & " (Error " & Err.Number & ")"
        wb.Close False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Validate worksheet
    If ws Is Nothing Then
        Debug.Print "ERROR: Worksheet is Nothing after setting"
        wb.Close False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    
    Debug.Print "Using first worksheet for data: " & ws.Name
    
    ' Use APP name from parameter (provided by user input)
    Dim appValueFromExcel As String
    appValueFromExcel = appName  ' Use the APP name passed as parameter
    Debug.Print "APP value from parameter: " & appValueFromExcel
    
    ' Find data range
    Debug.Print "Finding data range..."
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If Err.Number <> 0 Then
        Debug.Print "ERROR: Failed to find last row - " & Err.Description
        lastRow = 1
    End If
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Err.Number <> 0 Then
        Debug.Print "ERROR: Failed to find last column - " & Err.Description
        lastCol = 1
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "Last row: " & lastRow & ", Last column: " & lastCol
    
    ' Check if we have data
    If lastRow < 2 Then
        Debug.Print "No data found - lastRow < 2"
        wb.Close False
        ' Re-enable Excel features
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    
    ' Get data range (include header row) - ensure we have enough columns
    Debug.Print "Setting data range..."
    If lastCol < 12 Then
        lastCol = 12  ' Ensure we have at least 12 columns for all data
    End If
    
    ' Validate worksheet and range before setting
    If ws Is Nothing Then
        Debug.Print "ERROR: Worksheet is Nothing"
        wb.Close False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))  ' Include header row
    
    ' Validate dataRange before using
    If dataRange Is Nothing Then
        Debug.Print "ERROR: DataRange is Nothing"
        wb.Close False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        ImportSingleFileToPostgreSQLCore = False
        Exit Function
    End If
    
    ' Extract data
    Set bookData = New Collection
    Set transactionData = New Collection
    
    ExtractDataFromRange dataRange, bookData, transactionData
    
    ' Close workbook
    wb.Close False
    
    ' Import to PostgreSQL with APP value from Excel
    success = ImportDataToPostgreSQL(bookData, transactionData, GetFileName(filePath), appValueFromExcel)
    
    ' Re-enable Excel features
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ImportSingleFileToPostgreSQLCore = success
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in ImportSingleFileToPostgreSQLCore ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    
    ' Clean up all objects
    On Error Resume Next
    If Not dataRange Is Nothing Then
        Set dataRange = Nothing
    End If
    If Not ws Is Nothing Then
        Set ws = Nothing
    End If
    If Not wb Is Nothing Then
        wb.Close False
        Set wb = Nothing
    End If
    If Not bookData Is Nothing Then
        Set bookData = Nothing
    End If
    If Not transactionData Is Nothing Then
        Set transactionData = Nothing
    End If
    On Error GoTo 0
    
    ' Re-enable Excel features even if error occurs
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    ImportSingleFileToPostgreSQLCore = False
End Function

' Extract data from Excel range
Private Sub ExtractDataFromRange(dataRange As Range, ByRef bookData As Collection, ByRef transactionData As Collection)
    On Error GoTo ErrorHandler
    
    Dim row As Range
    Dim bookTitle As String
    Dim language As String
    Dim year As String
    Dim month As String
    Dim half As String
    Dim qty As String
    Dim avail As String
    Dim village As String
    Dim groupSector As String
    Dim contactName As String
    Dim bookKey As String
    Dim transactionKey As String
    Dim rowCount As Long
    Dim totalRows As Long
    
    ' Column mapping variables
    Dim columnMaps As ColumnMappings
    
    ' Find header row and map columns
    If Not FindColumnMappings(dataRange, columnMaps) Then
        Exit Sub
    End If
    
    rowCount = 0
    totalRows = dataRange.Rows.Count
    
    ' Process each row (skip header row)
    For Each row In dataRange.Rows
        rowCount = rowCount + 1
        
        ' Skip header row (first row)
        If rowCount = 1 Then
            Debug.Print "Skipping header row"
            GoTo NextRow
        End If
        
        ' Show progress every 100 rows to prevent Excel from becoming unresponsive
        If rowCount Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & rowCount & " of " & totalRows & " (" & Format(rowCount / totalRows, "0%") & ")"
            DoEvents
        End If
        
        ' Get values from mapped columns
        bookTitle = CStr(row.Cells(columnMaps.bookTitleCol).Value)
        language = CStr(row.Cells(columnMaps.languageCol).Value)
        year = CStr(row.Cells(columnMaps.yearCol).Value)
        month = CStr(row.Cells(columnMaps.monthCol).Value)
        
        ' Get half value - handle potential formula or merged cell issues
        On Error Resume Next
        half = CStr(row.Cells(columnMaps.halfCol).Value)
        ' If Value is empty, try Text property
        If half = "" Or IsNull(half) Then
            half = CStr(row.Cells(columnMaps.halfCol).Text)
        End If
        On Error GoTo ErrorHandler
        
        ' Debug: Print half value for first few rows
        If rowCount <= 5 Then
            Debug.Print "Row " & rowCount & " - Half value from Excel = '" & half & "' (Type: " & TypeName(row.Cells(columnMaps.halfCol).Value) & ")"
        End If
        
        qty = CStr(row.Cells(columnMaps.qtyCol).Value)
        avail = CStr(row.Cells(columnMaps.availCol).Value)
        village = CStr(row.Cells(columnMaps.villageCol).Value)
        groupSector = CStr(row.Cells(columnMaps.groupSectorCol).Value)
        contactName = CStr(row.Cells(columnMaps.contactNameCol).Value)
        
        ' Skip empty rows
        If bookTitle = "" Or bookTitle = "0" Then
            GoTo NextRow
        End If
        
        ' Create unique keys - BOOK TITLE + LANGUAGE combination
        bookKey = bookTitle & "|" & language  ' Book title + language as key
        ' For transactions, include row number to make each row unique
        transactionKey = bookTitle & "|" & year & "|" & month & "|" & half & "|" & village & "|" & rowCount
        
        ' Add to book collection (if not exists) - UNIQUE BOOK TITLE + LANGUAGE combination
        If Not CollectionContains(bookData, bookKey) Then
            bookData.Add Array(bookTitle, language), bookKey  ' Store book title and language
        End If
        
        ' Add to transaction collection
        If Not CollectionContains(transactionData, transactionKey) Then
            transactionData.Add Array(bookTitle, language, year, month, half, qty, avail, village, groupSector, contactName), transactionKey
        End If
        
NextRow:
    Next row
    
    ' Count total rows and skipped rows for summary
    Dim totalRowsProcessed As Integer
    Dim skippedRowsCount As Integer
    Dim tempRowCount As Long
    
    totalRowsProcessed = dataRange.Rows.Count - 1  ' Exclude header row
    skippedRowsCount = 0
    
    ' Count skipped rows (rows with empty book titles)
    tempRowCount = 0
    For Each row In dataRange.Rows
        tempRowCount = tempRowCount + 1
        If tempRowCount > 1 Then  ' Skip header row
            bookTitle = CStr(row.Cells(columnMaps.bookTitleCol).Value)
            If bookTitle = "" Or bookTitle = "0" Then
                skippedRowsCount = skippedRowsCount + 1
            End If
        End If
    Next row
    
    Debug.Print "=== IMPORT SUMMARY ==="
    Debug.Print "Total rows in Excel file: " & totalRowsProcessed
    Debug.Print "Rows processed successfully: " & transactionData.Count
    Debug.Print "Total books: " & bookData.Count
    Debug.Print "Total transactions: " & transactionData.Count
    Debug.Print "================================"
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in ExtractDataFromRange ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    
    ' Clean up objects
    On Error Resume Next
    If Not row Is Nothing Then
        Set row = Nothing
    End If
    On Error GoTo 0
End Sub

' Check if collection contains a key
Private Function CollectionContains(col As Collection, key As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim item As Variant
    item = col(key)
    CollectionContains = True
    Exit Function
    
ErrorHandler:
    CollectionContains = False
End Function

' Import data to PostgreSQL
Private Function ImportDataToPostgreSQL(bookData As Collection, transactionData As Collection, fileName As String, Optional appValueFromExcel As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim success As Boolean
    Dim bookId As Long
    Dim bookItem As Variant
    Dim transactionItem As Variant
    Dim transactionArray As Variant
    Dim appValue As String
    
    ' Use APP value from Excel A1 cell, or default to VIMARSH
    If appValueFromExcel <> "" Then
        appValue = appValueFromExcel
        Debug.Print "Using APP value from Excel: " & appValue
    Else
        ' Fallback to default if not found in Excel
        appValue = "VIMARSH"
        Debug.Print "APP value not found in Excel, using default: " & appValue
    End If
    
    ' Create connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL
    If Not ConnectToPostgreSQL(conn) Then
        ImportDataToPostgreSQL = False
        Exit Function
    End If
    
    ' Create tables if they don't exist
    If Not CreatePostgreSQLTables(conn) Then
        ImportDataToPostgreSQL = False
        Exit Function
    End If
    
    ' Create import history record
    Dim historyId As Long
    Dim historySql As String
    Dim period As String
    Dim totalBookQty As Integer
    Dim existingHistoryId As Long
    Dim existingHistoryRs As Object
    
    ' Calculate period and total book quantity from transaction data
    period = CalculatePeriodFromTransactions(transactionData)
    totalBookQty = CalculateTotalBookQty(transactionData)
    
    ' Check if same file name already exists for the same APP and period in import history
    ' Only check for duplicates if filename AND APP AND period all match - different APP can have same filename
    ' IMPORTANT: Check directly in import_history table, not through transactions join
    ' This ensures we check the actual APP value stored in import_history
    historySql = "SELECT h.history_id, h.file_name, h.period, COALESCE(h.app, 'VIMARSH') as app " & _
                 "FROM " & DB_SCHEMA & ".import_history h " & _
                 "WHERE h.file_name = '" & Replace(fileName, "'", "''") & "' " & _
                 "AND COALESCE(h.app, 'VIMARSH') = '" & Replace(appValue, "'", "''") & "' " & _
                 "AND h.period = '" & Replace(period, "'", "''") & "' " & _
                 "AND h.is_deleted = FALSE " & _
                 "LIMIT 1;"
    
    Set existingHistoryRs = conn.Execute(historySql)
    
    If Not existingHistoryRs.EOF Then
        ' Same file name exists for same APP and period in transaction history
        existingHistoryId = existingHistoryRs.Fields("history_id").Value
        Dim existingFileName As String
        Dim existingPeriod As String
        Dim existingApp As String
        existingFileName = existingHistoryRs.Fields("file_name").Value
        existingPeriod = existingHistoryRs.Fields("period").Value
        existingApp = existingHistoryRs.Fields("app").Value
        existingHistoryRs.Close
        Set existingHistoryRs = Nothing
        
        ' Ask user what to do
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("A file with the same name '" & fileName & "' has already been imported for APP '" & appValue & "' and period '" & period & "'." & vbCrLf & vbCrLf & _
                            "Existing History ID: " & existingHistoryId & vbCrLf & _
                            "Existing File: " & existingFileName & vbCrLf & _
                            "Existing APP: " & existingApp & vbCrLf & _
                            "Existing Period: " & existingPeriod & vbCrLf & vbCrLf & _
                            "Do you want to:" & vbCrLf & _
                            "• YES = Delete old data and import new data" & vbCrLf & _
                            "• NO = Cancel this import" & vbCrLf & vbCrLf & _
                            "Click YES to replace old data, or NO to cancel.", vbYesNo + vbQuestion, "Duplicate Import Detected")
        
        If userResponse = vbYes Then
            ' Delete existing data and continue with new import
            
            ' Delete existing import history (transactions will be deleted due to CASCADE)
            historySql = "DELETE FROM " & DB_SCHEMA & ".import_history WHERE history_id = " & existingHistoryId & ";"
            conn.Execute historySql
        Else
            ' User cancelled
            ImportDataToPostgreSQL = False
            Exit Function
        End If
    Else
        existingHistoryRs.Close
        Set existingHistoryRs = Nothing
    End If
    
    ' Now create new import history record (including APP)
    historySql = "INSERT INTO " & DB_SCHEMA & ".import_history (file_name, period, app, total_book_qty) " & _
                 "VALUES ('" & Replace(fileName, "'", "''") & "', '" & Replace(period, "'", "''") & "', '" & Replace(appValue, "'", "''") & "', " & totalBookQty & ") " & _
                 "RETURNING history_id;"
    Debug.Print "History SQL: " & historySql
    
    Dim historyRs As Object
    Set historyRs = conn.Execute(historySql)
    historyId = historyRs.Fields("history_id").Value
    historyRs.Close
    Set historyRs = Nothing
    
    Debug.Print "Import history created with ID: " & historyId
    
         ' Import books first
     Debug.Print "Starting book import..."
     For Each bookItem In bookData
         Debug.Print "Processing book item (array): " & TypeName(bookItem)
         bookId = InsertBookToPostgreSQL(conn, bookItem)
         If bookId = 0 Then
             Debug.Print "Failed to insert book"
             GoTo ErrorHandler
         End If
     Next bookItem
    
    ' Import transactions
    Dim transactionCount As Integer
    transactionCount = 0
    
    For Each transactionItem In transactionData
        transactionArray = transactionItem
        If InsertTransactionToPostgreSQL(conn, transactionArray, historyId, appValue) Then
            transactionCount = transactionCount + 1
        Else
            GoTo ErrorHandler
        End If
    Next transactionItem
    
    ' Close connection
    conn.Close
    Set conn = Nothing
    
    ImportDataToPostgreSQL = True
    Exit Function
    
ErrorHandler:
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    ImportDataToPostgreSQL = False
End Function

' Extract APP name from "village_namewise" sheet A1 cell
' Expected format: "PARISHKAAR DEC-2025(1-31)" or "VIMARSH NOV-2025(1-30)"
Private Function ExtractAppFromVillageNamewiseSheet(wb As Workbook) As String
    On Error GoTo ErrorHandler
    
    Dim villageWs As Worksheet
    Dim sheetFound As Boolean
    sheetFound = False
    Dim i As Long
    Dim sheetCount As Long
    
    ' Get total sheet count
    On Error Resume Next
    sheetCount = wb.Worksheets.Count
    On Error GoTo ErrorHandler
    
    ' Search for "village_namewise" sheet (case-insensitive)
    For i = 1 To sheetCount
        On Error Resume Next
        Dim tempWs As Worksheet
        Set tempWs = wb.Worksheets(i)
        If Err.Number = 0 Then
            Dim sheetName As String
            sheetName = LCase(Trim(tempWs.Name))
            Debug.Print "Checking sheet " & i & " for APP value: '" & tempWs.Name & "'"
            
            ' Check if sheet name contains "village" and "namewise" or exact match
            If sheetName = "village_namewise" Or sheetName = "village namewise" Or _
               (InStr(sheetName, "village") > 0 And InStr(sheetName, "namewise") > 0) Then
                Set villageWs = tempWs
                sheetFound = True
                Debug.Print "Found village_namewise sheet: '" & villageWs.Name & "'"
                Exit For
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' If village_namewise sheet found, extract APP from it
    If sheetFound And Not villageWs Is Nothing Then
        ExtractAppFromVillageNamewiseSheet = ExtractAppFromCell(villageWs)
    Else
        Debug.Print "village_namewise sheet not found, using default VIMARSH"
        ExtractAppFromVillageNamewiseSheet = "VIMARSH"
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error extracting APP from village_namewise sheet: " & Err.Description
    ExtractAppFromVillageNamewiseSheet = "VIMARSH"
End Function

' Extract APP name from Excel cell A1 (handles merged cells A1:G1)
' Expected format: "PARISHKAAR DEC-2025(1-31)" or "VIMARSH NOV-2025(1-30)"
Private Function ExtractAppFromCell(ws As Worksheet) As String
    On Error GoTo ErrorHandler
    
    Dim cellValue As String
    Dim appName As String
    Dim parts() As String
    Dim cellRange As Range
    
    ' Try to read A1 cell (works for both merged and unmerged cells)
    On Error Resume Next
    
    ' Method 1: Try Range("A1") - better for merged cells
    Set cellRange = ws.Range("A1")
    If Not cellRange Is Nothing Then
        cellValue = Trim(CStr(cellRange.Value))
    End If
    
    ' Method 2: If empty, try Cells(1,1)
    If cellValue = "" Or IsNull(cellValue) Then
        cellValue = Trim(CStr(ws.Cells(1, 1).Value))
    End If
    
    ' Method 3: If still empty, try reading from merged cell range
    If cellValue = "" Or IsNull(cellValue) Then
        On Error Resume Next
        Set cellRange = ws.Range("A1:G1")
        If Not cellRange Is Nothing Then
            cellValue = Trim(CStr(cellRange.Value))
        End If
        On Error GoTo ErrorHandler
    End If
    
    On Error GoTo ErrorHandler
    
    ' If cell is still empty, return default
    If cellValue = "" Or IsNull(cellValue) Or cellValue = "Empty" Then
        ExtractAppFromCell = "VIMARSH"
        Debug.Print "A1 cell is empty or could not read, using default VIMARSH"
        Exit Function
    End If
    
    Debug.Print "A1 cell value read: '" & cellValue & "'"
    Debug.Print "A1 cell value length: " & Len(cellValue)
    
    ' Extract first word (APP name)
    ' Split by space and get first part
    parts = Split(cellValue, " ")
    
    If UBound(parts) >= 0 Then
        appName = UCase(Trim(parts(0)))
        Debug.Print "First word extracted: '" & appName & "'"
        
        ' Validate APP name
        If appName = "VIMARSH" Or appName = "PARISHKAAR" Then
            ExtractAppFromCell = appName
            Debug.Print "*** SUCCESS: Extracted APP name: " & appName & " ***"
        Else
            ' Invalid APP name, use default
            ExtractAppFromCell = "VIMARSH"
            Debug.Print "*** WARNING: Invalid APP name in A1: '" & appName & "', using default VIMARSH ***"
            Debug.Print "Full cell value was: '" & cellValue & "'"
        End If
    Else
        ' No space found, check if entire value is APP name
        appName = UCase(Trim(cellValue))
        Debug.Print "No space found, checking if entire value is APP name: '" & appName & "'"
        If appName = "VIMARSH" Or appName = "PARISHKAAR" Then
            ExtractAppFromCell = appName
            Debug.Print "*** SUCCESS: Entire value is APP name: " & appName & " ***"
        Else
            ExtractAppFromCell = "VIMARSH"
            Debug.Print "*** WARNING: Could not extract APP name from A1, using default VIMARSH ***"
            Debug.Print "Full cell value was: '" & cellValue & "'"
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "*** ERROR extracting APP from A1: " & Err.Description & " (Error " & Err.Number & ") ***"
    ExtractAppFromCell = "VIMARSH"
End Function

' Connect to PostgreSQL with automatic database creation
Private Function ConnectToPostgreSQL(conn As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim connectionString As String
    Dim password As String
    Dim driverOptions() As String
    Dim i As Integer
    Dim databaseExists As Boolean
    
    Debug.Print "=== ConnectToPostgreSQL Started ==="
    
    ' Use hardcoded password
    password = DB_PASSWORD
    Debug.Print "Using hardcoded password, trying different driver options..."
    
    ' Define multiple driver options to try
    ReDim driverOptions(4)
    driverOptions(0) = "PostgreSQL ODBC Driver(UNICODE)"
    driverOptions(1) = "PostgreSQL UNICODE"
    driverOptions(2) = "PostgreSQL ODBC Driver"
    driverOptions(3) = "PostgreSQL"
    driverOptions(4) = "PostgreSQL ANSI"
    
    ' First, try to connect to the target database directly
    For i = 0 To 4
        Debug.Print "Trying driver: " & driverOptions(i)
        
        ' Build connection string for target database
        connectionString = "Driver={" & driverOptions(i) & "};" & _
                          "Server=" & DB_HOST & ";" & _
                          "Port=" & DB_PORT & ";" & _
                          "Database=" & DB_NAME & ";" & _
                          "Uid=" & DB_USER & ";" & _
                          "Pwd=" & password & ";"
        
        Debug.Print "Connection string: " & connectionString
        
        ' Try to open connection to target database
        On Error Resume Next
        conn.Open connectionString
        If Err.Number = 0 Then
            Debug.Print "Connection opened successfully to target database with driver: " & driverOptions(i)
            ConnectToPostgreSQL = True
            Exit Function
        Else
            Debug.Print "Failed to connect to target database with driver " & driverOptions(i) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' If target database connection failed, try to create database
    Debug.Print "Target database connection failed. Attempting to create database..."
    
    ' Try to connect to default postgres database to create target database
    For i = 0 To 4
        Debug.Print "Trying to connect to postgres database with driver: " & driverOptions(i)
        
        ' Build connection string for postgres database
        connectionString = "Driver={" & driverOptions(i) & "};" & _
                          "Server=" & DB_HOST & ";" & _
                          "Port=" & DB_PORT & ";" & _
                          "Database=postgres;" & _
                          "Uid=" & DB_USER & ";" & _
                          "Pwd=" & password & ";"
        
        Debug.Print "Postgres connection string: " & connectionString
        
        ' Try to open connection to postgres database
        On Error Resume Next
        conn.Open connectionString
        If Err.Number = 0 Then
            Debug.Print "Connected to postgres database successfully with driver: " & driverOptions(i)
            
            ' Create the target database
            Dim createDbSql As String
            createDbSql = "CREATE DATABASE " & DB_NAME & ";"
            Debug.Print "Creating database with SQL: " & createDbSql
            
            On Error Resume Next
            conn.Execute createDbSql
            If Err.Number = 0 Then
                Debug.Print "Database " & DB_NAME & " created successfully"
                databaseExists = True
            ElseIf Err.Number = -2147217900 Or InStr(Err.Description, "already exists") > 0 Then
                Debug.Print "Database " & DB_NAME & " already exists"
                databaseExists = True
                Err.Clear
            Else
                Debug.Print "Failed to create database: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' Close connection to postgres database
            conn.Close
            
            ' If database was created or exists, try to connect to target database
            If databaseExists Then
                ' Build connection string for target database
                connectionString = "Driver={" & driverOptions(i) & "};" & _
                                  "Server=" & DB_HOST & ";" & _
                                  "Port=" & DB_PORT & ";" & _
                                  "Database=" & DB_NAME & ";" & _
                                  "Uid=" & DB_USER & ";" & _
                                  "Pwd=" & password & ";"
                
                Debug.Print "Connecting to target database: " & connectionString
                
                On Error Resume Next
                conn.Open connectionString
                If Err.Number = 0 Then
                    Debug.Print "Connection opened successfully to target database after creation with driver: " & driverOptions(i)
                    ConnectToPostgreSQL = True
                    Exit Function
                Else
                    Debug.Print "Failed to connect to target database after creation: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
        Else
            Debug.Print "Failed to connect to postgres database with driver " & driverOptions(i) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' If we get here, all attempts failed
    Debug.Print "All connection and database creation attempts failed"
    MsgBox "Could not connect to PostgreSQL or create database." & vbCrLf & _
           "Please ensure:" & vbCrLf & _
           "1. PostgreSQL server is running" & vbCrLf & _
           "2. PostgreSQL ODBC driver is installed" & vbCrLf & _
           "3. User credentials are correct" & vbCrLf & _
           "4. User has permission to create databases" & vbCrLf & _
           "Download ODBC driver from: https://www.postgresql.org/ftp/odbc/", vbCritical
    ConnectToPostgreSQL = False
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in ConnectToPostgreSQL ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Line: " & Erl
    MsgBox "Error connecting to PostgreSQL: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Check Immediate Window (Ctrl+G) for details", vbCritical
    ConnectToPostgreSQL = False
End Function

' Check if database exists
Private Function DatabaseExists(databaseName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim connectionString As String
    Dim password As String
    Dim driverOptions() As String
    Dim i As Integer
    
    Debug.Print "=== DatabaseExists Started ==="
    Debug.Print "Checking if database exists: " & databaseName
    
    ' Create connection
    Set conn = CreateObject("ADODB.Connection")
    password = DB_PASSWORD
    
    ' Define multiple driver options to try
    ReDim driverOptions(4)
    driverOptions(0) = "PostgreSQL ODBC Driver(UNICODE)"
    driverOptions(1) = "PostgreSQL UNICODE"
    driverOptions(2) = "PostgreSQL ODBC Driver"
    driverOptions(3) = "PostgreSQL"
    driverOptions(4) = "PostgreSQL ANSI"
    
    ' Try to connect to postgres database to check if target database exists
    For i = 0 To 4
        ' Build connection string for postgres database
        connectionString = "Driver={" & driverOptions(i) & "};" & _
                          "Server=" & DB_HOST & ";" & _
                          "Port=" & DB_PORT & ";" & _
                          "Database=postgres;" & _
                          "Uid=" & DB_USER & ";" & _
                          "Pwd=" & password & ";"
        
        ' Try to open connection to postgres database
        On Error Resume Next
        conn.Open connectionString
        If Err.Number = 0 Then
            Debug.Print "Connected to postgres database to check database existence"
            
            ' Check if target database exists
            sql = "SELECT 1 FROM pg_database WHERE datname = '" & databaseName & "';"
            Set rs = conn.Execute(sql)
            
            If Not rs.EOF Then
                Debug.Print "Database " & databaseName & " exists"
                DatabaseExists = True
            Else
                Debug.Print "Database " & databaseName & " does not exist"
                DatabaseExists = False
            End If
            
            rs.Close
            conn.Close
            Set rs = Nothing
            Set conn = Nothing
            Exit Function
        Else
            Debug.Print "Failed to connect to postgres database with driver " & driverOptions(i) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' If we get here, could not connect to postgres database
    Debug.Print "Could not connect to postgres database to check database existence"
    DatabaseExists = False
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in DatabaseExists ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    DatabaseExists = False
End Function

' Create PostgreSQL tables
Private Function CreatePostgreSQLTables(conn As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    
    Debug.Print "=== CreatePostgreSQLTables Started ==="
    
    ' Create schema
    Debug.Print "Creating schema: " & DB_SCHEMA
    sql = "CREATE SCHEMA IF NOT EXISTS " & DB_SCHEMA & ";"
    Debug.Print "SQL: " & sql
    conn.Execute sql
    Debug.Print "Schema created/verified successfully"
    
         ' Create books table
     Debug.Print "Creating books table..."
     sql = "CREATE TABLE IF NOT EXISTS " & DB_SCHEMA & ".books (" & _
           "book_id SERIAL PRIMARY KEY," & _
           "book_name TEXT NOT NULL," & _
           "language TEXT NOT NULL," & _
           "created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP," & _
           "updated_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP," & _
           "created_by TEXT DEFAULT 'Excel Import'," & _
           "updated_by TEXT DEFAULT 'Excel Import'," & _
           "is_deleted BOOLEAN DEFAULT FALSE," & _
           "UNIQUE(book_name, language)" & _
           ");"
    Debug.Print "SQL: " & sql
    conn.Execute sql
    Debug.Print "Books table created/verified successfully"
    
    ' Create import history table
    Debug.Print "Creating import history table..."
    sql = "CREATE TABLE IF NOT EXISTS " & DB_SCHEMA & ".import_history (" & _
          "history_id SERIAL PRIMARY KEY," & _
          "file_name TEXT NOT NULL," & _
          "import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP," & _
          "period TEXT," & _
          "app TEXT DEFAULT 'VIMARSH'," & _
          "total_book_qty INTEGER DEFAULT 0," & _
          "status TEXT DEFAULT 'SUCCESS'," & _
          "created_by TEXT DEFAULT 'Excel Import'," & _
          "is_deleted BOOLEAN DEFAULT FALSE" & _
          ");"
    Debug.Print "SQL: " & sql
    conn.Execute sql
    Debug.Print "Import history table created/verified successfully"
    
    ' Ensure APP column exists in import_history table (for existing tables)
    On Error Resume Next
    sql = "ALTER TABLE " & DB_SCHEMA & ".import_history ADD COLUMN IF NOT EXISTS app TEXT DEFAULT 'VIMARSH';"
    conn.Execute sql
    On Error GoTo ErrorHandler
    Debug.Print "APP column verified/added to import_history table"
    
         ' Create transactions table
     Debug.Print "Creating transactions table..."
     sql = "CREATE TABLE IF NOT EXISTS " & DB_SCHEMA & ".transactions (" & _
           "importid SERIAL PRIMARY KEY," & _
           "history_id INTEGER REFERENCES " & DB_SCHEMA & ".import_history(history_id) ON DELETE CASCADE," & _
           "book_id INTEGER REFERENCES " & DB_SCHEMA & ".books(book_id)," & _
           "year INTEGER," & _
           "month INTEGER," & _
           "half TEXT," & _
           "qty INTEGER," & _
           "avail INTEGER," & _
           "village TEXT," & _
           "group_sector TEXT," & _
           "contact_name TEXT," & _
           "created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP," & _
           "updated_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP," & _
           "created_by TEXT DEFAULT 'Excel Import'," & _
           "updated_by TEXT DEFAULT 'Excel Import'," & _
           "is_deleted BOOLEAN DEFAULT FALSE" & _
           ");"
    Debug.Print "SQL: " & sql
    conn.Execute sql
    Debug.Print "Transactions table created/verified successfully"
    
    Debug.Print "=== CreatePostgreSQLTables Completed Successfully ==="
    CreatePostgreSQLTables = True
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in CreatePostgreSQLTables ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Line: " & Erl
    Debug.Print "Last SQL: " & sql
    MsgBox "Error creating PostgreSQL tables: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Check Immediate Window (Ctrl+G) for details", vbCritical
    CreatePostgreSQLTables = False
End Function

' Insert book to PostgreSQL
Private Function InsertBookToPostgreSQL(conn As Object, bookArray As Variant) As Long
    On Error GoTo ErrorHandler
    
         Dim sql As String
     Dim rs As Object
     Dim bookId As Long
     Dim bookName As String
     Dim language As String
     
     ' Extract book name and language from array
     If IsArray(bookArray) Then
         bookName = CStr(bookArray(0))
         language = CStr(bookArray(1))
     Else
         Debug.Print "ERROR: bookArray is not an array, it's: " & TypeName(bookArray)
         InsertBookToPostgreSQL = 0
         Exit Function
     End If
     
         ' Check if book already exists
    sql = "SELECT book_id FROM " & DB_SCHEMA & ".books WHERE book_name = '" & Replace(bookName, "'", "''") & "' AND language = '" & Replace(language, "'", "''") & "' AND is_deleted = FALSE;"
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        ' Book exists, return existing ID
        bookId = rs.Fields("book_id").Value
    Else
        ' Insert new book
        sql = "INSERT INTO " & DB_SCHEMA & ".books (book_name, language) VALUES ('" & Replace(bookName, "'", "''") & "', '" & Replace(language, "'", "''") & "') RETURNING book_id;"
        Set rs = conn.Execute(sql)
        bookId = rs.Fields("book_id").Value
    End If
    
    rs.Close
    Set rs = Nothing
    
    InsertBookToPostgreSQL = bookId
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in InsertBookToPostgreSQL ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Line: " & Erl
    Debug.Print "Book Name: " & bookName
    Debug.Print "Last SQL: " & sql
    InsertBookToPostgreSQL = 0
End Function

' Insert transaction to PostgreSQL
Private Function InsertTransactionToPostgreSQL(conn As Object, transactionArray As Variant, historyId As Long, Optional appValue As String = "VIMARSH") As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim bookName As String
    Dim language As String
    Dim year As String
    Dim month As String
    Dim half As String
    Dim qty As String
    Dim avail As String
    Dim village As String
    Dim groupSector As String
    Dim contactName As String
    Dim bookId As Long
    Dim rs As Object
    

    
         ' Extract values from array
     If IsArray(transactionArray) Then
         bookName = CStr(transactionArray(0))
         language = CStr(transactionArray(1))
         year = CStr(transactionArray(2))
         month = CStr(transactionArray(3))
         half = CStr(transactionArray(4))
         qty = CStr(transactionArray(5))
         avail = CStr(transactionArray(6))
         village = CStr(transactionArray(7))
         groupSector = CStr(transactionArray(8))
         contactName = CStr(transactionArray(9))
     Else
         Debug.Print "ERROR: transactionArray is not an array, it's: " & TypeName(transactionArray)
         InsertTransactionToPostgreSQL = False
         Exit Function
     End If
    
        ' Convert avail to integer (YES=1, NO=0, other values=0)
    Dim availInt As Integer
    If avail = "YES" Then
        availInt = 1
    ElseIf avail = "NO" Then
        availInt = 0
    Else
        availInt = 0
    End If
    
    ' Get book ID
     sql = "SELECT book_id FROM " & DB_SCHEMA & ".books WHERE book_name = '" & Replace(bookName, "'", "''") & "' AND language = '" & Replace(language, "'", "''") & "' AND is_deleted = FALSE;"
        Set rs = conn.Execute(sql)
    
    If rs.EOF Then
        Debug.Print "ERROR: Book not found in database: " & bookName
        rs.Close
        Set rs = Nothing
        InsertTransactionToPostgreSQL = False
        Exit Function
    End If
    
    bookId = rs.Fields("book_id").Value
    rs.Close
    Set rs = Nothing
    
    ' Insert transaction with APP value
    Dim appVal As String
    If appValue = "" Then
        appVal = "VIMARSH"
    Else
        appVal = appValue
    End If
    
     sql = "INSERT INTO " & DB_SCHEMA & ".transactions (" & _
           "history_id, book_id, year, month, half, qty, avail, village, group_sector, contact_name, app" & _
           ") VALUES (" & _
           historyId & ", " & _
           bookId & ", " & _
           IIf(year = "", "0", year) & ", " & _
           IIf(month = "", "0", month) & ", " & _
           "'" & Replace(half, "'", "''") & "', " & _
           IIf(qty = "", "0", qty) & ", " & _
           availInt & ", " & _
           "'" & Replace(village, "'", "''") & "', " & _
           "'" & Replace(groupSector, "'", "''") & "', " & _
           "'" & Replace(contactName, "'", "''") & "', " & _
           "'" & Replace(appVal, "'", "''") & "'" & _
           ");"
    
    conn.Execute sql
    
    InsertTransactionToPostgreSQL = True
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in InsertTransactionToPostgreSQL ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Line: " & Erl
    Debug.Print "Book Name: " & bookName
    Debug.Print "Last SQL: " & sql
    InsertTransactionToPostgreSQL = False
End Function

' Find column mappings based on header row
Private Function FindColumnMappings(dataRange As Range, ByRef columnMaps As ColumnMappings) As Boolean
    On Error GoTo ErrorHandler
    
    Dim headerRow As Range
    Dim cell As Range
    Dim col As Long
    Dim headerText As String
    
    ' Get the first row (header row)
    Set headerRow = dataRange.Rows(1)
    
    ' Initialize column numbers
    columnMaps.bookTitleCol = 0
    columnMaps.languageCol = 0
    columnMaps.yearCol = 0
    columnMaps.monthCol = 0
    columnMaps.halfCol = 0
    columnMaps.qtyCol = 0
    columnMaps.availCol = 0
    columnMaps.villageCol = 0
    columnMaps.groupSectorCol = 0
    columnMaps.contactNameCol = 0
    
    Debug.Print "Scanning header row for column mappings..."
    
    ' Scan each column in header row
    For Each cell In headerRow.Cells
        col = cell.Column - dataRange.Column + 1  ' Relative column number
        headerText = CStr(cell.Value)
        
        Debug.Print "Column " & col & ": '" & headerText & "'"
        
        ' Map headers to column numbers
        Select Case headerText
            Case "BOOK TITLE", "BOOK_TITLE", "BOOKTITLE", "TITLE"
                columnMaps.bookTitleCol = col
                Debug.Print "  -> Mapped to BOOK TITLE"
            Case "LANGUAGE", "LANG"
                columnMaps.languageCol = col
                Debug.Print "  -> Mapped to LANGUAGE"
            Case "YEAR"
                columnMaps.yearCol = col
                Debug.Print "  -> Mapped to YEAR"
            Case "MONTH"
                columnMaps.monthCol = col
                Debug.Print "  -> Mapped to MONTH"
            Case "HALF", "HALF PERIOD", "HALF_PERIOD"
                columnMaps.halfCol = col
                Debug.Print "  -> Mapped to HALF"
            Case "QTY", "QUANTITY", "QTY."
                columnMaps.qtyCol = col
                Debug.Print "  -> Mapped to QTY"
            Case "AVAIL", "AVAILABLE", "AVAIL."
                columnMaps.availCol = col
                Debug.Print "  -> Mapped to AVAIL"
            Case "VILLAGE"
                columnMaps.villageCol = col
                Debug.Print "  -> Mapped to VILLAGE"
            Case "GROUP(SECTOR)", "GROUP SECTOR", "GROUP_SECTOR", "SECTOR", "GROUP"
                columnMaps.groupSectorCol = col
                Debug.Print "  -> Mapped to GROUP(SECTOR)"
            Case "NAME(CONTACT NO.)", "NAME(CONTACT NO)", "NAME", "CONTACT NAME", "CONTACT", "NAME(CONTACT)"
                columnMaps.contactNameCol = col
                Debug.Print "  -> Mapped to NAME(CONTACT NO.)"
        End Select
    Next cell
    
    ' Check if all required columns were found
    If columnMaps.bookTitleCol = 0 Or columnMaps.languageCol = 0 Or columnMaps.yearCol = 0 Or columnMaps.monthCol = 0 Or columnMaps.halfCol = 0 Or columnMaps.qtyCol = 0 Or columnMaps.availCol = 0 Or columnMaps.villageCol = 0 Or columnMaps.groupSectorCol = 0 Or columnMaps.contactNameCol = 0 Then
        Debug.Print "ERROR: Missing required columns:"
        If columnMaps.bookTitleCol = 0 Then Debug.Print "  - BOOK TITLE"
        If columnMaps.languageCol = 0 Then Debug.Print "  - LANGUAGE"
        If columnMaps.yearCol = 0 Then Debug.Print "  - YEAR"
        If columnMaps.monthCol = 0 Then Debug.Print "  - MONTH"
        If columnMaps.halfCol = 0 Then Debug.Print "  - HALF"
        If columnMaps.qtyCol = 0 Then Debug.Print "  - QTY"
        If columnMaps.availCol = 0 Then Debug.Print "  - AVAIL"
        If columnMaps.villageCol = 0 Then Debug.Print "  - VILLAGE"
        If columnMaps.groupSectorCol = 0 Then Debug.Print "  - GROUP(SECTOR)"
        If columnMaps.contactNameCol = 0 Then Debug.Print "  - NAME(CONTACT NO.)"
        FindColumnMappings = False
        Exit Function
    End If
    
    Debug.Print "All required columns found successfully!"
    FindColumnMappings = True
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in FindColumnMappings: " & Err.Description
    FindColumnMappings = False
End Function

' Get file name from full path
Private Function GetFileName(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    fileName = Dir(filePath)
    GetFileName = fileName
    
    Exit Function
    
ErrorHandler:
    GetFileName = filePath
End Function

' Calculate period from transaction data (year-month-half format)
' For full month (1 & 2 or 1_2), format as "2025-NOV (1 To 30)"
' Otherwise format as "2025-11-1" or "2025-11-2"
Private Function CalculatePeriodFromTransactions(transactionData As Collection) As String
    On Error GoTo ErrorHandler
    
    Dim transactionItem As Variant
    Dim transactionArray As Variant
    Dim year As String
    Dim month As String
    Dim half As String
    Dim period As String
    Dim yearInt As Integer
    Dim monthInt As Integer
    Dim halfTrimmed As String
    
    ' Get first transaction to determine period
    If transactionData.Count > 0 Then
        transactionItem = transactionData.Item(1)
        transactionArray = transactionItem
        
        year = CStr(transactionArray(2))  ' year is at index 2
        month = CStr(transactionArray(3))  ' month is at index 3
        half = CStr(transactionArray(4))   ' half is at index 4
        
        ' Debug: Print the raw half value
        Debug.Print "CalculatePeriodFromTransactions: Raw half value = '" & half & "' (Length: " & Len(half) & ")"
        
        yearInt = CInt(year)
        monthInt = CInt(month)
        halfTrimmed = Trim(LCase(half))
        
        ' Handle HTML entities and other variations
        halfTrimmed = Replace(halfTrimmed, "&amp;", "&")
        halfTrimmed = Replace(halfTrimmed, "&AMP;", "&")
        
        ' Remove all spaces for comparison (to handle "1 & 2", "1&2", "1 &2", "1& 2", etc.)
        Dim halfNoSpaces As String
        halfNoSpaces = Replace(halfTrimmed, " ", "")
        halfNoSpaces = Replace(halfNoSpaces, "_", "")  ' Also remove underscores
        
        Debug.Print "CalculatePeriodFromTransactions: Raw Half='" & half & "', Trimmed='" & halfTrimmed & "', NoSpaces='" & halfNoSpaces & "'"
        
        ' Check if it's full month - handle multiple variations
        ' Check for: "1&2", "1 & 2", "1_2", "1 _ 2", "1& 2", "1 &2", etc.
        Dim isFullMonth As Boolean
        isFullMonth = False
        
        ' Check if it's full month (1 & 2 or 1_2) - must contain both 1 and 2 with separator
        If InStr(halfTrimmed, "1") > 0 And InStr(halfTrimmed, "2") > 0 Then
            ' Contains both 1 and 2, check for separator (&, _, or " and ")
            If InStr(halfTrimmed, "&") > 0 Or InStr(halfTrimmed, "_") > 0 Or InStr(halfTrimmed, " and ") > 0 Then
                isFullMonth = True
                Debug.Print "*** Full month detected: Contains 1 and 2 with separator ***"
            End If
        End If
        
        Debug.Print "CalculatePeriodFromTransactions: isFullMonth = " & isFullMonth
        
        If isFullMonth Then
            ' Full month - format as "2025-NOV (1 To 30)"
            Dim monthNameStr As String
            Dim lastDay As Integer
            monthNameStr = UCase(Left(MonthName(monthInt), 3))  ' Get first 3 letters: NOV, DEC, etc.
            lastDay = Day(DateSerial(yearInt, monthInt + 1, 0))  ' Get last day of month
            period = year & "-" & monthNameStr & " (1 To " & lastDay & ")"
            
            Debug.Print "*** CalculatePeriodFromTransactions: Full month detected ***"
            Debug.Print "  Year=" & year & ", Month=" & month & ", Half='" & half & "'"
            Debug.Print "  Period='" & period & "'"
        Else
            ' Format month to 2 digits for regular half periods
            If Len(month) = 1 Then
                month = "0" & month
            End If
            
            ' Convert half to 1 or 2 only
            If half = "1" Or half = "FIRST" Or half = "1ST" Or UCase(Trim(half)) = "1" Then
                half = "1"
            ElseIf half = "2" Or half = "SECOND" Or half = "2ND" Or UCase(Trim(half)) = "2" Then
                half = "2"
            Else
                half = "1"  ' Default to 1 if not recognized
            End If
            
            ' Create period in format: YYYY-MM-Half
            period = year & "-" & month & "-" & half
            
            Debug.Print "CalculatePeriodFromTransactions: Year=" & year & ", Month=" & month & ", Half=" & half & ", Period=" & period
        End If
    Else
        period = "Unknown"
        Debug.Print "CalculatePeriodFromTransactions: No transactions found, using 'Unknown'"
    End If
    
    CalculatePeriodFromTransactions = period
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in CalculatePeriodFromTransactions: " & Err.Description
    CalculatePeriodFromTransactions = "Unknown"
End Function

' Calculate total book quantity from transaction data
Private Function CalculateTotalBookQty(transactionData As Collection) As Integer
    On Error GoTo ErrorHandler
    
    Dim transactionItem As Variant
    Dim transactionArray As Variant
    Dim qty As String
    Dim totalQty As Integer
    Dim i As Integer
    
    totalQty = 0
    
    ' Sum up all quantities from transactions
    For i = 1 To transactionData.Count
        transactionItem = transactionData.Item(i)
        transactionArray = transactionItem
        
        qty = CStr(transactionArray(5))  ' qty is at index 5
        
        ' Convert to integer, default to 0 if not numeric
        If IsNumeric(qty) Then
            totalQty = totalQty + CInt(qty)
        End If
    Next i
    
    Debug.Print "CalculateTotalBookQty: Total quantity = " & totalQty
    CalculateTotalBookQty = totalQty
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in CalculateTotalBookQty: " & Err.Description
    CalculateTotalBookQty = 0
End Function

' Delete import history and related transactions
Public Sub DeleteImportHistory(historyId As Long)
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim sql As String
    Dim result As VbMsgBoxResult
    
    ' Confirm deletion
    result = MsgBox("Are you sure you want to delete this import history and all related transactions?" & vbCrLf & _
                   "History ID: " & historyId & vbCrLf & _
                   "This action cannot be undone!", vbYesNo + vbCritical, "Confirm Deletion")
    
    If result <> vbYes Then
        MsgBox "Deletion cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Create connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL
    If Not ConnectToPostgreSQL(conn) Then
        MsgBox "Failed to connect to PostgreSQL", vbCritical
        Exit Sub
    End If
    
         ' Delete import history record (transactions will be automatically deleted due to CASCADE)
     sql = "DELETE FROM " & DB_SCHEMA & ".import_history WHERE history_id = " & historyId & ";"
     Debug.Print "Deleting history (with CASCADE): " & sql
     conn.Execute sql
     
     MsgBox "Import history and all related transactions deleted successfully!" & vbCrLf & _
            "History ID: " & historyId & vbCrLf & _
            "(Related transactions were automatically deleted due to CASCADE)", vbInformation, "Deletion Complete"
        
    Cleanup:
        If Not conn Is Nothing Then
            conn.Close
            Set conn = Nothing
        End If
        Exit Sub
        
ErrorHandler:
    Debug.Print "=== ERROR in DeleteImportHistory ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Line: " & Erl
    
         If Not conn Is Nothing Then
         conn.Close
         Set conn = Nothing
     End If
    
    MsgBox "Error deleting import history: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

' Show import history
Public Sub ShowImportHistory()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== ShowImportHistory Started ==="
    Debug.Print "Showing import history..."
    
    ' Show import history using UserForm
    ImportHistoryForm.Show
    
    Debug.Print "=== ShowImportHistory Completed Successfully ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in ShowImportHistory ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Line: " & Erl
    
    MsgBox "Error showing import history: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

' Optimize Excel settings for import process to prevent blank screen
Private Sub OptimizeExcelForImport()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== OptimizeExcelForImport Started ==="
    
    ' Disable Excel features that can cause blank screen during import
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = False
    
    ' Disable automatic refresh and updates
    Application.AskToUpdateLinks = False
    Application.DisplayFormulaBar = False
    
    Debug.Print "Excel optimized for import process"
    Debug.Print "=== OptimizeExcelForImport Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in OptimizeExcelForImport ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
End Sub

' Restore Excel settings after import process
Private Sub RestoreExcelSettings()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== RestoreExcelSettings Started ==="
    
    ' Re-enable Excel features
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.DisplayFormulaBar = True
    
    ' Force Excel to refresh and prevent blank screen
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    
    Debug.Print "Excel settings restored"
    Debug.Print "=== RestoreExcelSettings Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in RestoreExcelSettings ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
End Sub

' Force Excel to refresh and prevent blank screen
Private Sub ForceExcelRefresh()
    On Error GoTo ErrorHandler
    
    ' Force Excel to process events and refresh display
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Force screen update
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in ForceExcelRefresh ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
End Sub

' Manually create database - can be called separately if needed
Public Function CreateDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim connectionString As String
    Dim password As String
    Dim driverOptions() As String
    Dim i As Integer
    Dim createDbSql As String
    
    Debug.Print "=== CreateDatabase Started ==="
    Debug.Print "Attempting to create database: " & DB_NAME
    
    ' Create connection
    Set conn = CreateObject("ADODB.Connection")
    password = DB_PASSWORD
    
    ' Define multiple driver options to try
    ReDim driverOptions(4)
    driverOptions(0) = "PostgreSQL ODBC Driver(UNICODE)"
    driverOptions(1) = "PostgreSQL UNICODE"
    driverOptions(2) = "PostgreSQL ODBC Driver"
    driverOptions(3) = "PostgreSQL"
    driverOptions(4) = "PostgreSQL ANSI"
    
    ' Try to connect to postgres database to create target database
    For i = 0 To 4
        Debug.Print "Trying to connect to postgres database with driver: " & driverOptions(i)
        
        ' Build connection string for postgres database
        connectionString = "Driver={" & driverOptions(i) & "};" & _
                          "Server=" & DB_HOST & ";" & _
                          "Port=" & DB_PORT & ";" & _
                          "Database=postgres;" & _
                          "Uid=" & DB_USER & ";" & _
                          "Pwd=" & password & ";"
        
        ' Try to open connection to postgres database
        On Error Resume Next
        conn.Open connectionString
        If Err.Number = 0 Then
            Debug.Print "Connected to postgres database successfully"
            
            ' Create the target database
            createDbSql = "CREATE DATABASE " & DB_NAME & ";"
            Debug.Print "Creating database with SQL: " & createDbSql
            
            On Error Resume Next
            conn.Execute createDbSql
            If Err.Number = 0 Then
                Debug.Print "Database " & DB_NAME & " created successfully"
                conn.Close
                Set conn = Nothing
                MsgBox "Database '" & DB_NAME & "' created successfully!" & vbCrLf & _
                       "You can now use the VIMARS import system.", vbInformation, "Database Created"
                CreateDatabase = True
                Exit Function
            ElseIf Err.Number = -2147217900 Or InStr(Err.Description, "already exists") > 0 Then
                Debug.Print "Database " & DB_NAME & " already exists"
                conn.Close
                Set conn = Nothing
                MsgBox "Database '" & DB_NAME & "' already exists." & vbCrLf & _
                       "You can proceed with using the VIMARS import system.", vbInformation, "Database Exists"
                CreateDatabase = True
                Exit Function
            Else
                Debug.Print "Failed to create database: " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            conn.Close
        Else
            Debug.Print "Failed to connect to postgres database with driver " & driverOptions(i) & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' If we get here, all attempts failed
    Debug.Print "All database creation attempts failed"
    Set conn = Nothing
    MsgBox "Could not create database '" & DB_NAME & "'." & vbCrLf & _
           "Please ensure:" & vbCrLf & _
           "1. PostgreSQL server is running" & vbCrLf & _
           "2. PostgreSQL ODBC driver is installed" & vbCrLf & _
           "3. User credentials are correct" & vbCrLf & _
           "4. User has permission to create databases", vbCritical, "Database Creation Failed"
    CreateDatabase = False
    Exit Function
    
ErrorHandler:
    Debug.Print "=== ERROR in CreateDatabase ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    MsgBox "Error creating database: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Database Creation Error"
    CreateDatabase = False
End Function


