Option Explicit

Private conn As Object
Private rs As Object
Private historyData() As Variant  ' Array to store history data

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== UserForm_Initialize Started ==="
    
    ' Set form properties - Fixed size
    Me.Caption = "Import History - VIMARS"
    Me.Width = 800
    Me.Height = 550
    
    Debug.Print "Form properties set"
    
    ' Set ListView column widths using common function
    SetListViewColumnWidths
    
    Debug.Print "ListView columns configured: " & ListView1.ColumnHeaders.count & " columns"
    
    ' Initialize connection
    Debug.Print "Creating ADODB connection..."
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL first
    Debug.Print "Attempting to connect to PostgreSQL..."
    If Not ConnectToPostgreSQL(conn) Then
        Debug.Print "Failed to connect to PostgreSQL"
        Unload Me
        Exit Sub
    End If
    
    Debug.Print "Successfully connected to PostgreSQL"
    
    ' Initialize APP dropdown (ComboBox2) - create if it doesn't exist
    Debug.Print "Initializing APP dropdown..."
    InitializeAppDropdown
    
    ' Load APP dropdown
    Debug.Print "Loading APP dropdown..."
    LoadAppDropdown
    
    ' Load periods for dropdown
    Debug.Print "Loading periods dropdown..."
    LoadPeriodsDropdown
    
    ' Load import history
    Debug.Print "Loading import history..."
    LoadImportHistory
    
    Debug.Print "=== UserForm_Initialize Completed Successfully ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error initializing form: " & Err.description
    Unload Me
End Sub

Private Sub LoadImportHistory()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadImportHistory Started ==="
    
    ' Check if connection is open
    If conn Is Nothing Or conn.State <> 1 Then
        Debug.Print "Connection check failed - conn is Nothing or State <> 1"
        lblStatus.Caption = "Database connection not available"
        Exit Sub
    End If
    
    Debug.Print "Connection check passed"
    
    Dim sql As String
    Dim row As Integer
    Dim selectedPeriod As String
    Dim selectedApp As String
    Dim listItem As listItem
    
    ' Get selected period from dropdown
    If ComboBox1.ListIndex > 0 Then  ' Skip "All" option
        selectedPeriod = ComboBox1.text
    End If
    
    ' Get selected APP from dropdown
    On Error Resume Next
    If Me.Controls("ComboBox2").ListIndex > 0 Then  ' Skip "All" option
        selectedApp = Me.Controls("ComboBox2").text
    End If
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    ListView1.ListItems.Clear
    
    ' Ensure ListView columns are properly set up
    SetupListViewColumns
    
    ' Get import history with period and APP filter (including APP column)
    sql = "SELECT history_id, file_name, import_date, period, COALESCE(app, 'VIMARSH') as app, total_book_qty, status " & _
          "FROM vimars.import_history " & _
          "WHERE is_deleted = FALSE "
    
    ' Add period filter if selected
    If selectedPeriod <> "" Then
        sql = sql & "AND period = '" & Replace(selectedPeriod, "'", "''") & "' "
    End If
    
    ' Add APP filter if selected
    If selectedApp <> "" Then
        sql = sql & "AND COALESCE(app, 'VIMARSH') = '" & Replace(selectedApp, "'", "''") & "' "
    End If
    
    sql = sql & "ORDER BY import_date DESC;"
    
    Debug.Print "Executing SQL: " & sql
    Set rs = conn.Execute(sql)
    Debug.Print "SQL executed successfully"
    
    If rs.EOF Then
        Debug.Print "No records found in result set"
        lblStatus.Caption = "No import history found"
        Exit Sub
    End If
    
    Debug.Print "Records found, starting to process data"
    
    ' Count records first
    Dim recordCount As Integer
    recordCount = 0
    Debug.Print "Starting to count records..."
    Do While Not rs.EOF
        recordCount = recordCount + 1
        rs.MoveNext
    Loop
    
    Debug.Print "Found " & recordCount & " records"
    
    ' Resize array (including APP column)
    ReDim historyData(recordCount - 1, 6)  ' 7 columns: ID, File, Date, Period, APP, Qty, Status
    Debug.Print "Array resized for " & recordCount & " records"
    
    ' Reset recordset
    rs.MoveFirst
    Debug.Print "Reset recordset to first record"
    
    ' Add data to ListView and array
    row = 0
    Debug.Print "Starting to process data rows..."
    Do While Not rs.EOF
        Debug.Print "Processing row " & (row + 1) & " of " & recordCount
        
        ' Store in array (including APP)
        historyData(row, 0) = rs.Fields("history_id").Value
        historyData(row, 1) = rs.Fields("file_name").Value
        historyData(row, 2) = rs.Fields("import_date").Value
        historyData(row, 3) = rs.Fields("period").Value
        historyData(row, 4) = rs.Fields("app").Value
        historyData(row, 5) = rs.Fields("total_book_qty").Value
        historyData(row, 6) = rs.Fields("status").Value
        
        ' Create formatted display text for ListView
        Dim importDate As Date
        importDate = rs.Fields("import_date").Value
        Dim idStr As String, fileNameStr As String, dateStr As String, periodStr As String, appStr As String, qtyStr As String, statusStr As String
        
        idStr = CStr(rs.Fields("history_id").Value)
        fileNameStr = CStr(rs.Fields("file_name").Value)
        dateStr = CStr(Day(importDate)) & "-" & CStr(month(importDate)) & "-" & CStr(year(importDate))
        periodStr = CStr(rs.Fields("period").Value)
        appStr = CStr(rs.Fields("app").Value)
        qtyStr = CStr(rs.Fields("total_book_qty").Value)
        statusStr = CStr(rs.Fields("status").Value)
        
        ' Add to ListView (including APP column)
        Set listItem = ListView1.ListItems.Add(, , idStr)
        listItem.SubItems(1) = fileNameStr
        listItem.SubItems(2) = dateStr
        listItem.SubItems(3) = periodStr
        listItem.SubItems(4) = appStr
        listItem.SubItems(5) = qtyStr
        listItem.SubItems(6) = statusStr
        
        ' Verify all columns are populated
        Debug.Print "Row " & (row + 1) & " - Columns: " & _
                   "ID=" & listItem.text & ", " & _
                   "File=" & listItem.SubItems(1) & ", " & _
                   "Date=" & listItem.SubItems(2) & ", " & _
                   "Period=" & listItem.SubItems(3) & ", " & _
                   "APP=" & listItem.SubItems(4) & ", " & _
                   "Qty=" & listItem.SubItems(5) & ", " & _
                   "Status=" & listItem.SubItems(6)
        
        Debug.Print "Added item to ListView: " & idStr & " - " & Left(fileNameStr, 30) & "..."
        
        row = row + 1
        rs.MoveNext
    Loop
    
    Debug.Print "Completed processing " & row & " rows"
    
    ' Update status
    lblStatus.Caption = "Total Records: " & recordCount
    Debug.Print "Updated status label"
    
    ' Check ListView properties
    Debug.Print "ListView ListItems.Count: " & ListView1.ListItems.count
    Debug.Print "ListView Visible: " & ListView1.Visible
    Debug.Print "ListView Enabled: " & ListView1.Enabled
    
    ' Check ComboBox properties
    Debug.Print "ComboBox1 (Period) ListCount: " & ComboBox1.ListCount
    On Error Resume Next
    Debug.Print "ComboBox2 (APP) ListCount: " & Me.Controls("ComboBox2").ListCount
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadImportHistory Completed Successfully ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in LoadImportHistory ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    Debug.Print "Error Line: " & Erl
    lblStatus.Caption = "Error loading import history"
End Sub

' Initialize APP dropdown control (ComboBox2)
Private Sub InitializeAppDropdown()
    On Error Resume Next
    Dim cboApp As MSForms.ComboBox
    Set cboApp = Me.Controls("ComboBox2")
    On Error GoTo 0
    
    If cboApp Is Nothing Then
        ' Try to add the control programmatically
        On Error Resume Next
        Set cboApp = Me.Controls.Add("Forms.ComboBox.1", "ComboBox2", True)
        If Not cboApp Is Nothing Then
            ' Position it near ComboBox1 (adjust position as needed)
            With cboApp
                .Left = ComboBox1.Left + ComboBox1.Width + 10  ' Next to ComboBox1
                .Top = ComboBox1.Top
                .Width = ComboBox1.Width
                .Height = ComboBox1.Height
            End With
        End If
        On Error GoTo 0
    End If
End Sub

' Load APP values for dropdown
Private Sub LoadAppDropdown()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadAppDropdown Started ==="
    
    ' Check if connection is open
    If conn Is Nothing Or conn.State <> 1 Then
        Debug.Print "Connection check failed in LoadAppDropdown"
        lblStatus.Caption = "Database connection not available"
        Exit Sub
    End If
    
    Debug.Print "Connection check passed in LoadAppDropdown"
    
    Dim sql As String
    Dim appRs As Object
    
    ' Clear dropdown (use Me.Controls for safety)
    On Error Resume Next
    Me.Controls("ComboBox2").Clear
    On Error GoTo ErrorHandler
    Debug.Print "ComboBox2 (APP) cleared"
    
    ' Add "All" option as default
    On Error Resume Next
    Me.Controls("ComboBox2").AddItem "All"
    Me.Controls("ComboBox2").ListIndex = 0  ' Select "All" by default
    On Error GoTo ErrorHandler
    Debug.Print "Added 'All' option to APP dropdown"
    
    ' Get unique APP values from database
    sql = "SELECT DISTINCT COALESCE(app, 'VIMARSH') as app FROM vimars.import_history " & _
          "WHERE is_deleted = FALSE AND app IS NOT NULL " & _
          "ORDER BY app;"
    
    Debug.Print "Executing SQL: " & sql
    Set appRs = conn.Execute(sql)
    
    ' Add APP values to dropdown (VIMARSH first if available)
    Dim appCount As Integer
    Dim hasVimarsh As Boolean
    hasVimarsh = False
    appCount = 0
    
    ' First pass: Find and add VIMARSH first
    appRs.MoveFirst
    Do While Not appRs.EOF
        If UCase(Trim(appRs.Fields("app").Value)) = "VIMARSH" Then
            On Error Resume Next
            Me.Controls("ComboBox2").AddItem appRs.Fields("app").Value
            On Error GoTo ErrorHandler
            hasVimarsh = True
            appCount = appCount + 1
            Debug.Print "Added APP to ComboBox2: " & appRs.Fields("app").Value & " (first)"
            Exit Do
        End If
        appRs.MoveNext
    Loop
    
    ' Second pass: Add all other APP values
    appRs.MoveFirst
    Do While Not appRs.EOF
        If UCase(Trim(appRs.Fields("app").Value)) <> "VIMARSH" Then
            On Error Resume Next
            Me.Controls("ComboBox2").AddItem appRs.Fields("app").Value
            On Error GoTo ErrorHandler
            appCount = appCount + 1
            Debug.Print "Added APP to ComboBox2: " & appRs.Fields("app").Value
        End If
        appRs.MoveNext
    Loop
    
    Debug.Print "Added " & appCount & " APP values to dropdown"
    
    appRs.Close
    Set appRs = Nothing
    
    ' If no APP values found, add default
    If appCount = 0 Then
        On Error Resume Next
        Me.Controls("ComboBox2").AddItem "VIMARSH"
        Me.Controls("ComboBox2").AddItem "PARISHKAAR"
        On Error GoTo ErrorHandler
        Debug.Print "Added default APP values: VIMARSH, PARISHKAAR"
    End If
    
    Debug.Print "=== LoadAppDropdown Completed Successfully ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in LoadAppDropdown ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    ' Add default values on error
    On Error Resume Next
    Me.Controls("ComboBox2").AddItem "VIMARSH"
    Me.Controls("ComboBox2").AddItem "PARISHKAAR"
    On Error GoTo 0
    lblStatus.Caption = "Error loading APP dropdown"
End Sub

' Load periods for dropdown
Private Sub LoadPeriodsDropdown()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadPeriodsDropdown Started ==="
    
    ' Check if connection is open
    If conn Is Nothing Or conn.State <> 1 Then
        Debug.Print "Connection check failed in LoadPeriodsDropdown"
        lblStatus.Caption = "Database connection not available"
        Exit Sub
    End If
    
    Debug.Print "Connection check passed in LoadPeriodsDropdown"
    
    Dim sql As String
    Dim periodRs As Object
    
    ' Clear dropdown
    ComboBox1.Clear
    Debug.Print "ComboBox1 (Period) cleared"
    
    ' Add "All" option as default
    ComboBox1.AddItem "All"
    ComboBox1.ListIndex = 0  ' Select "All" by default
    Debug.Print "Added 'All' option to period dropdown"
    
    ' Get unique periods from database
    sql = "SELECT DISTINCT period FROM vimars.import_history " & _
          "WHERE is_deleted = FALSE AND period IS NOT NULL AND period != '' " & _
          "ORDER BY period;"
    
    Debug.Print "Executing SQL: " & sql
    Set periodRs = conn.Execute(sql)
    
    ' Add periods to dropdown
    Dim periodCount As Integer
    periodCount = 0
    Do While Not periodRs.EOF
        Dim periodValue As String
        periodValue = periodRs.Fields("period").Value
        ComboBox1.AddItem periodValue
        Debug.Print "Added period to ComboBox1: " & periodValue
        periodCount = periodCount + 1
        periodRs.MoveNext
    Loop
    
    Debug.Print "Added " & periodCount & " periods to dropdown"
    
    periodRs.Close
    Set periodRs = Nothing
    
    Debug.Print "=== LoadPeriodsDropdown Completed Successfully ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in LoadPeriodsDropdown ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    lblStatus.Caption = "Error loading periods"
End Sub

' Period dropdown change event
Private Sub ComboBox1_Change()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== ComboBox1_Change Started ==="
    Debug.Print "Selected period: " & ComboBox1.text
    Debug.Print "ListIndex: " & ComboBox1.ListIndex
    
    ' Reload data when period selection changes
    LoadImportHistory
    
    Debug.Print "=== ComboBox1_Change Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in ComboBox1_Change ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
End Sub

' Note: ComboBox2_Change event handler cannot be defined if ComboBox2 doesn't exist in form
' To handle APP change, add ComboBox2 to form manually OR refresh manually using btnRefresh button
' If you want automatic refresh, add ComboBox2 control to the form in VBA form designer first

Private Sub btnDelete_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedIndices() As Integer
    Dim selectedCount As Integer
    Dim i As Integer
    Dim historyId As Long
    Dim result As VbMsgBoxResult
    Dim deleteCount As Integer
    Dim sql As String
    
    ' Get selected items
    selectedCount = GetSelectedItems(selectedIndices)
    
    ' Check if any item is selected
    If selectedCount = 0 Then
        lblStatus.Caption = "Please select records to delete"
        Exit Sub
    End If
    
    ' Confirm deletion
    Dim confirmMsg As String
    If selectedCount = 1 Then
        confirmMsg = "Are you sure you want to delete this import history and all related transactions?" & vbCrLf & _
                   "History ID: " & historyData(selectedIndices(0), 0) & vbCrLf & _
                   "File: " & historyData(selectedIndices(0), 1) & vbCrLf & _
                   "APP: " & historyData(selectedIndices(0), 4) & vbCrLf & _
                   "This action cannot be undone!"
    Else
        confirmMsg = "Are you sure you want to delete " & selectedCount & " import history records and all related transactions?" & vbCrLf & _
                   "This action cannot be undone!"
    End If
    
    result = MsgBox(confirmMsg, vbYesNo + vbCritical, "Confirm Deletion")
    
    If result <> vbYes Then
        Exit Sub
    End If
    
    ' Delete selected records from database
    deleteCount = 0
    For i = 0 To selectedCount - 1
        historyId = historyData(selectedIndices(i), 0)
        sql = "DELETE FROM vimars.import_history WHERE history_id = " & historyId & ";"
        conn.Execute sql
        deleteCount = deleteCount + 1
    Next i
    
    ' Reload data
    LoadImportHistory
    
    ' Update status label
    If deleteCount = 1 Then
        lblStatus.Caption = "1 record deleted successfully"
    Else
        lblStatus.Caption = deleteCount & " records deleted successfully"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error deleting record(s): " & Err.description
    lblStatus.Caption = "Error deleting records"
End Sub

Private Sub btnRefresh_Click()
    LoadImportHistory
End Sub

' Test button to verify ListView columns
Private Sub btnTestListView_Click()
    TestListViewDisplay
End Sub

' Force display all columns button
Private Sub btnForceDisplay_Click()
    ForceDisplayAllColumns
End Sub

' Quick fix button for ListView columns
Private Sub btnFixColumns_Click()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== btnFixColumns_Click Started ==="
    
    ' Clear existing data
    ListView1.ListItems.Clear
    
    ' Use SetListViewColumnWidths function for consistent column widths
    SetListViewColumnWidths
    
    ' Add test data (including APP)
    Dim testItem As listItem
    Set testItem = ListView1.ListItems.Add(, , "1")
    testItem.SubItems(1) = "Test.xlsx"
    testItem.SubItems(2) = "01-01-2024"
    testItem.SubItems(3) = "January"
    testItem.SubItems(4) = "VIMARSH"
    testItem.SubItems(5) = "100"
    testItem.SubItems(6) = "SUCCESS"
    
    lblStatus.Caption = "Columns fixed - " & ListView1.ColumnHeaders.count & " columns, " & ListView1.ListItems.count & " items"
    
    Debug.Print "=== btnFixColumns_Click Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in btnFixColumns_Click ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    lblStatus.Caption = "Error fixing columns"
End Sub

' Test ListView functionality
Private Sub TestListViewDisplay()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== TestListViewDisplay Started ==="
    
    ' Clear existing data
    ListView1.ListItems.Clear
    
    ' Setup columns properly
    SetupListViewColumns
    
    ' Verify columns are set up
    Debug.Print "Column count: " & ListView1.ColumnHeaders.count
    
    ' Add test data with all columns (including APP)
    Dim testItem As listItem
    
    ' Test row 1
    Set testItem = ListView1.ListItems.Add(, , "TEST001")
    testItem.SubItems(1) = "TestFile1.xlsx"
    testItem.SubItems(2) = "01-01-2024"
    testItem.SubItems(3) = "January"
    testItem.SubItems(4) = "VIMARSH"
    testItem.SubItems(5) = "150"
    testItem.SubItems(6) = "Completed"
    
    ' Test row 2
    Set testItem = ListView1.ListItems.Add(, , "TEST002")
    testItem.SubItems(1) = "TestFile2.xlsx"
    testItem.SubItems(2) = "02-01-2024"
    testItem.SubItems(3) = "February"
    testItem.SubItems(4) = "PARISHKAAR"
    testItem.SubItems(5) = "200"
    testItem.SubItems(6) = "Pending"
    
    ' Test row 3
    Set testItem = ListView1.ListItems.Add(, , "TEST003")
    testItem.SubItems(1) = "TestFile3.xlsx"
    testItem.SubItems(2) = "03-01-2024"
    testItem.SubItems(3) = "March"
    testItem.SubItems(4) = "VIMARSH"
    testItem.SubItems(5) = "300"
    testItem.SubItems(6) = "Failed"
    
    Debug.Print "Added test data: " & ListView1.ListItems.count & " items"
    
    ' Verify each row has all columns
    Dim i As Integer
    For i = 1 To ListView1.ListItems.count
        Debug.Print "Row " & i & " - ID: " & ListView1.ListItems(i).text & _
                   ", File: " & ListView1.ListItems(i).SubItems(1) & _
                   ", Date: " & ListView1.ListItems(i).SubItems(2) & _
                   ", Period: " & ListView1.ListItems(i).SubItems(3) & _
                   ", APP: " & ListView1.ListItems(i).SubItems(4) & _
                   ", Qty: " & ListView1.ListItems(i).SubItems(5) & _
                   ", Status: " & ListView1.ListItems(i).SubItems(6)
    Next i
    
    ' Update status
    lblStatus.Caption = "Test Data: " & ListView1.ListItems.count & " records"
    
    ' Force refresh
    ListView1.Refresh
    DoEvents
    
    lblStatus.Caption = "Test data added - " & ListView1.ColumnHeaders.count & " columns, " & ListView1.ListItems.count & " items"
    
    Debug.Print "=== TestListViewDisplay Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in TestListViewDisplay ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    lblStatus.Caption = "Error testing ListView"
End Sub

Private Sub btnSelectAll_Click()
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim maxItems As Integer
    
    maxItems = ListView1.ListItems.count
    
    ' Select all items
    For i = 1 To maxItems
        ListView1.ListItems(i).Selected = True
    Next i
    
    ' Update status
    lblStatus.Caption = "Selected: " & maxItems & " of " & maxItems & " records"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error selecting all items: " & Err.description
    lblStatus.Caption = "Error selecting items"
End Sub

Private Sub btnClearSelection_Click()
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim maxItems As Integer
    
    maxItems = ListView1.ListItems.count
    
    ' Clear all selections
    For i = 1 To maxItems
        ListView1.ListItems(i).Selected = False
    Next i
    
    ' Update status
    lblStatus.Caption = "No records selected"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error clearing selection: " & Err.description
    lblStatus.Caption = "Error clearing selection"
End Sub

Private Sub btnClose_Click()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== btnClose_Click Started ==="
    
    ' Cleanup before closing
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
        Debug.Print "Recordset closed"
    End If
    
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
        Debug.Print "Connection closed"
    End If
    
    Debug.Print "Unloading form..."
    Unload Me
    
    Debug.Print "=== btnClose_Click Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in btnClose_Click ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    ' Try to unload anyway
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    ' Cleanup
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

' Connect to PostgreSQL (copied from PostgreSQLImportModule)
Private Function ConnectToPostgreSQL(conn As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim connectionString As String
    Dim password As String
    Dim driverOptions() As String
    Dim i As Integer
    
    ' Use hardcoded password
    password = "Ketan@757399"
    
    ' Define multiple driver options to try
    ReDim driverOptions(4)
    driverOptions(0) = "PostgreSQL ODBC Driver(UNICODE)"
    driverOptions(1) = "PostgreSQL UNICODE"
    driverOptions(2) = "PostgreSQL ODBC Driver"
    driverOptions(3) = "PostgreSQL"
    driverOptions(4) = "PostgreSQL ANSI"
    
    ' Try each driver option
    For i = 0 To 4
        ' Build connection string
        connectionString = "Driver={" & driverOptions(i) & "};" & _
                          "Server=localhost;" & _
                          "Port=5432;" & _
                          "Database=vimarshbooks;" & _
                          "Uid=postgres;" & _
                          "Pwd=" & password & ";"
        
        ' Try to open connection
        On Error Resume Next
        conn.Open connectionString
        If Err.Number = 0 Then
            ConnectToPostgreSQL = True
            Exit Function
        Else
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' If we get here, all drivers failed
    Debug.Print "Could not connect to PostgreSQL with any available driver."
    ConnectToPostgreSQL = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Error connecting to PostgreSQL: " & Err.description
    ConnectToPostgreSQL = False
End Function

' Get selected items from ListView
Private Function GetSelectedItems(ByRef selectedIndices() As Integer) As Integer
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim count As Integer
    Dim maxItems As Integer
    
    ' Count selected items first
    count = 0
    maxItems = ListView1.ListItems.count
    
    For i = 1 To maxItems
        If ListView1.ListItems(i).Selected Then
            count = count + 1
        End If
    Next i
    
    ' Resize array
    If count > 0 Then
        ReDim selectedIndices(count - 1)
        
        ' Fill array with selected indices
        count = 0
        For i = 1 To maxItems
            If ListView1.ListItems(i).Selected Then
                selectedIndices(count) = i - 1  ' Convert to 0-based index for historyData array
                count = count + 1
            End If
        Next i
    Else
        ReDim selectedIndices(0)
    End If
    
    GetSelectedItems = count
    Exit Function
    
ErrorHandler:
    GetSelectedItems = 0
End Function

' Common function to set ListView column widths - used everywhere (including APP column)
Private Sub SetListViewColumnWidths()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== SetListViewColumnWidths Started ==="
    
    ' Clear existing columns first
    ListView1.ColumnHeaders.Clear
    
    ' Set ListView properties
    ListView1.View = lvwReport
    ListView1.FullRowSelect = True
    ListView1.Gridlines = True
    ListView1.MultiSelect = True
    
    ' Add columns with perfect widths - ALL COLUMNS INCLUDING APP
    ' Period column width increased to accommodate full month format like "2025-DEC (1 To 31)"
    ListView1.ColumnHeaders.Add , "ID", "ID", 60
    ListView1.ColumnHeaders.Add , "FileName", "File Name", 200
    ListView1.ColumnHeaders.Add , "ImportDate", "Import Date", 80
    ListView1.ColumnHeaders.Add , "Period", "Period", 180
    ListView1.ColumnHeaders.Add , "APP", "APP", 80
    ListView1.ColumnHeaders.Add , "TotalQty", "Total Qty", 80
    ListView1.ColumnHeaders.Add , "Status", "Status", 80
    
    Debug.Print "ListView columns configured: " & ListView1.ColumnHeaders.count & " columns (including APP)"
    
    ' Force refresh
    ListView1.Refresh
    DoEvents
    
    Debug.Print "=== SetListViewColumnWidths Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in SetListViewColumnWidths ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    lblStatus.Caption = "Error setting column widths"
End Sub

' Ensure ListView columns are properly set up
Private Sub SetupListViewColumns()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== SetupListViewColumns Started ==="
    
    ' Use common function to set column widths
    SetListViewColumnWidths
    
    Debug.Print "=== SetupListViewColumns Completed Successfully ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in SetupListViewColumns ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    lblStatus.Caption = "Error setting up columns"
End Sub

' Force display all columns
Private Sub ForceDisplayAllColumns()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== ForceDisplayAllColumns Started ==="
    
    ' Clear and reset columns
    SetupListViewColumns
    
    ' Refresh ListView
    ListView1.Refresh
    DoEvents
    
    lblStatus.Caption = "All columns displayed - " & ListView1.ColumnHeaders.count & " columns"
    
    Debug.Print "=== ForceDisplayAllColumns Completed ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "=== ERROR in ForceDisplayAllColumns ==="
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    lblStatus.Caption = "Error displaying columns"
End Sub
