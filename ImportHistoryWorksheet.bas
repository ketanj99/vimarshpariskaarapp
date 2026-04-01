Option Explicit

' Show import history in Excel worksheet (Better than ListView)
Public Sub ShowImportHistoryInWorksheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sql As String
    Dim rs As Object
    Dim conn As Object
    Dim row As Integer
    Dim lastRow As Long
    
    ' Create new worksheet or use existing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Import History")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Import History"
    End If
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    ws.Cells.Clear
    
    ' Connect to database
    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        MsgBox "Failed to connect to PostgreSQL", vbCritical
        Exit Sub
    End If
    
    ' Get import history data (including APP column)
    sql = "SELECT history_id, file_name, import_date, period, COALESCE(app, 'VIMARSH') as app, total_book_qty, status " & _
          "FROM vimars.import_history " & _
          "WHERE is_deleted = FALSE " & _
          "ORDER BY import_date DESC;"
    
    Set rs = conn.Execute(sql)
    
    If rs.EOF Then
        MsgBox "No import history found.", vbInformation
        GoTo Cleanup
    End If
    
    ' Add title
    ws.Cells(1, 1).Value = "VIMARS - Import History"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Interior.Color = RGB(70, 130, 180)
    ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    ws.Range("A1:G1").Merge
    ws.Range("A1:G1").HorizontalAlignment = xlCenter
    
    ' Add headers (including APP column)
    ws.Cells(3, 1).Value = "ID"
    ws.Cells(3, 2).Value = "File Name"
    ws.Cells(3, 3).Value = "Import Date"
    ws.Cells(3, 4).Value = "Period"
    ws.Cells(3, 5).Value = "APP"
    ws.Cells(3, 6).Value = "Total Qty"
    ws.Cells(3, 7).Value = "Status"
    
    ' Format headers
    ws.Range("A3:G3").Font.Bold = True
    ws.Range("A3:G3").Interior.Color = RGB(220, 220, 220)
    ws.Range("A3:G3").HorizontalAlignment = xlCenter
    ws.Range("A3:G3").Borders.LineStyle = xlContinuous
    
    ' Add data (including APP column)
    row = 4
    Do While Not rs.EOF
        ws.Cells(row, 1).Value = rs.Fields("history_id").Value
        ws.Cells(row, 2).Value = rs.Fields("file_name").Value
        ws.Cells(row, 3).Value = Format(rs.Fields("import_date").Value, "dd-mm-yyyy")
        ws.Cells(row, 4).Value = rs.Fields("period").Value
        ws.Cells(row, 5).Value = rs.Fields("app").Value
        ws.Cells(row, 6).Value = rs.Fields("total_book_qty").Value
        ws.Cells(row, 7).Value = rs.Fields("status").Value
        
        ' Format data row
        ws.Range("A" & row & ":G" & row).Borders.LineStyle = xlContinuous
        ws.Range("A" & row & ":G" & row).HorizontalAlignment = xlCenter
        
        ' Alternate row colors
        If row Mod 2 = 0 Then
            ws.Range("A" & row & ":G" & row).Interior.Color = RGB(245, 245, 245)
        End If
        
        row = row + 1
        rs.MoveNext
    Loop
    
    ' Auto-fit columns
    ws.Columns("A:G").AutoFit
    
    ' Add borders to entire table
    lastRow = row - 1
    ws.Range("A3:G" & lastRow).Borders.LineStyle = xlContinuous
    ws.Range("A3:G" & lastRow).Borders.Weight = xlThin
    
    ' Add summary
    ws.Cells(lastRow + 2, 1).Value = "Total Records: " & (lastRow - 3)
    ws.Cells(lastRow + 2, 1).Font.Bold = True
    
    ' Add instructions
    ws.Cells(lastRow + 4, 1).Value = "Instructions:"
    ws.Cells(lastRow + 4, 1).Font.Bold = True
    ws.Cells(lastRow + 5, 1).Value = "• To delete a record, select the row and run 'DeleteSelectedHistory' macro"
    ws.Cells(lastRow + 6, 1).Value = "• To refresh data, run 'RefreshImportHistory' macro"
    
    ' Activate worksheet
    ws.Activate
    
    ' Select first data cell
    ws.Cells(4, 1).Select
    
    MsgBox "Import history displayed in worksheet!" & vbCrLf & _
           "Total records: " & (lastRow - 3) & vbCrLf & _
           "You can now view, sort, and filter the data.", vbInformation
    
Cleanup:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing import history: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub

' Delete selected history record
Public Sub DeleteSelectedHistory()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim historyId As Long
    Dim fileName As String
    Dim result As VbMsgBoxResult
    Dim conn As Object
    Dim sql As String
    
    ' Get active worksheet
    Set ws = ActiveSheet
    If ws.Name <> "Import History" Then
        MsgBox "Please select a row in the Import History worksheet.", vbExclamation
        Exit Sub
    End If
    
    ' Get selected row
    selectedRow = Selection.Row
    If selectedRow < 4 Then
        MsgBox "Please select a data row (not header).", vbExclamation
        Exit Sub
    End If
    
    ' Get data from selected row
    historyId = ws.Cells(selectedRow, 1).Value
    fileName = ws.Cells(selectedRow, 2).Value
    
    If historyId = "" Then
        MsgBox "Please select a valid data row.", vbExclamation
        Exit Sub
    End If
    
    ' Confirm deletion
    result = MsgBox("Are you sure you want to delete this import history and all related transactions?" & vbCrLf & _
                   "History ID: " & historyId & vbCrLf & _
                   "File: " & fileName & vbCrLf & _
                   "This action cannot be undone!", vbYesNo + vbCritical, "Confirm Deletion")
    
    If result <> vbYes Then
        Exit Sub
    End If
    
    ' Connect to database and delete
    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        MsgBox "Failed to connect to PostgreSQL", vbCritical
        Exit Sub
    End If
    
    sql = "DELETE FROM vimars.import_history WHERE history_id = " & historyId & ";"
    conn.Execute sql
    
    ' Refresh the worksheet
    ShowImportHistoryInWorksheet
    
    MsgBox "Import history and all related transactions deleted successfully!", vbInformation
    
Cleanup:
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error deleting record: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub

' Refresh import history worksheet
Public Sub RefreshImportHistory()
    ShowImportHistoryInWorksheet
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
    MsgBox "Could not connect to PostgreSQL with any available driver.", vbCritical
    ConnectToPostgreSQL = False
    Exit Function
    
ErrorHandler:
    MsgBox "Error connecting to PostgreSQL: " & Err.Description, vbCritical
    ConnectToPostgreSQL = False
End Function
