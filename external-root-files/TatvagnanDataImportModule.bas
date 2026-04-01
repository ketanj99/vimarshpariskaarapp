Option Explicit

' Simple Tatvagnan Import Module (fresh clean version)
' - Asks user once for ACTION (ADD / REMOVE) and PUSHP NO
' - Lets user pick one or many Excel files
' - For each file, finds headers in row 1:
'   NAME, MEMBER, LANGUAGE, MOBILE_NUMBER, PIN_CODE, ADDRESS_LINE1, ADDRESS_LINE2,
'   KENDRACODE, TRANSACTIONDATE, PUSHP, REMARKS, DISTRICT, TALUKO, VILLAGE, GROUP
' - Imports rows to a dedicated PostgreSQL table: vimars.tatvagnan_simple_data

Public Const DB_HOST As String = "localhost"
Public Const DB_PORT As String = "5432"
Public Const DB_NAME As String = "vimarshbooks"
Public Const DB_USER As String = "postgres"
Public Const DB_PASSWORD As String = "Ketan@757399"
Public Const DB_SCHEMA As String = "vimars"

' ===================== PUBLIC ENTRY =====================

Public Sub Tatvagnan_Import()
    TatvagnanForm.Show
End Sub

' Entry for UserForm: pass READY values (no InputBox/MsgBox asked for action/pushp)
Public Sub Tatvagnan_ImportFromUI(ByVal actionType As String, _
                                  ByVal pushpNo As String, _
                                  Optional ByVal fileList As String = "")
    On Error GoTo ErrHandler

    Dim paths() As String
    Dim i As Long
    Dim conn As Object

    actionType = UCase$(Trim$(actionType))
    If actionType <> "ADD" And actionType <> "REMOVE" Then Exit Sub
    If Trim$(pushpNo) = "" Then Exit Sub

    If fileList = "" Then
        fileList = SelectTatvagnanFiles()
        If fileList = "" Then Exit Sub
    End If

    paths = Split(fileList, "|")

    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then Exit Sub
    If Not EnsureTatvagnanSimpleTable(conn) Then
        conn.Close
        Set conn = Nothing
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True

    For i = LBound(paths) To UBound(paths)
        Dim p As String
        p = Trim$(paths(i))
        If Len(p) > 0 Then
            Application.StatusBar = "Tatvagnan import (UI): " & (i - LBound(paths) + 1) & _
                                   " of " & (UBound(paths) - LBound(paths) + 1) & _
                                   " - " & GetFileNameSafe(p)
            DoEvents
            ImportTatvagnanFile p, actionType, pushpNo, conn
        End If
    Next i

    conn.Close
    Set conn = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    Set conn = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' ===================== CORE IMPORT =====================

Private Sub ImportTatvagnanFile(ByVal filePath As String, _
                                ByVal actionType As String, _
                                ByVal pushpNo As String, _
                                ByVal conn As Object)
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerMap As Object
    Dim r As Long

    If Dir$(filePath) = "" Then Exit Sub

    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then GoTo CleanUp   ' no data

    Set headerMap = BuildHeaderMap(ws, 1, lastCol)
    If headerMap Is Nothing Then GoTo CleanUp

    ' Require at least NAME column
    If Not headerMap.Exists("NAME") Then
        MsgBox "Header 'NAME' not found in file: " & GetFileNameSafe(filePath), vbExclamation
        GoTo CleanUp
    End If

    For r = 2 To lastRow
        Dim nameVal As String
        nameVal = GetCellString(ws.Cells(r, headerMap("NAME")))
        If Len(Trim$(nameVal)) = 0 Then GoTo NextRow

        InsertTatvagnanRow ws, r, headerMap, actionType, pushpNo, conn
NextRow:
    Next r

CleanUp:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set ws = Nothing
    Set headerMap = Nothing
    Exit Sub

ErrHandler:
    Resume CleanUp
End Sub

Private Sub InsertTatvagnanRow(ws As Worksheet, ByVal rowIndex As Long, _
                               ByVal headerMap As Object, _
                               ByVal actionType As String, _
                               ByVal pushpNo As String, _
                               ByVal conn As Object)
    On Error GoTo ErrHandler

    Dim vals(1 To 16) As String   ' 2 custom + 14 from sheet
    Dim sql As String

    ' 1–2: custom inputs
    vals(1) = SafeSql(actionType)
    vals(2) = SafeSql(pushpNo)

    ' 3–16: columns from sheet (if header not present -> blank)
    vals(3) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "NAME"))
    vals(4) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "MEMBER"))
    vals(5) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "LANGUAGE"))
    vals(6) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "MOBILE_NUMBER"))
    vals(7) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "PIN_CODE"))
    vals(8) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "ADDRESS_LINE1"))
    vals(9) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "ADDRESS_LINE2"))
    vals(10) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "KENDRACODE"))
    vals(11) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "TRANSACTIONDATE"))
    vals(12) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "PUSHP"))
    vals(13) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "REMARKS"))
    vals(14) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "DISTRICT"))
    vals(15) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "TALUKO"))
    vals(16) = SafeSql(GetByHeader(ws, rowIndex, headerMap, "VILLAGE"))

    ' GROUP column (reserved word in SQL, quote with double-quotes)
    Dim grp As String
    grp = SafeSql(GetByHeader(ws, rowIndex, headerMap, "GROUP"))

    sql = "INSERT INTO " & DB_SCHEMA & ".tatvagnan_simple_data (" & _
          "action_type, pushp_no, " & _
          "name, member, language, mobile_number, pin_code, " & _
          "address_line1, address_line2, kendracode, transactiondate, " & _
          "pushp, remarks, district, taluko, village, ""group"") VALUES (" & _
          "'" & vals(1) & "'," & _
          "'" & vals(2) & "'," & _
          "'" & vals(3) & "'," & _
          "'" & vals(4) & "'," & _
          "'" & vals(5) & "'," & _
          "'" & vals(6) & "'," & _
          "'" & vals(7) & "'," & _
          "'" & vals(8) & "'," & _
          "'" & vals(9) & "'," & _
          "'" & vals(10) & "'," & _
          "'" & vals(11) & "'," & _
          "'" & vals(12) & "'," & _
          "'" & vals(13) & "'," & _
          "'" & vals(14) & "'," & _
          "'" & vals(15) & "'," & _
          "'" & vals(16) & "'," & _
          "'" & grp & "');"

    conn.Execute sql
    Exit Sub

ErrHandler:
    ' Skip bad row, continue
End Sub

' ===================== HEADER / CELL HELPERS =====================

Private Function BuildHeaderMap(ws As Worksheet, ByVal headerRow As Long, ByVal lastCol As Long) As Object
    On Error GoTo ErrHandler

    Dim dict As Object
    Dim c As Long
    Dim rawVal As String

    Set dict = CreateObject("Scripting.Dictionary")

    For c = 1 To lastCol
        rawVal = UCase$(Trim$(GetCellString(ws.Cells(headerRow, c))))
        Select Case rawVal
            Case "NAME": dict("NAME") = c
            Case "MEMBER": dict("MEMBER") = c
            Case "LANGUAGE": dict("LANGUAGE") = c
            Case "MOBILE_NUMBER", "MOBILE NUMBER", "MOBILE": dict("MOBILE_NUMBER") = c
            Case "PIN_CODE", "PIN CODE", "PINCODE": dict("PIN_CODE") = c
            Case "ADDRESS_LINE1", "ADDRESS LINE1", "ADDRESS LINE 1", "ADDRESS1": dict("ADDRESS_LINE1") = c
            Case "ADDRESS_LINE2", "ADDRESS LINE2", "ADDRESS LINE 2", "ADDRESS2": dict("ADDRESS_LINE2") = c
            Case "KENDRACODE", "KENDRA", "KENDRA CODE": dict("KENDRACODE") = c
            Case "TRANSACTIONDATE", "TRANSACTION DATE", "DATE": dict("TRANSACTIONDATE") = c
            Case "PUSHP", "PUSH": dict("PUSHP") = c
            Case "REMARKS": dict("REMARKS") = c
            Case "DISTRICT": dict("DISTRICT") = c
            Case "TALUKO", "TALUK": dict("TALUKO") = c
            Case "VILLAGE": dict("VILLAGE") = c
            Case "GROUP", "GROUP_NAME": dict("GROUP") = c
        End Select
    Next c

    Set BuildHeaderMap = dict
    Exit Function

ErrHandler:
    Set BuildHeaderMap = Nothing
End Function

Private Function GetByHeader(ws As Worksheet, ByVal rowIndex As Long, _
                             ByVal headerMap As Object, _
                             ByVal key As String) As String
    On Error GoTo ErrHandler

    If Not headerMap.Exists(key) Then
        GetByHeader = ""
    Else
        GetByHeader = GetCellString(ws.Cells(rowIndex, headerMap(key)))
    End If
    Exit Function

ErrHandler:
    GetByHeader = ""
End Function

Private Function GetCellString(cell As Range) As String
    On Error Resume Next
    Dim v As Variant
    v = cell.Value
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        GetCellString = ""
    Else
        GetCellString = Trim$(CStr(v))
    End If
End Function

Private Function SafeSql(ByVal s As String) As String
    SafeSql = Replace(s, "'", "''")
End Function

Private Function GetFileNameSafe(ByVal fullPath As String) As String
    On Error Resume Next
    Dim parts() As String
    parts = Split(fullPath, "\")
    If UBound(parts) >= 0 Then
        GetFileNameSafe = parts(UBound(parts))
    Else
        GetFileNameSafe = fullPath
    End If
End Function

' ===================== PG CONNECTION & TABLE =====================

Public Function ConnectToPostgreSQL(conn As Object) As Boolean
    On Error GoTo ErrHandler

    Dim cs As String
    cs = "Driver={PostgreSQL UNICODE};" & _
         "Server=" & DB_HOST & ";" & _
         "Port=" & DB_PORT & ";" & _
         "Database=" & DB_NAME & ";" & _
         "Uid=" & DB_USER & ";" & _
         "Pwd=" & DB_PASSWORD & ";"

    conn.Open cs
    ConnectToPostgreSQL = True
    Exit Function

ErrHandler:
    ConnectToPostgreSQL = False
End Function

Private Function EnsureTatvagnanSimpleTable(conn As Object) As Boolean
    On Error GoTo ErrHandler

    Dim sql As String

    sql = "CREATE SCHEMA IF NOT EXISTS " & DB_SCHEMA & ";"
    conn.Execute sql

    sql = "CREATE TABLE IF NOT EXISTS " & DB_SCHEMA & ".tatvagnan_simple_data (" & _
          "id SERIAL PRIMARY KEY," & _
          "action_type TEXT NOT NULL," & _
          "pushp_no TEXT," & _
          "name TEXT," & _
          "member TEXT," & _
          "language TEXT," & _
          "mobile_number TEXT," & _
          "pin_code TEXT," & _
          "address_line1 TEXT," & _
          "address_line2 TEXT," & _
          "kendracode TEXT," & _
          "transactiondate TEXT," & _
          "pushp TEXT," & _
          "remarks TEXT," & _
          "district TEXT," & _
          "taluko TEXT," & _
          "village TEXT," & _
          """group"" TEXT," & _
          "created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP" & _
          ");"
    conn.Execute sql

    sql = "CREATE INDEX IF NOT EXISTS idx_tatvagnan_simple_action ON " & DB_SCHEMA & ".tatvagnan_simple_data(action_type);"
    conn.Execute sql

    EnsureTatvagnanSimpleTable = True
    Exit Function

ErrHandler:
    EnsureTatvagnanSimpleTable = False
End Function

' ===================== HISTORY LISTVIEW (FOR FORM) =====================

' Fill a ListView-like control with Tatvagnan history
' Call from UserForm code: Tatvagnan_LoadHistory Me.ListView1
Public Sub Tatvagnan_LoadHistory(ByVal target As Object)
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim sql As String

    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then Exit Sub

    ' Summary by group: each (pushp_no + action + group) = one row, own total
    sql = "SELECT action_type, pushp_no, COALESCE(""group"", '') AS grp, MIN(created_at) AS import_date, COUNT(*) AS total_no " & _
          "FROM " & DB_SCHEMA & ".tatvagnan_simple_data " & _
          "GROUP BY action_type, pushp_no, ""group"" " & _
          "ORDER BY MIN(created_at) DESC;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1

    ' Expecting MSComctlLib.ListView or similar – use late binding
    On Error Resume Next
    With target
        .View = 3          ' lvwReport
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , , "Import Date", 120
        .ColumnHeaders.Add , , "Pushp No", 100
        .ColumnHeaders.Add , , "Action", 80
        .ColumnHeaders.Add , , "Group", 100
        .ColumnHeaders.Add , , "Total No.", 70
    End With
    On Error GoTo ErrHandler

    If rs.EOF Then GoTo CleanUp

    Dim itm As Object
    Dim d As Variant
    Do While Not rs.EOF
        d = rs.Fields("import_date").Value
        Set itm = target.ListItems.Add(, , NzStr(d))
        itm.SubItems(1) = NzStr(rs.Fields("pushp_no").Value)
        itm.SubItems(2) = NzStr(rs.Fields("action_type").Value)
        itm.SubItems(3) = NzStr(rs.Fields("grp").Value)
        itm.SubItems(4) = NzStr(rs.Fields("total_no").Value)
        rs.MoveNext
    Loop
    ' Column order: 0=Import Date, 1=Pushp No, 2=Action, 3=Group, 4=Total No.

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

ErrHandler:
    GoTo CleanUp
End Sub

' Delete selected summary row(s) from DB (by Pushp No + Action + Group)
Public Sub Tatvagnan_DeleteSelected(ByVal target As Object)
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim itm As Object
    Dim pushpNo As String
    Dim actionType As String
    Dim grp As String
    Dim sql As String
    Dim n As Long

    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        MsgBox "Database connection failed.", vbCritical
        Exit Sub
    End If

    n = 0
    On Error Resume Next
    For Each itm In target.ListItems
        If itm.Selected Then
            pushpNo = Trim$(itm.SubItems(1))
            actionType = Trim$(itm.SubItems(2))
            grp = Trim$(itm.SubItems(3))
            If Len(pushpNo) > 0 And Len(actionType) > 0 Then
                sql = "DELETE FROM " & DB_SCHEMA & ".tatvagnan_simple_data " & _
                      "WHERE pushp_no = '" & Replace(pushpNo, "'", "''") & "' " & _
                      "AND action_type = '" & Replace(actionType, "'", "''") & "' " & _
                      "AND COALESCE(""group"", '') = '" & Replace(grp, "'", "''") & "';"
                conn.Execute sql
                n = n + 1
            End If
        End If
    Next itm
    On Error GoTo ErrHandler

    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set conn = Nothing

    Tatvagnan_LoadHistory target
    MsgBox IIf(n > 0, "Deleted " & n & " selected batch(es).", "No selection to delete."), vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set conn = Nothing
    MsgBox "Delete error: " & Err.Description, vbCritical
End Sub

' ===================== UI HELPERS =====================

Private Function GetTatvagnanAction() As String
    Dim v As String
    Do
        v = UCase$(Trim$(InputBox("Tatvagnan ACTION (ADD / REMOVE):", "Tatvagnan Import")))
        If v = "" Then
            GetTatvagnanAction = ""
            Exit Function
        End If
        If v = "ADD" Or v = "REMOVE" Then
            GetTatvagnanAction = v
            Exit Function
        End If
        MsgBox "Please enter exactly ADD or REMOVE.", vbExclamation
    Loop
End Function

Private Function SelectTatvagnanFiles() As String
    On Error GoTo ErrHandler

    Dim fd As Object
    Dim i As Long
    Dim result As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Tatvagnan Excel File(s)"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = True
        If .Show <> -1 Then
            SelectTatvagnanFiles = ""
            GoTo CleanUp
        End If
        For i = 1 To .SelectedItems.Count
            If result <> "" Then result = result & "|"
            result = result & .SelectedItems(i)
        Next i
    End With

    SelectTatvagnanFiles = result

CleanUp:
    Set fd = Nothing
    Exit Function

ErrHandler:
    SelectTatvagnanFiles = ""
End Function

' ===================== SHOW TATVAGNAN REPORT FORM =====================

Public Sub ShowTatvagnanReportForm()
    On Error Resume Next
    TatvagnanForm.Show
    If Err.Number <> 0 Then
        MsgBox "Error opening Tatvagnan Report Form: " & Err.Description, vbCritical
    End If
End Sub

