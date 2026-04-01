Option Explicit

' Tatvagnan report / PDF generation module
' Uses DB_* constants and ConnectToPostgreSQL from TatvagnanDataImportModule

Public Function NzStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Public Function LangToCode(ByVal lang As String) As String
    Dim u As String
    u = UCase$(Trim$(lang))
    Select Case u
        Case "GUJARATI", "GUJRATI", "GUJ"
            LangToCode = "GJ"
        Case "HINDI", "HIN"
            LangToCode = "HN"
        Case "MARATHI", "MARATHI ", "MAR"
            LangToCode = "MH"
        Case Else
            If Len(u) >= 2 Then
                LangToCode = Left$(u, 2)
            Else
                LangToCode = u
            End If
    End Select
End Function

' Natural sort function to properly sort alphanumeric strings (E1, E2, ..., E9, E10, E11)
Public Sub NaturalSortArray(ByRef arr() As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim swapped As Boolean
    
    If UBound(arr) <= LBound(arr) Then Exit Sub
    
    ' Bubble sort with natural comparison
    For i = LBound(arr) To UBound(arr) - 1
        swapped = False
        For j = LBound(arr) To UBound(arr) - i - 1
            If NaturalCompare(CStr(arr(j)), CStr(arr(j + 1))) > 0 Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i
End Sub

' Compare two strings naturally (numbers within strings are compared numerically)
Private Function NaturalCompare(ByVal s1 As String, ByVal s2 As String) As Integer
    Dim i1 As Long, i2 As Long
    Dim ch1 As String, ch2 As String
    Dim num1 As Long, num2 As Long
    Dim isNum1 As Boolean, isNum2 As Boolean
    
    i1 = 1: i2 = 1
    
    Do While i1 <= Len(s1) Or i2 <= Len(s2)
        If i1 > Len(s1) Then
            NaturalCompare = -1
            Exit Function
        End If
        If i2 > Len(s2) Then
            NaturalCompare = 1
            Exit Function
        End If
        
        ch1 = Mid$(s1, i1, 1)
        ch2 = Mid$(s2, i2, 1)
        
        isNum1 = (ch1 >= "0" And ch1 <= "9")
        isNum2 = (ch2 >= "0" And ch2 <= "9")
        
        If isNum1 And isNum2 Then
            ' Both are numbers - extract full number
            num1 = 0
            Do While i1 <= Len(s1) And Mid$(s1, i1, 1) >= "0" And Mid$(s1, i1, 1) <= "9"
                num1 = num1 * 10 + CLng(Mid$(s1, i1, 1))
                i1 = i1 + 1
            Loop
            
            num2 = 0
            Do While i2 <= Len(s2) And Mid$(s2, i2, 1) >= "0" And Mid$(s2, i2, 1) <= "9"
                num2 = num2 * 10 + CLng(Mid$(s2, i2, 1))
                i2 = i2 + 1
            Loop
            
            If num1 < num2 Then
                NaturalCompare = -1
                Exit Function
            ElseIf num1 > num2 Then
                NaturalCompare = 1
                Exit Function
            End If
        Else
            ' Character comparison
            If ch1 < ch2 Then
                NaturalCompare = -1
                Exit Function
            ElseIf ch1 > ch2 Then
                NaturalCompare = 1
                Exit Function
            End If
            i1 = i1 + 1
            i2 = i2 + 1
        End If
    Loop
    
    NaturalCompare = 0
End Function

' ===================== SUMMARY PDF (TOTAL) =====================

Public Sub Tatvagnan_GenerateSummaryPdf(Optional ByVal pushpNo As String = "")
    On Error GoTo ErrHandler

    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dictGroups As Object
    Dim dictLangs As Object
    Dim dictOld As Object
    Dim dictNew As Object
    Dim dictRmv As Object
    Dim key As String
    Dim grp As String, lang As String
    Dim code As String

    If Trim$(pushpNo) = "" Then
        pushpNo = InputBox("Enter Tatvagnan Pushp No. for summary PDF:", "Tatvagnan Summary PDF")
        If Trim$(pushpNo) = "" Then Exit Sub
    End If

    Set conn = CreateObject("ADODB.Connection")
    If Not TatvagnanDataimportModule.ConnectToPostgreSQL(conn) Then
        MsgBox "Database connection failed.", vbCritical
        Exit Sub
    End If

    ' Dictionaries
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictLangs = CreateObject("Scripting.Dictionary")
    Dim dictLangNames As Object  ' Maps code to full language name
    Set dictLangNames = CreateObject("Scripting.Dictionary")
    Set dictOld = CreateObject("Scripting.Dictionary")
    Set dictNew = CreateObject("Scripting.Dictionary")
    Set dictRmv = CreateObject("Scripting.Dictionary")

    ' 1) OLD = all previous pushp_no < current (member+lang+mobile pair cancel logic)
    sql = "WITH base AS (" & _
          "SELECT UPPER(TRIM(COALESCE(""group"", ''))) AS grp, UPPER(TRIM(COALESCE(language, ''))) AS lang, " & _
          "UPPER(TRIM(COALESCE(member, ''))) AS member_name, REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile, " & _
          "UPPER(TRIM(COALESCE(village, ''))) AS village_name, action_type " & _
          "FROM " & DB_SCHEMA & ".tatvagnan_simple_data " & _
          "WHERE pushp_no::integer < " & Replace(pushpNo, "'", "''") & "), " & _
          "matched AS (" & _
          "SELECT grp, lang, member_name, mobile, village_name, " & _
          "SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt, " & _
          "SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt " & _
          "FROM base GROUP BY grp, lang, member_name, mobile, village_name), " & _
          "net AS (" & _
          "SELECT grp, lang, GREATEST(add_cnt - rmv_cnt, 0) AS net_add, GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv " & _
          "FROM matched) " & _
          "SELECT grp, lang, SUM(net_add) AS add_cnt, SUM(net_rmv) AS rmv_cnt " & _
          "FROM net GROUP BY grp, lang ORDER BY grp, lang;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1
    If Not rs.EOF Then
        Do While Not rs.EOF
            grp = NzStr(rs.Fields("grp").Value)
            lang = NzStr(rs.Fields("lang").Value)
            code = LangToCode(lang)

            If Not dictGroups.Exists(grp) Then dictGroups.Add grp, grp
            If Not dictLangs.Exists(code) Then
                dictLangs.Add code, code
                dictLangNames.Add code, lang  ' Store full language name
            End If

            key = grp & "|" & code
            dictOld(key) = CLng(NzStr(rs.Fields("add_cnt").Value)) - CLng(NzStr(rs.Fields("rmv_cnt").Value))

            rs.MoveNext
        Loop
    End If
    rs.Close

    ' 2) NEW / RMV = only current pushp_no (member+lang+mobile pair cancel logic)
    sql = "WITH base AS (" & _
          "SELECT UPPER(TRIM(COALESCE(""group"", ''))) AS grp, UPPER(TRIM(COALESCE(language, ''))) AS lang, " & _
          "UPPER(TRIM(COALESCE(member, ''))) AS member_name, REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile, " & _
          "UPPER(TRIM(COALESCE(village, ''))) AS village_name, action_type " & _
          "FROM " & DB_SCHEMA & ".tatvagnan_simple_data " & _
          "WHERE pushp_no = '" & Replace(pushpNo, "'", "''") & "'), " & _
          "matched AS (" & _
          "SELECT grp, lang, member_name, mobile, village_name, " & _
          "SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt, " & _
          "SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt " & _
          "FROM base GROUP BY grp, lang, member_name, mobile, village_name), " & _
          "net AS (" & _
          "SELECT grp, lang, GREATEST(add_cnt - rmv_cnt, 0) AS net_add, GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv " & _
          "FROM matched) " & _
          "SELECT grp, lang, SUM(net_add) AS new_cnt, SUM(net_rmv) AS rmv_cnt " & _
          "FROM net GROUP BY grp, lang ORDER BY grp, lang;"

    rs.Open sql, conn, 0, 1
    If rs.EOF Then
        MsgBox "No data found for Pushp No. " & pushpNo, vbInformation
        GoTo CleanUp
    End If

    Do While Not rs.EOF
        grp = NzStr(rs.Fields("grp").Value)
        lang = NzStr(rs.Fields("lang").Value)
        code = LangToCode(lang)

        If Not dictGroups.Exists(grp) Then dictGroups.Add grp, grp
        If Not dictLangs.Exists(code) Then
            dictLangs.Add code, code
            dictLangNames.Add code, lang  ' Store full language name
        End If

        key = grp & "|" & code
        dictNew(key) = CLng(NzStr(rs.Fields("new_cnt").Value))
        dictRmv(key) = CLng(NzStr(rs.Fields("rmv_cnt").Value))

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Tatvagnan Summary"

    Dim row As Long, col As Long
    Dim iGrp As Long, iLang As Long
    Dim grpKeys() As Variant, langKeys() As Variant
    Dim oldCnt As Long, newCnt As Long, rmvCnt As Long, totLang As Long
    Dim grandColsEnd As Long

    grpKeys = dictGroups.Keys
    langKeys = dictLangs.Keys
    
    ' Sort groups naturally (E1, E2, ..., E9, E10, E11)
    NaturalSortArray grpKeys

    ' Set font to Calibri for entire sheet
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    
    row = 1
    ' Title row (will be formatted later after grandColsEnd is determined)
    row = row + 2

    ' Header rows
    Dim headerRow As Long
    headerRow = row
    
    ws.Cells(row, 1).Value = "Sr."
    ws.Range(ws.Cells(row, 1), ws.Cells(row + 1, 1)).Merge
    ws.Cells(row, 1).HorizontalAlignment = xlCenter
    ws.Cells(row, 1).VerticalAlignment = xlCenter
    ws.Cells(row, 1).Font.Bold = True
    
    ws.Cells(row, 2).Value = "Group"
    ws.Range(ws.Cells(row, 2), ws.Cells(row + 1, 2)).Merge
    ws.Cells(row, 2).HorizontalAlignment = xlCenter
    ws.Cells(row, 2).VerticalAlignment = xlCenter
    ws.Cells(row, 2).Font.Bold = True

    ' Calculate previous Pushp number for column header
    Dim prevPushpNo As Long
    prevPushpNo = CLng(pushpNo) - 1
    
    col = 3
    For iLang = LBound(langKeys) To UBound(langKeys)
        code = CStr(langKeys(iLang))
        ' Display full language name instead of code
        Dim fullLangName As String
        fullLangName = code  ' Default to code if not found
        If dictLangNames.Exists(code) Then
            fullLangName = dictLangNames(code)
        End If
        ws.Cells(row, col).Value = fullLangName
        ws.Range(ws.Cells(row, col), ws.Cells(row, col + 3)).Merge
        ws.Cells(row, col).HorizontalAlignment = xlCenter
        ws.Cells(row, col).VerticalAlignment = xlCenter
        ws.Cells(row, col).Font.Bold = True
        
        ' Old ki jagah "P-" + previous pushp number
        ws.Cells(row + 1, col).Value = "P-" & prevPushpNo
        ws.Cells(row + 1, col).HorizontalAlignment = xlCenter
        ws.Cells(row + 1, col).VerticalAlignment = xlCenter
        ws.Cells(row + 1, col).Font.Bold = True
        
        ws.Cells(row + 1, col + 1).Value = "New"
        ws.Cells(row + 1, col + 1).HorizontalAlignment = xlCenter
        ws.Cells(row + 1, col + 1).VerticalAlignment = xlCenter
        ws.Cells(row + 1, col + 1).Font.Bold = True
        
        ' RMV ki jagah "Remove"
        ws.Cells(row + 1, col + 2).Value = "Remove"
        ws.Cells(row + 1, col + 2).HorizontalAlignment = xlCenter
        ws.Cells(row + 1, col + 2).VerticalAlignment = xlCenter
        ws.Cells(row + 1, col + 2).Font.Bold = True
        
        ws.Cells(row + 1, col + 3).Value = "Total"
        ws.Cells(row + 1, col + 3).HorizontalAlignment = xlCenter
        ws.Cells(row + 1, col + 3).VerticalAlignment = xlCenter
        ws.Cells(row + 1, col + 3).Font.Bold = True
        
        col = col + 4
    Next iLang

    ws.Cells(row, col).Value = "Total"
    ws.Range(ws.Cells(row, col), ws.Cells(row + 1, col)).Merge
    ws.Cells(row, col).HorizontalAlignment = xlCenter
    ws.Cells(row, col).VerticalAlignment = xlCenter
    ws.Cells(row, col).Font.Bold = True
    
    ' Add Sign and Date columns
    col = col + 1
    ws.Cells(row, col).Value = "Sign"
    ws.Range(ws.Cells(row, col), ws.Cells(row + 1, col)).Merge
    ws.Cells(row, col).HorizontalAlignment = xlCenter
    ws.Cells(row, col).VerticalAlignment = xlCenter
    ws.Cells(row, col).Font.Bold = True
    
    col = col + 1
    ws.Cells(row, col).Value = "Date"
    ws.Range(ws.Cells(row, col), ws.Cells(row + 1, col)).Merge
    ws.Cells(row, col).HorizontalAlignment = xlCenter
    ws.Cells(row, col).VerticalAlignment = xlCenter
    ws.Cells(row, col).Font.Bold = True
    
    grandColsEnd = col
    
    ' Set header row heights (25% increase: 20 ? 25)
    ws.Rows(headerRow).RowHeight = 25
    ws.Rows(headerRow + 1).RowHeight = 25
    
    ' Add light dark background to header rows
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow + 1, grandColsEnd)).Interior.Color = RGB(217, 217, 217)

    ' Body
    row = row + 2
    Dim sr As Long
    sr = 1

    For iGrp = LBound(grpKeys) To UBound(grpKeys)
        grp = CStr(grpKeys(iGrp))
        col = 3
        Dim rowTotal As Long
        rowTotal = 0

        For iLang = LBound(langKeys) To UBound(langKeys)
            code = CStr(langKeys(iLang))
            oldCnt = 0
            newCnt = 0
            rmvCnt = 0

            key = grp & "|" & code
            If dictOld.Exists(key) Then oldCnt = dictOld(key)
            If dictNew.Exists(key) Then newCnt = dictNew(key)
            If dictRmv.Exists(key) Then rmvCnt = dictRmv(key)

            totLang = oldCnt + newCnt - rmvCnt

            If oldCnt <> 0 Then
                ws.Cells(row, col).Value = oldCnt
                ws.Cells(row, col).HorizontalAlignment = xlCenter
                ws.Cells(row, col).VerticalAlignment = xlCenter
            End If
            If newCnt <> 0 Then
                ws.Cells(row, col + 1).Value = newCnt
                ws.Cells(row, col + 1).HorizontalAlignment = xlCenter
                ws.Cells(row, col + 1).VerticalAlignment = xlCenter
            End If
            If rmvCnt <> 0 Then
                ws.Cells(row, col + 2).Value = rmvCnt
                ws.Cells(row, col + 2).HorizontalAlignment = xlCenter
                ws.Cells(row, col + 2).VerticalAlignment = xlCenter
            End If
            If totLang <> 0 Then
                ws.Cells(row, col + 3).Value = totLang
                ws.Cells(row, col + 3).HorizontalAlignment = xlCenter
                ws.Cells(row, col + 3).VerticalAlignment = xlCenter
            End If
            rowTotal = rowTotal + totLang

            col = col + 4
        Next iLang

        If rowTotal <> 0 Then
            ws.Cells(row, 1).Value = sr
            ws.Cells(row, 1).HorizontalAlignment = xlCenter
            ws.Cells(row, 1).VerticalAlignment = xlCenter
            ws.Cells(row, 2).Value = grp
            ws.Cells(row, 2).VerticalAlignment = xlCenter
            ws.Cells(row, grandColsEnd - 2).Value = rowTotal
            ws.Cells(row, grandColsEnd - 2).HorizontalAlignment = xlCenter
            ws.Cells(row, grandColsEnd - 2).VerticalAlignment = xlCenter
            
            sr = sr + 1
            row = row + 1
        Else
            ' If this group is fully cancelled out, remove the row from output.
            ws.Range(ws.Cells(row, 1), ws.Cells(row, grandColsEnd)).ClearContents
        End If
    Next iGrp

    ' Grand total row
    Dim totalRow As Long
    Dim firstDataRow As Long
    firstDataRow = headerRow + 2  ' First data row after header
    totalRow = row + 1
    ws.Cells(totalRow, 1).Value = "Total"
    ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, 2)).Merge
    ws.Cells(totalRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(totalRow, 1).VerticalAlignment = xlCenter
    ws.Cells(totalRow, 1).Font.Bold = True

    For col = 3 To grandColsEnd - 2
        ' Fixed: Use firstDataRow to row-1 for correct total calculation
        ws.Cells(totalRow, col).FormulaR1C1 = "=SUM(R" & firstDataRow & "C" & col & ":R" & (row - 1) & "C" & col & ")"
        ws.Cells(totalRow, col).HorizontalAlignment = xlCenter
        ws.Cells(totalRow, col).VerticalAlignment = xlCenter
        ws.Cells(totalRow, col).Font.Bold = True
    Next col
    
    ' Add light background to total row
    ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, grandColsEnd)).Interior.Color = RGB(217, 217, 217)
    
    ' Add dark background to entire Total columns (header to total row) for visual separation
    Dim totalCol As Long
    totalCol = 6  ' First language Total column (3 + 3)
    For iLang = LBound(langKeys) To UBound(langKeys)
        ' Apply to entire column: header + all data rows + total row
        ws.Range(ws.Cells(headerRow + 1, totalCol), ws.Cells(totalRow, totalCol)).Interior.Color = RGB(230, 230, 230)
        totalCol = totalCol + 4  ' Move to next language Total column
    Next iLang

    ' Predefined column widths for PDF
    ws.columns(1).ColumnWidth = 4      ' Sr. column
    ws.columns(2).ColumnWidth = 12     ' Group column (reduced from 15)
    
    ' Set width for all language columns (Old, New, RMV, Total)
    For col = 3 To grandColsEnd - 3
        ws.columns(col).ColumnWidth = 6  ' Reduced from 7
    Next col
    
    ws.columns(grandColsEnd - 2).ColumnWidth = 7   ' Grand Total column (reduced)
    ws.columns(grandColsEnd - 1).ColumnWidth = 18  ' Sign column (increased from 12)
    ws.columns(grandColsEnd).ColumnWidth = 10      ' Date column
    
    ' Format title - centered without border
    Dim titleRange As Range
    Set titleRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, grandColsEnd))
    titleRange.Merge
    titleRange.Value = "TATVAGNAN SUMMARY - Pushp No. " & pushpNo
    titleRange.Font.Bold = True
    titleRange.Font.Size = 14
    titleRange.HorizontalAlignment = xlCenter
    titleRange.VerticalAlignment = xlCenter
    ws.Rows(1).RowHeight = 31  ' 25% increase: 25 ? 31
    
    ' Set row heights for data rows (25% increase: 18 ? 23)
    For row = headerRow + 2 To totalRow
        ws.Rows(row).RowHeight = 23
    Next row
    
    ' Apply borders only to table (not title)
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(totalRow, grandColsEnd)).Borders.LineStyle = xlContinuous

    ' Page setup: landscape, fit full width on one page
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
    End With

    Dim pdfPath As Variant
    pdfPath = Application.GetSaveAsFilename( _
        InitialFileName:="Tatvagnan_Summary_Pushp_" & pushpNo & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If VarType(pdfPath) <> vbBoolean And Len(CStr(pdfPath)) > 0 Then
        wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=CStr(pdfPath), _
                               OpenAfterPublish:=True
    End If

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set rs = Nothing
    Set conn = Nothing
    Set wb = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error generating Tatvagnan summary PDF: " & Err.description, vbCritical
    GoTo CleanUp
End Sub

' ===================== GROUP WISE SUMMARY PDF =====================

Public Sub Tatvagnan_GenerateGroupWisePdf(Optional ByVal pushpNo As String = "")
    On Error GoTo ErrHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dictGroups As Object
    Dim dictData As Object
    Dim key As String
    Dim grp As String, lang As String, code As String
    Dim grpKeys() As Variant
    Dim iGrp As Long
    
    If Trim$(pushpNo) = "" Then
        pushpNo = InputBox("Enter Tatvagnan Pushp No. for group-wise PDF:", "Tatvagnan Group-Wise PDF")
        If Trim$(pushpNo) = "" Then Exit Sub
    End If
    
    Set conn = CreateObject("ADODB.Connection")
    If Not TatvagnanDataImportModule.ConnectToPostgreSQL(conn) Then
        MsgBox "Database connection failed.", vbCritical
        Exit Sub
    End If
    
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictData = CreateObject("Scripting.Dictionary")
    
    Call LogTatvagnanNettingDebug(conn, pushpNo)
    
    ' Get OLD data (all previous pushp_no) with member+lang+mobile pair cancel logic
    sql = "WITH base AS (" & _
          "SELECT UPPER(TRIM(COALESCE(""group"", ''))) AS grp, UPPER(TRIM(COALESCE(language, ''))) AS lang, " & _
          "UPPER(TRIM(COALESCE(member, ''))) AS member_name, REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile, " & _
          "UPPER(TRIM(COALESCE(village, ''))) AS village_name, action_type " & _
          "FROM " & DB_SCHEMA & ".tatvagnan_simple_data " & _
          "WHERE pushp_no::integer < " & Replace(pushpNo, "'", "''") & "), " & _
          "matched AS (" & _
          "SELECT grp, lang, member_name, mobile, village_name, " & _
          "SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt, " & _
          "SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt " & _
          "FROM base GROUP BY grp, lang, member_name, mobile, village_name), " & _
          "net AS (" & _
          "SELECT grp, lang, GREATEST(add_cnt - rmv_cnt, 0) AS net_add, GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv " & _
          "FROM matched) " & _
          "SELECT grp, lang, SUM(net_add) AS add_cnt, SUM(net_rmv) AS rmv_cnt " & _
          "FROM net GROUP BY grp, lang;"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            grp = NzStr(rs.Fields("grp").Value)
            lang = NzStr(rs.Fields("lang").Value)
            
            If Not dictGroups.Exists(grp) Then dictGroups.Add grp, grp
            
            key = grp & "|" & lang & "|OLD"
            dictData(key) = CLng(NzStr(rs.Fields("add_cnt").Value)) - CLng(NzStr(rs.Fields("rmv_cnt").Value))
            
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    ' Get NEW and RMV data (current pushp_no) with member+lang+mobile pair cancel logic
    sql = "WITH base AS (" & _
          "SELECT UPPER(TRIM(COALESCE(""group"", ''))) AS grp, UPPER(TRIM(COALESCE(language, ''))) AS lang, " & _
          "UPPER(TRIM(COALESCE(member, ''))) AS member_name, REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile, " & _
          "UPPER(TRIM(COALESCE(village, ''))) AS village_name, action_type " & _
          "FROM " & DB_SCHEMA & ".tatvagnan_simple_data " & _
          "WHERE pushp_no = '" & Replace(pushpNo, "'", "''") & "'), " & _
          "matched AS (" & _
          "SELECT grp, lang, member_name, mobile, village_name, " & _
          "SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt, " & _
          "SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt " & _
          "FROM base GROUP BY grp, lang, member_name, mobile, village_name), " & _
          "net AS (" & _
          "SELECT grp, lang, GREATEST(add_cnt - rmv_cnt, 0) AS net_add, GREATEST(rmv_cnt - add_cnt, 0) AS net_rmv " & _
          "FROM matched) " & _
          "SELECT grp, lang, SUM(net_add) AS new_cnt, SUM(net_rmv) AS rmv_cnt " & _
          "FROM net GROUP BY grp, lang;"
    
    rs.Open sql, conn, 0, 1
    
    If rs.EOF Then
        MsgBox "No data found for Pushp No. " & pushpNo, vbInformation
        GoTo CleanUp
    End If
    
    Do While Not rs.EOF
        grp = NzStr(rs.Fields("grp").Value)
        lang = NzStr(rs.Fields("lang").Value)
        
        If Not dictGroups.Exists(grp) Then dictGroups.Add grp, grp
        
        key = grp & "|" & lang & "|NEW"
        dictData(key) = CLng(NzStr(rs.Fields("new_cnt").Value))
        
        key = grp & "|" & lang & "|RMV"
        dictData(key) = CLng(NzStr(rs.Fields("rmv_cnt").Value))
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    grpKeys = dictGroups.Keys
    NaturalSortArray grpKeys
    
    Set wb = Workbooks.Add
    
    ' Create sheets with 4 groups per sheet
    Dim sheetNum As Long
    Dim groupsInSheet As Long
    Dim startRow As Long
    sheetNum = 1
    
    For iGrp = LBound(grpKeys) To UBound(grpKeys)
        groupsInSheet = ((iGrp - LBound(grpKeys)) Mod 4)
        
        ' Create new sheet for every 4 groups
        If groupsInSheet = 0 Then
            If iGrp = LBound(grpKeys) Then
                Set ws = wb.Worksheets(1)
            Else
                Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
                sheetNum = sheetNum + 1
            End If
            ws.Name = "Page_" & sheetNum
            startRow = 1
        End If
        
        grp = CStr(grpKeys(iGrp))
        
        ' Generate group table at current position
        Call GenerateGroupTable(ws, pushpNo, grp, dictData, startRow)
        
        ' Move to next position - spacing for max 9 rows + headers + total + sign
        ' Title(1) + Header(1) + Data(9) + Total(1) + Sign(1) + Gap(1) = 14 rows
        startRow = startRow + 18  ' Compact spacing for 4 tables per page
    Next iGrp
    
    ' Set page setup for all sheets - Portrait A4 optimized for 4 tables
    Dim wSheet As Worksheet
    For Each wSheet In wb.Worksheets
        With wSheet.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False  ' Let it flow to multiple pages if needed
            .LeftMargin = Application.CentimetersToPoints(1.2)     ' Reduced
            .RightMargin = Application.CentimetersToPoints(1.2)    ' Reduced
            .TopMargin = Application.CentimetersToPoints(0.8)      ' Reduced
            .BottomMargin = Application.CentimetersToPoints(0.8)   ' Reduced
            .HeaderMargin = Application.CentimetersToPoints(0.3)
            .FooterMargin = Application.CentimetersToPoints(0.3)
        End With
    Next wSheet
    
    Dim pdfPath As Variant
    pdfPath = Application.GetSaveAsFilename( _
        InitialFileName:="Tatvagnan_GroupWise_Pushp_" & pushpNo & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")
    
    If VarType(pdfPath) <> vbBoolean And Len(CStr(pdfPath)) > 0 Then
        wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=CStr(pdfPath), _
                               OpenAfterPublish:=True
    End If
    
CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set rs = Nothing
    Set conn = Nothing
    Set wb = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error generating group-wise PDF: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub

Private Sub LogTatvagnanNettingDebug(ByVal conn As Object, ByVal pushpNo As String)
    On Error GoTo ErrHandler
    
    Dim rsDbg As Object
    Dim sqlDbg As String
    Dim n As Long
    
    sqlDbg = "WITH base AS (" & _
             "SELECT UPPER(TRIM(COALESCE(""group"", ''))) AS grp, UPPER(TRIM(COALESCE(village, ''))) AS village_name, " & _
             "UPPER(TRIM(COALESCE(language, ''))) AS lang, UPPER(TRIM(COALESCE(member, ''))) AS member_name, " & _
             "REGEXP_REPLACE(COALESCE(mobile_number, ''), '[^0-9]', '', 'g') AS mobile, action_type " & _
             "FROM " & DB_SCHEMA & ".tatvagnan_simple_data WHERE pushp_no = '" & Replace(pushpNo, "'", "''") & "'), " & _
             "agg AS (" & _
             "SELECT grp, village_name, lang, member_name, mobile, " & _
             "SUM(CASE WHEN action_type = 'ADD' THEN 1 ELSE 0 END) AS add_cnt, " & _
             "SUM(CASE WHEN action_type = 'REMOVE' THEN 1 ELSE 0 END) AS rmv_cnt " & _
             "FROM base GROUP BY grp, village_name, lang, member_name, mobile) " & _
             "SELECT grp, village_name, lang, member_name, mobile, add_cnt, rmv_cnt, " & _
             "GREATEST(add_cnt-rmv_cnt,0) AS net_add, GREATEST(rmv_cnt-add_cnt,0) AS net_rmv " & _
             "FROM agg WHERE add_cnt > 0 OR rmv_cnt > 0 ORDER BY grp, village_name, member_name, lang;"
    
    Set rsDbg = CreateObject("ADODB.Recordset")
    rsDbg.Open sqlDbg, conn, 0, 1
    
    Debug.Print "===== Tatvagnan Netting Debug (Pushp " & pushpNo & ") ====="
    n = 0
    Do While Not rsDbg.EOF
        n = n + 1
        Debug.Print rsDbg.Fields("grp").Value & " | " & rsDbg.Fields("village_name").Value & " | " & _
                    rsDbg.Fields("member_name").Value & " | " & rsDbg.Fields("lang").Value & " | " & _
                    rsDbg.Fields("mobile").Value & " || ADD=" & rsDbg.Fields("add_cnt").Value & _
                    ", RMV=" & rsDbg.Fields("rmv_cnt").Value & ", NET_ADD=" & rsDbg.Fields("net_add").Value & _
                    ", NET_RMV=" & rsDbg.Fields("net_rmv").Value
        rsDbg.MoveNext
    Loop
    Debug.Print "===== End Netting Debug (" & n & " keys) ====="
    
    rsDbg.Close
    Set rsDbg = Nothing
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    Debug.Print "Netting debug failed: " & Err.Description
    If Not rsDbg Is Nothing Then
        If rsDbg.State <> 0 Then rsDbg.Close
    End If
    Set rsDbg = Nothing
End Sub

Private Sub GenerateGroupTable(ByVal ws As Worksheet, ByVal pushpNo As String, ByVal grp As String, ByVal dictData As Object, ByVal startRow As Long)
    Dim row As Long, col As Long
    Dim key As String
    Dim oldCnt As Long, newCnt As Long, rmvCnt As Long, totCnt As Long
    Dim langList As Object
    Dim lang As Variant
    Dim sr As Long
    
    ' Get all unique languages for this group with their totals
    Set langList = CreateObject("Scripting.Dictionary")
    Dim langTotals As Object
    Set langTotals = CreateObject("Scripting.Dictionary")
    
    Dim k As Variant
    For Each k In dictData.Keys
        If InStr(CStr(k), grp & "|") = 1 Then
            Dim parts() As String
            parts = Split(CStr(k), "|")
            If UBound(parts) >= 1 Then
                Dim tempLang As String
                tempLang = parts(1)
                If Not langList.Exists(tempLang) Then
                    langList.Add tempLang, tempLang
                    
                    ' Calculate total for this language
                    Dim tempOld As Long, tempNew As Long, tempRmv As Long, tempTotal As Long
                    tempOld = 0: tempNew = 0: tempRmv = 0
                    
                    If dictData.Exists(grp & "|" & tempLang & "|OLD") Then
                        tempOld = dictData(grp & "|" & tempLang & "|OLD")
                        If tempOld < 0 Then tempOld = 0
                    End If
                    If dictData.Exists(grp & "|" & tempLang & "|NEW") Then
                        tempNew = dictData(grp & "|" & tempLang & "|NEW")
                    End If
                    If dictData.Exists(grp & "|" & tempLang & "|RMV") Then
                        tempRmv = dictData(grp & "|" & tempLang & "|RMV")
                    End If
                    
                    tempTotal = tempOld + tempNew - tempRmv
                    If tempTotal <> 0 Then
                        langTotals(tempLang) = tempTotal
                    Else
                        langList.Remove tempLang
                    End If
                End If
            End If
        End If
    Next k
    
    If langList.Count = 0 Then Exit Sub
    
    ' Sort languages by total (highest first)
    Dim langArray() As Variant
    ReDim langArray(0 To langList.Count - 1, 0 To 1)
    
    Dim i As Long
    i = 0
    For Each lang In langList.Keys
        langArray(i, 0) = CStr(lang)
        langArray(i, 1) = langTotals(CStr(lang))
        i = i + 1
    Next lang
    
    ' Bubble sort by total (descending)
    Dim j As Long
    Dim tempStr As String
    Dim tempVal As Long
    For i = 0 To UBound(langArray, 1) - 1
        For j = i + 1 To UBound(langArray, 1)
            If langArray(j, 1) > langArray(i, 1) Then
                tempStr = langArray(i, 0)
                tempVal = langArray(i, 1)
                langArray(i, 0) = langArray(j, 0)
                langArray(i, 1) = langArray(j, 1)
                langArray(j, 0) = tempStr
                langArray(j, 1) = tempVal
            End If
        Next j
    Next i
    
    ' Set font to Calibri (only on first call) - compact size for 4 tables
    If startRow = 1 Then
        ws.Cells.Font.Name = "Calibri"
        ws.Cells.Font.Size = 9
    End If
    
    row = startRow
    
    ' Title row - Left: Pushp No, Right: Group Name (with left margin)
    ws.Cells(row, 2).Value = "Tatvagnan Pushp No. " & pushpNo
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 10
    ws.Cells(row, 2).HorizontalAlignment = xlLeft
    
    ws.Cells(row, 6).Value = grp
    ws.Cells(row, 6).Font.Bold = True
    ws.Cells(row, 6).Font.Size = 10
    ws.Cells(row, 6).HorizontalAlignment = xlRight
    ws.Rows(row).RowHeight = 15  ' Extra compact - 1 point reduced
    
    row = row + 1
    
    ' Header row (starting from column 2 for left margin)
    Dim headerRow As Long
    headerRow = row
    
    ws.Cells(row, 2).Value = "Sr."
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).HorizontalAlignment = xlCenter
    ws.Cells(row, 2).VerticalAlignment = xlCenter
    
    ws.Cells(row, 3).Value = "Language"
    ws.Cells(row, 3).Font.Bold = True
    ws.Cells(row, 3).HorizontalAlignment = xlCenter
    ws.Cells(row, 3).VerticalAlignment = xlCenter
    
    ' Old ki jagah "Pushp - (input-1)" likho
    Dim prevPushpNo As Long
    prevPushpNo = CLng(pushpNo) - 1
    ws.Cells(row, 4).Value = "Pushp - " & prevPushpNo
    ws.Cells(row, 4).Font.Bold = True
    ws.Cells(row, 4).HorizontalAlignment = xlCenter
    ws.Cells(row, 4).VerticalAlignment = xlCenter
    
    ws.Cells(row, 5).Value = "New"
    ws.Cells(row, 5).Font.Bold = True
    ws.Cells(row, 5).HorizontalAlignment = xlCenter
    ws.Cells(row, 5).VerticalAlignment = xlCenter
    
    ' RMV ki jagah "Remove" likho
    ws.Cells(row, 6).Value = "Remove"
    ws.Cells(row, 6).Font.Bold = True
    ws.Cells(row, 6).HorizontalAlignment = xlCenter
    ws.Cells(row, 6).VerticalAlignment = xlCenter
    
    ws.Cells(row, 7).Value = "Total"
    ws.Cells(row, 7).Font.Bold = True
    ws.Cells(row, 7).HorizontalAlignment = xlCenter
    ws.Cells(row, 7).VerticalAlignment = xlCenter
    
    ws.Range(ws.Cells(headerRow, 2), ws.Cells(headerRow, 7)).Interior.Color = RGB(217, 217, 217)
    ws.Rows(headerRow).RowHeight = 14  ' Extra compact - 1 point reduced
    
    row = row + 1
    
    ' Data rows (starting from column 2 for left margin) - sorted by total
    sr = 1
    For i = 0 To UBound(langArray, 1)
        lang = langArray(i, 0)
        ws.Cells(row, 2).Value = sr
        ws.Cells(row, 2).HorizontalAlignment = xlCenter
        ws.Cells(row, 2).VerticalAlignment = xlCenter
        
        ws.Cells(row, 3).Value = CStr(lang)
        ws.Cells(row, 3).VerticalAlignment = xlCenter
        
        key = grp & "|" & lang & "|OLD"
        oldCnt = 0
        If dictData.Exists(key) Then oldCnt = dictData(key)
        If oldCnt < 0 Then oldCnt = 0
        
        key = grp & "|" & lang & "|NEW"
        newCnt = 0
        If dictData.Exists(key) Then newCnt = dictData(key)
        
        key = grp & "|" & lang & "|RMV"
        rmvCnt = 0
        If dictData.Exists(key) Then rmvCnt = dictData(key)
        
        totCnt = oldCnt + newCnt - rmvCnt
        
        ' Only show value if not 0, otherwise leave blank
        If oldCnt <> 0 Then
            ws.Cells(row, 4).Value = oldCnt
        End If
        ws.Cells(row, 4).HorizontalAlignment = xlCenter
        ws.Cells(row, 4).VerticalAlignment = xlCenter
        
        If newCnt <> 0 Then
            ws.Cells(row, 5).Value = newCnt
        End If
        ws.Cells(row, 5).HorizontalAlignment = xlCenter
        ws.Cells(row, 5).VerticalAlignment = xlCenter
        
        If rmvCnt <> 0 Then
            ws.Cells(row, 6).Value = rmvCnt
        End If
        ws.Cells(row, 6).HorizontalAlignment = xlCenter
        ws.Cells(row, 6).VerticalAlignment = xlCenter
        
        If totCnt <> 0 Then
            ws.Cells(row, 7).Value = totCnt
        End If
        ws.Cells(row, 7).HorizontalAlignment = xlCenter
        ws.Cells(row, 7).VerticalAlignment = xlCenter
        
        ws.Rows(row).RowHeight = 13  ' Extra compact - 1 point reduced
        sr = sr + 1
        row = row + 1
    Next i
    
    ' Total row (starting from column 2 for left margin)
    Dim totalRow As Long
    totalRow = row
    
    ws.Cells(totalRow, 2).Value = "Total"
    ws.Range(ws.Cells(totalRow, 2), ws.Cells(totalRow, 3)).Merge
    ws.Cells(totalRow, 2).Font.Bold = True
    ws.Cells(totalRow, 2).HorizontalAlignment = xlCenter
    ws.Cells(totalRow, 2).VerticalAlignment = xlCenter
    
    For col = 4 To 7
        ws.Cells(totalRow, col).FormulaR1C1 = "=SUM(R" & (headerRow + 1) & "C" & col & ":R" & (totalRow - 1) & "C" & col & ")"
        ws.Cells(totalRow, col).Font.Bold = True
        ws.Cells(totalRow, col).HorizontalAlignment = xlCenter
        ws.Cells(totalRow, col).VerticalAlignment = xlCenter
    Next col
    
    ws.Range(ws.Cells(totalRow, 2), ws.Cells(totalRow, 7)).Interior.Color = RGB(217, 217, 217)
    ws.Rows(totalRow).RowHeight = 13  ' Extra compact - 1 point reduced
    
    row = totalRow + 1
    
    ' Sign and Date at bottom (without lines, with left margin)
    ws.Cells(row, 2).Value = "Sign:"
    ws.Cells(row, 2).Font.Size = 8
    
    ws.Cells(row, 5).Value = "Date:"
    ws.Cells(row, 5).Font.Size = 8
    ws.Rows(row).RowHeight = 12  ' Extra compact - 1 point reduced
    
    ' Column widths optimized for portrait with left margin (only set once)
    If startRow = 1 Then
        ws.Columns(1).ColumnWidth = 6      ' Left margin (equal to Sr. width)
        ws.Columns(2).ColumnWidth = 6      ' Sr. - थोड़ा बड़ा
        ws.Columns(3).ColumnWidth = 20     ' Language - सबसे बड़ा
        ws.Columns(4).ColumnWidth = 10     ' Old
        ws.Columns(5).ColumnWidth = 10     ' New
        ws.Columns(6).ColumnWidth = 10     ' RMV
        ws.Columns(7).ColumnWidth = 12     ' Total - थोड़ा बड़ा
    End If
    
    ' Apply borders only to data table (header to total row, excluding sign/date)
    ws.Range(ws.Cells(headerRow, 2), ws.Cells(totalRow, 7)).Borders.LineStyle = xlContinuous
    
    ' Thicker outer border for better visibility
    With ws.Range(ws.Cells(headerRow, 2), ws.Cells(totalRow, 7))
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
End Sub

