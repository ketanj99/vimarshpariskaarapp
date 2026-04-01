Option Explicit

' SummaryReportPdf.bas
' VBA equivalent of Next.js Book_Summary (generateBookGroupSummaryPdf)
' Book Group Summary: Landscape, 6 cols Sr., Group, Total Qty, Name, Mobile No., Date
' Filename: Book_Summary_Org_Period.pdf

Private Const CONN_STR As String = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"

' Main entry point - Book Group Summary (Next.js Book_Summary format)
Public Sub GenerateSummaryReportPdf(bookGroupData As Collection, period As String, org As String, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
    On Error GoTo ErrorHandler
    Dim fileName As String
    fileName = SafeFileName(org) & "_Book_Summary_" & SafeFileName(period) & ".pdf"
    GenerateBookGroupSummaryPDF bookGroupData, period, org, customFolder, fileName
    Exit Sub
ErrorHandler:
    MsgBox "Error generating Book Summary PDF: " & Err.Description, vbCritical
End Sub

' Get Book Group Summary data - group_sector, total_qty per group for period
Public Function GetSummaryReportData(period As String, groupSector As String, Optional selectedBookIds As Collection = Nothing, Optional orgName As String = "VIMARSH") As Collection
    On Error GoTo ErrorHandler
    Set GetSummaryReportData = GetBookGroupSummaryData(period, orgName)
    Exit Function
ErrorHandler:
    Set GetSummaryReportData = New Collection
End Function

' Form integration - call from ReportGenerationForm when chkAllReport is selected
Public Sub GenerateSummaryReportPdfFromForm(period As String, groupSector As String, selectedBookIds As Collection, customFolder As String, orgName As String)
    On Error GoTo ErrorHandler
    Dim bookGroupData As Collection
    Set bookGroupData = GetBookGroupSummaryData(period, orgName)
    If bookGroupData Is Nothing Or bookGroupData.Count = 0 Then
        MsgBox "No data found for the selected period.", vbInformation, "No Data"
        Exit Sub
    End If
    GenerateSummaryReportPdf bookGroupData, period, orgName, customFolder, orgName
    Exit Sub
ErrorHandler:
    MsgBox "Error generating Book Summary Report: " & Err.Description, vbCritical
End Sub

' Fetch group_sector, SUM(qty) from transactions for period and org
Private Function GetBookGroupSummaryData(period As String, orgName As String) As Collection
    On Error GoTo ErrorHandler
    Dim conn As Object, rs As Object, sql As String
    Dim result As Collection
    Set result = New Collection
    Set conn = CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open CONN_STR
    If Err.Number <> 0 Then
        Set GetBookGroupSummaryData = result
        Exit Function
    End If
    On Error GoTo ErrorHandler
    sql = "SELECT COALESCE(group_sector, '') AS ""group"", SUM(qty)::bigint AS total_qty " & _
          "FROM vimars.transactions " & _
          "WHERE is_deleted = FALSE AND app = '" & Replace(orgName, "'", "''") & "' "
    If period <> "" And period <> "All" Then
        Dim parts() As String
        parts = Split(period, "-")
        If UBound(parts) >= 2 Then
            Dim halfVal As String
            halfVal = Replace(Trim(parts(2)), "'", "''")
            sql = sql & "AND year = " & CLng(parts(0)) & " AND month = " & CLng(parts(1)) & _
                  " AND (REPLACE(LOWER(COALESCE(half,'')), ' ', '') IN ('1&2','1and2') OR half = '" & halfVal & "') "
        End If
    End If
    sql = sql & "GROUP BY group_sector ORDER BY group_sector;"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        Dim rec(1) As Variant
        rec(0) = rs.Fields("group").Value
        rec(1) = rs.Fields("total_qty").Value
        result.Add rec
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    Set GetBookGroupSummaryData = result
    Exit Function
ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set GetBookGroupSummaryData = result
End Function

' Generate Book Group Summary PDF - Landscape, 6 cols
Private Sub GenerateBookGroupSummaryPDF(bookGroupData As Collection, period As String, org As String, customFolder As String, fileName As String)
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Dim pdfWb As Workbook, pdfWs As Worksheet
    Set pdfWb = Workbooks.Add
    Set pdfWs = pdfWb.Sheets(1)
    pdfWs.Name = "Book Summary"
    pdfWs.PageSetup.Orientation = xlLandscape
    pdfWs.PageSetup.PaperSize = xlPaperA4
    pdfWs.PageSetup.FitToPagesWide = 1
    pdfWs.PageSetup.Zoom = False
    pdfWs.PageSetup.LeftMargin = Application.InchesToPoints(0.55)
    pdfWs.PageSetup.RightMargin = Application.InchesToPoints(0.55)
    pdfWs.PageSetup.TopMargin = Application.InchesToPoints(0.4)
    pdfWs.PageSetup.BottomMargin = Application.InchesToPoints(0.4)
    Dim currentRow As Integer, i As Integer, grandTotal As Long
    currentRow = 1
    grandTotal = 0
    Dim periodDisplay As String
    periodDisplay = FormatPeriodForDisplay(period)
    pdfWs.Cells(currentRow, 1).Value = org & " - " & periodDisplay
    pdfWs.Cells(currentRow, 1).Font.Bold = True
    pdfWs.Cells(currentRow, 1).Font.Size = 22
    pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter
    pdfWs.Cells(currentRow, 1).VerticalAlignment = xlCenter
    pdfWs.Range("A1:G1").Merge
    pdfWs.Rows(1).RowHeight = 36
    pdfWs.Cells(2, 7).Value = Format(Now, "DD-MM-YYYY HH:MM")
    pdfWs.Cells(2, 7).Font.Bold = True
    pdfWs.Cells(2, 7).Font.Size = 22
    pdfWs.Cells(2, 7).HorizontalAlignment = xlRight
    currentRow = 3
    pdfWs.Cells(currentRow, 1).Value = "Sr."
    pdfWs.Cells(currentRow, 2).Value = "Group"
    pdfWs.Cells(currentRow, 3).Value = "Qty"
    pdfWs.Cells(currentRow, 4).Value = "Name"
    pdfWs.Cells(currentRow, 5).Value = "Mobile No."
    pdfWs.Cells(currentRow, 6).Value = "Date"
    pdfWs.Cells(currentRow, 7).Value = "Remark"
    With pdfWs.Range("A" & currentRow & ":G" & currentRow)
        .Font.Bold = True
        .Font.Size = 22
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    pdfWs.Rows(currentRow).RowHeight = 36
    currentRow = currentRow + 1
    Dim groupOrder As Collection
    Set groupOrder = NaturalSortGroups(bookGroupData)
    For i = 1 To groupOrder.Count
        Dim rec As Variant
        rec = groupOrder.Item(i)
        pdfWs.Cells(currentRow, 1).Value = i
        pdfWs.Cells(currentRow, 2).Value = rec(0)
        pdfWs.Cells(currentRow, 3).Value = rec(1)
        pdfWs.Cells(currentRow, 4).Value = ""
        pdfWs.Cells(currentRow, 5).Value = ""
        pdfWs.Cells(currentRow, 6).Value = ""
        pdfWs.Cells(currentRow, 7).Value = ""
        grandTotal = grandTotal + CLng(rec(1))
        With pdfWs.Range("A" & currentRow & ":G" & currentRow)
            .Font.Size = 22
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        pdfWs.Rows(currentRow).RowHeight = 36
        currentRow = currentRow + 1
    Next i
    pdfWs.Cells(currentRow, 1).Value = ""
    pdfWs.Cells(currentRow, 2).Value = "Total"
    pdfWs.Cells(currentRow, 3).Value = grandTotal
    pdfWs.Cells(currentRow, 4).Value = ""
    pdfWs.Cells(currentRow, 5).Value = ""
    pdfWs.Cells(currentRow, 6).Value = ""
    pdfWs.Cells(currentRow, 7).Value = ""
    With pdfWs.Range("A" & currentRow & ":G" & currentRow)
        .Font.Bold = True
        .Font.Size = 22
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    pdfWs.Rows(currentRow).RowHeight = 36
    pdfWs.Columns("A").ColumnWidth = 6
    pdfWs.Columns("B").ColumnWidth = 15
    pdfWs.Columns("C").ColumnWidth = 10
    pdfWs.Columns("D").ColumnWidth = 55
    pdfWs.Columns("E").ColumnWidth = 18
    pdfWs.Columns("F").ColumnWidth = 18
    pdfWs.Columns("G").ColumnWidth = 55
    Dim pdfPath As String
    If customFolder <> "" Then
        pdfPath = customFolder & "\" & fileName
    Else
        pdfPath = ThisWorkbook.Path & "\" & fileName
    End If
    pdfWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, OpenAfterPublish:=False
    pdfWb.Close False
    Application.DisplayAlerts = True
    Exit Sub
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Error generating Book Summary PDF: " & Err.Description, vbCritical
End Sub

Private Function NaturalSortGroups(data As Collection) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim arr() As Variant
    ReDim arr(1 To data.Count)
    Dim i As Long
    For i = 1 To data.Count
        arr(i) = data.Item(i)
    Next i
    Dim j As Long
    Dim temp As Variant
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If NaturalCompare(CStr(arr(i)(0)), CStr(arr(j)(0))) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    For i = 1 To UBound(arr)
        result.Add arr(i)
    Next i
    Set NaturalSortGroups = result
End Function

Private Function NaturalCompare(a As String, b As String) As Integer
    Dim regex As Object, ma As Object, mb As Object
    Dim i As Long, ap As String, bp As String
    On Error GoTo UseStrCompare
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "(\d+)|(\D+)"
    Set ma = regex.Execute(a)
    Set mb = regex.Execute(b)
    Dim mx As Long
    mx = ma.Count: If mb.Count > mx Then mx = mb.Count
    For i = 0 To mx - 1
        If i < ma.Count Then ap = ma(i).Value Else ap = ""
        If i < mb.Count Then bp = mb(i).Value Else bp = ""
        If IsNumeric(ap) And IsNumeric(bp) Then
            Dim d As Long
            d = CLng(ap) - CLng(bp)
            If d <> 0 Then NaturalCompare = IIf(d < 0, -1, 1): Exit Function
        Else
            If ap < bp Then NaturalCompare = -1: Exit Function
            If ap > bp Then NaturalCompare = 1: Exit Function
        End If
    Next i
    NaturalCompare = 0
    Exit Function
UseStrCompare:
    If a < b Then NaturalCompare = -1: Exit Function
    If a > b Then NaturalCompare = 1: Exit Function
    NaturalCompare = 0
End Function

Private Function FormatPeriodForDisplay(period As String) As String
    If period = "" Or period = "All" Then
        FormatPeriodForDisplay = "All Periods"
        Exit Function
    End If
    Dim parts() As String
    parts = Split(period, "-")
    If UBound(parts) < 2 Then
        FormatPeriodForDisplay = period
        Exit Function
    End If
    Dim monNames As Variant
    monNames = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    Dim mn As Long
    mn = CLng(parts(1))
    If mn < 1 Or mn > 12 Then
        FormatPeriodForDisplay = period
        Exit Function
    End If
    Dim half As String
    half = Replace(LCase(Trim(parts(2))), " ", "")
    Dim rng As String
    If half = "1&2" Or half = "1and2" Then
        rng = "1 to " & DaysInMonth(mn, CLng(parts(0)))
    ElseIf half = "1" Then
        rng = "1 to 15"
    ElseIf half = "2" Then
        rng = "16 to " & DaysInMonth(mn, CLng(parts(0)))
    Else
        rng = "1 to " & DaysInMonth(mn, CLng(parts(0)))
    End If
    FormatPeriodForDisplay = parts(0) & "-" & monNames(mn - 1) & "(" & rng & ")"
End Function

Private Function DaysInMonth(monthNum As Long, yearNum As Long) As Long
    Select Case monthNum
        Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
        Case 4, 6, 9, 11: DaysInMonth = 30
        Case 2
            If (yearNum Mod 4 = 0 And yearNum Mod 100 <> 0) Or (yearNum Mod 400 = 0) Then
                DaysInMonth = 29
            Else
                DaysInMonth = 28
            End If
        Case Else: DaysInMonth = 30
    End Select
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim i As Long, c As String, result As String
    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or c = "_" Or c = "-" Then
            result = result & c
        Else
            result = result & "_"
        End If
    Next i
    SafeFileName = result
End Function
