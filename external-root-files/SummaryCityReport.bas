Option Explicit

' All Summary Report Function (4th Report)
Public Function GetAllSummaryDataForPDF(period As String, Optional selectedBookIds As Collection = Nothing, Optional appValue As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    Dim conn As Object
    Dim summaryData As Collection
    Set summaryData = New Collection
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Simple connection string
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"
    
    ' Try to open connection
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        Set GetAllSummaryDataForPDF = New Collection
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Build SQL query to get all data for the period (no sector filter)
    sql = "SELECT b.language, " & _
          "b.book_name, " & _
          "CASE WHEN t.avail = 1 THEN 'YES' WHEN t.avail = 0 THEN 'NS' WHEN t.avail = 2 THEN 'B-YES' ELSE 'UNKNOWN' END AS availability_text, " & _
          "SUM(t.qty) as total_qty " & _
          "FROM vimars.transactions t " & _
          "INNER JOIN vimars.books b ON t.book_id = b.book_id " & _
          "WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE "
    
    ' Add period filter
    If period <> "" And period <> "All" Then
        Dim periodParts() As String
        periodParts = Split(period, "-")
        If UBound(periodParts) >= 2 Then
            ' Escape single quotes in half value for SQL
            Dim halfValue As String
            halfValue = Replace(Trim(periodParts(2)), "'", "''")
            sql = sql & "AND t.year = " & periodParts(0) & _
                      " AND t.month = " & CLng(periodParts(1)) & _
                      " AND t.half = '" & halfValue & "' "
        End If
    End If
    
    If Not selectedBookIds Is Nothing Then
        If selectedBookIds.Count > 0 Then
            Dim bookFilter As String
            bookFilter = SectorVillageDetailsReport.BuildNumericInClause(selectedBookIds)
            If bookFilter <> "" Then
                sql = sql & "AND b.book_id IN (" & bookFilter & ") "
            Else
                GoTo Cleanup
            End If
        Else
            GoTo Cleanup
        End If
    End If
    
    ' Add APP filter when provided
    If appValue <> "" Then
        sql = sql & "AND t.app = '" & Replace(appValue, "'", "''") & "' "
    End If
    
    sql = sql & "GROUP BY b.language, t.avail, b.book_name " & _
                "ORDER BY t.avail DESC, b.language ASC, b.book_name ASC;"
    
    Set rs = conn.Execute(sql)
    
    Dim recordCount As Integer
    recordCount = 0
    
    While Not rs.EOF
        Dim record(4) As Variant
        record(0) = rs.Fields("language").Value
        record(1) = rs.Fields("book_name").Value
        record(2) = rs.Fields("availability_text").Value
        record(3) = rs.Fields("total_qty").Value
        record(4) = ""  ' Empty field for consistency
        
        summaryData.Add record
        recordCount = recordCount + 1
        rs.MoveNext
    Wend
    
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    On Error GoTo 0
    Set GetAllSummaryDataForPDF = summaryData
    Exit Function
    
ErrorHandler:
    Set summaryData = New Collection
    Resume Cleanup
End Function

' Test the All Summary function
Public Sub TestAllSummary()
    On Error GoTo ErrorHandler
    
    Dim result As Collection
    Set result = GetAllSummaryDataForPDF("2025-05-1")
    
    MsgBox "All Summary test completed. Check Immediate Window.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "All Summary test error: " & Err.Description, vbCritical
End Sub

' Generate PDF All Summary Report (1-page City Summary format)
' customFileName: if provided use it (e.g. Next.js style Org_Sector_Period_Summary.pdf)
Public Sub GenerateAllSummaryPDF(summaryData As Collection, period As String, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH", Optional customFileName As String = "")
    On Error GoTo ErrorHandler
    
    ' Store organization name for use in helper functions
    Dim orgName As String
    If organizationName = "" Then
        orgName = "VIMARSH"
    Else
        orgName = organizationName
    End If
    
    ' Disable Excel alerts to prevent merge warnings
    Application.DisplayAlerts = False
    
    ' Create temporary workbook for PDF generation
    Dim pdfWb As Workbook
    Dim pdfWs As Worksheet
    
    Set pdfWb = Workbooks.Add
    Set pdfWs = pdfWb.Sheets(1)
    pdfWs.Name = "City Summary"
    
    ' Initialize variables
    Dim i As Integer
    
    ' Generate formatted summary report
    Dim currentRow As Integer
    currentRow = 1
    
    ' Add City Summary on left and VIMARSH period on right (same line)
    pdfWs.Cells(currentRow, 1).Value = "City Summary"
    pdfWs.Cells(currentRow, 1).Font.Bold = True
    pdfWs.Cells(currentRow, 1).Font.Size = 12
    pdfWs.Cells(currentRow, 1).Font.Name = "Calibri"
    pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlLeft
    
    ' Add VIMARSH period on right side
    If period <> "" And period <> "All" Then
        Dim periodParts() As String
        periodParts = Split(period, "-")
        If UBound(periodParts) >= 2 Then
            pdfWs.Cells(currentRow, 5).Value = orgName & ": " & GetShortMonthName(CLng(periodParts(1))) & "-" & periodParts(0) & "(" & GetHalfRange(periodParts(2), CLng(periodParts(1)), CLng(periodParts(0))) & ") - Summary"
        Else
            pdfWs.Cells(currentRow, 5).Value = orgName & ": " & period & " - Summary"
        End If
    Else
        pdfWs.Cells(currentRow, 5).Value = orgName & ": All Periods - Summary"
    End If
    pdfWs.Cells(currentRow, 5).Font.Bold = True
    pdfWs.Cells(currentRow, 5).Font.Size = 12
    pdfWs.Cells(currentRow, 5).Font.Name = "Calibri"
    pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlRight
    
         ' Format the header row
     pdfWs.Range("A" & currentRow & ":E" & currentRow).Interior.Color = RGB(240, 240, 240)
     pdfWs.Rows(currentRow).RowHeight = 18.15  ' Fixed 18.15 pixels height for header row
     currentRow = currentRow + 1
    
    ' Add table headers
    pdfWs.Cells(currentRow, 1).Value = "LG"
    pdfWs.Cells(currentRow, 2).Value = "BOOK TITLE"
    pdfWs.Cells(currentRow, 3).Value = "AVAIL"
    pdfWs.Cells(currentRow, 4).Value = "QTY"
    pdfWs.Cells(currentRow, 5).Value = "SIGN"
    
    ' Format headers
    With pdfWs.Range("A" & currentRow & ":E" & currentRow)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Interior.Color = RGB(220, 220, 220)
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
         ' Set specific alignments
     pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter ' LG - Center
     pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlLeft   ' BOOK TITLE - Left
     pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter ' AVAIL - Center
     pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter ' QTY - Center
     pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter ' SIGN - Center
     pdfWs.Rows(currentRow).RowHeight = 18.15  ' Fixed 18.15 pixels height for table header row
     currentRow = currentRow + 1
    
    ' Sort records by availability, language, and book title
    Dim sortedRecords As Collection
    Set sortedRecords = New Collection
    
    ' Create separate collections for each availability type
    Dim yesRecords As Collection
    Dim notInStockRecords As Collection
    Dim bYesRecords As Collection
    Dim dataRecord As Variant
    Set yesRecords = New Collection
    Set notInStockRecords = New Collection
    Set bYesRecords = New Collection
    
    ' Separate records by availability
    For i = 1 To summaryData.Count
        dataRecord = summaryData.Item(i)
        
        ' Convert "NOT IN STOCK" to "NS" for consistency
        If dataRecord(2) = "NOT IN STOCK" Then
            dataRecord(2) = "NS"
        End If
        
        If dataRecord(2) = "YES" Then
            yesRecords.Add dataRecord
        ElseIf dataRecord(2) = "NS" Then
            notInStockRecords.Add dataRecord
        ElseIf dataRecord(2) = "B-YES" Then
            bYesRecords.Add dataRecord
        End If
    Next i
    
    ' Sort each availability group by language and book title
    Call SortRecordsByLanguageAndBook(yesRecords)
    Call SortRecordsByLanguageAndBook(notInStockRecords)
    Call SortRecordsByLanguageAndBook(bYesRecords)
    
    ' Add YES records first
    For i = 1 To yesRecords.Count
        sortedRecords.Add yesRecords.Item(i)
    Next i
    
    ' Add NOT IN STOCK records second
    For i = 1 To notInStockRecords.Count
        sortedRecords.Add notInStockRecords.Item(i)
    Next i
    
    ' Add B-YES records third
    For i = 1 To bYesRecords.Count
        sortedRecords.Add bYesRecords.Item(i)
    Next i
    
    ' Process records (YES first, then NS)
    Dim yesStartRow As Integer
    yesStartRow = 0
    Dim yesTotalQty As Integer
    yesTotalQty = 0
    Dim notInStockStartRow As Integer
    notInStockStartRow = 0
    Dim notInStockTotalQty As Integer
    notInStockTotalQty = 0
    Dim totalQty As Integer
    totalQty = 0
    Dim currentLanguage As String
    currentLanguage = ""
    
    For i = 1 To sortedRecords.Count
        dataRecord = sortedRecords.Item(i)
        
        ' Get language short code
        Dim currentRecordLangCode As String
        currentRecordLangCode = GetLanguageShortCode(CStr(dataRecord(0)))
        
        ' Check if language has changed (compare short codes)
        If currentRecordLangCode <> currentLanguage Then
            ' If we were processing records, merge the previous language group
            If currentLanguage <> "" Then
                ' Merge LG column for the previous language group
                If currentRow > 1 Then
                    Dim languageStartRow As Integer
                    languageStartRow = currentRow - 1
                    ' Find the start of current language group (compare short codes)
                    While languageStartRow > 1 And pdfWs.Cells(languageStartRow - 1, 1).Value = currentLanguage
                        languageStartRow = languageStartRow - 1
                    Wend
                    If languageStartRow < currentRow - 1 Then
                        pdfWs.Range("A" & languageStartRow & ":A" & (currentRow - 1)).Merge
                        pdfWs.Range("A" & languageStartRow & ":A" & (currentRow - 1)).HorizontalAlignment = xlCenter
                        pdfWs.Range("A" & languageStartRow & ":A" & (currentRow - 1)).VerticalAlignment = xlCenter
                    End If
                End If
            End If
            currentLanguage = currentRecordLangCode
        End If
        
        pdfWs.Cells(currentRow, 1).Value = currentRecordLangCode  ' Language short code
        pdfWs.Cells(currentRow, 2).Value = dataRecord(1)  ' Book Name
        pdfWs.Cells(currentRow, 3).Value = dataRecord(2)  ' Availability
        pdfWs.Cells(currentRow, 4).Value = dataRecord(3)  ' Quantity
        pdfWs.Cells(currentRow, 5).Value = ""             ' Sign (empty)
        
        ' Make NS text light/gray in individual cells
        If dataRecord(2) = "NS" Then
            pdfWs.Cells(currentRow, 3).Font.Color = RGB(150, 150, 150)
            pdfWs.Cells(currentRow, 3).Font.Bold = False
        End If
        
        ' Track totals for each availability type
        If dataRecord(2) = "YES" Then
            yesTotalQty = yesTotalQty + CLng(dataRecord(3))
            If yesStartRow = 0 Then
                yesStartRow = currentRow
            End If
            ' Don't merge NS cells - just reset the tracking
            notInStockStartRow = 0
            notInStockTotalQty = 0
        Else
            notInStockTotalQty = notInStockTotalQty + CLng(dataRecord(3))
            If notInStockStartRow = 0 Then
                notInStockStartRow = currentRow
            End If
            ' If we were in a YES sequence, merge YES section
            If yesStartRow > 0 And yesStartRow < currentRow Then
                ' Merge AVAIL column for YES section
                pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).Merge
                pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).HorizontalAlignment = xlCenter
                pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).VerticalAlignment = xlCenter
                pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).Value = "YES"
                
                ' Merge SIGN column for YES section and add total quantity in brackets
                pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).Merge
                pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).Value = "[" & yesTotalQty & "]"
                pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).HorizontalAlignment = xlLeft
                pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).VerticalAlignment = xlCenter
                
                ' Note: Language merging is handled separately when language changes
            End If
            yesStartRow = 0
            yesTotalQty = 0
        End If
        
        ' Format data row
        With pdfWs.Range("A" & currentRow & ":E" & currentRow)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
        End With
        
        ' Set specific alignments
        pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter ' LG - Center
        pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlLeft   ' BOOK TITLE - Left
        pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter ' AVAIL - Center
        pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter ' QTY - Center
        pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter ' SIGN - Center
        
        pdfWs.Rows(currentRow).RowHeight = 18.15  ' Fixed 18.15 pixels height for data rows
        
        totalQty = totalQty + CLng(dataRecord(3))
        currentRow = currentRow + 1
    Next i
    
    ' Handle merge cells for YES values at the end
    If yesStartRow > 0 And yesStartRow < currentRow Then
        ' Merge AVAIL column for YES section
        pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).Merge
        pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).HorizontalAlignment = xlCenter
        pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).VerticalAlignment = xlCenter
        pdfWs.Range("C" & yesStartRow & ":C" & (currentRow - 1)).Value = "YES"
        
        ' Merge SIGN column for YES section and add total quantity in brackets
        pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).Merge
        pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).Value = "[" & yesTotalQty & "]"
        pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).HorizontalAlignment = xlLeft
        pdfWs.Range("E" & yesStartRow & ":E" & (currentRow - 1)).VerticalAlignment = xlCenter
    End If
    
    ' Note: NS cells are NOT merged - they remain individual cells with light font
    
    ' Handle merge cells for the last language group at the end (compare short codes)
    If currentLanguage <> "" And currentRow > 1 Then
        Dim lastLanguageStartRow As Integer
        lastLanguageStartRow = currentRow - 1
        ' Find the start of last language group (compare short codes)
        While lastLanguageStartRow > 1 And pdfWs.Cells(lastLanguageStartRow - 1, 1).Value = currentLanguage
            lastLanguageStartRow = lastLanguageStartRow - 1
        Wend
        If lastLanguageStartRow < currentRow - 1 Then
            pdfWs.Range("A" & lastLanguageStartRow & ":A" & (currentRow - 1)).Merge
            pdfWs.Range("A" & lastLanguageStartRow & ":A" & (currentRow - 1)).HorizontalAlignment = xlCenter
            pdfWs.Range("A" & lastLanguageStartRow & ":A" & (currentRow - 1)).VerticalAlignment = xlCenter
        End If
    End If
    
    ' Add total line at the bottom
    pdfWs.Cells(currentRow, 3).Value = "Total Qty:"
    pdfWs.Cells(currentRow, 3).Font.Bold = True
    pdfWs.Cells(currentRow, 3).Font.Name = "Calibri"
    pdfWs.Cells(currentRow, 3).Font.Size = 10
    pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlRight
    
    pdfWs.Cells(currentRow, 4).Value = totalQty
    pdfWs.Cells(currentRow, 4).Font.Bold = True
    pdfWs.Cells(currentRow, 4).Font.Name = "Calibri"
    pdfWs.Cells(currentRow, 4).Font.Size = 10
    pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter
    
         ' Format the total row
     With pdfWs.Range("A" & currentRow & ":E" & currentRow)
         .Interior.Color = RGB(240, 240, 240)
         .Borders.LineStyle = xlContinuous
         .Borders.Weight = xlThin
         .VerticalAlignment = xlCenter
     End With
     pdfWs.Rows(currentRow).RowHeight = 18.15  ' Fixed 18.15 pixels height for total row
    
    ' Set column widths to match details report (exact pixel values)
    pdfWs.Columns("A").ColumnWidth = 8   ' LNG (62 pixels)
    pdfWs.Columns("B").ColumnWidth = 35  ' BOOK TITLE (230 pixels)
    pdfWs.Columns("C").ColumnWidth = 14  ' AVAIL (96 pixels)
    pdfWs.Columns("D").ColumnWidth = 8   ' QTY (50 pixels)
    pdfWs.Columns("E").ColumnWidth = 25  ' SIGN (317 pixels) - Made wider
    
    ' Set page format to A4 Portrait with minimal margins
    With pdfWs.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .TopMargin = Application.InchesToPoints(0.1)    ' Minimal margin at top
        .BottomMargin = Application.InchesToPoints(0.1)  ' Minimal margin at bottom
        .LeftMargin = Application.InchesToPoints(0.5)    ' More margin on left for punching
        .RightMargin = Application.InchesToPoints(0.1)   ' Minimal margin on right
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    
    ' Set print area
    Dim lastRow As Long
    lastRow = pdfWs.Cells(pdfWs.Rows.Count, 1).End(xlUp).Row
    pdfWs.PageSetup.PrintArea = "A1:E" & lastRow
    
    ' Save as PDF
    Dim pdfPath As String
    Dim fileName As String
    If customFileName <> "" Then
        fileName = customFileName
    Else
        Dim safeOrgSc As String
        safeOrgSc = Replace(Replace(orgName, " ", "_"), "\", "_")
        fileName = safeOrgSc & "_city_summary_" & period & ".pdf"
    End If
    If customFolder <> "" Then
        pdfPath = customFolder & "\" & fileName
    Else
        pdfPath = ThisWorkbook.Path & "\" & fileName
    End If
    
    pdfWb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False
    
    ' Close temporary workbook
    pdfWb.Close False
    
    ' Re-enable Excel alerts
    Application.DisplayAlerts = True
    
    ' Individual success message removed - only final message will be shown
    
    Exit Sub
    
ErrorHandler:
    ' Re-enable Excel alerts even if error occurs
    Application.DisplayAlerts = True
    MsgBox "Error generating all summary PDF: " & Err.Description, vbCritical
End Sub

' Helper function to get month name
Private Function GetMonthName(monthNum As Long) As String
    Dim monthNames() As String
    monthNames = Split("JANUARY,FEBRUARY,MARCH,APRIL,MAY,JUNE,JULY,AUGUST,SEPTEMBER,OCTOBER,NOVEMBER,DECEMBER", ",")
    GetMonthName = monthNames(monthNum - 1)
End Function

' Helper function to get short month name (3 characters)
Private Function GetShortMonthName(monthNum As Long) As String
    Dim monthNames() As String
    monthNames = Split("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC", ",")
    GetShortMonthName = monthNames(monthNum - 1)
End Function

' Helper function to get half range
Private Function GetHalfRange(half As String, monthNum As Long, Optional yearNum As Long = 0) As String
    ' Calculate number of days in the month based on year and month
    Dim daysInMonth As Integer
    Select Case monthNum
        Case 1, 3, 5, 7, 8, 10, 12
            daysInMonth = 31
        Case 4, 6, 9, 11
            daysInMonth = 30
        Case 2
            ' February - calculate based on leap year
            If yearNum > 0 Then
                ' Check if leap year: divisible by 4, but not by 100 unless also divisible by 400
                If (yearNum Mod 4 = 0 And yearNum Mod 100 <> 0) Or (yearNum Mod 400 = 0) Then
                    daysInMonth = 29
                Else
                    daysInMonth = 28
                End If
            Else
                ' If year not provided, default to 28
                daysInMonth = 28
            End If
        Case Else
            daysInMonth = 30
    End Select
    
    ' When half is "1 & 2", show full month range
    If Trim(LCase(half)) = "1 & 2" Or Trim(LCase(half)) = "1&2" Then
        GetHalfRange = "1-" & daysInMonth
    ElseIf half = "1" Then
        ' First half: 1-15
        GetHalfRange = "1-15"
    ElseIf half = "2" Then
        ' Second half: 16 to last day of month
        GetHalfRange = "16-" & daysInMonth
    Else
        ' Default fallback - use calculated days in month
        GetHalfRange = "1-" & daysInMonth
    End If
End Function

' Helper function to get language short code
Private Function GetLanguageShortCode(ByVal langName As String) As String
    Select Case LCase(Trim(langName))
        Case "gujarati"
            GetLanguageShortCode = "GJ"
        Case "hindi"
            GetLanguageShortCode = "HN"
        Case "marathi"
            GetLanguageShortCode = "MR"
        Case "telugu", "telagu"
            GetLanguageShortCode = "TL"
        Case "tamil"
            GetLanguageShortCode = "TM"
        Case Else
            GetLanguageShortCode = Left(langName, 2)
    End Select
End Function

' Helper function to sort records by language and book title
Private Sub SortRecordsByLanguageAndBook(records As Collection)
    If records.Count <= 1 Then 
        Exit Sub 
    End If
    
    ' Convert collection to array for easier sorting
    Dim recordArray() As Variant
    ReDim recordArray(1 To records.Count)
    
    Dim i As Integer
    For i = 1 To records.Count
        recordArray(i) = records.Item(i)
    Next i
    
    ' Simple bubble sort on array
    Dim j As Integer
    Dim tempRecord As Variant
    
    For i = 1 To records.Count - 1
        For j = i + 1 To records.Count
            ' Compare language first (index 0)
            If recordArray(i)(0) > recordArray(j)(0) Then
                ' Swap records
                tempRecord = recordArray(i)
                recordArray(i) = recordArray(j)
                recordArray(j) = tempRecord
            ElseIf recordArray(i)(0) = recordArray(j)(0) Then
                ' If language is same, compare book title (index 1)
                If recordArray(i)(1) > recordArray(j)(1) Then
                    ' Swap records
                    tempRecord = recordArray(i)
                    recordArray(i) = recordArray(j)
                    recordArray(j) = tempRecord
                End If
            End If
        Next j
    Next i
    
    ' Clear collection and add sorted records back
    Set records = New Collection
    For i = 1 To UBound(recordArray)
        records.Add recordArray(i)
    Next i
End Sub
