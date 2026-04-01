Option Explicit

' Sector-Wise Stock PDF Generation Module
' This module generates sector-wise PDF reports for books with P No.
' Groups by book_id + period (each book_id appears once per period)
' Then organizes by sector for PDF output

Public Sub GenerateSectorWiseStockPDF()
    On Error GoTo ErrorHandler
    
    ' Declare variables at function level
    Dim i As Long
    Dim pNoFilter As String
    Dim userInput As Variant
    Dim pNoArray() As String
    Dim pNoCount As Long
    Dim pNoFilterClause As String
    
    ' Get P No. filter from user (comma-separated, e.g., "44,45,51")
    userInput = Application.InputBox("Enter P No. to filter (comma-separated, e.g., 44,45,51):" & vbCrLf & _
                          "Leave empty to include all P No.", _
                          "Filter P No.", "", Type:=2)
    
    ' If user cancelled (Application.InputBox returns False on cancel), exit
    If userInput = False Then
        Exit Sub
    End If
    
    pNoFilter = CStr(userInput)
    
    ' If empty input, show error and exit
    If Trim(pNoFilter) = "" Then
        MsgBox "P No. filter is required. Please enter at least one P No.", vbExclamation, "Input Required"
        Exit Sub
    End If
    
    ' Parse P No. filter into array
    pNoFilterClause = ""
    
    If Trim(pNoFilter) <> "" Then
        ' Split comma-separated values
        pNoArray = Split(Replace(Replace(pNoFilter, " ", ""), vbTab, ""), ",")
        pNoCount = UBound(pNoArray) + 1
        
        ' Build SQL IN clause
        If pNoCount > 0 Then
            pNoFilterClause = "AND t.stock_number IN ("
            For i = 0 To pNoCount - 1
                If Trim(pNoArray(i)) <> "" Then
                    If i > 0 Then pNoFilterClause = pNoFilterClause & ", "
                    pNoFilterClause = pNoFilterClause & "'" & Replace(Trim(pNoArray(i)), "'", "''") & "'"
                End If
            Next i
            pNoFilterClause = pNoFilterClause & ") "
        End If
    End If
    
    ' Get save file path
    Dim savePath As String
    Dim defaultFileName As String
    defaultFileName = "StockReport_SectorWise_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
    
    ' Connect to database to get accurate sector-wise data with P No.
    ' Group by book_id + period (same as form) - each book_id appears once per period
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"
    
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        MsgBox "Connection failed: " & Err.description, vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Query database: Group by book_id + period + sector + stock_number
    ' Only include transactions with P No. (stock_number)
    ' Each unique combination of book_id + period + sector + P No. appears once
    ' If same book+period+sector has different P No., they appear as separate rows
    Dim sql As String
    sql = "SELECT " & _
          "b.book_id, " & _
          "b.book_name, " & _
          "b.language, " & _
          "t.year, " & _
          "t.month, " & _
          "t.half, " & _
          "t.group_sector, " & _
          "SUM(t.qty) as total_qty, " & _
          "t.stock_number " & _
          "FROM vimars.transactions t " & _
          "INNER JOIN vimars.books b ON t.book_id = b.book_id " & _
          "WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE " & _
          "AND t.avail = 0 " & _
          "AND COALESCE(t.stock_number, '') <> '' " & _
          "AND t.group_sector IS NOT NULL AND t.group_sector <> '' " & _
          pNoFilterClause & _
          "GROUP BY b.book_id, b.book_name, b.language, t.year, t.month, t.half, t.group_sector, t.stock_number " & _
          "ORDER BY t.group_sector ASC, b.language ASC, b.book_name ASC, t.year DESC, t.month DESC, t.half DESC;"
    
    Dim rs As Object
    Set rs = conn.Execute(sql)
    
    ' Group data by sector for PDF organization
    Dim sectorGroups As Object
    Set sectorGroups = CreateObject("Scripting.Dictionary")
    Dim sectorBooks As Collection
    Dim bookEntry(9) As Variant  ' bookName(0), language(1), period(2), qty(3), pNo(4), sector(5), book_id(6), year(7), month(8), half(9)
    Dim bookName As String
    Dim language As String
    Dim pNo As String
    Dim bookId As Long
    
    While Not rs.EOF
        Dim sector As String
        Dim year As Integer
        Dim monthNum As Integer
        Dim half As String
        Dim qtyVal As Long
        Dim periodDisplay As String
        
        sector = rs.Fields("group_sector").Value & ""
        bookName = rs.Fields("book_name").Value & ""
        language = rs.Fields("language").Value & ""
        year = CInt(rs.Fields("year").Value)
        monthNum = CInt(rs.Fields("month").Value)
        half = CStr(rs.Fields("half").Value)
        qtyVal = CLng(rs.Fields("total_qty").Value)
        pNo = rs.Fields("stock_number").Value & ""
        bookId = CLng(rs.Fields("book_id").Value)
        
        ' Format period display
        Dim monthNameStr As String
        monthNameStr = Left(MonthName(monthNum), 3)
        
        ' Handle "1 & 2" for full month
        Dim halfTrimmed As String
        halfTrimmed = Trim(LCase(half))
        
        If halfTrimmed = "1 & 2" Or halfTrimmed = "1&2" Then
            ' Full month range
            Dim lastDayFull As Integer
            lastDayFull = Day(DateSerial(year, monthNum + 1, 0))
            periodDisplay = year & "-" & monthNameStr & " (1 To " & lastDayFull & ")"
        ElseIf half = "1" Then
            periodDisplay = year & "-" & monthNameStr & " (1 To 15)"
        ElseIf half = "2" Then
            Dim lastDay As Integer
            lastDay = Day(DateSerial(year, monthNum + 1, 0))
            periodDisplay = year & "-" & monthNameStr & " (16 To " & lastDay & ")"
        Else
            periodDisplay = year & "-" & monthNameStr & " (" & year & "-" & Right("0" & monthNum, 2) & "-" & half & ")"
        End If
        
        ' Add to sector group
        If Not sectorGroups.Exists(sector) Then
            Set sectorBooks = New Collection
            sectorGroups.Add sector, sectorBooks
        Else
            Set sectorBooks = sectorGroups(sector)
        End If
        
        ' Store book data (including book_id for village query)
        bookEntry(0) = bookName
        bookEntry(1) = language
        bookEntry(2) = periodDisplay
        bookEntry(3) = CStr(qtyVal)
        bookEntry(4) = pNo
        bookEntry(5) = sector
        bookEntry(6) = bookId  ' Store book_id for village query
        bookEntry(7) = year  ' Store year for village query
        bookEntry(8) = monthNum  ' Store month for village query
        bookEntry(9) = half  ' Store half for village query
        sectorBooks.Add bookEntry
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    ' Check if we have any data
    If sectorGroups.count = 0 Then
        MsgBox "No data available with P No. to generate PDF.", vbExclamation
        Exit Sub
    End If
    
    ' Use Excel to create PDF
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    ' Set default vertical alignment to center for all cells - use xlCenter constant
    xlWorksheet.Cells.VerticalAlignment = xlCenter
    
    ' i is already declared at function level
    Dim sectorKey As String
    Dim period As String
    Dim qty As String
    Dim bookEntryVar As Variant  ' Variant to hold array from collection
    Dim todayDateStr As String  ' Today's date for sector header
    
    ' Generate PDF grouped by sector
    Dim rowNum As Long
    rowNum = 1
    Dim rowsPerPage As Long
    rowsPerPage = 50  ' Approximate rows per A4 page
    
    ' Set today's date once
    todayDateStr = Format(Date, "dd-mm-yyyy")
    
    ' Get sorted sector keys
    Dim sectorKeys As Variant
    sectorKeys = sectorGroups.Keys
    
    ' Sort sectors alphabetically
    Dim sortedSectors As Collection
    Set sortedSectors = New Collection
    Dim tempArray() As String
    ReDim tempArray(0 To sectorGroups.count - 1)
    Dim idx As Long
    idx = 0
    Dim sectorKeyItem As Variant
    For Each sectorKeyItem In sectorKeys
        tempArray(idx) = CStr(sectorKeyItem)
        idx = idx + 1
    Next
    
    ' Simple bubble sort
    Dim k As Long
    Dim l As Long
    Dim tempStr As String
    For k = 0 To UBound(tempArray) - 1
        For l = k + 1 To UBound(tempArray)
            If tempArray(k) > tempArray(l) Then
                tempStr = tempArray(k)
                tempArray(k) = tempArray(l)
                tempArray(l) = tempStr
            End If
        Next l
    Next k
    
    ' Track page start row and sectors per page for "Summary: Sector" section
    Dim currentPageStartRow As Long
    currentPageStartRow = 1
    Dim sectorsOnCurrentPage As Integer
    sectorsOnCurrentPage = 0
    Dim currentPageUsedRows As Long
    currentPageUsedRows = 0
    Dim maxRowsPerPage As Long
    maxRowsPerPage = 45  ' Same as SectorSummaryReport.bas
    Dim previousSectorRowCount As Long
    previousSectorRowCount = 0
    
    ' Process each sector
    For k = 0 To UBound(tempArray)
        sectorKey = tempArray(k)
        Set sectorBooks = sectorGroups(sectorKey)
        
        ' Sort books within this sector: first by language, then by book name (A to Z)
        Dim sortedSectorBooks As Collection
        Set sortedSectorBooks = New Collection
        Dim bookArray() As Variant
        ReDim bookArray(1 To sectorBooks.count)
        Dim bookIdx As Long
        bookIdx = 1
        For bookIdx = 1 To sectorBooks.count
            bookArray(bookIdx) = sectorBooks.item(bookIdx)
        Next bookIdx
        
        ' Sort by language, then book name
        Dim sortI As Long
        Dim sortJ As Long
        Dim tempBook As Variant
        For sortI = 1 To UBound(bookArray) - 1
            For sortJ = sortI + 1 To UBound(bookArray)
                Dim book1 As Variant
                Dim book2 As Variant
                book1 = bookArray(sortI)
                book2 = bookArray(sortJ)
                
                Dim lang1 As String
                Dim lang2 As String
                Dim name1 As String
                Dim name2 As String
                lang1 = CStr(book1(1))
                lang2 = CStr(book2(1))
                name1 = CStr(book1(0))
                name2 = CStr(book2(0))
                
                ' Compare: first by language (using short code like SectorSummaryReport), then by book name
                Dim langCode1 As String
                Dim langCode2 As String
                langCode1 = GetLanguageShortCode(lang1)
                langCode2 = GetLanguageShortCode(lang2)
                
                Dim shouldSwap As Boolean
                shouldSwap = False
                If langCode1 > langCode2 Then
                    shouldSwap = True
                ElseIf langCode1 = langCode2 Then
                    If name1 > name2 Then
                        shouldSwap = True
                    End If
                End If
                
                If shouldSwap Then
                    tempBook = bookArray(sortI)
                    bookArray(sortI) = bookArray(sortJ)
                    bookArray(sortJ) = tempBook
                End If
            Next sortJ
        Next sortI
        
        ' Create sorted collection
        For bookIdx = 1 To UBound(bookArray)
            sortedSectorBooks.Add bookArray(bookIdx)
        Next bookIdx
        
        ' Use sorted collection
        Set sectorBooks = sortedSectorBooks
        
        ' Calculate estimated table size (1st copy only for now)
        Dim estimatedTableSize As Long
        estimatedTableSize = 3 + sectorBooks.count  ' Header + headers + data + total
        
        ' Check if we need a new page (max 2 sectors per page for "Summary: Sector" section)
        Dim needNewPage As Boolean
        needNewPage = False
        
        If rowNum > 1 Then  ' Not the first sector
            ' If we already have 2 sectors on this page, force new page
            If sectorsOnCurrentPage >= 2 Then
                needNewPage = True
            ElseIf sectorsOnCurrentPage = 1 Then
                ' Allow 2nd sector on same page, but check size limits
                If previousSectorRowCount > 23 Then
                    needNewPage = True   ' Force new page because previous was too large (>23 rows)
                ElseIf estimatedTableSize > 23 Then
                    needNewPage = True   ' Force new page because current is too large (>23 rows)
                Else
                    needNewPage = False  ' Keep on same page (both =23 rows)
                End If
            End If
        End If
        
        ' Add page break if needed
        If needNewPage Then
            xlWorksheet.HPageBreaks.Add xlWorksheet.Cells(rowNum, 1)
            sectorsOnCurrentPage = 0
            currentPageStartRow = rowNum
            currentPageUsedRows = 0  ' Reset page usage counter
        ElseIf sectorsOnCurrentPage = 1 Then
            ' If this is the 2nd table on the same page, position header at row 25 (relative to page start)
            Dim targetRow As Long
            targetRow = currentPageStartRow + 24   ' Row 25 of the page
            
            Dim spacingRows As Long
            spacingRows = targetRow - rowNum
            
            ' Add spacing if needed to reach the target row
            If spacingRows > 0 Then
                Dim spacingIdx As Long
                For spacingIdx = 1 To spacingRows
                    xlWorksheet.Rows(rowNum).RowHeight = 17.5
                    rowNum = rowNum + 1
                Next spacingIdx
                currentPageUsedRows = currentPageUsedRows + spacingRows
            End If
        End If
        
        ' Track sector start row
        Dim sectorStartRow As Long
        sectorStartRow = rowNum
        
        ' VIMARSH only - no period value
        Dim sectorPeriodDisplay As String
        sectorPeriodDisplay = "VIMARSH"
        
        ' Add sector header (1st copy) - matching SectorSummaryReport.bas format
        ' Left side: Sector name, Center: VIMARSH period, Right side: Summary: Sector
        xlWorksheet.Cells(rowNum, 1).Value = sectorKey
        xlWorksheet.Cells(rowNum, 1).Font.Bold = True
        xlWorksheet.Cells(rowNum, 1).Font.Size = 12
        xlWorksheet.Cells(rowNum, 1).Font.Name = "Calibri"
        xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
        
        ' Center: VIMARSH period
        xlWorksheet.Cells(rowNum, 3).Value = sectorPeriodDisplay
        xlWorksheet.Cells(rowNum, 3).Font.Bold = True
        xlWorksheet.Cells(rowNum, 3).Font.Size = 12
        xlWorksheet.Cells(rowNum, 3).Font.Name = "Calibri"
        xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4108  ' xlCenter
        
        ' Right side: Pending -> Clear Summary : Sector
        xlWorksheet.Cells(rowNum, 7).Value = "Pending -> Clear Summary : Sector"
        xlWorksheet.Cells(rowNum, 7).Font.Bold = True
        xlWorksheet.Cells(rowNum, 7).Font.Size = 12
        xlWorksheet.Cells(rowNum, 7).Font.Name = "Calibri"
        xlWorksheet.Cells(rowNum, 7).HorizontalAlignment = -4152  ' xlRight
        
        ' Format the header row with background color
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
            .Interior.Color = RGB(240, 240, 240)
            .VerticalAlignment = xlCenter  ' xlCenter - Middle align
        End With
        xlWorksheet.Rows(rowNum).RowHeight = 17.5
        rowNum = rowNum + 1
        
        ' Add table headers for this sector (1st copy)
        ' Column order: P No., Book Name, Language, Period, Qty, Sign, VK Books
        xlWorksheet.Cells(rowNum, 1).Value = "P No."
        xlWorksheet.Cells(rowNum, 2).Value = "Book Name"
        xlWorksheet.Cells(rowNum, 3).Value = "LG"
        xlWorksheet.Cells(rowNum, 4).Value = "Period"
        xlWorksheet.Cells(rowNum, 5).Value = "Qty"
        xlWorksheet.Cells(rowNum, 6).Value = "Sign"
        xlWorksheet.Cells(rowNum, 7).Value = "VK Books"
        
        ' Format header row - exactly like SectorVillageDetailsReport.bas AddTableHeaders
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Interior.Color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlHairline
            .VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
            ' Bold outer border for header
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin
        End With
        
        ' Set header alignments
        xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4108  ' xlCenter - P No.
        xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4131  ' xlLeft - Book Name
        xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4108  ' xlCenter - Language
        xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4131  ' xlLeft - Period
        xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4108  ' xlCenter - Qty
        xlWorksheet.Cells(rowNum, 6).HorizontalAlignment = -4131  ' xlLeft - Sign
        xlWorksheet.Cells(rowNum, 7).HorizontalAlignment = -4131  ' xlLeft - VK Books
        
        rowNum = rowNum + 1
        
        ' Add books for this sector (1st copy) - keep all rows, will merge cells later
        Dim sectorTotalQty As Long
        sectorTotalQty = 0
        Dim dataStartRow As Long
        dataStartRow = rowNum  ' Track start of data rows for this sector
        Dim villageTotalDict As Object
        Set villageTotalDict = CreateObject("Scripting.Dictionary")  ' Track village totals for this table
        
        For i = 1 To sectorBooks.count
            bookEntryVar = sectorBooks.item(i)
            
            bookName = CStr(bookEntryVar(0))
            language = CStr(bookEntryVar(1))
            period = CStr(bookEntryVar(2))
            qty = CStr(bookEntryVar(3))
            pNo = CStr(bookEntryVar(4))
            
            ' Add to total
            sectorTotalQty = sectorTotalQty + CLng(qty)
            
            ' Get village breakdown for this book+period+sector
            Dim bookIdVal As Long
            Dim yearVal As Integer
            Dim monthVal As Integer
            Dim halfVal As String
            Dim sectorVal As String
            Dim villageBreakdown As String
            
            ' Safely extract values from bookEntryVar array
            On Error Resume Next
            If IsNumeric(bookEntryVar(6)) Then
                bookIdVal = CLng(bookEntryVar(6))
            Else
                bookIdVal = 0
            End If
            
            If IsNumeric(bookEntryVar(7)) Then
                yearVal = CInt(bookEntryVar(7))
            Else
                yearVal = 0
            End If
            
            If IsNumeric(bookEntryVar(8)) Then
                monthVal = CInt(bookEntryVar(8))
            Else
                monthVal = 0
            End If
            
            halfVal = CStr(bookEntryVar(9) & "")
            sectorVal = CStr(bookEntryVar(5) & "")
            On Error GoTo 0
            
            ' Only get village breakdown if we have valid data
            If bookIdVal > 0 And yearVal > 0 And monthVal > 0 Then
                villageBreakdown = GetVillageBreakdown(bookIdVal, yearVal, monthVal, halfVal, sectorVal, pNo)
            Else
                villageBreakdown = ""
            End If
            
            ' Add book entry - Column order: P No., Book Name, Language, Period, Qty, Sign, VK Books
            xlWorksheet.Cells(rowNum, 1).Value = pNo
            xlWorksheet.Cells(rowNum, 2).Value = bookName
            xlWorksheet.Cells(rowNum, 3).Value = GetLanguageShortCode(language)  ' Language short code (like LG in SectorSummaryReport)
            xlWorksheet.Cells(rowNum, 4).Value = period
            xlWorksheet.Cells(rowNum, 5).Value = qty
            xlWorksheet.Cells(rowNum, 6).Value = ""  ' Sign column - will be filled after merging
            xlWorksheet.Cells(rowNum, 7).Value = villageBreakdown  ' VK Books - village breakdown
            
            xlWorksheet.Cells(rowNum, 5).Font.Bold = True  ' Qty bold
            
            ' Format data row - matching SectorVillageDetailsReport.bas pattern
            With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlHairline
                .VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
            End With
            
            ' Set specific horizontal alignments for data cells (vertical already set on range above)
            xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4108  ' xlCenter - P No.
            xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4131  ' xlLeft - Book Name
            xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4108  ' xlCenter - Language
            xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4131  ' xlLeft - Period
            xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4108  ' xlCenter - Qty
            xlWorksheet.Cells(rowNum, 6).HorizontalAlignment = -4131  ' xlLeft - Sign
            xlWorksheet.Cells(rowNum, 7).HorizontalAlignment = -4131  ' xlLeft - VK Books
            
            ' Set font size 8 and ShrinkToFit for VK Books column
            xlWorksheet.Cells(rowNum, 7).Font.Size = 8
            xlWorksheet.Cells(rowNum, 7).ShrinkToFit = True
            
            ' Set ShrinkToFit for Book Name column
            xlWorksheet.Cells(rowNum, 2).ShrinkToFit = True
            
            ' Parse and accumulate village breakdown for total row
            If villageBreakdown <> "" Then
                Call ParseAndAccumulateVillageBreakdown(villageBreakdown, villageTotalDict)
            End If
            
            rowNum = rowNum + 1
        Next i
        
        Dim dataEndRow As Long
        dataEndRow = rowNum - 1  ' Track end of data rows (before total row)
        
        ' Apply cell merging for same book name + language (like SectorVillageDetailsReport.bas)
        Call ApplyCellMergingForSectorTable(xlWorksheet, dataStartRow, dataEndRow)
        
        ' Build village total breakdown string
        Dim villageTotalBreakdown As String
        villageTotalBreakdown = BuildVillageTotalBreakdown(villageTotalDict)
        
        ' Add total row for 1st copy
        xlWorksheet.Cells(rowNum, 1).Value = ""
        xlWorksheet.Cells(rowNum, 2).Value = ""
        xlWorksheet.Cells(rowNum, 3).Value = ""
        xlWorksheet.Cells(rowNum, 4).Value = "Total Qty"
        xlWorksheet.Cells(rowNum, 5).Value = sectorTotalQty
        xlWorksheet.Cells(rowNum, 6).Value = ""  ' Sign column - will be merged with VK Books
        xlWorksheet.Cells(rowNum, 7).Value = villageTotalBreakdown  ' VK Books - village totals
        
        ' Merge Sign (column 6) and VK Books (column 7) cells in total row
        xlWorksheet.Range(xlWorksheet.Cells(rowNum, 6), xlWorksheet.Cells(rowNum, 7)).Merge
        xlWorksheet.Cells(rowNum, 6).Value = villageTotalBreakdown  ' Put village totals in merged cell
        xlWorksheet.Cells(rowNum, 6).HorizontalAlignment = -4108  ' xlCenter - center align
        xlWorksheet.Cells(rowNum, 6).VerticalAlignment = xlCenter  ' xlCenter - center align vertically
        xlWorksheet.Cells(rowNum, 6).Font.Size = 8
        xlWorksheet.Cells(rowNum, 6).ShrinkToFit = True
        
        ' Format total row - exactly like SectorVillageDetailsReport.bas pattern
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
            .Interior.Color = RGB(240, 240, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlHairline
            .VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
        End With
        
        xlWorksheet.Cells(rowNum, 4).Font.Bold = True
        xlWorksheet.Cells(rowNum, 5).Font.Bold = True
        xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4131  ' xlLeft
        xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4108  ' xlCenter
        
        rowNum = rowNum + 1
        
        ' Store the actual row count used by this sector (only 1 copy in "Summary: Sector" section)
        Dim actualSectorRows As Long
        actualSectorRows = rowNum - sectorStartRow
        previousSectorRowCount = actualSectorRows
        
        ' Update page usage counter
        currentPageUsedRows = currentPageUsedRows + actualSectorRows
        
        ' Handle page breaks - ensure next sector starts on new page if we already have 2 sectors on current page
        ' Check if there's a next sector
        If k < UBound(tempArray) Then
            ' Check if we need a new page (max 2 sectors per page for "Summary: Sector" section)
            If sectorsOnCurrentPage >= 2 Then
                ' Already have 2 sectors on this page, force new page
                xlWorksheet.HPageBreaks.Add xlWorksheet.Cells(rowNum, 1)
                sectorsOnCurrentPage = 0
                currentPageStartRow = rowNum
                currentPageUsedRows = 0  ' Reset page usage counter
            End If
        End If
        
        ' Increment sector counter for this page
        sectorsOnCurrentPage = sectorsOnCurrentPage + 1
        
        ' Update firstTableStartRow for next iteration (if this was first table on page)
        If sectorsOnCurrentPage = 1 Then
            currentPageStartRow = sectorStartRow
        End If
        
    Next k
    
    ' Now add duplicate tables for all sectors with "Summary: Sector - City" format
    ' Multiple sectors continue (like SectorSummaryReport.bas)
    Dim previousSectorKeyForDup As String
    previousSectorKeyForDup = ""
    Dim currentPageUsedRowsDup As Long
    currentPageUsedRowsDup = 0
    
    ' Page limit for duplicate section (allowing multiple sectors to continue)
    Dim maxRowsPerPageDup As Long
    maxRowsPerPageDup = 45  ' Same as SectorSummaryReport.bas
    
    ' Add page break before duplicate section
    If rowNum > 1 Then
        xlWorksheet.HPageBreaks.Add xlWorksheet.Cells(rowNum, 1)
    End If
    
    ' Process all sectors again for "Summary: Sector - City" format
    For k = 0 To UBound(tempArray)
        sectorKey = tempArray(k)
        Set sectorBooks = sectorGroups(sectorKey)
        
        ' Calculate table size
        Dim dupTableSize As Long
        dupTableSize = 3 + sectorBooks.count  ' Header + headers + data + total
        
        ' Add 2 rows spacing between sectors (except first sector)
        Dim spacingRowsNeededDup As Long
        spacingRowsNeededDup = 0
        If previousSectorKeyForDup <> "" And sectorKey <> previousSectorKeyForDup Then
            spacingRowsNeededDup = 2
        End If
        
        ' Calculate total rows needed including spacing
        Dim totalRowsNeededDup As Long
        totalRowsNeededDup = spacingRowsNeededDup + dupTableSize
        
        ' Check if we need a new page (optimize to fit as many as possible)
        If rowNum > 1 Then
            If (currentPageUsedRowsDup + totalRowsNeededDup) >= maxRowsPerPageDup Then
                ' Table won't fit completely - add page break BEFORE starting
                xlWorksheet.HPageBreaks.Add xlWorksheet.Cells(rowNum, 1)
                currentPageUsedRowsDup = 0  ' Reset page usage counter
                spacingRowsNeededDup = 0  ' No spacing needed after page break
            End If
        End If
        
        ' Add spacing rows between sectors (except first sector and after page break)
        If spacingRowsNeededDup > 0 Then
            Dim spacingIdxDup As Long
            For spacingIdxDup = 1 To spacingRowsNeededDup
                xlWorksheet.Rows(rowNum).RowHeight = 17.5
                rowNum = rowNum + 1
            Next spacingIdxDup
            currentPageUsedRowsDup = currentPageUsedRowsDup + spacingRowsNeededDup  ' Track spacing rows
        End If
        
        ' Track table start row
        Dim dupTableStartRow As Long
        dupTableStartRow = rowNum
        
        ' Add duplicate table header with "Summary: Sector - City" format
        xlWorksheet.Cells(rowNum, 1).Value = sectorKey
        xlWorksheet.Cells(rowNum, 1).Font.Bold = True
        xlWorksheet.Cells(rowNum, 1).Font.Size = 12
        xlWorksheet.Cells(rowNum, 1).Font.Name = "Calibri"
        xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
        
        ' Center: VIMARSH only
        xlWorksheet.Cells(rowNum, 3).Value = "VIMARSH"
        xlWorksheet.Cells(rowNum, 3).Font.Bold = True
        xlWorksheet.Cells(rowNum, 3).Font.Size = 12
        xlWorksheet.Cells(rowNum, 3).Font.Name = "Calibri"
        xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4108  ' xlCenter
        
        ' Right side: Pending -> Clear Summary : Sector - City
        xlWorksheet.Cells(rowNum, 7).Value = "Pending -> Clear Summary : Sector - City"
        xlWorksheet.Cells(rowNum, 7).Font.Bold = True
        xlWorksheet.Cells(rowNum, 7).Font.Size = 12
        xlWorksheet.Cells(rowNum, 7).Font.Name = "Calibri"
        xlWorksheet.Cells(rowNum, 7).HorizontalAlignment = -4152  ' xlRight
        
        ' Format the header row with background color
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
            .Interior.Color = RGB(240, 240, 240)
            .VerticalAlignment = xlCenter  ' xlCenter - Middle align
        End With
        xlWorksheet.Rows(rowNum).RowHeight = 17.5
        rowNum = rowNum + 1
        
        ' Add table headers
        xlWorksheet.Cells(rowNum, 1).Value = "P No."
        xlWorksheet.Cells(rowNum, 2).Value = "Book Name"
        xlWorksheet.Cells(rowNum, 3).Value = "LG"
        xlWorksheet.Cells(rowNum, 4).Value = "Period"
        xlWorksheet.Cells(rowNum, 5).Value = "Qty"
        xlWorksheet.Cells(rowNum, 6).Value = "Sign"
        xlWorksheet.Cells(rowNum, 7).Value = "VK Books"
        
        ' Format header row - exactly like SectorVillageDetailsReport.bas AddTableHeaders
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Interior.Color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlHairline
            .VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
            ' Bold outer border for header
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin
        End With
        
        ' Set header horizontal alignments (vertical already set on range above)
        xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4108  ' xlCenter - P No.
        xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4131  ' xlLeft - Book Name
        xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4108  ' xlCenter - Language
        xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4131  ' xlLeft - Period
        xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4108  ' xlCenter - Qty
        xlWorksheet.Cells(rowNum, 6).HorizontalAlignment = -4131  ' xlLeft - Sign
        xlWorksheet.Cells(rowNum, 7).HorizontalAlignment = -4131  ' xlLeft - VK Books
        xlWorksheet.Rows(rowNum).RowHeight = 17.5
        
        rowNum = rowNum + 1
        
        ' Add books for duplicate table - keep all rows, will merge cells later
        Dim dupSectorTotalQty As Long
        dupSectorTotalQty = 0
        Dim dupDataStartRow As Long
        dupDataStartRow = rowNum  ' Track start of data rows for duplicate section
        Dim dupVillageTotalDict As Object
        Set dupVillageTotalDict = CreateObject("Scripting.Dictionary")  ' Track village totals for duplicate table
        
        For i = 1 To sectorBooks.count
            bookEntryVar = sectorBooks.item(i)
            
            bookName = CStr(bookEntryVar(0))
            language = CStr(bookEntryVar(1))
            period = CStr(bookEntryVar(2))
            qty = CStr(bookEntryVar(3))
            pNo = CStr(bookEntryVar(4))
            
            ' Add to total
            dupSectorTotalQty = dupSectorTotalQty + CLng(qty)
            
            ' Get village breakdown for this book+period+sector
            Dim bookIdValDup As Long
            Dim yearValDup As Integer
            Dim monthValDup As Integer
            Dim halfValDup As String
            Dim sectorValDup As String
            Dim villageBreakdownDup As String
            
            ' Safely extract values from bookEntryVar array
            On Error Resume Next
            If IsNumeric(bookEntryVar(6)) Then
                bookIdValDup = CLng(bookEntryVar(6))
            Else
                bookIdValDup = 0
            End If
            
            If IsNumeric(bookEntryVar(7)) Then
                yearValDup = CInt(bookEntryVar(7))
            Else
                yearValDup = 0
            End If
            
            If IsNumeric(bookEntryVar(8)) Then
                monthValDup = CInt(bookEntryVar(8))
            Else
                monthValDup = 0
            End If
            
            halfValDup = CStr(bookEntryVar(9) & "")
            sectorValDup = CStr(bookEntryVar(5) & "")
            On Error GoTo 0
            
            ' Only get village breakdown if we have valid data
            If bookIdValDup > 0 And yearValDup > 0 And monthValDup > 0 Then
                villageBreakdownDup = GetVillageBreakdown(bookIdValDup, yearValDup, monthValDup, halfValDup, sectorValDup, pNo)
            Else
                villageBreakdownDup = ""
            End If
            
            ' Add book entry - Column order: P No., Book Name, Language, Period, Qty, Sign, VK Books
            xlWorksheet.Cells(rowNum, 1).Value = pNo
            xlWorksheet.Cells(rowNum, 2).Value = bookName
            xlWorksheet.Cells(rowNum, 3).Value = GetLanguageShortCode(language)  ' Language short code (like LG in SectorSummaryReport)
            xlWorksheet.Cells(rowNum, 4).Value = period
            xlWorksheet.Cells(rowNum, 5).Value = qty
            xlWorksheet.Cells(rowNum, 6).Value = ""  ' Sign column - will be filled after merging
            xlWorksheet.Cells(rowNum, 7).Value = villageBreakdownDup  ' VK Books - village breakdown
            
            xlWorksheet.Cells(rowNum, 5).Font.Bold = True  ' Qty bold
            
            ' Format data row - matching SectorVillageDetailsReport.bas pattern
            With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlHairline
                .VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
            End With
            
            ' Set specific horizontal alignments for data cells (vertical already set on range above)
            xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4108  ' xlCenter - P No.
            xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4131  ' xlLeft - Book Name
            xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4108  ' xlCenter - Language
            xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4131  ' xlLeft - Period
            xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4108  ' xlCenter - Qty
            xlWorksheet.Cells(rowNum, 6).HorizontalAlignment = -4131  ' xlLeft - Sign
            xlWorksheet.Cells(rowNum, 7).HorizontalAlignment = -4131  ' xlLeft - VK Books
            
            ' Set font size 8 and ShrinkToFit for VK Books column
            xlWorksheet.Cells(rowNum, 7).Font.Size = 8
            xlWorksheet.Cells(rowNum, 7).ShrinkToFit = True
            
            ' Set ShrinkToFit for Book Name column
            xlWorksheet.Cells(rowNum, 2).ShrinkToFit = True
            
            ' Parse and accumulate village breakdown for total row
            If villageBreakdownDup <> "" Then
                Call ParseAndAccumulateVillageBreakdown(villageBreakdownDup, dupVillageTotalDict)
            End If
            
            xlWorksheet.Rows(rowNum).RowHeight = 17.5
            
            rowNum = rowNum + 1
        Next i
        
        Dim dupDataEndRow As Long
        dupDataEndRow = rowNum - 1  ' Track end of data rows (before total row)
        
        ' Apply cell merging for same book name + language (duplicate section)
        Call ApplyCellMergingForSectorTable(xlWorksheet, dupDataStartRow, dupDataEndRow)
        
        ' Build village total breakdown string for duplicate table
        Dim dupVillageTotalBreakdown As String
        dupVillageTotalBreakdown = BuildVillageTotalBreakdown(dupVillageTotalDict)
        
        ' Add total row for duplicate table
        xlWorksheet.Cells(rowNum, 1).Value = ""
        xlWorksheet.Cells(rowNum, 2).Value = ""
        xlWorksheet.Cells(rowNum, 3).Value = ""
        xlWorksheet.Cells(rowNum, 4).Value = "Total Qty"
        xlWorksheet.Cells(rowNum, 5).Value = dupSectorTotalQty
        xlWorksheet.Cells(rowNum, 6).Value = ""  ' Sign column - will be merged with VK Books
        xlWorksheet.Cells(rowNum, 7).Value = dupVillageTotalBreakdown  ' VK Books - village totals
        
        ' Merge Sign (column 6) and VK Books (column 7) cells in total row
        xlWorksheet.Range(xlWorksheet.Cells(rowNum, 6), xlWorksheet.Cells(rowNum, 7)).Merge
        xlWorksheet.Cells(rowNum, 6).Value = dupVillageTotalBreakdown  ' Put village totals in merged cell
        xlWorksheet.Cells(rowNum, 6).HorizontalAlignment = -4108  ' xlCenter - center align
        xlWorksheet.Cells(rowNum, 6).VerticalAlignment = xlCenter  ' xlCenter - center align vertically
        xlWorksheet.Cells(rowNum, 6).Font.Size = 8
        xlWorksheet.Cells(rowNum, 6).ShrinkToFit = True
        
        ' Format total row - exactly like SectorVillageDetailsReport.bas pattern
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 7))
            .Interior.Color = RGB(240, 240, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlHairline
            .VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
        End With
        
        xlWorksheet.Cells(rowNum, 4).Font.Bold = True
        xlWorksheet.Cells(rowNum, 5).Font.Bold = True
        xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4131  ' xlLeft
        xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4108  ' xlCenter
        
        ' Ensure ShrinkToFit is set on merged cell after all formatting (duplicate section)
        xlWorksheet.Cells(rowNum, 6).ShrinkToFit = True
        
        rowNum = rowNum + 1
        
        ' Update page usage counter after table is completely added
        currentPageUsedRowsDup = currentPageUsedRowsDup + dupTableSize
        
        ' Reset for next duplicate table
        previousSectorKeyForDup = sectorKey
        
    Next k
    
    
    ' Set column widths - matching SectorSummaryReport.bas
    ' Column order: P No., Book Name, Language, Period, Qty, Sign, VK Books
    xlWorksheet.Columns("A").ColumnWidth = 5   ' P No.
    xlWorksheet.Columns("B").ColumnWidth = 20' Book Name
    xlWorksheet.Columns("C").ColumnWidth = 4   ' Language (matching LG column width from SectorSummaryReport)
    xlWorksheet.Columns("D").ColumnWidth = 15   ' Period
    xlWorksheet.Columns("E").ColumnWidth = 5   ' Qty
    xlWorksheet.Columns("F").ColumnWidth = 20   ' Sign
    xlWorksheet.Columns("G").ColumnWidth = 26 ' VK Books - village breakdown
    
    ' Set row heights (slightly reduced)
    Dim lastDataRow As Long
    lastDataRow = rowNum - 1
    If lastDataRow > 0 Then
        xlWorksheet.Rows("1:" & lastDataRow).RowHeight = 17.5
        ' Set middle align (center vertical alignment) on entire sheet - use xlCenter constant
        ' This ensures all data including merged cells are center-aligned
        xlWorksheet.UsedRange.VerticalAlignment = xlCenter  ' Use xlCenter constant like SectorVillageDetailsReport.bas
    End If
    
    ' Set page format to match SectorSummaryReport.bas
    With xlWorksheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .TopMargin = Application.InchesToPoints(0.1)    ' Minimal margin at top
        .BottomMargin = Application.InchesToPoints(0.1)  ' Minimal margin at bottom
        .LeftMargin = Application.InchesToPoints(0.3)     ' More margin on left for punching
        .RightMargin = Application.InchesToPoints(0.1)   ' Minimal margin on right
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    
    ' Set print area
    xlWorksheet.PageSetup.PrintArea = "A1:G" & (rowNum - 1)
    
    ' Get save path from user
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Sector PDF As")
    
    If savePath = "False" Then
        xlWorkbook.Close False
        xlApp.Quit
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing
        Exit Sub
    End If
    
    ' Export as PDF
    xlWorksheet.ExportAsFixedFormat 0, savePath, 0, True, False, , , True
    
    ' Close and cleanup
    xlWorkbook.Close False
    xlApp.Quit
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Sector PDF generated successfully!" & vbCrLf & savePath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating sector PDF: " & Err.description, vbCritical
    
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
    End If
    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    On Error GoTo 0
End Sub

' Helper function to get short month name (3 characters)
Private Function GetShortMonthName(monthNum As Long) As String
    Dim monthNames() As String
    monthNames = Split("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC", ",")
    If monthNum >= 1 And monthNum <= 12 Then
        GetShortMonthName = monthNames(monthNum - 1)
    Else
        GetShortMonthName = "UNK" ' Unknown or error month
    End If
End Function

' Helper function to get half range
Private Function GetHalfRange(half As String, monthNum As Long) As String
    ' When half is "1 & 2", show full month range
    If Trim(LCase(half)) = "1 & 2" Or Trim(LCase(half)) = "1&2" Then
        ' Calculate number of days in the month
        Dim daysInMonth As Integer
        Select Case monthNum
            Case 1, 3, 5, 7, 8, 10, 12
                daysInMonth = 31
            Case 4, 6, 9, 11
                daysInMonth = 30
            Case 2
                ' February - assume 30 days for simplicity (or could use 28/29)
                daysInMonth = 30
            Case Else
                daysInMonth = 30
        End Select
        GetHalfRange = "1-" & daysInMonth
    ElseIf half = "1" Then
        ' First half: 1-15
        GetHalfRange = "1-15"
    ElseIf half = "2" Then
        ' Second half: 16-31
        GetHalfRange = "16-31"
    Else
        ' Default fallback
        GetHalfRange = "1-30"
    End If
End Function

' Apply cell merging for same book name + language (like SectorVillageDetailsReport.bas)
' Column order: A=P No., B=Book Name, C=Language/LG, D=Period, E=Qty, F=Sign
Private Sub ApplyCellMergingForSectorTable(xlWs As Worksheet, startRow As Long, endRow As Long)
    ' First, merge P No. column independently - merge consecutive same P No. values
    Dim pNoStartRow As Long
    Dim pNoEndRow As Long
    Dim currentPNo As String
    Dim pNoRow As Long
    
    pNoStartRow = startRow
    currentPNo = xlWs.Cells(pNoStartRow, 1).Value  ' Column A = P No.
    
    For pNoRow = startRow + 1 To endRow + 1
        Dim nextPNo As String
        If pNoRow <= endRow Then
            nextPNo = xlWs.Cells(pNoRow, 1).Value
        Else
            nextPNo = ""
        End If
        
        ' Check if P No. changed
        If nextPNo <> currentPNo Or pNoRow > endRow Then
            pNoEndRow = pNoRow - 1
            
            ' Merge P No. cells if more than one row with same P No.
            If pNoEndRow > pNoStartRow Then
                xlWs.Range(xlWs.Cells(pNoStartRow, 1), xlWs.Cells(pNoEndRow, 1)).Merge
                xlWs.Cells(pNoStartRow, 1).VerticalAlignment = xlCenter  ' Use xlCenter constant
            End If
            
            pNoStartRow = pNoRow
            If pNoRow <= endRow Then
                currentPNo = xlWs.Cells(pNoRow, 1).Value
            End If
        End If
    Next pNoRow
    
    ' Now merge Book Name, Language, and Sign columns based on book name + language
    Dim mergeStartRow As Long
    Dim mergeEndRow As Long
    Dim currentBookName As String
    Dim currentLanguage As String
    
    mergeStartRow = startRow
    currentBookName = xlWs.Cells(startRow, 2).Value  ' Column B = Book Name
    currentLanguage = xlWs.Cells(startRow, 3).Value  ' Column C = Language
    
    Dim k As Long
    For k = startRow + 1 To endRow + 1
        Dim nextBookName As String
        Dim nextLanguage As String
        
        If k <= endRow Then
            nextBookName = xlWs.Cells(k, 2).Value
            nextLanguage = xlWs.Cells(k, 3).Value
        Else
            nextBookName = ""
            nextLanguage = ""
        End If
        
        ' Check if book name + language changed
        If (nextBookName <> currentBookName Or nextLanguage <> currentLanguage) Or k > endRow Then
            mergeEndRow = k - 1
            
            ' Merge book name and language cells if more than one row
            If mergeEndRow > mergeStartRow Then
                ' Merge Book Name column (B) - exactly like SectorVillageDetailsReport.bas
                xlWs.Range(xlWs.Cells(mergeStartRow, 2), xlWs.Cells(mergeEndRow, 2)).Merge
                xlWs.Cells(mergeStartRow, 2).VerticalAlignment = xlCenter  ' Use xlCenter constant
                
                ' Merge Language column (C) - exactly like SectorVillageDetailsReport.bas
                xlWs.Range(xlWs.Cells(mergeStartRow, 3), xlWs.Cells(mergeEndRow, 3)).Merge
                xlWs.Cells(mergeStartRow, 3).VerticalAlignment = xlCenter  ' Use xlCenter constant
                
                ' Merge Sign column (F) and calculate total quantity - exactly like SectorVillageDetailsReport.bas
                xlWs.Range(xlWs.Cells(mergeStartRow, 6), xlWs.Cells(mergeEndRow, 6)).Merge
                xlWs.Cells(mergeStartRow, 6).VerticalAlignment = xlCenter  ' Use xlCenter constant
                
                ' VK Books column (G) - DO NOT MERGE, keep each row's village breakdown separate (period-wise)
                ' Each row maintains its own village breakdown for its specific period
                
                ' Calculate total quantity for this merged group
                Dim totalMergedQty As Long
                totalMergedQty = 0
                Dim m As Long
                For m = mergeStartRow To mergeEndRow
                    totalMergedQty = totalMergedQty + CLng(xlWs.Cells(m, 5).Value)  ' Column E = Qty
                Next m
                
                ' Set the merged Sign cell value to [total] - only horizontal alignment on individual cell
                xlWs.Cells(mergeStartRow, 6).Value = "[" & totalMergedQty & "]"
                xlWs.Cells(mergeStartRow, 6).HorizontalAlignment = -4131  ' xlLeft
            End If
            
            mergeStartRow = k
            If k <= endRow Then
                currentBookName = xlWs.Cells(k, 2).Value
                currentLanguage = xlWs.Cells(k, 3).Value
            End If
        End If
    Next k
End Sub

' Helper function to get village breakdown for a book+period+sector combination
' Returns format: "village- qty, village- qty" (e.g., "terr- 3, pooi - 2, assd- 2")
Private Function GetVillageBreakdown(bookId As Long, year As Integer, month As Integer, half As String, sector As String, pNo As String) As String
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim result As String
    Dim villageList As Object
    Set villageList = CreateObject("Scripting.Dictionary")
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"
    
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        GetVillageBreakdown = ""
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Query to get village-wise quantities for this book+period+sector+pNo
    sql = "SELECT " & _
          "t.village, " & _
          "SUM(t.qty) as village_qty " & _
          "FROM vimars.transactions t " & _
          "INNER JOIN vimars.books b ON t.book_id = b.book_id " & _
          "WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE " & _
          "AND t.book_id = " & bookId & " " & _
          "AND t.year = " & year & " " & _
          "AND t.month = " & month & " " & _
          "AND t.half = '" & Replace(half, "'", "''") & "' " & _
          "AND t.group_sector = '" & Replace(sector, "'", "''") & "' " & _
          "AND t.stock_number = '" & Replace(pNo, "'", "''") & "' " & _
          "AND t.avail = 0 " & _
          "AND COALESCE(t.village, '') <> '' " & _
          "GROUP BY t.village " & _
          "ORDER BY t.village ASC;"
    
    Set rs = conn.Execute(sql)
    
    ' First, collect all villages with quantities
    Dim villageData As Object
    Set villageData = CreateObject("Scripting.Dictionary")
    
    While Not rs.EOF
        Dim villageName As String
        Dim villageQty As Long
        On Error Resume Next
        villageName = Trim(rs.Fields("village").Value & "")
        If IsNumeric(rs.Fields("village_qty").Value) Then
            villageQty = CLng(rs.Fields("village_qty").Value)
        Else
            villageQty = 0
        End If
        On Error GoTo ErrorHandler
        
        If villageName <> "" And villageQty > 0 Then
            ' Store village name and quantity
            If villageData.Exists(villageName) Then
                villageData(villageName) = villageData(villageName) + villageQty
            Else
                villageData.Add villageName, villageQty
            End If
        End If
        
        rs.MoveNext
    Wend
    
    ' Now process village names to create short names (6 chars, with suffix if duplicate)
    Dim shortNameDict As Object
    Set shortNameDict = CreateObject("Scripting.Dictionary")
    Dim villageShortNames As Object
    Set villageShortNames = CreateObject("Scripting.Dictionary")
    
    ' First pass: create 6-char short names and track duplicates
    Dim villageKey As Variant
    For Each villageKey In villageData.Keys
        Dim fullName As String
        fullName = CStr(villageKey)
        Dim shortName As String
        shortName = Left(ToCamelCase(fullName), 6)
        
        ' Track how many times this short name appears
        If shortNameDict.Exists(shortName) Then
            shortNameDict(shortName) = shortNameDict(shortName) + 1
        Else
            shortNameDict.Add shortName, 1
        End If
        
        ' Store mapping of full name to short name
        villageShortNames.Add fullName, shortName
    Next villageKey
    
    ' Second pass: add suffix to duplicates
    Dim shortNameCount As Object
    Set shortNameCount = CreateObject("Scripting.Dictionary")
    
    For Each villageKey In villageData.Keys
        Dim fullName2 As String
        fullName2 = CStr(villageKey)
        Dim shortName2 As String
        shortName2 = villageShortNames(fullName2)
        
        ' If this short name appears more than once, add suffix
        If shortNameDict(shortName2) > 1 Then
            ' Count how many times we've seen this short name
            If shortNameCount.Exists(shortName2) Then
                shortNameCount(shortName2) = shortNameCount(shortName2) + 1
            Else
                shortNameCount.Add shortName2, 1
            End If
            
            ' Add suffix from the end of full name to make it unique
            Dim suffixLen As Long
            suffixLen = Len(fullName2) - 6
            If suffixLen > 0 Then
                ' Take 1-2 more characters from the end to make it unique
                Dim suffixChars As Long
                suffixChars = IIf(suffixLen >= 2, 2, suffixLen)
                Dim suffix As String
                suffix = Mid(fullName2, 7, suffixChars)
                shortName2 = Left(ToCamelCase(fullName2), 6) & LCase(suffix)
            End If
            
            ' Update the mapping
            villageShortNames(fullName2) = shortName2
        End If
    Next villageKey
    
    ' Build result string with short names
    result = ""
    For Each villageKey In villageData.Keys
        Dim finalShortName As String
        finalShortName = villageShortNames(villageKey)
        Dim finalQty As Long
        finalQty = villageData(villageKey)
        
        If result <> "" Then
            result = result & ", "
        End If
        result = result & finalShortName & "- " & finalQty
    Next villageKey
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    GetVillageBreakdown = result
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    GetVillageBreakdown = ""
End Function

' Helper function to parse village breakdown string and accumulate quantities
' Format: "Village- qty, Village- qty"
Private Sub ParseAndAccumulateVillageBreakdown(villageBreakdown As String, villageDict As Object)
    If villageBreakdown = "" Then
        Exit Sub
    End If
    
    ' Split by comma to get individual village entries
    Dim villageParts() As String
    villageParts = Split(villageBreakdown, ", ")
    
    Dim partIdx As Long
    For partIdx = 0 To UBound(villageParts)
        Dim villagePart As String
        villagePart = Trim(villageParts(partIdx))
        
        If villagePart <> "" Then
            ' Extract village name and qty (format: "Village- qty")
            Dim dashPos As Long
            dashPos = InStr(villagePart, "-")
            
            If dashPos > 0 Then
                Dim vName As String
                Dim vQty As Long
                Dim vQtyStr As String
                
                vName = Trim(Left(villagePart, dashPos - 1))
                vQtyStr = Trim(Mid(villagePart, dashPos + 1))
                
                On Error Resume Next
                If IsNumeric(vQtyStr) Then
                    vQty = CLng(vQtyStr)
                Else
                    vQty = 0
                End If
                On Error GoTo 0
                
                If vName <> "" And vQty > 0 Then
                    ' Accumulate quantity for this village
                    If villageDict.Exists(vName) Then
                        villageDict(vName) = villageDict(vName) + vQty
                    Else
                        villageDict.Add vName, vQty
                    End If
                End If
            End If
        End If
    Next partIdx
End Sub

' Helper function to build village total breakdown string from dictionary
Private Function BuildVillageTotalBreakdown(villageDict As Object) As String
    If villageDict.Count = 0 Then
        BuildVillageTotalBreakdown = ""
        Exit Function
    End If
    
    Dim result As String
    result = ""
    
    Dim villageKey As Variant
    For Each villageKey In villageDict.Keys
        If result <> "" Then
            result = result & ", "
        End If
        result = result & villageKey & "- " & villageDict(villageKey)
    Next villageKey
    
    BuildVillageTotalBreakdown = result
End Function

' Helper function to convert string to camel case (first letter uppercase, rest lowercase)
Private Function ToCamelCase(ByVal text As String) As String
    If Len(text) = 0 Then
        ToCamelCase = ""
        Exit Function
    End If
    
    Dim trimmedText As String
    trimmedText = Trim(text)
    
    If Len(trimmedText) = 1 Then
        ToCamelCase = UCase(trimmedText)
    Else
        ToCamelCase = UCase(Left(trimmedText, 1)) & LCase(Mid(trimmedText, 2))
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





