Option Explicit

' Sector Summary Report Function
Public Function GetSectorSummaryDataForPDF(period As String, groupSector As String, Optional selectedBookIds As Collection = Nothing, Optional appValue As String = "") As Collection
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
        Set GetSectorSummaryDataForPDF = New Collection
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Build SQL query to get sector data and organize by availability
    sql = "SELECT t.group_sector, " & _
          "b.language, " & _
          "b.book_name, " & _
          "CASE WHEN t.avail = 1 THEN 'YES' WHEN t.avail = 0 THEN 'NS' WHEN t.avail = 2 THEN 'B-YES' ELSE 'UNKNOWN' END AS availability_text, " & _
          "SUM(t.qty) as total_qty " & _
          "FROM vimars.transactions t " & _
          "INNER JOIN vimars.books b ON t.book_id = b.book_id " & _
          "WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE "
    
    ' Add filters
    If groupSector <> "" And groupSector <> "All" Then
        sql = sql & "AND t.group_sector = '" & Replace(groupSector, "'", "''") & "' "
    End If
    ' Note: When "All" is selected, we don't add any filter, so it shows all sectors
    
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
        Dim bookFilter As String
        bookFilter = SectorVillageDetailsReport.BuildNumericInClause(selectedBookIds)
        If bookFilter = "" Then
            GoTo Cleanup
        End If
        sql = sql & "AND b.book_id IN (" & bookFilter & ") "
    End If
    
    ' Add APP filter when provided
    If appValue <> "" Then
        sql = sql & "AND t.app = '" & Replace(appValue, "'", "''") & "' "
    End If
    
    sql = sql & "GROUP BY t.group_sector, b.language, t.avail, b.book_name " & _
                "ORDER BY t.group_sector, t.avail DESC, b.language ASC, b.book_name ASC;"
    
    
    Set rs = conn.Execute(sql)
    
    Dim recordCount As Integer
    recordCount = 0
    
    While Not rs.EOF
        Dim record(4) As Variant
        record(0) = rs.Fields("group_sector").Value
        record(1) = rs.Fields("language").Value
        record(2) = rs.Fields("book_name").Value
        record(3) = rs.Fields("availability_text").Value
        record(4) = rs.Fields("total_qty").Value
        
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
    Set GetSectorSummaryDataForPDF = summaryData
    Exit Function
    
ErrorHandler:
    Set summaryData = New Collection
    Resume Cleanup
End Function

' Test the sector function
Public Sub TestSectorSummary()
    On Error GoTo ErrorHandler
    
    
    Dim result As Collection
    Set result = GetSectorSummaryDataForPDF("2025-05-1", "All")
    
    
    If result.Count > 0 Then
    End If
    
    MsgBox "Sector test completed. Check Immediate Window.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Sector test error: " & Err.Description, vbCritical
End Sub

' Generate PDF Sector Summary Report
Public Sub GenerateSectorSummaryPDF(summaryData As Collection, period As String, groupSector As String, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
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
    pdfWs.Name = "Sector Summary"
    
    ' Initialize variables
    Dim i As Integer
    
    ' Generate formatted summary report
    Dim currentRow As Integer
    currentRow = 1
    
    ' Group data by sector
    Dim groupedData As Object
    Set groupedData = CreateObject("Scripting.Dictionary")
    
    ' Group the data by sector
    For i = 1 To summaryData.Count
        Dim dataRecord As Variant
        dataRecord = summaryData.Item(i)
        Dim sectorKeyStr As String
        sectorKeyStr = dataRecord(0)  ' group_sector only
        
        If Not groupedData.Exists(sectorKeyStr) Then
            Dim newGroup As Collection
            Set newGroup = New Collection
            groupedData.Add sectorKeyStr, newGroup
        End If
        
        groupedData(sectorKeyStr).Add dataRecord
    Next i
    
    ' Add data rows grouped by sector
    Dim totalQty As Integer
    totalQty = 0
    Dim currentLanguage As String
    currentLanguage = ""
    
    ' Separate tables into small (≤20 data rows) and large (>20 data rows)
    Dim smallTables As Object
    Dim largeTables As Object
    Set smallTables = CreateObject("Scripting.Dictionary")
    Set largeTables = CreateObject("Scripting.Dictionary")
    
    Dim sectorKeys As Variant
    Dim sectorKey As Variant
    Dim skIdx As Long
    sectorKeys = groupedData.Keys
    NaturalSort.SortKeysNatural sectorKeys
    
    ' Categorize tables by data row count (E1, E2, ... E9, E10 order)
    For skIdx = LBound(sectorKeys) To UBound(sectorKeys)
        sectorKey = sectorKeys(skIdx)
        Dim sectorRecords As Collection
        Set sectorRecords = groupedData(sectorKey)
        Dim dataRowCount As Integer
        dataRowCount = sectorRecords.Count
        
        If dataRowCount <= 20 Then
            ' Small table (≤20 data rows)
            smallTables.Add sectorKey, groupedData(sectorKey)
        Else
            ' Large table (>20 data rows)
            largeTables.Add sectorKey, groupedData(sectorKey)
        End If
    Next skIdx
    
    ' Page break variables for smart layout
    Dim sectorsOnCurrentPage As Integer
    sectorsOnCurrentPage = 0
    Dim currentPageStartRow As Integer
    currentPageStartRow = 1
    Dim firstTableStartRow As Integer
    firstTableStartRow = 1
    Dim previousSectorRowCount As Integer
    previousSectorRowCount = 0
    Dim previousSectorName As String
    previousSectorName = ""
    
    ' First process all small tables (≤20 data rows) with optimization to minimize pages
    Dim smallTableKeys As Variant
    smallTableKeys = smallTables.Keys
    NaturalSort.SortKeysNatural smallTableKeys
    
    ' Optimization: Pre-calculate table sizes for smart grouping
    Dim smallTableSizes As Object
    Set smallTableSizes = CreateObject("Scripting.Dictionary")
    Dim smallTableOrder As Collection
    Set smallTableOrder = New Collection
    
    ' Calculate size for each small table
    For Each sectorKey In smallTableKeys
        Dim sectorRecordsTemp As Collection
        Set sectorRecordsTemp = smallTables(sectorKey)
        Dim smallTableSize As Integer
        smallTableSize = 3 + sectorRecordsTemp.Count  ' Header + headers + data + total
        smallTableSizes.Add sectorKey, smallTableSize
        smallTableOrder.Add sectorKey
    Next sectorKey
    
    ' Page limit for small tables (allowing 2 small tables per page with spacing)
    Dim maxRowsPerPageSmall As Integer
    maxRowsPerPageSmall = 45  ' Same as duplicate section for consistency
    Dim currentPageUsedRowsSmall As Integer
    currentPageUsedRowsSmall = 0
    
    For i = 1 To smallTableOrder.Count
        sectorKey = smallTableOrder.Item(i)
        Dim sectorRecordsSmall As Collection
        Set sectorRecordsSmall = smallTables(sectorKey)
        
        ' Get pre-calculated table size
        Dim estimatedRowsForSector As Integer
        estimatedRowsForSector = smallTableSizes(sectorKey)
        
        ' Check if we need a new page (max 2 small tables per page)
        Dim needNewPage As Boolean
        needNewPage = False
        
        If currentRow > 1 Then  ' Not the first sector
            ' If we already have 2 sectors on this page, force new page
            If sectorsOnCurrentPage >= 2 Then
                needNewPage = True
            ElseIf sectorsOnCurrentPage = 1 Then
                ' Allow different sectors on same page, but check size limits
                If previousSectorRowCount > 23 Then
                    needNewPage = True   ' Force new page because previous was too large (>23 rows)
                ElseIf estimatedRowsForSector > 23 Then
                    needNewPage = True   ' Force new page because current is too large (>23 rows)
                Else
                    needNewPage = False  ' Keep on same page (both ≤23 rows including headers, even if different sectors)
                End If
            End If
        End If
        
        ' Add page break if needed
        If needNewPage Then
            pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
            sectorsOnCurrentPage = 0
            currentPageStartRow = currentRow
            firstTableStartRow = currentRow
            currentPageUsedRowsSmall = 0  ' Reset page usage counter
        ElseIf sectorsOnCurrentPage = 1 Then
            ' If this is the 2nd table on the same page, position header at row 25 (relative to page start)
            ' Adding spacing rows to reach the target row
            Dim targetRow As Integer
            targetRow = currentPageStartRow + 24   ' Row 25 of the page (pageStartRow is row 1 of the page)
            
            Dim spacingRows As Integer
            spacingRows = targetRow - currentRow
            
            ' Add spacing if needed to reach the target row
            If spacingRows > 0 Then
                Dim spacingIdx As Integer
                For spacingIdx = 1 To spacingRows
                    pdfWs.Rows(currentRow).RowHeight = 18.15
                    currentRow = currentRow + 1
                Next spacingIdx
                currentPageUsedRowsSmall = currentPageUsedRowsSmall + spacingRows  ' Track spacing rows
            End If
        End If
        
        ' Update page usage counter (will be updated after table is added)
        
        ' Add sector on left and VIMARSH period on right (same line)
        pdfWs.Cells(currentRow, 1).Value = sectorKey
        pdfWs.Cells(currentRow, 1).Font.Bold = True
        pdfWs.Cells(currentRow, 1).Font.Size = 12
        pdfWs.Cells(currentRow, 1).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlLeft
        
        ' Add VIMARSH period in AVAIL column (C) - center aligned
        If period <> "" And period <> "All" Then
            Dim periodParts() As String
            periodParts = Split(period, "-")
            If UBound(periodParts) >= 2 Then
                pdfWs.Cells(currentRow, 3).Value = orgName & ": " & GetShortMonthName(CLng(periodParts(1))) & "-" & periodParts(0) & "(" & GetHalfRange(periodParts(2), CLng(periodParts(1)), CLng(periodParts(0))) & ")"
            Else
                pdfWs.Cells(currentRow, 3).Value = orgName & ": " & period
            End If
        Else
            pdfWs.Cells(currentRow, 3).Value = orgName & ": All Periods"
        End If
        pdfWs.Cells(currentRow, 3).Font.Bold = True
        pdfWs.Cells(currentRow, 3).Font.Size = 12
        pdfWs.Cells(currentRow, 3).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
        
        ' Add "Summary: Sector" in DATE column (F) - right aligned
        pdfWs.Cells(currentRow, 6).Value = "Summary: Sector"
        pdfWs.Cells(currentRow, 6).Font.Bold = True
        pdfWs.Cells(currentRow, 6).Font.Size = 12
        pdfWs.Cells(currentRow, 6).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlRight
        
        ' Format the header row
        pdfWs.Range("A" & currentRow & ":F" & currentRow).Interior.Color = RGB(240, 240, 240)
        pdfWs.Rows(currentRow).RowHeight = 18.15
        currentRow = currentRow + 1
        
        ' Add table headers for each sector
        pdfWs.Cells(currentRow, 1).Value = "LG"
        pdfWs.Cells(currentRow, 2).Value = "BOOK TITLE"
        pdfWs.Cells(currentRow, 3).Value = "AVAIL"
        pdfWs.Cells(currentRow, 4).Value = "QTY"
        pdfWs.Cells(currentRow, 5).Value = "SIGN"
        pdfWs.Cells(currentRow, 6).Value = "DATE"
        
        ' Format headers
        With pdfWs.Range("A" & currentRow & ":F" & currentRow)
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
        pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter ' DATE - Center
        pdfWs.Rows(currentRow).RowHeight = 18.15
        currentRow = currentRow + 1
        
        ' Process records for this sector (YES first, then NS)
        Dim yesStartRow As Integer
        yesStartRow = 0
        Dim yesTotalQty As Integer
        yesTotalQty = 0
        Dim notInStockStartRow As Integer
        notInStockStartRow = 0
        Dim notInStockTotalQty As Integer
        notInStockTotalQty = 0
        Dim lastAvailability As String
        lastAvailability = ""
        
        ' Sort records within this sector group by availability, language, and book title
        Dim sortedRecords As Collection
        Set sortedRecords = New Collection
        
        ' Create separate collections for each availability type
        Dim yesRecords As Collection
        Dim notInStockRecords As Collection
        Dim bYesRecords As Collection
        Set yesRecords = New Collection
        Set notInStockRecords = New Collection
        Set bYesRecords = New Collection
        
        ' Separate records by availability
        Dim recordIdx As Integer
        For recordIdx = 1 To sectorRecordsSmall.Count
            dataRecord = sectorRecordsSmall.Item(recordIdx)
            If dataRecord(3) = "YES" Then
                yesRecords.Add dataRecord
            ElseIf dataRecord(3) = "NS" Then
                notInStockRecords.Add dataRecord
            ElseIf dataRecord(3) = "B-YES" Then
                bYesRecords.Add dataRecord
            End If
        Next recordIdx
        
        ' Sort each availability group by language and book title
        Call SortRecordsByLanguageAndBook(yesRecords)
        Call SortRecordsByLanguageAndBook(notInStockRecords)
        Call SortRecordsByLanguageAndBook(bYesRecords)
        
        ' Add YES records first
        For recordIdx = 1 To yesRecords.Count
            sortedRecords.Add yesRecords.Item(recordIdx)
        Next recordIdx
        
        ' Add NS records second
        For recordIdx = 1 To notInStockRecords.Count
            sortedRecords.Add notInStockRecords.Item(recordIdx)
        Next recordIdx
        
        ' Add B-YES records third
        For recordIdx = 1 To bYesRecords.Count
            sortedRecords.Add bYesRecords.Item(recordIdx)
        Next recordIdx
        
        ' Now process the sorted records
        For recordIdx = 1 To sortedRecords.Count
            dataRecord = sortedRecords.Item(recordIdx)
            
            ' Check if language has changed (compare short codes)
            Dim currentRecordLangCode As String
            currentRecordLangCode = GetLanguageShortCode(CStr(dataRecord(1)))
            If currentRecordLangCode <> currentLanguage Then
                ' If we were processing records, merge the previous language group
                If currentLanguage <> "" Then
                    ' Merge LNG column for the previous language group
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
            pdfWs.Cells(currentRow, 2).Value = dataRecord(2)  ' Book Name
            pdfWs.Cells(currentRow, 3).Value = dataRecord(3)  ' Availability
            pdfWs.Cells(currentRow, 4).Value = dataRecord(4)  ' Quantity
            pdfWs.Cells(currentRow, 5).Value = ""             ' Sign (empty)
            pdfWs.Cells(currentRow, 6).Value = ""             ' Date (empty)
            
            ' Make NS text light/gray in individual cells
            If dataRecord(3) = "NS" Then
                pdfWs.Cells(currentRow, 3).Font.Color = RGB(150, 150, 150)
                pdfWs.Cells(currentRow, 3).Font.Bold = False
            End If
            
            ' Track totals for each availability type
            If dataRecord(3) = "YES" Then
                yesTotalQty = yesTotalQty + CLng(dataRecord(4))
                If yesStartRow = 0 Then
                    yesStartRow = currentRow
                End If
                ' Don't merge NS cells - just reset the tracking
                notInStockStartRow = 0
                notInStockTotalQty = 0
            Else
                notInStockTotalQty = notInStockTotalQty + CLng(dataRecord(4))
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
                    
                    ' Merge DATE column for YES section (empty for now)
                    pdfWs.Range("F" & yesStartRow & ":F" & (currentRow - 1)).Merge
                    pdfWs.Range("F" & yesStartRow & ":F" & (currentRow - 1)).HorizontalAlignment = xlCenter
                    pdfWs.Range("F" & yesStartRow & ":F" & (currentRow - 1)).VerticalAlignment = xlCenter
                    
                    ' Note: Language merging is handled separately when language changes
                End If
                yesStartRow = 0
                yesTotalQty = 0
            End If
            
            lastAvailability = dataRecord(3)
            
            ' Format data row
            With pdfWs.Range("A" & currentRow & ":F" & currentRow)
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
            pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter ' DATE - Center
            
            pdfWs.Rows(currentRow).RowHeight = 18.15
            
            totalQty = totalQty + CLng(dataRecord(4))
            currentRow = currentRow + 1
        Next recordIdx
        
        ' Handle merge cells for YES values at the end of sector
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
            
            ' Merge DATE column for YES section (empty for now)
            pdfWs.Range("F" & yesStartRow & ":F" & (currentRow - 1)).Merge
            pdfWs.Range("F" & yesStartRow & ":F" & (currentRow - 1)).HorizontalAlignment = xlCenter
            pdfWs.Range("F" & yesStartRow & ":F" & (currentRow - 1)).VerticalAlignment = xlCenter
        End If
        
        ' NS cells are not merged - no action needed
        
        ' Handle merge cells for the last language group at the end of sector (compare short codes)
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
        
        ' Add total line at the bottom of this sector table
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
        With pdfWs.Range("A" & currentRow & ":F" & currentRow)
            .Interior.Color = RGB(240, 240, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
        End With
        pdfWs.Rows(currentRow).RowHeight = 18.15
        
        currentRow = currentRow + 1
        
        ' Reset total for next sector
        totalQty = 0
        
        ' Store the actual row count used by this table (for next iteration)
        Dim actualTableRows As Integer
        actualTableRows = currentRow - firstTableStartRow + 1
        previousSectorRowCount = actualTableRows
        previousSectorName = sectorKey
        
        ' Update page usage counter
        currentPageUsedRowsSmall = currentPageUsedRowsSmall + actualTableRows
        
        ' Increment sector counter for this page
        sectorsOnCurrentPage = sectorsOnCurrentPage + 1
        
        ' Update firstTableStartRow for next iteration (if this was first table on page)
        If sectorsOnCurrentPage = 1 Then
            firstTableStartRow = currentRow - actualTableRows + 1
        End If
        
    Next i
    
    ' Now process all large tables (>20 data rows) - each on new page
    Dim largeTableKeys As Variant
    largeTableKeys = largeTables.Keys
    NaturalSort.SortKeysNatural largeTableKeys
    
    ' Reset page tracking for large tables
    sectorsOnCurrentPage = 0
    If currentRow > 1 Then
        pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
        currentPageStartRow = currentRow
        firstTableStartRow = currentRow
    End If
    
    For Each sectorKey In largeTableKeys
        Dim sectorRecordsLarge As Collection
        Set sectorRecordsLarge = largeTables(sectorKey)
        
        ' Large tables always start on new page
        If currentRow > 1 Then
            pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
            sectorsOnCurrentPage = 0
            currentPageStartRow = currentRow
            firstTableStartRow = currentRow
        End If
        
        ' Add sector on left and VIMARSH period on right (same line)
        pdfWs.Cells(currentRow, 1).Value = sectorKey
        pdfWs.Cells(currentRow, 1).Font.Bold = True
        pdfWs.Cells(currentRow, 1).Font.Size = 12
        pdfWs.Cells(currentRow, 1).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlLeft
        
        ' Add VIMARSH period in AVAIL column (C) - center aligned
        If period <> "" And period <> "All" Then
            Dim periodPartsLarge() As String
            periodPartsLarge = Split(period, "-")
            If UBound(periodPartsLarge) >= 2 Then
                pdfWs.Cells(currentRow, 3).Value = orgName & ": " & GetShortMonthName(CLng(periodPartsLarge(1))) & "-" & periodPartsLarge(0) & "(" & GetHalfRange(periodPartsLarge(2), CLng(periodPartsLarge(1)), CLng(periodPartsLarge(0))) & ")"
            Else
                pdfWs.Cells(currentRow, 3).Value = orgName & ": " & period
            End If
        Else
            pdfWs.Cells(currentRow, 3).Value = orgName & ": All Periods"
        End If
        pdfWs.Cells(currentRow, 3).Font.Bold = True
        pdfWs.Cells(currentRow, 3).Font.Size = 12
        pdfWs.Cells(currentRow, 3).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
        
        ' Add "Summary: Sector" in DATE column (F) - right aligned
        pdfWs.Cells(currentRow, 6).Value = "Summary: Sector"
        pdfWs.Cells(currentRow, 6).Font.Bold = True
        pdfWs.Cells(currentRow, 6).Font.Size = 12
        pdfWs.Cells(currentRow, 6).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlRight
        
        ' Format the header row
        pdfWs.Range("A" & currentRow & ":F" & currentRow).Interior.Color = RGB(240, 240, 240)
        pdfWs.Rows(currentRow).RowHeight = 18.15
        currentRow = currentRow + 1
        
        ' Add table headers
        pdfWs.Cells(currentRow, 1).Value = "LG"
        pdfWs.Cells(currentRow, 2).Value = "BOOK TITLE"
        pdfWs.Cells(currentRow, 3).Value = "AVAIL"
        pdfWs.Cells(currentRow, 4).Value = "QTY"
        pdfWs.Cells(currentRow, 5).Value = "SIGN"
        pdfWs.Cells(currentRow, 6).Value = "DATE"
        
        ' Format headers
        With pdfWs.Range("A" & currentRow & ":F" & currentRow)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Interior.Color = RGB(220, 220, 220)
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' Set specific alignments
        pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlLeft
        pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter
        pdfWs.Rows(currentRow).RowHeight = 18.15
        currentRow = currentRow + 1
        
        ' Process records for large table
        Dim yesStartRowLarge As Integer
        yesStartRowLarge = 0
        Dim yesTotalQtyLarge As Integer
        yesTotalQtyLarge = 0
        Dim notInStockStartRowLarge As Integer
        notInStockStartRowLarge = 0
        Dim notInStockTotalQtyLarge As Integer
        notInStockTotalQtyLarge = 0
        Dim lastAvailabilityLarge As String
        lastAvailabilityLarge = ""
        
        ' Sort records within this sector group
        Dim sortedRecordsLarge As Collection
        Set sortedRecordsLarge = New Collection
        
        Dim yesRecordsLarge As Collection
        Dim notInStockRecordsLarge As Collection
        Dim bYesRecordsLarge As Collection
        Set yesRecordsLarge = New Collection
        Set notInStockRecordsLarge = New Collection
        Set bYesRecordsLarge = New Collection
        
        ' Separate records by availability
        For i = 1 To sectorRecordsLarge.Count
            Dim dataRecordLarge As Variant
            dataRecordLarge = sectorRecordsLarge.Item(i)
            If dataRecordLarge(3) = "YES" Then
                yesRecordsLarge.Add dataRecordLarge
            ElseIf dataRecordLarge(3) = "NS" Then
                notInStockRecordsLarge.Add dataRecordLarge
            ElseIf dataRecordLarge(3) = "B-YES" Then
                bYesRecordsLarge.Add dataRecordLarge
            End If
        Next i
        
        ' Sort each availability group
        Call SortRecordsByLanguageAndBook(yesRecordsLarge)
        Call SortRecordsByLanguageAndBook(notInStockRecordsLarge)
        Call SortRecordsByLanguageAndBook(bYesRecordsLarge)
        
        ' Add YES records first
        For i = 1 To yesRecordsLarge.Count
            sortedRecordsLarge.Add yesRecordsLarge.Item(i)
        Next i
        
        ' Add NS records second
        For i = 1 To notInStockRecordsLarge.Count
            sortedRecordsLarge.Add notInStockRecordsLarge.Item(i)
        Next i
        
        ' Add B-YES records third
        For i = 1 To bYesRecordsLarge.Count
            sortedRecordsLarge.Add bYesRecordsLarge.Item(i)
        Next i
        
        ' Process sorted records
        Dim currentLanguageLarge As String
        currentLanguageLarge = ""
        
        For i = 1 To sortedRecordsLarge.Count
            Dim dataRecordLarge2 As Variant
            dataRecordLarge2 = sortedRecordsLarge.Item(i)
            
            ' Check if language has changed
            Dim currentRecordLangCodeLarge As String
            currentRecordLangCodeLarge = GetLanguageShortCode(CStr(dataRecordLarge2(1)))
            If currentRecordLangCodeLarge <> currentLanguageLarge Then
                If currentLanguageLarge <> "" Then
                    If currentRow > 1 Then
                        Dim languageStartRowLarge As Integer
                        languageStartRowLarge = currentRow - 1
                        While languageStartRowLarge > 1 And pdfWs.Cells(languageStartRowLarge - 1, 1).Value = currentLanguageLarge
                            languageStartRowLarge = languageStartRowLarge - 1
                        Wend
                        If languageStartRowLarge < currentRow - 1 Then
                            pdfWs.Range("A" & languageStartRowLarge & ":A" & (currentRow - 1)).Merge
                            pdfWs.Range("A" & languageStartRowLarge & ":A" & (currentRow - 1)).HorizontalAlignment = xlCenter
                            pdfWs.Range("A" & languageStartRowLarge & ":A" & (currentRow - 1)).VerticalAlignment = xlCenter
                        End If
                    End If
                End If
                currentLanguageLarge = currentRecordLangCodeLarge
            End If
            
            pdfWs.Cells(currentRow, 1).Value = currentRecordLangCodeLarge
            pdfWs.Cells(currentRow, 2).Value = dataRecordLarge2(2)
            pdfWs.Cells(currentRow, 3).Value = dataRecordLarge2(3)
            pdfWs.Cells(currentRow, 4).Value = dataRecordLarge2(4)
            pdfWs.Cells(currentRow, 5).Value = ""
            pdfWs.Cells(currentRow, 6).Value = ""
            
            ' Make NS text light/gray
            If dataRecordLarge2(3) = "NS" Then
                pdfWs.Cells(currentRow, 3).Font.Color = RGB(150, 150, 150)
                pdfWs.Cells(currentRow, 3).Font.Bold = False
            End If
            
            If dataRecordLarge2(3) = "YES" Then
                yesTotalQtyLarge = yesTotalQtyLarge + CLng(dataRecordLarge2(4))
                If yesStartRowLarge = 0 Then
                    yesStartRowLarge = currentRow
                End If
                notInStockStartRowLarge = 0
                notInStockTotalQtyLarge = 0
            Else
                notInStockTotalQtyLarge = notInStockTotalQtyLarge + CLng(dataRecordLarge2(4))
                If notInStockStartRowLarge = 0 Then
                    notInStockStartRowLarge = currentRow
                End If
                If yesStartRowLarge > 0 And yesStartRowLarge < currentRow Then
                    pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).Merge
                    pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).HorizontalAlignment = xlCenter
                    pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).VerticalAlignment = xlCenter
                    pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).Value = "YES"
                    
                    pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).Merge
                    pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).Value = "[" & yesTotalQtyLarge & "]"
                    pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).HorizontalAlignment = xlLeft
                    pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).VerticalAlignment = xlCenter
                    
                    pdfWs.Range("F" & yesStartRowLarge & ":F" & (currentRow - 1)).Merge
                    pdfWs.Range("F" & yesStartRowLarge & ":F" & (currentRow - 1)).HorizontalAlignment = xlCenter
                    pdfWs.Range("F" & yesStartRowLarge & ":F" & (currentRow - 1)).VerticalAlignment = xlCenter
                End If
                yesStartRowLarge = 0
                yesTotalQtyLarge = 0
            End If
            
            lastAvailabilityLarge = dataRecordLarge2(3)
            
            ' Format data row
            With pdfWs.Range("A" & currentRow & ":F" & currentRow)
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .VerticalAlignment = xlCenter
            End With
            
            pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlLeft
            pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter
            
            pdfWs.Rows(currentRow).RowHeight = 18.15
            
            totalQty = totalQty + CLng(dataRecordLarge2(4))
            currentRow = currentRow + 1
        Next i
        
        ' Handle merge cells for YES values at the end
        If yesStartRowLarge > 0 And yesStartRowLarge < currentRow Then
            pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).Merge
            pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).HorizontalAlignment = xlCenter
            pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).VerticalAlignment = xlCenter
            pdfWs.Range("C" & yesStartRowLarge & ":C" & (currentRow - 1)).Value = "YES"
            
            pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).Merge
            pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).Value = "[" & yesTotalQtyLarge & "]"
            pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).HorizontalAlignment = xlLeft
            pdfWs.Range("E" & yesStartRowLarge & ":E" & (currentRow - 1)).VerticalAlignment = xlCenter
            
            pdfWs.Range("F" & yesStartRowLarge & ":F" & (currentRow - 1)).Merge
            pdfWs.Range("F" & yesStartRowLarge & ":F" & (currentRow - 1)).HorizontalAlignment = xlCenter
            pdfWs.Range("F" & yesStartRowLarge & ":F" & (currentRow - 1)).VerticalAlignment = xlCenter
        End If
        
        ' Handle merge cells for the last language group
        If currentLanguageLarge <> "" And currentRow > 1 Then
            Dim lastLanguageStartRowLarge As Integer
            lastLanguageStartRowLarge = currentRow - 1
            While lastLanguageStartRowLarge > 1 And pdfWs.Cells(lastLanguageStartRowLarge - 1, 1).Value = currentLanguageLarge
                lastLanguageStartRowLarge = lastLanguageStartRowLarge - 1
            Wend
            If lastLanguageStartRowLarge < currentRow - 1 Then
                pdfWs.Range("A" & lastLanguageStartRowLarge & ":A" & (currentRow - 1)).Merge
                pdfWs.Range("A" & lastLanguageStartRowLarge & ":A" & (currentRow - 1)).HorizontalAlignment = xlCenter
                pdfWs.Range("A" & lastLanguageStartRowLarge & ":A" & (currentRow - 1)).VerticalAlignment = xlCenter
            End If
        End If
        
        ' Add total line
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
        With pdfWs.Range("A" & currentRow & ":F" & currentRow)
            .Interior.Color = RGB(240, 240, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
        End With
        pdfWs.Rows(currentRow).RowHeight = 18.15
        
        currentRow = currentRow + 1
        
        ' Reset total for next sector
        totalQty = 0
        
        sectorsOnCurrentPage = sectorsOnCurrentPage + 1
    Next sectorKey
    
    ' Now add duplicate tables for all "Summary: Sector" tables
    ' Collect all tables (both small and large) for duplicates
    Dim allTablesDict As Object
    Set allTablesDict = CreateObject("Scripting.Dictionary")
    
    ' Add small tables
    For Each sectorKey In smallTableKeys
        allTablesDict.Add sectorKey, smallTables(sectorKey)
    Next sectorKey
    
    ' Add large tables
    For Each sectorKey In largeTableKeys
        allTablesDict.Add sectorKey, largeTables(sectorKey)
    Next sectorKey
    
    ' Process duplicate tables with optimization logic to minimize pages
    ' Pre-calculate table sizes and group them efficiently
    Dim allKeys As Variant
    allKeys = allTablesDict.Keys
    NaturalSort.SortKeysNatural allKeys
    
    ' Reset page tracking
    sectorsOnCurrentPage = 0
    If currentRow > 1 Then
        pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
        currentPageStartRow = currentRow
        firstTableStartRow = currentRow
    End If
    
    ' Page height approximately 50-55 rows (A4 portrait with margins)
    ' Use 45 instead of 50 to leave more buffer and prevent table cutting
    ' This ensures tables never get cut, especially on last page
    Dim maxRowsPerPage As Integer
    maxRowsPerPage = 45
    
    ' Optimization: Pre-calculate all table sizes for smart grouping
    Dim tableSizes As Object
    Set tableSizes = CreateObject("Scripting.Dictionary")
    Dim tableOrder As Collection
    Set tableOrder = New Collection
    
    ' Calculate size for each table and create ordered list
    For Each sectorKey In allKeys
        Dim dupRecordsTemp As Collection
        Set dupRecordsTemp = allTablesDict(sectorKey)
        Dim tableSize As Integer
        tableSize = 3 + dupRecordsTemp.Count  ' Header + headers + data + total
        tableSizes.Add sectorKey, tableSize
        tableOrder.Add sectorKey
    Next sectorKey
    
    ' Process tables with optimization: try to fit as many as possible on each page
    Dim previousSectorKeyForDup As String
    previousSectorKeyForDup = ""
    Dim currentPageUsedRows As Integer
    currentPageUsedRows = 0
    
    For i = 1 To tableOrder.Count
        sectorKey = tableOrder.Item(i)
        Dim dupRecords As Collection
        Set dupRecords = allTablesDict(sectorKey)
        
        ' Get pre-calculated table size (optimization: already calculated)
        Dim dupTableRows As Integer
        dupTableRows = tableSizes(sectorKey)
        
        ' Add 1 row spacing between sectors (except first sector)
        Dim spacingRowsNeeded As Integer
        spacingRowsNeeded = 0
        If previousSectorKeyForDup <> "" And sectorKey <> previousSectorKeyForDup Then
            spacingRowsNeeded = 1
        End If
        
        ' Calculate total rows needed including spacing
        Dim totalRowsNeeded As Integer
        totalRowsNeeded = spacingRowsNeeded + dupTableRows
        
        ' Optimization: Use tracked page usage for accurate fit calculation
        ' This minimizes pages by maximizing utilization
        If currentRow > 1 Then
            If (currentPageUsedRows + totalRowsNeeded) >= maxRowsPerPage Then
                ' Table won't fit completely - add page break BEFORE starting
                pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
                sectorsOnCurrentPage = 0
                currentPageStartRow = currentRow
                firstTableStartRow = currentRow
                currentPageUsedRows = 0  ' Reset page usage counter
                spacingRowsNeeded = 0  ' No spacing needed after page break
            End If
        End If
        
        ' Add spacing row between sectors (except first sector and after page break)
        If spacingRowsNeeded > 0 Then
            pdfWs.Rows(currentRow).RowHeight = 18.15
            currentRow = currentRow + 1
            currentPageUsedRows = currentPageUsedRows + 1  ' Track spacing row
        End If
        
        ' Note: currentPageUsedRows will be updated after table is completely added (at end of table)
        
        ' Add duplicate table header
        pdfWs.Cells(currentRow, 1).Value = sectorKey
        pdfWs.Cells(currentRow, 1).Font.Bold = True
        pdfWs.Cells(currentRow, 1).Font.Size = 12
        pdfWs.Cells(currentRow, 1).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlLeft
        
        ' Add VIMARSH period in AVAIL column (C) - center aligned
        If period <> "" And period <> "All" Then
            Dim periodPartsDup() As String
            periodPartsDup = Split(period, "-")
            If UBound(periodPartsDup) >= 2 Then
                pdfWs.Cells(currentRow, 3).Value = orgName & ": " & GetShortMonthName(CLng(periodPartsDup(1))) & "-" & periodPartsDup(0) & "(" & GetHalfRange(periodPartsDup(2), CLng(periodPartsDup(1)), CLng(periodPartsDup(0))) & ")"
            Else
                pdfWs.Cells(currentRow, 3).Value = orgName & ": " & period
            End If
        Else
            pdfWs.Cells(currentRow, 3).Value = orgName & ": All Periods"
        End If
        pdfWs.Cells(currentRow, 3).Font.Bold = True
        pdfWs.Cells(currentRow, 3).Font.Size = 12
        pdfWs.Cells(currentRow, 3).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
        
        ' Add "Summary: Sector - City" in DATE column (F) - right aligned
        pdfWs.Cells(currentRow, 6).Value = "Summary: Sector - City"
        pdfWs.Cells(currentRow, 6).Font.Bold = True
        pdfWs.Cells(currentRow, 6).Font.Size = 12
        pdfWs.Cells(currentRow, 6).Font.Name = "Calibri"
        pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlRight
        
        ' Format the header row
        pdfWs.Range("A" & currentRow & ":F" & currentRow).Interior.Color = RGB(240, 240, 240)
        pdfWs.Rows(currentRow).RowHeight = 18.15
        currentRow = currentRow + 1
        
        ' Add table headers
        pdfWs.Cells(currentRow, 1).Value = "LG"
        pdfWs.Cells(currentRow, 2).Value = "BOOK TITLE"
        pdfWs.Cells(currentRow, 3).Value = "AVAIL"
        pdfWs.Cells(currentRow, 4).Value = "QTY"
        pdfWs.Cells(currentRow, 5).Value = "SIGN"
        pdfWs.Cells(currentRow, 6).Value = "DATE"
        
        ' Format headers
        With pdfWs.Range("A" & currentRow & ":F" & currentRow)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 9
            .Interior.Color = RGB(220, 220, 220)
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlLeft
        pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter
        pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter
        pdfWs.Rows(currentRow).RowHeight = 18.15
        currentRow = currentRow + 1
        
        ' Process records for duplicate table
        Dim dupYesStartRow As Integer
        dupYesStartRow = 0
        Dim dupYesTotalQty As Integer
        dupYesTotalQty = 0
        Dim dupCurrentLanguage As String
        dupCurrentLanguage = ""
        
        ' Sort records
        Dim dupSortedRecords As Collection
        Set dupSortedRecords = New Collection
        
        Dim dupYesRecords As Collection
        Dim dupNotInStockRecords As Collection
        Dim dupBYesRecords As Collection
        Set dupYesRecords = New Collection
        Set dupNotInStockRecords = New Collection
        Set dupBYesRecords = New Collection
        
        ' Separate records by availability
        Dim j As Integer
        For j = 1 To dupRecords.Count
            Dim dupDataRecord As Variant
            dupDataRecord = dupRecords.Item(j)
            If dupDataRecord(3) = "YES" Then
                dupYesRecords.Add dupDataRecord
            ElseIf dupDataRecord(3) = "NS" Then
                dupNotInStockRecords.Add dupDataRecord
            ElseIf dupDataRecord(3) = "B-YES" Then
                dupBYesRecords.Add dupDataRecord
            End If
        Next j
        
        ' Sort each availability group
        Call SortRecordsByLanguageAndBook(dupYesRecords)
        Call SortRecordsByLanguageAndBook(dupNotInStockRecords)
        Call SortRecordsByLanguageAndBook(dupBYesRecords)
        
        ' Add YES records first
        For j = 1 To dupYesRecords.Count
            dupSortedRecords.Add dupYesRecords.Item(j)
        Next j
        
        ' Add NS records second
        For j = 1 To dupNotInStockRecords.Count
            dupSortedRecords.Add dupNotInStockRecords.Item(j)
        Next j
        
        ' Add B-YES records third
        For j = 1 To dupBYesRecords.Count
            dupSortedRecords.Add dupBYesRecords.Item(j)
        Next j
        
        ' Process sorted records
        For j = 1 To dupSortedRecords.Count
            Dim dupDataRecord2 As Variant
            dupDataRecord2 = dupSortedRecords.Item(j)
            
            ' Check if language has changed
            Dim dupCurrentRecordLangCode As String
            dupCurrentRecordLangCode = GetLanguageShortCode(CStr(dupDataRecord2(1)))
            If dupCurrentRecordLangCode <> dupCurrentLanguage Then
                If dupCurrentLanguage <> "" Then
                    If currentRow > 1 Then
                        Dim dupLanguageStartRow As Integer
                        dupLanguageStartRow = currentRow - 1
                        While dupLanguageStartRow > 1 And pdfWs.Cells(dupLanguageStartRow - 1, 1).Value = dupCurrentLanguage
                            dupLanguageStartRow = dupLanguageStartRow - 1
                        Wend
                        If dupLanguageStartRow < currentRow - 1 Then
                            pdfWs.Range("A" & dupLanguageStartRow & ":A" & (currentRow - 1)).Merge
                            pdfWs.Range("A" & dupLanguageStartRow & ":A" & (currentRow - 1)).HorizontalAlignment = xlCenter
                            pdfWs.Range("A" & dupLanguageStartRow & ":A" & (currentRow - 1)).VerticalAlignment = xlCenter
                        End If
                    End If
                End If
                dupCurrentLanguage = dupCurrentRecordLangCode
            End If
            
            pdfWs.Cells(currentRow, 1).Value = dupCurrentRecordLangCode
            pdfWs.Cells(currentRow, 2).Value = dupDataRecord2(2)
            pdfWs.Cells(currentRow, 3).Value = dupDataRecord2(3)
            pdfWs.Cells(currentRow, 4).Value = dupDataRecord2(4)
            pdfWs.Cells(currentRow, 5).Value = ""
            pdfWs.Cells(currentRow, 6).Value = ""
            
            ' Make NS text light/gray
            If dupDataRecord2(3) = "NS" Then
                pdfWs.Cells(currentRow, 3).Font.Color = RGB(150, 150, 150)
                pdfWs.Cells(currentRow, 3).Font.Bold = False
            End If
            
            If dupDataRecord2(3) = "YES" Then
                dupYesTotalQty = dupYesTotalQty + CLng(dupDataRecord2(4))
                If dupYesStartRow = 0 Then
                    dupYesStartRow = currentRow
                End If
            Else
                If dupYesStartRow > 0 And dupYesStartRow < currentRow Then
                    pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).Merge
                    pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).HorizontalAlignment = xlCenter
                    pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).VerticalAlignment = xlCenter
                    pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).Value = "YES"
                    
                    pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).Merge
                    pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).Value = "[" & dupYesTotalQty & "]"
                    pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).HorizontalAlignment = xlLeft
                    pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).VerticalAlignment = xlCenter
                    
                    pdfWs.Range("F" & dupYesStartRow & ":F" & (currentRow - 1)).Merge
                    pdfWs.Range("F" & dupYesStartRow & ":F" & (currentRow - 1)).HorizontalAlignment = xlCenter
                    pdfWs.Range("F" & dupYesStartRow & ":F" & (currentRow - 1)).VerticalAlignment = xlCenter
                End If
                dupYesStartRow = 0
                dupYesTotalQty = 0
            End If
            
            ' Format data row
            With pdfWs.Range("A" & currentRow & ":F" & currentRow)
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .VerticalAlignment = xlCenter
            End With
            
            pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlLeft
            pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter
            pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter
            
            pdfWs.Rows(currentRow).RowHeight = 18.15
            
            totalQty = totalQty + CLng(dupDataRecord2(4))
            currentRow = currentRow + 1
        Next j
        
        ' Handle merge cells for YES values at the end
        If dupYesStartRow > 0 And dupYesStartRow < currentRow Then
            pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).Merge
            pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).HorizontalAlignment = xlCenter
            pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).VerticalAlignment = xlCenter
            pdfWs.Range("C" & dupYesStartRow & ":C" & (currentRow - 1)).Value = "YES"
            
            pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).Merge
            pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).Value = "[" & dupYesTotalQty & "]"
            pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).HorizontalAlignment = xlLeft
            pdfWs.Range("E" & dupYesStartRow & ":E" & (currentRow - 1)).VerticalAlignment = xlCenter
            
            pdfWs.Range("F" & dupYesStartRow & ":F" & (currentRow - 1)).Merge
            pdfWs.Range("F" & dupYesStartRow & ":F" & (currentRow - 1)).HorizontalAlignment = xlCenter
            pdfWs.Range("F" & dupYesStartRow & ":F" & (currentRow - 1)).VerticalAlignment = xlCenter
        End If
        
        ' Handle merge cells for the last language group
        If dupCurrentLanguage <> "" And currentRow > 1 Then
            Dim dupLastLanguageStartRow As Integer
            dupLastLanguageStartRow = currentRow - 1
            While dupLastLanguageStartRow > 1 And pdfWs.Cells(dupLastLanguageStartRow - 1, 1).Value = dupCurrentLanguage
                dupLastLanguageStartRow = dupLastLanguageStartRow - 1
            Wend
            If dupLastLanguageStartRow < currentRow - 1 Then
                pdfWs.Range("A" & dupLastLanguageStartRow & ":A" & (currentRow - 1)).Merge
                pdfWs.Range("A" & dupLastLanguageStartRow & ":A" & (currentRow - 1)).HorizontalAlignment = xlCenter
                pdfWs.Range("A" & dupLastLanguageStartRow & ":A" & (currentRow - 1)).VerticalAlignment = xlCenter
            End If
        End If
        
        ' Add total line
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
        With pdfWs.Range("A" & currentRow & ":F" & currentRow)
            .Interior.Color = RGB(240, 240, 240)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
        End With
        pdfWs.Rows(currentRow).RowHeight = 18.15
        
        currentRow = currentRow + 1
        
        ' Update page usage counter after table is completely added
        currentPageUsedRows = currentPageUsedRows + dupTableRows
        
        ' Reset total for next duplicate table
        totalQty = 0
        
        previousSectorKeyForDup = sectorKey
    Next i
    
    ' Set column widths to match details report (exact pixel values)
    pdfWs.Columns("A").ColumnWidth = 4   ' LG (62 pixels)
    pdfWs.Columns("B").ColumnWidth = 35  ' BOOK TITLE (230 pixels)
    pdfWs.Columns("C").ColumnWidth = 8  ' AVAIL (96 pixels)
    pdfWs.Columns("D").ColumnWidth = 5   ' QTY (50 pixels)
    pdfWs.Columns("E").ColumnWidth = 25  ' SIGN (317 pixels) - Made wider
    pdfWs.Columns("F").ColumnWidth = 12  ' DATE
    
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
    pdfWs.PageSetup.PrintArea = "A1:F" & lastRow
    
    ' Save as PDF
    Dim pdfPath As String
    Dim safeOrgSs As String
    safeOrgSs = Replace(Replace(orgName, " ", "_"), "\", "_")
    If customFolder <> "" Then
        pdfPath = customFolder & "\" & safeOrgSs & "_sector_summary_" & period & ".pdf"
    Else
        pdfPath = ThisWorkbook.Path & "\" & safeOrgSs & "_sector_summary_" & period & ".pdf"
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
    MsgBox "Error generating sector summary PDF: " & Err.Description, vbCritical
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
    If monthNum >= 1 And monthNum <= 12 Then
        GetShortMonthName = monthNames(monthNum - 1)
    Else
        GetShortMonthName = "UNK" ' Unknown or error month
    End If
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
    If records.Count <= 1 Then Exit Sub
    
    Dim i As Integer, j As Integer
    Dim tempRecord As Variant
    
    ' Simple bubble sort
    For i = 1 To records.Count - 1
        For j = i + 1 To records.Count
            ' Compare language first
            If records.Item(i)(1) > records.Item(j)(1) Then
                ' Swap records
                tempRecord = records.Item(i)
                records.Remove i
                records.Add tempRecord, , , j
            ElseIf records.Item(i)(1) = records.Item(j)(1) Then
                ' If language is same, compare book title
                If records.Item(i)(2) > records.Item(j)(2) Then
                    ' Swap records
                    tempRecord = records.Item(i)
                    records.Remove i
                    records.Add tempRecord, , , j
                End If
            End If
        Next j
    Next i
End Sub

