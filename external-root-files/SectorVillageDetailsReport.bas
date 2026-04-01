Option Explicit

' Sector Village Details Report Module for VIMARS
' This module handles all report generation functionality

Private Const DB_HOST As String = "localhost"
Private Const DB_PORT As String = "5432"
Private Const DB_NAME As String = "vimarshbooks"
Private Const DB_USER As String = "postgres"
Private Const DB_PASSWORD As String = "Ketan@757399"
Private Const DB_SCHEMA As String = "vimars"

' ============================================================================
' CROSS-MODULE HELPERS
' ============================================================================

Public Function GetSectorSummaryDataForPDF(period As String, groupSector As String, Optional selectedBookIds As Collection = Nothing, Optional appValue As String = "") As Collection
    ' Call the function directly instead of using Application.Run to preserve Collection objects
    If selectedBookIds Is Nothing Then
        Set GetSectorSummaryDataForPDF = SectorSummaryReport.GetSectorSummaryDataForPDF(period, groupSector, Nothing, appValue)
    Else
        Set GetSectorSummaryDataForPDF = SectorSummaryReport.GetSectorSummaryDataForPDF(period, groupSector, selectedBookIds, appValue)
    End If
End Function

Public Function GetAllSummaryDataForPDF(period As String, Optional selectedBookIds As Collection = Nothing, Optional appValue As String = "") As Collection
    ' Call the function directly instead of using Application.Run to preserve Collection objects
    If selectedBookIds Is Nothing Then
        Set GetAllSummaryDataForPDF = SummaryCityReport.GetAllSummaryDataForPDF(period, Nothing, appValue)
    Else
        Set GetAllSummaryDataForPDF = SummaryCityReport.GetAllSummaryDataForPDF(period, selectedBookIds, appValue)
    End If
End Function

' ============================================================================
' PDF GENERATION WRAPPERS
' ============================================================================

Public Sub GenerateSectorSummaryPDFWrapper(summaryData As Collection, period As String, groupSector As String, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
    ' Call the function directly instead of using Application.Run to preserve Collection objects
    If customFolder = "" Then
        Call SectorSummaryReport.GenerateSectorSummaryPDF(summaryData, period, groupSector, "", organizationName)
    Else
        Call SectorSummaryReport.GenerateSectorSummaryPDF(summaryData, period, groupSector, customFolder, organizationName)
    End If
End Sub

Public Sub GenerateAllSummaryPDFWrapper(summaryData As Collection, period As String, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
    ' Call the function directly instead of using Application.Run to preserve Collection objects
    If customFolder = "" Then
        Call SummaryCityReport.GenerateAllSummaryPDF(summaryData, period, "", organizationName)
    Else
        Call SummaryCityReport.GenerateAllSummaryPDF(summaryData, period, customFolder, organizationName)
    End If
End Sub

Private Function RunModuleCollection(funcName As String, args As Variant) As Collection
    Dim result As Variant
    result = RunModuleProcedure(funcName, args, True)
    
    If IsObject(result) Then
        Set RunModuleCollection = result
    Else
        Set RunModuleCollection = New Collection
    End If
End Function

Private Function RunModuleProcedure(funcName As String, args As Variant, Optional showError As Boolean = False) As Variant
    On Error GoTo ErrorHandler
    
    Dim macroVariants As Collection
    Set macroVariants = New Collection
    
    macroVariants.Add funcName
    
    If Not ThisWorkbook Is Nothing Then
        macroVariants.Add "'" & ThisWorkbook.Name & "'!" & funcName
    End If
    
    If Not ActiveWorkbook Is Nothing Then
        macroVariants.Add "'" & ActiveWorkbook.Name & "'!" & funcName
    End If
    
    Dim macroName As Variant
    For Each macroName In macroVariants
        On Error Resume Next
        Select Case UBound(args)
            Case -1
                RunModuleProcedure = Application.Run(CStr(macroName))
            Case 0
                RunModuleProcedure = Application.Run(CStr(macroName), args(0))
            Case 1
                RunModuleProcedure = Application.Run(CStr(macroName), args(0), args(1))
            Case 2
                RunModuleProcedure = Application.Run(CStr(macroName), args(0), args(1), args(2))
            Case 3
                RunModuleProcedure = Application.Run(CStr(macroName), args(0), args(1), args(2), args(3))
            Case Else
                RunModuleProcedure = Application.Run(CStr(macroName), args)
        End Select
        If Err.Number = 0 Then
            Exit Function
        End If
    Next macroName
    
    Err.Clear
    If showError Then
        MsgBox "Error running function '" & funcName & "': " & Err.Description, vbCritical
    End If
    RunModuleProcedure = Null
End Function

' ============================================================================
' Show report generation form
Public Sub ShowReportGenerationForm()
    On Error GoTo ErrorHandler
    
    ' Show report generation form
    ReportGenerationForm.Show
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing report generation form: " & Err.description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

' Connect to PostgreSQL
Private Function ConnectToPostgreSQL(conn As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim connectionString As String
    Dim password As String
    Dim driverOptions() As String
    Dim i As Integer
    
    ' Use hardcoded password
    password = DB_PASSWORD
    
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
                          "Server=" & DB_HOST & ";" & _
                          "Port=" & DB_PORT & ";" & _
                          "Database=" & DB_NAME & ";" & _
                          "Uid=" & DB_USER & ";" & _
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
    MsgBox "Could not connect to PostgreSQL with any available driver." & vbCrLf & _
           "Please ensure PostgreSQL ODBC driver is installed." & vbCrLf & _
           "You can download it from: https://www.postgresql.org/ftp/odbc/", vbCritical
    ConnectToPostgreSQL = False
    Exit Function
    
ErrorHandler:
    MsgBox "Error connecting to PostgreSQL: " & Err.description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
    ConnectToPostgreSQL = False
End Function

' Report Generation Form Code
Public Sub InitializeReportForm()
    ' This function will be called from the ReportGenerationForm
    ' All form initialization code is in the ReportGenerationForm.frm file
End Sub

' Generate filtered data for reports
Public Function GetFilteredReportData(groupSector As String, period As String, availYes As Boolean, availNotInStock As Boolean, availBYes As Boolean, Optional selectedBookIds As Collection = Nothing, Optional appValue As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    Dim conn As Object
    Dim reportData As Collection
    Set reportData = New Collection
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL
    If Not ConnectToPostgreSQL(conn) Then
        Set GetFilteredReportData = New Collection
        Exit Function
    End If
    
    ' Build WHERE clause based on filters
    Dim whereClause As String
    whereClause = "WHERE t.is_deleted = FALSE"
    
         ' Debug.Print "=== GetFilteredReportData Debug ==="
     ' Debug.Print "Group Sector: " & groupSector
     ' Debug.Print "Period: " & period
     ' Debug.Print "Avail Yes: " & availYes
     ' Debug.Print "Avail Not In Stock: " & availNotInStock
     ' Debug.Print "Avail B-Yes: " & availBYes
    
         ' Add group sector filter
     If groupSector <> "All" Then
         whereClause = whereClause & " AND t.group_sector = '" & Replace(groupSector, "'", "''") & "'"
         ' Debug.Print "Added group sector filter: " & groupSector
     End If
    
         ' Add period filter
     If period <> "" And period <> "All" Then
         Dim periodParts() As String
         periodParts = Split(period, "-")
         If UBound(periodParts) >= 2 Then
             ' Escape single quotes in half value for SQL
             Dim halfValue As String
             halfValue = Replace(Trim(periodParts(2)), "'", "''")
             whereClause = whereClause & " AND t.year = " & periodParts(0) & _
                          " AND t.month = " & CLng(periodParts(1)) & _
                          " AND t.half = '" & halfValue & "'"
             ' Debug.Print "Added period filter: " & period & " (Year: " & periodParts(0) & ", Month: " & periodParts(1) & ", Half: " & periodParts(2) & ")"
         End If
     End If
    
    ' Add availability filter
    Dim availConditions() As String
    Dim availCount As Integer
    availCount = 0
    
    If availYes Then
        ReDim Preserve availConditions(availCount)
        availConditions(availCount) = "t.avail = 1"
        availCount = availCount + 1
    End If
    
    If availNotInStock Then
        ReDim Preserve availConditions(availCount)
        availConditions(availCount) = "t.avail = 0"
        availCount = availCount + 1
    End If
    
    If availBYes Then
        ReDim Preserve availConditions(availCount)
        availConditions(availCount) = "t.avail = 2"
        availCount = availCount + 1
    End If
    
         If availCount > 0 Then
         whereClause = whereClause & " AND (" & Join(availConditions, " OR ") & ")"
         ' Debug.Print "Added availability filter: " & Join(availConditions, " OR ")
     End If
    
    ' Add book selection filter when provided
    If Not selectedBookIds Is Nothing Then
        Dim bookFilter As String
        bookFilter = BuildNumericInClause(selectedBookIds)
        If bookFilter = "" Then
            GoTo Cleanup
        End If
        whereClause = whereClause & " AND b.book_id IN (" & bookFilter & ")"
    End If
    
    ' Add APP filter when provided
    If appValue <> "" Then
        whereClause = whereClause & " AND t.app = '" & Replace(appValue, "'", "''") & "'"
    End If
    
    ' Build complete SQL query with consolidation of duplicate book names by contact
    ' Only merge if same book_name, village, and sector (village and sector must match for merge)
    sql = "SELECT " & _
          "MIN(t.importid) as importid, " & _
          "b.book_name, " & _
          "b.language, " & _
          "t.year, " & _
          "t.month, " & _
          "t.half, " & _
          "SUM(t.qty) as qty, " & _
          "t.avail, " & _
          "t.village, " & _
          "t.group_sector, " & _
          "t.contact_name, " & _
          "CASE WHEN t.avail = 1 THEN 'Yes' WHEN t.avail = 0 THEN 'Not In Stock' WHEN t.avail = 2 THEN 'B-Yes' ELSE 'Unknown' END AS availability_text " & _
          "FROM " & DB_SCHEMA & ".transactions t " & _
          "INNER JOIN " & DB_SCHEMA & ".books b ON t.book_id = b.book_id " & _
          whereClause & " AND b.is_deleted = FALSE " & _
          "GROUP BY b.book_name, b.language, t.year, t.month, t.half, t.avail, t.group_sector, t.contact_name, t.village " & _
          "ORDER BY t.group_sector, t.village, t.contact_name ASC, b.language ASC, b.book_name ASC;"
    
         ' Debug.Print "Report SQL: " & sql
     ' Debug.Print "Final WHERE clause: " & whereClause
    
    Set rs = conn.Execute(sql)
    
         ' Debug.Print "Records found in query: " & rs.RecordCount
    
    While Not rs.EOF
        ' Create record array
        Dim record(11) As Variant
        record(0) = rs.Fields("importid").Value
        record(1) = rs.Fields("book_name").Value
        record(2) = rs.Fields("language").Value
        record(3) = rs.Fields("year").Value
        record(4) = rs.Fields("month").Value
        record(5) = rs.Fields("half").Value
        record(6) = rs.Fields("qty").Value
        record(7) = rs.Fields("avail").Value
        record(8) = rs.Fields("village").Value
        record(9) = rs.Fields("group_sector").Value
        record(10) = rs.Fields("contact_name").Value
        record(11) = rs.Fields("availability_text").Value
        
        reportData.Add record
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
    Set GetFilteredReportData = reportData
    Exit Function
    
ErrorHandler:
    Debug.Print "Error getting filtered data: " & Err.description
    Set reportData = New Collection
    Resume Cleanup
End Function

   ' Generate PDF Report
  Public Sub GeneratePDFReport(reportData As Collection, groupSector As String, period As String, availabilityText As String, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
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
      pdfWs.Name = "VIMARS Report"
     
     ' Group data by Sector-Village only (not by availability)
     Dim groupedData As Object
     Set groupedData = CreateObject("Scripting.Dictionary")
     
     Dim record As Variant
     Dim i As Integer
     Dim key As String
     Dim sectorVillage As String
     
     ' Group the data by sector-village only
     For i = 1 To reportData.count
         record = reportData.item(i)
         sectorVillage = record(9) & " - " & record(8)  ' Group Sector - Village
         
         If Not groupedData.Exists(sectorVillage) Then
             Dim newGroup As Collection
             Set newGroup = New Collection
             groupedData.Add sectorVillage, newGroup
         End If
         
         groupedData(sectorVillage).Add record
     Next i
     
    ' Generate formatted report
    Dim currentRow As Integer
    currentRow = 1
    
    ' Separate tables into small (≤20 data rows) and large (>20 data rows)
    Dim smallTables As Object
    Dim largeTables As Object
    Set smallTables = CreateObject("Scripting.Dictionary")
    Set largeTables = CreateObject("Scripting.Dictionary")
    
    Dim groupKeys As Variant
    Dim groupKey As Variant
    Dim gkIdx As Long
    groupKeys = groupedData.Keys
    NaturalSort.SortKeysNatural groupKeys
    
    ' Categorize tables by data row count (E1, E2, ... E9, E10 order)
    For gkIdx = LBound(groupKeys) To UBound(groupKeys)
        groupKey = groupKeys(gkIdx)
        Dim groupRecordsTemp As Collection
        Set groupRecordsTemp = groupedData(groupKey)
        Dim dataRowCount As Integer
        dataRowCount = groupRecordsTemp.Count
        
        If dataRowCount <= 20 Then
            ' Small table (≤20 data rows)
            smallTables.Add groupKey, groupedData(groupKey)
        Else
            ' Large table (>20 data rows)
            largeTables.Add groupKey, groupedData(groupKey)
        End If
    Next gkIdx
    
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
    
    ' First process all small tables (≤20 data rows) - 2 per page,
    ' and after finishing small tables for a sector, immediately place its large tables.
    Dim smallTableKeys As Variant
    Dim stkIdx As Long
    smallTableKeys = smallTables.Keys
    NaturalSort.SortKeysNatural smallTableKeys
    
    For stkIdx = LBound(smallTableKeys) To UBound(smallTableKeys)
        groupKey = smallTableKeys(stkIdx)
        ' Extract sector name from sector-village key (e.g., "E8 - UTRAN" -> "E8")
        Dim currentSectorName As String
        If InStr(groupKey, " - ") > 0 Then
            currentSectorName = Split(groupKey, " - ")(0)
        Else
            currentSectorName = groupKey
        End If
        
        ' If sector changed, flush any pending large tables for the previous sector first
        If previousSectorName <> "" And currentSectorName <> previousSectorName Then
            Call ProcessLargeTablesForSector( _
                pdfWs, largeTables, previousSectorName, groupedData, _
                currentRow, sectorsOnCurrentPage, currentPageStartRow, firstTableStartRow, period, orgName)
        End If
        
        Dim groupRecords As Collection
        Set groupRecords = smallTables(groupKey)
        
        ' Calculate how many rows this sector-village will take
        Dim estimatedRowsForSector As Integer
        estimatedRowsForSector = 3  ' Header + table headers + total row
        estimatedRowsForSector = estimatedRowsForSector + groupRecords.Count  ' Data rows
        
        ' Check if we need a new page
        Dim needNewPage As Boolean
        needNewPage = False
        
        If currentRow > 1 Then  ' Not the first sector
            ' If we already have 2 sectors on this page, force new page
            If sectorsOnCurrentPage >= 2 Then
                needNewPage = True
            ElseIf sectorsOnCurrentPage = 1 Then
                ' Check if sector name is different - if so, force new page
                If currentSectorName <> previousSectorName Then
                    needNewPage = True   ' Force new page because sector is different
                ElseIf previousSectorRowCount > 23 Then
                    needNewPage = True   ' Force new page because previous was too large (>23 rows)
                ElseIf estimatedRowsForSector > 23 Then
                    needNewPage = True   ' Force new page because current is too large (>23 rows)
                Else
                    needNewPage = False  ' Keep on same page (same sector, both ≤23 rows including headers)
                End If
            End If
        End If
        
        ' Add page break if needed
        If needNewPage Then
            pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
            sectorsOnCurrentPage = 0
            currentPageStartRow = currentRow
            firstTableStartRow = currentRow
        ElseIf sectorsOnCurrentPage = 1 Then
            ' If this is the 2nd table on the same page, position header at row 25 (relative to page start)
            ' Adding 1 blank row between tables
            Dim targetRow As Integer
            targetRow = currentPageStartRow + 24   ' Row 25 of the page (pageStartRow is row 1 of the page)
            
            Dim spacingRows As Integer
            spacingRows = targetRow - currentRow
            
            ' Add spacing if needed to reach the target row
            If spacingRows > 0 Then
                For i = 1 To spacingRows
                    pdfWs.Rows(currentRow).RowHeight = 18.15
                    currentRow = currentRow + 1
                Next i
            End If
        End If
        
        ' Group records by availability within this sector-village
        Dim availGroups As Object
        Set availGroups = CreateObject("Scripting.Dictionary")
        
        Dim j As Integer
        For j = 1 To groupRecords.count
            Dim dataRecord As Variant
            dataRecord = groupRecords.item(j)
            Dim availKey As String
            availKey = dataRecord(11)  ' Availability text
            
            If Not availGroups.Exists(availKey) Then
                Dim availGroup As Collection
                Set availGroup = New Collection
                availGroups.Add availKey, availGroup
            End If
            
            availGroups(availKey).Add dataRecord
        Next j
        
        ' Process all availability groups in a single combined table
        Dim orderedAvailKeys As Collection
        Set orderedAvailKeys = New Collection
        
        ' Add "Yes" first if it exists
        If availGroups.Exists("Yes") Then
            orderedAvailKeys.Add "Yes"
        End If
        
        ' Add "Not In Stock" second if it exists
        If availGroups.Exists("Not In Stock") Then
            orderedAvailKeys.Add "Not In Stock"
        End If
        
        ' Add "B-Yes" third if it exists
        If availGroups.Exists("B-Yes") Then
            orderedAvailKeys.Add "B-Yes"
        End If
        
        Dim sectionTotalQty As Integer
        sectionTotalQty = 0
        
        ' Process combined table for this sector-village
        Call ProcessCombinedTable(pdfWs, availGroups, orderedAvailKeys, currentRow, CStr(groupKey), period, orgName)
        
        ' Calculate section total
        Dim availKey2 As Variant
        Dim availRecords As Collection
        For Each availKey2 In orderedAvailKeys
            Set availRecords = availGroups(availKey2)
            Dim totalQty As Integer
            totalQty = 0
            Dim k As Integer
            For k = 1 To availRecords.count
                totalQty = totalQty + CLng(availRecords.item(k)(6))
            Next k
            sectionTotalQty = sectionTotalQty + totalQty
        Next availKey2
        
        ' Track previous sector info for next iteration
        previousSectorRowCount = estimatedRowsForSector
        previousSectorName = currentSectorName
        sectorsOnCurrentPage = sectorsOnCurrentPage + 1
    Next stkIdx
    
    ' After finishing all small tables, flush any remaining large tables for the last sector processed
    If previousSectorName <> "" Then
        Call ProcessLargeTablesForSector( _
            pdfWs, largeTables, previousSectorName, groupedData, _
            currentRow, sectorsOnCurrentPage, currentPageStartRow, firstTableStartRow, period, orgName)
    End If
    
    ' Fallback: if any large tables are still left (sectors that had only large tables),
    ' process them now without sector grouping (old behaviour for those cases).
    If Not largeTables Is Nothing Then
        If largeTables.Count > 0 Then
            Call ProcessLargeTablesForSector( _
                pdfWs, largeTables, "", groupedData, _
                currentRow, sectorsOnCurrentPage, currentPageStartRow, firstTableStartRow, period, orgName)
        End If
    End If
      
      ' Set column widths for portrait A4 (8.3" x 11.7")
       pdfWs.Columns("A").ColumnWidth = 22  ' NAME
       pdfWs.Columns("B").ColumnWidth = 10  ' CONTACT NO.
       pdfWs.Columns("C").ColumnWidth = 3   ' LG
       pdfWs.Columns("D").ColumnWidth = 22  ' BOOK TITLE
       pdfWs.Columns("E").ColumnWidth = 5   ' AVL
       pdfWs.Columns("F").ColumnWidth = 3   ' QTY
       pdfWs.Columns("G").ColumnWidth = 18  ' SIGN
       pdfWs.Columns("H").ColumnWidth = 8   ' Date
       
       ' Enable Shrink to Fit for Name and Book Title columns
       pdfWs.Columns("A").ShrinkToFit = True   ' NAME
       pdfWs.Columns("D").ShrinkToFit = True   ' BOOK TITLE
     
       ' Set page format to A4 Portrait with optimized margins for punching
       With pdfWs.PageSetup
           .PaperSize = xlPaperA4
           .Orientation = xlPortrait
           .FitToPagesWide = 1
           .FitToPagesTall = False  ' Don't force to 1 page tall
           .Zoom = False
           .TopMargin = Application.InchesToPoints(0.1)    ' Minimal margin at top
           .BottomMargin = Application.InchesToPoints(0.1)  ' Minimal margin at bottom
           .LeftMargin = Application.InchesToPoints(0.4)    ' More margin on left for punching
           .RightMargin = Application.InchesToPoints(0.1)   ' Minimal margin on right
           .HeaderMargin = Application.InchesToPoints(0.1)
           .FooterMargin = Application.InchesToPoints(0.1)
           .PrintGridlines = False
           .PrintHeadings = False
       End With
      
               ' Set print area to include all data AFTER all processing is complete
        Dim lastRow As Long
        lastRow = pdfWs.Cells(pdfWs.Rows.count, 1).End(xlUp).row
        

        
        pdfWs.PageSetup.PrintArea = "A1:H" & lastRow
      
          ' Save as PDF (Org first: VIMARSH_sector_village_details_period.pdf)
    Dim pdfPath As String
    Dim safeOrg As String
    safeOrg = Replace(Replace(orgName, " ", "_"), "\", "_")
    If customFolder <> "" Then
        pdfPath = customFolder & "\" & safeOrg & "_sector_village_details_" & period & ".pdf"
    Else
        pdfPath = ThisWorkbook.Path & "\" & safeOrg & "_sector_village_details_" & period & ".pdf"
    End If
      
      pdfWb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False
     
           ' Close temporary workbook
      pdfWb.Close False
      
      ' Re-enable Excel alerts
      Application.DisplayAlerts = True
      
          ' Individual success message removed - only final message will be shown
      
      Exit Sub
      
ErrorHandler:
      ' Re-enable Excel alerts even if error occurs
      Application.DisplayAlerts = True
      MsgBox "Error generating PDF: " & Err.description, vbCritical
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

' Helper function to convert text to Camel Case
Private Function ToCamelCase(text As Variant) As String
    If text = "" Or IsNull(text) Then
        ToCamelCase = ""
        Exit Function
    End If
    
    Dim textStr As String
    textStr = CStr(text)
    
    Dim words() As String
    Dim result As String
    Dim i As Integer
    
    ' Split by spaces
    words = Split(Trim(textStr), " ")
    result = ""
    
    For i = 0 To UBound(words)
        If words(i) <> "" Then
            ' Convert first letter to uppercase, rest to lowercase
            If i > 0 Then result = result & " "
            result = result & UCase(Left(words(i), 1)) & LCase(Mid(words(i), 2))
        End If
    Next i
    
    ToCamelCase = result
End Function

' Helper function to extract contact number from contact name
 Private Function GetContactNumber(contactName As String) As String
     ' Extract contact number from contact name
     ' Assuming format: "Name (Number)" or "Name - Number"
     Dim contactNumber As String
     
     ' Check if contact name contains parentheses
     If InStr(contactName, "(") > 0 And InStr(contactName, ")") > 0 Then
         ' Extract number from parentheses
         Dim startPos As Integer
         Dim endPos As Integer
         startPos = InStr(contactName, "(") + 1
         endPos = InStr(contactName, ")") - 1
         contactNumber = Mid(contactName, startPos, endPos - startPos + 1)
     ElseIf InStr(contactName, "-") > 0 Then
         ' Extract number after dash
         contactNumber = Trim(Split(contactName, "-")(1))
     Else
         ' If no special format, return empty
         contactNumber = ""
     End If
     
     ' Format: add space after 5 digits (e.g., "98765 43210")
     If Len(contactNumber) = 10 Then
         contactNumber = Left(contactNumber, 5) & " " & Right(contactNumber, 5)
     End If
     
     GetContactNumber = contactNumber
 End Function
 
 ' Helper function to extract contact name only
 Private Function GetContactNameOnly(contactName As String) As String
     ' Extract only the name part from contact name
     Dim nameOnly As String
     
     ' Check if contact name contains parentheses
     If InStr(contactName, "(") > 0 Then
         ' Extract name before parentheses
         nameOnly = Trim(Left(contactName, InStr(contactName, "(") - 1))
     ElseIf InStr(contactName, "-") > 0 Then
         ' Extract name before dash
         nameOnly = Trim(Split(contactName, "-")(0))
     Else
         ' If no special format, return as is
         nameOnly = contactName
     End If
     
     GetContactNameOnly = nameOnly
 End Function

' Process combined table for all availability types in one table
Private Sub ProcessCombinedTable(pdfWs As Worksheet, availGroups As Object, orderedAvailKeys As Collection, ByRef currentRow As Integer, groupKey As String, period As String, Optional organizationName As String = "VIMARSH")
    On Error GoTo ErrorHandler
    
    ' Add header line just above the table
    Call AddTableHeader(pdfWs, currentRow, groupKey, period, organizationName)
    currentRow = currentRow + 1
    
    ' Add table headers
    Dim tableStartRow As Integer
    tableStartRow = currentRow
    Call AddTableHeaders(pdfWs, currentRow)
    currentRow = currentRow + 1
    
    Dim availKey As Variant
    Dim availRecords As Collection
    Dim isFirstGroup As Boolean
    isFirstGroup = True
    Dim lastAvailEndRow As Integer
    lastAvailEndRow = currentRow - 1
    
    ' Track Yes section rows for merging
    Dim yesStartRow As Integer
    Dim yesEndRow As Integer
    yesStartRow = 0
    yesEndRow = 0
    
    ' Process each availability group
    For Each availKey In orderedAvailKeys
        Set availRecords = availGroups(availKey)
        
        Dim groupStartRow As Integer
        groupStartRow = currentRow
        
        ' Add bold separator line between availability groups (not before first)
        If Not isFirstGroup Then
            ' Add bold bottom border to previous row to separate groups
            pdfWs.Range("A" & lastAvailEndRow & ":H" & lastAvailEndRow).Borders(xlEdgeBottom).Weight = xlThin
        End If
        isFirstGroup = False
        
        ' Group records by contact name
        Dim contactGroups As Object
        Set contactGroups = CreateObject("Scripting.Dictionary")
        
        Dim i As Integer
        For i = 1 To availRecords.count
            Dim record As Variant
            record = availRecords.item(i)
            Dim contactName As String
            contactName = record(10)
            
            If Not contactGroups.Exists(contactName) Then
                Dim contactGroup As Collection
                Set contactGroup = New Collection
                contactGroups.Add contactName, contactGroup
            End If
            
            contactGroups(contactName).Add record
        Next i
        
        ' Process each contact group
        Dim contactKey As Variant
        For Each contactKey In contactGroups.Keys
            Dim contactRecords As Collection
            Set contactRecords = contactGroups(contactKey)
            
            Dim j As Integer
            For j = 1 To contactRecords.count
                Dim dataRecord As Variant
                dataRecord = contactRecords.item(j)
                
                pdfWs.Cells(currentRow, 1).Value = ToCamelCase(GetContactNameOnly(CStr(dataRecord(10))))
                pdfWs.Cells(currentRow, 1).Font.Size = 9
                pdfWs.Cells(currentRow, 2).Value = GetContactNumber(CStr(dataRecord(10)))
                pdfWs.Cells(currentRow, 3).Value = GetLanguageShortCode(CStr(dataRecord(2)))
                pdfWs.Cells(currentRow, 4).Value = dataRecord(1)
                Call FormatAvailabilityCell(pdfWs.Cells(currentRow, 5), CStr(dataRecord(11)))
                pdfWs.Cells(currentRow, 6).Value = dataRecord(6)
                pdfWs.Cells(currentRow, 7).Value = ""
                
                Call FormatDataRow(pdfWs, currentRow)
                currentRow = currentRow + 1
            Next j
        Next contactKey
        
        lastAvailEndRow = currentRow - 1
        
        ' Track Yes section for full merging
        If availKey = "Yes" Then
            yesStartRow = groupStartRow
            yesEndRow = lastAvailEndRow
        End If
        
        ' Apply merging based on availability type
        If availKey = "Yes" Then
            ' Full merge for Yes (NAME, Mo.No., LG, AVL, SIGN, DATE)
            Call ApplyCellMergingForYesTable(pdfWs, groupStartRow, lastAvailEndRow)
        Else
            ' Limited merge for NS/B-Yes (only NAME, Mo.No., LG - no AVL, SIGN, DATE merge)
            Call ApplyCellMergingForRegularTable(pdfWs, groupStartRow, lastAvailEndRow)
        End If
    Next availKey
    
    ' Add bold outer border to entire table
    With pdfWs.Range("A" & tableStartRow & ":H" & lastAvailEndRow)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    ' Add total row + compute total data lines for smart shrink
    Dim grandTotal As Integer
    Dim totalDataRows As Integer
    grandTotal = 0
    totalDataRows = 0
    For Each availKey In orderedAvailKeys
        Set availRecords = availGroups(availKey)
        totalDataRows = totalDataRows + availRecords.count
        For i = 1 To availRecords.count
            grandTotal = grandTotal + CLng(availRecords.item(i)(6))
        Next i
    Next availKey
    
    ' Add total row at currentRow
    Call AddTotalRow(pdfWs, currentRow, grandTotal)
    
    ' --- Smart row height adjustment for near-overflow full-page tables ---
    ' If this table has data lines >42 and ≤44, slightly reduce row height JUST for this table
    ' so it fits on one page. This does not affect other tables/pages.
    Const MAX_DATA_ROWS_SINGLE_PAGE As Integer = 42
    Const DEFAULT_ROW_HEIGHT As Double = 18.15
    Dim totalLinesThisTable As Integer
    Dim totalRowIndex As Integer
    totalRowIndex = currentRow
    ' Approx: header row + data rows + total row = totalDataRows + 2
    totalLinesThisTable = totalDataRows + 2
    
    If totalDataRows > MAX_DATA_ROWS_SINGLE_PAGE And totalDataRows <= MAX_DATA_ROWS_SINGLE_PAGE + 2 Then
        Dim adjustedHeight As Double
        adjustedHeight = DEFAULT_ROW_HEIGHT * (MAX_DATA_ROWS_SINGLE_PAGE / totalLinesThisTable)
        If adjustedHeight < DEFAULT_ROW_HEIGHT And adjustedHeight > 10# Then
            Dim adjRow As Long
            For adjRow = tableStartRow To totalRowIndex
                pdfWs.Rows(adjRow).RowHeight = adjustedHeight
            Next adjRow
        End If
    End If
    
    currentRow = currentRow + 2
    
    Exit Sub
    
ErrorHandler:
End Sub

' Process "Yes" table with page break handling
Private Sub ProcessYesTableWithPageBreaks(pdfWs As Worksheet, availRecords As Collection, ByRef currentRow As Integer, groupKey As String, period As String)
    On Error GoTo ErrorHandler
    
    ' Add header line just above the table
    Call AddTableHeader(pdfWs, currentRow, groupKey, period)
    currentRow = currentRow + 1
    
    ' Add table headers
    Call AddTableHeaders(pdfWs, currentRow)
    currentRow = currentRow + 1
    
    ' Group records by contact name to keep merged cells together
    Dim contactGroups As Object
    Set contactGroups = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 1 To availRecords.count
        Dim record As Variant
        record = availRecords.item(i)
        Dim contactName As String
        contactName = record(10)  ' Contact name
        
        If Not contactGroups.Exists(contactName) Then
            Dim contactGroup As Collection
            Set contactGroup = New Collection
            contactGroups.Add contactName, contactGroup
        End If
        
        contactGroups(contactName).Add record
    Next i
    
         ' Process each contact group with much smarter page break logic
     Dim contactKey As Variant
     Dim contactRecords As Collection
     Dim totalQty As Integer
     Dim contactGroupsList As Collection
     Set contactGroupsList = New Collection
     totalQty = 0
     
     ' First, collect all contact groups to analyze them together
     For Each contactKey In contactGroups.Keys
         contactGroupsList.Add contactKey
     Next contactKey
     
     ' Now process contact groups with much smarter page break logic
     Dim groupIndex As Integer
     Dim currentPageStartRow As Integer
     Dim accumulatedRows As Integer
     Dim totalRowsForPage As Integer
     currentPageStartRow = currentRow
     accumulatedRows = 0
     totalRowsForPage = 0
     
     For groupIndex = 1 To contactGroupsList.count
         Set contactRecords = contactGroups(contactGroupsList(groupIndex))
         
         ' Calculate space needed for this contact group
         Dim rowsNeeded As Integer
         rowsNeeded = contactRecords.count
         
         ' Calculate total space needed including total row and spacing
         Dim totalSpaceNeeded As Integer
         totalSpaceNeeded = rowsNeeded + 3  ' +3 for total row and spacing
         
                   ' Check if adding this contact group would exceed page limit
          ' Only add page break if we're really close to the page limit (allow much more content per page)
          If currentRow + totalSpaceNeeded > 80 Then  ' Increased limit to 80 to allow much more content per page
              ' Debug.Print "  *** YES TABLE PAGE BREAK *** - Adding page break at row " & currentRow
              ' Add page break
              pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
              
              ' Add header again on new page
              Call AddTableHeader(pdfWs, currentRow, groupKey, period)
              currentRow = currentRow + 1
              Call AddTableHeaders(pdfWs, currentRow)
              currentRow = currentRow + 1
              ' Debug.Print "  Headers added on new page, currentRow now: " & currentRow
          End If
         
         ' Add all records for this contact
         Dim startRow As Integer
         startRow = currentRow
         
         For i = 1 To contactRecords.count
             Dim dataRecord As Variant
             dataRecord = contactRecords.item(i)
             
             pdfWs.Cells(currentRow, 1).Value = ToCamelCase(GetContactNameOnly(CStr(dataRecord(10))))  ' Name only
             pdfWs.Cells(currentRow, 1).Font.Size = 9  ' Smaller font for Name
             pdfWs.Cells(currentRow, 2).Value = GetContactNumber(CStr(dataRecord(10)))  ' Contact No.
             pdfWs.Cells(currentRow, 3).Value = GetLanguageShortCode(CStr(dataRecord(2)))   ' Language
             pdfWs.Cells(currentRow, 4).Value = dataRecord(1)   ' Book Name
             Call FormatAvailabilityCell(pdfWs.Cells(currentRow, 5), CStr(dataRecord(11)))  ' Availability
             pdfWs.Cells(currentRow, 6).Value = dataRecord(6)   ' Quantity
             pdfWs.Cells(currentRow, 7).Value = ""              ' Sign (will be filled after merging)
             
             ' Format data row
             Call FormatDataRow(pdfWs, currentRow)
             
             totalQty = totalQty + CLng(dataRecord(6))
             currentRow = currentRow + 1
         Next i
         
         ' Sort data by Language and Book Title within each contact group
         If currentRow - 1 >= startRow Then
             pdfWs.Range("A" & startRow & ":H" & currentRow - 1).Sort _
                 Key1:=pdfWs.Range("B" & startRow), Order1:=xlAscending, _
                 Key2:=pdfWs.Range("C" & startRow), Order2:=xlAscending, _
                 Header:=xlNo
         End If
         
         ' Apply cell merging for this contact group
         Call ApplyCellMergingForYesTable(pdfWs, startRow, currentRow - 1)
     Next groupIndex
    
    ' Add total row
    Call AddTotalRow(pdfWs, currentRow, totalQty)
    currentRow = currentRow + 2  ' Add space after table
    
    Exit Sub
    
ErrorHandler:
End Sub

' Process regular availability table (Not In Stock, B-Yes)
Private Sub ProcessAvailabilityTable(pdfWs As Worksheet, availRecords As Collection, ByRef currentRow As Integer, groupKey As String, period As String, availType As String)
    On Error GoTo ErrorHandler
    
    ' Add header line just above the table
    Call AddTableHeader(pdfWs, currentRow, groupKey, period)
    currentRow = currentRow + 1
    
    ' Add table headers
    Call AddTableHeaders(pdfWs, currentRow)
    currentRow = currentRow + 1
    
    ' Group records by contact name to keep merged cells together
    Dim contactGroups As Object
    Set contactGroups = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 1 To availRecords.count
        Dim record As Variant
        record = availRecords.item(i)
        Dim contactName As String
        contactName = record(10)  ' Contact name
        
        If Not contactGroups.Exists(contactName) Then
            Dim contactGroup As Collection
            Set contactGroup = New Collection
            contactGroups.Add contactName, contactGroup
        End If
        
        contactGroups(contactName).Add record
    Next i
    
    ' Process each contact group
    Dim contactKey As Variant
    Dim contactRecords As Collection
    Dim totalQty As Integer
    totalQty = 0
    
    For Each contactKey In contactGroups.Keys
        Set contactRecords = contactGroups(contactKey)
        
        ' Add all records for this contact
        Dim startRow As Integer
        startRow = currentRow
        
        For i = 1 To contactRecords.count
            Dim dataRecord As Variant
            dataRecord = contactRecords.item(i)
            
            pdfWs.Cells(currentRow, 1).Value = ToCamelCase(GetContactNameOnly(CStr(dataRecord(10))))  ' Name only
            pdfWs.Cells(currentRow, 1).Font.Size = 9  ' Smaller font for Name
            pdfWs.Cells(currentRow, 2).Value = GetContactNumber(CStr(dataRecord(10)))  ' Contact No.
            pdfWs.Cells(currentRow, 3).Value = GetLanguageShortCode(CStr(dataRecord(2)))   ' Language
            pdfWs.Cells(currentRow, 4).Value = dataRecord(1)   ' Book Name
            Call FormatAvailabilityCell(pdfWs.Cells(currentRow, 5), CStr(dataRecord(11)))  ' Availability
            pdfWs.Cells(currentRow, 6).Value = dataRecord(6)   ' Quantity
            pdfWs.Cells(currentRow, 7).Value = ""              ' Sign (empty)
            
            ' Format data row
            Call FormatDataRow(pdfWs, currentRow)
            
            totalQty = totalQty + CLng(dataRecord(6))
            currentRow = currentRow + 1
        Next i
        
        ' Sort data by Language and Book Title within each contact group
        If currentRow - 1 >= startRow Then
            pdfWs.Range("A" & startRow & ":H" & currentRow - 1).Sort _
                Key1:=pdfWs.Range("B" & startRow), Order1:=xlAscending, _
                Key2:=pdfWs.Range("C" & startRow), Order2:=xlAscending, _
                Header:=xlNo
        End If
        
        ' Apply cell merging for this contact group
        Call ApplyCellMergingForRegularTable(pdfWs, startRow, currentRow - 1)
    Next contactKey
    
    ' Add total row
    Call AddTotalRow(pdfWs, currentRow, totalQty)
    currentRow = currentRow + 2  ' Add space after table
    
    Exit Sub
    
ErrorHandler:
End Sub

' Add table header (sector-village and period)
Private Sub AddTableHeader(pdfWs As Worksheet, currentRow As Integer, groupKey As String, period As String, Optional organizationName As String = "VIMARSH")
    ' Add the sector-village and organization header line
    Dim sectorVillageParts() As String
    sectorVillageParts = Split(groupKey, " - ")
    pdfWs.Cells(currentRow, 1).Value = sectorVillageParts(0) & " - " & sectorVillageParts(1)
    pdfWs.Cells(currentRow, 1).Font.Bold = True
    pdfWs.Cells(currentRow, 1).Font.Size = 12
    pdfWs.Cells(currentRow, 1).Font.Name = "Calibri"
    pdfWs.Range("A" & currentRow & ":C" & currentRow).Merge
    pdfWs.Range("A" & currentRow & ":C" & currentRow).HorizontalAlignment = xlLeft
    
    ' Add period header on the same row
    Dim periodParts() As String
    Dim orgName As String
    If organizationName = "" Then
        orgName = "VIMARSH"
    Else
        orgName = organizationName
    End If
    periodParts = Split(period, "-")
    pdfWs.Cells(currentRow, 4).Value = orgName & ": " & GetShortMonthName(CLng(periodParts(1))) & "-" & periodParts(0) & "(" & GetHalfRange(periodParts(2), CLng(periodParts(1)), CLng(periodParts(0))) & ")"
    pdfWs.Cells(currentRow, 4).Font.Bold = True
    pdfWs.Cells(currentRow, 4).Font.Size = 12
    pdfWs.Cells(currentRow, 4).Font.Name = "Calibri"
    pdfWs.Range("D" & currentRow & ":F" & currentRow).Merge
    pdfWs.Range("D" & currentRow & ":F" & currentRow).HorizontalAlignment = xlCenter
    
    ' Add "Details Name" in Date column (H)
    pdfWs.Cells(currentRow, 8).Value = "Details Name"
    pdfWs.Cells(currentRow, 8).Font.Bold = True
    pdfWs.Cells(currentRow, 8).Font.Size = 12
    pdfWs.Cells(currentRow, 8).Font.Name = "Calibri"
    pdfWs.Cells(currentRow, 8).HorizontalAlignment = xlRight
End Sub

' Add table headers
Private Sub AddTableHeaders(pdfWs As Worksheet, currentRow As Integer)
    pdfWs.Cells(currentRow, 1).Value = "Name"
    pdfWs.Cells(currentRow, 2).Value = "Mo.No."
    pdfWs.Cells(currentRow, 3).Value = "LG"
    pdfWs.Cells(currentRow, 4).Value = "Book Name"
    pdfWs.Cells(currentRow, 5).Value = "AVL"
    pdfWs.Cells(currentRow, 6).Value = "QTY"
    pdfWs.Cells(currentRow, 7).Value = "Sign"
    pdfWs.Cells(currentRow, 8).Value = "Date"
    
    ' Format headers with borders and Calibri font
    With pdfWs.Range("A" & currentRow & ":H" & currentRow)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Interior.Color = RGB(220, 220, 220)
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline
        ' Bold outer border for header
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    ' Set specific alignments for headers
    pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlLeft   ' NAME - Left
    pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlCenter ' CONTACT NO. - Center
    pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter ' LG - Center
    pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlLeft   ' BOOK TITLE - Left
    pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter ' AVL - Center
    pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter ' QTY - Center
    pdfWs.Cells(currentRow, 7).HorizontalAlignment = xlCenter ' SIGN - Center
    pdfWs.Cells(currentRow, 8).HorizontalAlignment = xlCenter ' Date - Center
    
    pdfWs.Rows(currentRow).RowHeight = 18.15
End Sub

' Format data row
Private Sub FormatDataRow(pdfWs As Worksheet, currentRow As Integer)
    ' Add borders to data row and set row height to 21 pixels with Calibri font
    With pdfWs.Range("A" & currentRow & ":H" & currentRow)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline
        .VerticalAlignment = xlCenter
    End With
    
    ' Set specific alignments for data cells
    pdfWs.Cells(currentRow, 1).HorizontalAlignment = xlLeft   ' NAME - Left
    pdfWs.Cells(currentRow, 2).HorizontalAlignment = xlCenter ' CONTACT NO. - Center
    pdfWs.Cells(currentRow, 3).HorizontalAlignment = xlCenter ' LG - Center
    pdfWs.Cells(currentRow, 4).HorizontalAlignment = xlLeft   ' BOOK TITLE - Left
    pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter ' AVL - Center
    pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter ' QTY - Center
    pdfWs.Cells(currentRow, 7).HorizontalAlignment = xlCenter ' SIGN - Center
    pdfWs.Cells(currentRow, 8).HorizontalAlignment = xlCenter ' Date - Center
    
    pdfWs.Rows(currentRow).RowHeight = 18.15
End Sub

' Apply cell merging for Yes table (NAME, LNG, AVAIL, SIGN with total)
Private Sub ApplyCellMergingForYesTable(pdfWs As Worksheet, startRow As Integer, endRow As Integer)
    ' Merge cells for same contact names
    Dim mergeStartRow As Integer
    Dim mergeEndRow As Integer
    Dim currentContact As String
    Dim currentLanguage As String
    Dim currentAvailability As String
    
    mergeStartRow = startRow
    currentContact = pdfWs.Cells(startRow, 1).Value
    currentLanguage = pdfWs.Cells(startRow, 3).Value
    currentAvailability = pdfWs.Cells(startRow, 5).Value
    
    Dim k As Integer
    For k = startRow + 1 To endRow + 1
        Dim nextContact As String
        Dim nextLanguage As String
        Dim nextAvailability As String
        
        If k <= endRow Then
            nextContact = pdfWs.Cells(k, 1).Value
            nextLanguage = pdfWs.Cells(k, 3).Value
            nextAvailability = pdfWs.Cells(k, 4).Value
        Else
            nextContact = ""
            nextLanguage = ""
            nextAvailability = ""
        End If
        
        ' Check if contact name changed
        If nextContact <> currentContact Or k > endRow Then
            mergeEndRow = k - 1
            
            ' Merge contact name cells if more than one row
            If mergeEndRow > mergeStartRow Then
                pdfWs.Range("A" & mergeStartRow & ":A" & mergeEndRow).Merge
                pdfWs.Cells(mergeStartRow, 1).VerticalAlignment = xlCenter
            End If
            
            ' Merge Mo.No. column (B) same as NAME column
            If mergeEndRow > mergeStartRow Then
                pdfWs.Range("B" & mergeStartRow & ":B" & mergeEndRow).Merge
                pdfWs.Cells(mergeStartRow, 2).VerticalAlignment = xlCenter
            End If
            
            ' Check for language merging within this contact group (column C)
            Dim langStartRow As Integer
            Dim langEndRow As Integer
            Dim m As Integer
            langStartRow = mergeStartRow
            currentLanguage = pdfWs.Cells(mergeStartRow, 3).Value
            
            For m = mergeStartRow + 1 To mergeEndRow + 1
                If m <= mergeEndRow Then
                    nextLanguage = pdfWs.Cells(m, 3).Value
                Else
                    nextLanguage = ""
                End If
                
                If nextLanguage <> currentLanguage Or m > mergeEndRow Then
                    langEndRow = m - 1
                    
                    ' Merge language cells if more than one row
                    If langEndRow > langStartRow Then
                        pdfWs.Range("C" & langStartRow & ":C" & langEndRow).Merge
                        pdfWs.Cells(langStartRow, 3).VerticalAlignment = xlCenter
                    End If
                    
                    langStartRow = m
                    If m <= mergeEndRow Then
                        currentLanguage = pdfWs.Cells(m, 3).Value
                    End If
                End If
            Next m
            
            ' Check for availability merging within this contact group (column E - AVL)
            Dim availStartRow As Integer
            Dim availEndRow As Integer
            availStartRow = mergeStartRow
            currentAvailability = pdfWs.Cells(mergeStartRow, 5).Value
            
            For m = mergeStartRow + 1 To mergeEndRow + 1
                If m <= mergeEndRow Then
                    nextAvailability = pdfWs.Cells(m, 5).Value
                Else
                    nextAvailability = ""
                End If
                
                If nextAvailability <> currentAvailability Or m > mergeEndRow Then
                    availEndRow = m - 1
                    
                    ' Merge availability cells if more than one row
                    If availEndRow > availStartRow Then
                        pdfWs.Range("E" & availStartRow & ":E" & availEndRow).Merge
                        pdfWs.Cells(availStartRow, 5).VerticalAlignment = xlCenter
                    End If
                    
                    availStartRow = m
                    If m <= mergeEndRow Then
                        currentAvailability = pdfWs.Cells(m, 5).Value
                    End If
                End If
            Next m
            
            ' Merge SIGN (G) and Date (H) columns same as NAME column and add total quantity
            If mergeEndRow > mergeStartRow Then
                pdfWs.Range("G" & mergeStartRow & ":G" & mergeEndRow).Merge
                pdfWs.Cells(mergeStartRow, 7).VerticalAlignment = xlCenter
                pdfWs.Range("H" & mergeStartRow & ":H" & mergeEndRow).Merge
                pdfWs.Cells(mergeStartRow, 8).VerticalAlignment = xlCenter
                
                ' Calculate total quantity for this contact
                Dim totalContactQty As Integer
                totalContactQty = 0
                For m = mergeStartRow To mergeEndRow
                    totalContactQty = totalContactQty + CLng(pdfWs.Cells(m, 6).Value)
                Next m
                
                ' Set the merged SIGN cell value to [total]
                pdfWs.Cells(mergeStartRow, 7).Value = "[" & totalContactQty & "]"
                pdfWs.Cells(mergeStartRow, 7).HorizontalAlignment = xlLeft
            End If
            
            mergeStartRow = k
            If k <= endRow Then
                currentContact = pdfWs.Cells(k, 1).Value
                currentLanguage = pdfWs.Cells(k, 3).Value
                currentAvailability = pdfWs.Cells(k, 5).Value
            End If
        End If
    Next k
End Sub

' Apply cell merging for regular tables (only NAME and LNG)
Private Sub ApplyCellMergingForRegularTable(pdfWs As Worksheet, startRow As Integer, endRow As Integer)
    ' Merge cells for same contact names
    Dim mergeStartRow As Integer
    Dim mergeEndRow As Integer
    Dim currentContact As String
    Dim currentLanguage As String
    
    mergeStartRow = startRow
    currentContact = pdfWs.Cells(startRow, 1).Value
    currentLanguage = pdfWs.Cells(startRow, 3).Value
    
    Dim k As Integer
    For k = startRow + 1 To endRow + 1
        Dim nextContact As String
        Dim nextLanguage As String
        
        If k <= endRow Then
            nextContact = pdfWs.Cells(k, 1).Value
            nextLanguage = pdfWs.Cells(k, 3).Value
        Else
            nextContact = ""
            nextLanguage = ""
        End If
        
        ' Check if contact name changed
        If nextContact <> currentContact Or k > endRow Then
            mergeEndRow = k - 1
            
            ' Merge contact name cells if more than one row
            If mergeEndRow > mergeStartRow Then
                pdfWs.Range("A" & mergeStartRow & ":A" & mergeEndRow).Merge
                pdfWs.Cells(mergeStartRow, 1).VerticalAlignment = xlCenter
            End If
            
            ' Merge Mo.No. column (B) same as NAME column
            If mergeEndRow > mergeStartRow Then
                pdfWs.Range("B" & mergeStartRow & ":B" & mergeEndRow).Merge
                pdfWs.Cells(mergeStartRow, 2).VerticalAlignment = xlCenter
            End If
            
            ' Check for language merging within this contact group (column C)
            Dim langStartRow As Integer
            Dim langEndRow As Integer
            Dim m As Integer
            langStartRow = mergeStartRow
            currentLanguage = pdfWs.Cells(mergeStartRow, 3).Value
            
            For m = mergeStartRow + 1 To mergeEndRow + 1
                If m <= mergeEndRow Then
                    nextLanguage = pdfWs.Cells(m, 3).Value
                Else
                    nextLanguage = ""
                End If
                
                If nextLanguage <> currentLanguage Or m > mergeEndRow Then
                    langEndRow = m - 1
                    
                    ' Merge language cells if more than one row
                    If langEndRow > langStartRow Then
                        pdfWs.Range("C" & langStartRow & ":C" & langEndRow).Merge
                        pdfWs.Cells(langStartRow, 3).VerticalAlignment = xlCenter
                    End If
                    
                    langStartRow = m
                    If m <= mergeEndRow Then
                        currentLanguage = pdfWs.Cells(m, 3).Value
                    End If
                End If
            Next m
            
            mergeStartRow = k
            If k <= endRow Then
                currentContact = pdfWs.Cells(k, 1).Value
                currentLanguage = pdfWs.Cells(k, 3).Value
            End If
        End If
    Next k
End Sub

' Process all large tables for a given sector so that full-page villages for that sector
' come immediately after that sector's half-page (small) villages.
Private Sub ProcessLargeTablesForSector( _
    pdfWs As Worksheet, _
    largeTables As Object, _
    ByVal targetSectorName As String, _
    groupedData As Object, _
    ByRef currentRow As Integer, _
    ByRef sectorsOnCurrentPage As Integer, _
    ByRef currentPageStartRow As Integer, _
    ByRef firstTableStartRow As Integer, _
    ByVal period As String, _
    ByVal orgName As String)
    
    On Error GoTo ErrorHandler
    
    If largeTables Is Nothing Then Exit Sub
    If largeTables.Count = 0 Then Exit Sub
    
    Dim keys As Variant
    keys = largeTables.Keys
    NaturalSort.SortKeysNatural keys
    
    Dim idx As Long
    For idx = LBound(keys) To UBound(keys)
        Dim groupKey As Variant
        groupKey = keys(idx)
        
        Dim sectorNameFromKey As String
        If InStr(CStr(groupKey), " - ") > 0 Then
            sectorNameFromKey = Split(CStr(groupKey), " - ")(0)
        Else
            sectorNameFromKey = CStr(groupKey)
        End If
        
        ' If targetSectorName is empty, process all remaining large tables (fallback)
        If (targetSectorName = "" Or sectorNameFromKey = targetSectorName) And largeTables.Exists(groupKey) Then
            Dim groupRecordsLarge As Collection
            Set groupRecordsLarge = largeTables(groupKey)
            
            ' Large tables always start on new page
            If currentRow > 1 Then
                pdfWs.HPageBreaks.Add pdfWs.Rows(currentRow)
                sectorsOnCurrentPage = 0
                currentPageStartRow = currentRow
                firstTableStartRow = currentRow
            End If
            
            ' Group records by availability within this sector-village
            Dim availGroupsLarge As Object
            Set availGroupsLarge = CreateObject("Scripting.Dictionary")
            
            Dim jLarge As Integer
            For jLarge = 1 To groupRecordsLarge.count
                Dim dataRecordLarge As Variant
                dataRecordLarge = groupRecordsLarge.item(jLarge)
                Dim availKeyLarge As String
                availKeyLarge = dataRecordLarge(11)  ' Availability text
                
                If Not availGroupsLarge.Exists(availKeyLarge) Then
                    Dim availGroupLarge As Collection
                    Set availGroupLarge = New Collection
                    availGroupsLarge.Add availKeyLarge, availGroupLarge
                End If
                
                availGroupsLarge(availKeyLarge).Add dataRecordLarge
            Next jLarge
            
            ' Process all availability groups in a single combined table
            Dim orderedAvailKeysLarge As Collection
            Set orderedAvailKeysLarge = New Collection
            
            ' Add "Yes" first if it exists
            If availGroupsLarge.Exists("Yes") Then
                orderedAvailKeysLarge.Add "Yes"
            End If
            
            ' Add "Not In Stock" second if it exists
            If availGroupsLarge.Exists("Not In Stock") Then
                orderedAvailKeysLarge.Add "Not In Stock"
            End If
            
            ' Add "B-Yes" third if it exists
            If availGroupsLarge.Exists("B-Yes") Then
                orderedAvailKeysLarge.Add "B-Yes"
            End If
            
            Dim sectionTotalQtyLarge As Integer
            sectionTotalQtyLarge = 0
            
            ' Process combined table for this sector-village
            Call ProcessCombinedTable(pdfWs, availGroupsLarge, orderedAvailKeysLarge, currentRow, CStr(groupKey), period, orgName)
            
            ' Calculate section total (kept for parity with original logic, though not used directly here)
            Dim availKey2Large As Variant
            Dim availRecordsLarge As Collection
            For Each availKey2Large In orderedAvailKeysLarge
                Set availRecordsLarge = availGroupsLarge(availKey2Large)
                Dim totalQtyLarge As Integer
                totalQtyLarge = 0
                Dim kLarge As Integer
                For kLarge = 1 To availRecordsLarge.count
                    totalQtyLarge = totalQtyLarge + CLng(availRecordsLarge.item(kLarge)(6))
                Next kLarge
                sectionTotalQtyLarge = sectionTotalQtyLarge + totalQtyLarge
            Next availKey2Large
            
            sectorsOnCurrentPage = sectorsOnCurrentPage + 1
            
            ' Remove from dictionary so it is not processed again
            largeTables.Remove groupKey
        End If
    Next idx
    
    Exit Sub
    
ErrorHandler:
    ' Swallow errors to avoid breaking overall PDF generation; layout will just be less optimal
End Sub

      ' Add total row
   Private Sub AddTotalRow(pdfWs As Worksheet, currentRow As Integer, totalQty As Integer)
       
     
     ' Put "Total Qty" in the AVL column (column E)
     pdfWs.Cells(currentRow, 5).Value = "Total Qty"
     pdfWs.Cells(currentRow, 5).Font.Bold = True
     pdfWs.Cells(currentRow, 5).Font.Name = "Calibri"
     pdfWs.Cells(currentRow, 5).Font.Size = 8
     pdfWs.Cells(currentRow, 5).HorizontalAlignment = xlCenter
     
     ' Put total quantity in the QTY column (column F)
     pdfWs.Cells(currentRow, 6).Value = totalQty
     pdfWs.Cells(currentRow, 6).Font.Bold = True
     pdfWs.Cells(currentRow, 6).Font.Name = "Calibri"
     pdfWs.Cells(currentRow, 6).Font.Size = 8
     pdfWs.Cells(currentRow, 6).HorizontalAlignment = xlCenter
     
     ' Add borders to the total row
     With pdfWs.Range("A" & currentRow & ":H" & currentRow)
         .Borders.LineStyle = xlContinuous
         .Borders.Weight = xlHairline
         .VerticalAlignment = xlCenter
     End With
     
    ' Set row height for total row
    pdfWs.Rows(currentRow).RowHeight = 18.15
    
          ' Debug.Print "*** TOTAL ROW ADDED SUCCESSFULLY ***"
End Sub
 
 ' Add continued header (for condition 4)
 Private Sub AddContinuedHeader(pdfWs As Worksheet, currentRow As Integer, groupKey As String, period As String, Optional organizationName As String = "VIMARSH")
    ' Add the sector-village and organization header line with "CONTINUED"
    Dim sectorVillageParts() As String
    sectorVillageParts = Split(groupKey, " - ")
    pdfWs.Cells(currentRow, 1).Value = sectorVillageParts(0) & " - " & sectorVillageParts(1) & " (CONTINUED)"
     pdfWs.Cells(currentRow, 1).Font.Bold = True
     pdfWs.Cells(currentRow, 1).Font.Size = 14
     pdfWs.Cells(currentRow, 1).Font.Name = "Calibri"
     pdfWs.Range("A" & currentRow & ":C" & currentRow).Merge
     pdfWs.Range("A" & currentRow & ":C" & currentRow).HorizontalAlignment = xlLeft
     
     ' Add period header on the same row
     Dim periodParts() As String
     Dim orgName As String
     If organizationName = "" Then
         orgName = "VIMARSH"
     Else
         orgName = organizationName
     End If
     periodParts = Split(period, "-")
     pdfWs.Cells(currentRow, 4).Value = orgName & ": " & GetShortMonthName(CLng(periodParts(1))) & "-" & periodParts(0) & "(" & GetHalfRange(periodParts(2), CLng(periodParts(1))) & ")"
     pdfWs.Cells(currentRow, 4).Font.Bold = True
     pdfWs.Cells(currentRow, 4).Font.Size = 14
     pdfWs.Cells(currentRow, 4).Font.Name = "Calibri"
     pdfWs.Range("D" & currentRow & ":F" & currentRow).Merge
     pdfWs.Range("D" & currentRow & ":F" & currentRow).HorizontalAlignment = xlCenter
     
     ' Add "Details Name" in Date column (H)
     pdfWs.Cells(currentRow, 8).Value = "Details Name"
     pdfWs.Cells(currentRow, 8).Font.Bold = True
     pdfWs.Cells(currentRow, 8).Font.Size = 14
     pdfWs.Cells(currentRow, 8).Font.Name = "Calibri"
     pdfWs.Cells(currentRow, 8).HorizontalAlignment = xlRight
 End Sub
 
 ' Process "Yes" table without page breaks (for Condition 2)
 Private Sub ProcessYesTableSimple(pdfWs As Worksheet, availRecords As Collection, ByRef currentRow As Integer, groupKey As String, period As String)
     On Error GoTo ErrorHandler
     
     ' Add header line just above the table
     Call AddTableHeader(pdfWs, currentRow, groupKey, period)
     currentRow = currentRow + 1
     
     ' Add table headers
     Call AddTableHeaders(pdfWs, currentRow)
     currentRow = currentRow + 1
     
     ' Group records by contact name to keep merged cells together
     Dim contactGroups As Object
     Set contactGroups = CreateObject("Scripting.Dictionary")
     
     Dim i As Integer
     For i = 1 To availRecords.count
         Dim record As Variant
         record = availRecords.item(i)
         Dim contactName As String
         contactName = record(10)  ' Contact name
         
         If Not contactGroups.Exists(contactName) Then
             Dim contactGroup As Collection
             Set contactGroup = New Collection
             contactGroups.Add contactName, contactGroup
         End If
         
         contactGroups(contactName).Add record
     Next i
     
     ' Process each contact group WITHOUT page breaks
     Dim contactKey As Variant
     Dim contactRecords As Collection
     Dim totalQty As Integer
     totalQty = 0
     
     For Each contactKey In contactGroups.Keys
         Set contactRecords = contactGroups(contactKey)
         
         ' Add all records for this contact
         Dim startRow As Integer
         startRow = currentRow
         
         For i = 1 To contactRecords.count
             Dim dataRecord As Variant
             dataRecord = contactRecords.item(i)
             
             pdfWs.Cells(currentRow, 1).Value = ToCamelCase(GetContactNameOnly(CStr(dataRecord(10))))  ' Name only
             pdfWs.Cells(currentRow, 1).Font.Size = 9  ' Smaller font for Name
             pdfWs.Cells(currentRow, 2).Value = GetContactNumber(CStr(dataRecord(10)))  ' Contact No.
             pdfWs.Cells(currentRow, 3).Value = GetLanguageShortCode(CStr(dataRecord(2)))   ' Language
             pdfWs.Cells(currentRow, 4).Value = dataRecord(1)   ' Book Name
             Call FormatAvailabilityCell(pdfWs.Cells(currentRow, 5), CStr(dataRecord(11)))  ' Availability
             pdfWs.Cells(currentRow, 6).Value = dataRecord(6)   ' Quantity
             pdfWs.Cells(currentRow, 7).Value = ""              ' Sign (will be filled after merging)
             
             ' Format data row
             Call FormatDataRow(pdfWs, currentRow)
             
             totalQty = totalQty + CLng(dataRecord(6))
             currentRow = currentRow + 1
         Next i
         
         ' Sort data by Language and Book Title within each contact group
         If currentRow - 1 >= startRow Then
             pdfWs.Range("A" & startRow & ":H" & currentRow - 1).Sort _
                 Key1:=pdfWs.Range("B" & startRow), Order1:=xlAscending, _
                 Key2:=pdfWs.Range("C" & startRow), Order2:=xlAscending, _
                 Header:=xlNo
         End If
         
         ' Apply cell merging for this contact group
         Call ApplyCellMergingForYesTable(pdfWs, startRow, currentRow - 1)
     Next contactKey
     
     ' Add total row
     Call AddTotalRow(pdfWs, currentRow, totalQty)
     currentRow = currentRow + 2  ' Add space after table
     
     Exit Sub
     
ErrorHandler:
 End Sub

 

' Load Group Sector Data for dropdown
Public Function LoadGroupSectorData() As Collection
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    Dim conn As Object
    Dim groupSectors As Collection
    Set groupSectors = New Collection
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL
    If Not ConnectToPostgreSQL(conn) Then
        Set LoadGroupSectorData = groupSectors
        Exit Function
    End If
    
    ' Get unique group_sector values
    sql = "SELECT DISTINCT group_sector FROM " & DB_SCHEMA & ".transactions " & _
          "WHERE is_deleted = FALSE AND group_sector IS NOT NULL AND group_sector != '' " & _
          "ORDER BY group_sector;"
    
    Set rs = conn.Execute(sql)
    Dim tempArr() As String
    Dim cnt As Long
    cnt = 0
    While Not rs.EOF
        cnt = cnt + 1
        ReDim Preserve tempArr(1 To cnt)
        tempArr(cnt) = CStr(rs.Fields("group_sector").Value)
        rs.MoveNext
    Wend
    rs.Close
    If cnt > 0 Then
        NaturalSort.SortKeysNatural tempArr
        Dim k As Long
        For k = 1 To cnt
            groupSectors.Add tempArr(k)
        Next k
    End If
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Set LoadGroupSectorData = groupSectors
    Exit Function
    
ErrorHandler:
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    Set LoadGroupSectorData = New Collection
End Function

' Load Period Data for dropdown
Public Function LoadPeriodData(Optional appValue As String = "") As Collection
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    Dim conn As Object
    Dim periods As Collection
    Set periods = New Collection
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL
    If Not ConnectToPostgreSQL(conn) Then
        Set LoadPeriodData = periods
        Exit Function
    End If
    
    ' Get unique period values (year-month-half format) in descending order (latest first)
    sql = "SELECT DISTINCT year, month, half, " & _
          "CONCAT(year, '-', " & _
          "CASE WHEN month < 10 THEN CONCAT('0', month) ELSE CAST(month AS TEXT) END, '-', " & _
          "CAST(half AS TEXT)) AS period " & _
          "FROM " & DB_SCHEMA & ".transactions " & _
          "WHERE is_deleted = FALSE AND year IS NOT NULL AND month IS NOT NULL "
    
    ' Add APP filter when provided
    If appValue <> "" Then
        sql = sql & "AND app = '" & Replace(appValue, "'", "''") & "' "
    End If
    
    sql = sql & "ORDER BY year DESC, month DESC, half DESC;"
    
    Set rs = conn.Execute(sql)
    
    While Not rs.EOF
        Dim periodValue As String
        periodValue = rs.Fields("period").Value
        periods.Add periodValue
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Set LoadPeriodData = periods
    Exit Function
    
ErrorHandler:
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    Set LoadPeriodData = New Collection
End Function

' Load available books for selection
Public Function LoadAvailableBooks() As Collection
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    Dim conn As Object
    Dim books As Collection
    Set books = New Collection
    
    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        Set LoadAvailableBooks = books
        Exit Function
    End If
    
    sql = "SELECT b.book_id, b.book_name, b.language, " & _
          "COALESCE(SUM(CASE WHEN t.is_deleted = FALSE THEN t.qty ELSE 0 END), 0) AS total_qty " & _
          "FROM " & DB_SCHEMA & ".books b " & _
          "LEFT JOIN " & DB_SCHEMA & ".transactions t ON t.book_id = b.book_id " & _
          "WHERE b.is_deleted = FALSE " & _
          "GROUP BY b.book_id, b.book_name, b.language " & _
          "ORDER BY b.language ASC, b.book_name ASC;"
    
    Set rs = conn.Execute(sql)
    
    While Not rs.EOF
        Dim bookInfo(3) As Variant
        bookInfo(0) = CLng(rs.Fields("book_id").Value)
        bookInfo(1) = rs.Fields("book_name").Value & ""
        bookInfo(2) = rs.Fields("language").Value & ""
        If IsNull(rs.Fields("total_qty").Value) Then
            bookInfo(3) = 0
        Else
            bookInfo(3) = CLng(rs.Fields("total_qty").Value)
        End If
        books.Add bookInfo
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Set LoadAvailableBooks = books
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    Set LoadAvailableBooks = New Collection
End Function

' Build comma-separated numeric list for SQL IN clause
Public Function BuildNumericInClause(selectedIds As Collection) As String
    Dim buffer As String
    Dim item As Variant
    
    If selectedIds Is Nothing Then
        BuildNumericInClause = ""
        Exit Function
    End If
    
    For Each item In selectedIds
        If IsNumeric(item) Then
            If buffer <> "" Then
                buffer = buffer & ","
            End If
            buffer = buffer & CStr(CLng(item))
        End If
    Next item
    
    BuildNumericInClause = buffer
End Function

' Get total quantities for each book based on current filters
Public Function GetBookTotalsForFilters(groupSector As String, period As String, availYes As Boolean, availNotInStock As Boolean, availBYes As Boolean, Optional appValue As String = "") As Object
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim totals As Object
    Dim sql As String
    Dim whereClause As String
    
    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        Set GetBookTotalsForFilters = Nothing
        Exit Function
    End If
    
    whereClause = "WHERE t.is_deleted = FALSE"
    
    If groupSector <> "" And groupSector <> "All" Then
        whereClause = whereClause & " AND t.group_sector = '" & Replace(groupSector, "'", "''") & "'"
    End If
    
    If period <> "" And period <> "All" Then
        Dim periodParts() As String
        periodParts = Split(period, "-")
        If UBound(periodParts) >= 2 Then
            whereClause = whereClause & " AND t.year = " & periodParts(0) & _
                         " AND t.month = " & CLng(periodParts(1)) & _
                         " AND t.half = '" & periodParts(2) & "'"
        End If
    End If
    
    Dim availConditions() As String
    Dim availCount As Integer
    availCount = 0
    
    If availYes Then
        ReDim Preserve availConditions(availCount)
        availConditions(availCount) = "t.avail = 1"
        availCount = availCount + 1
    End If
    
    If availNotInStock Then
        ReDim Preserve availConditions(availCount)
        availConditions(availCount) = "t.avail = 0"
        availCount = availCount + 1
    End If
    
    If availBYes Then
        ReDim Preserve availConditions(availCount)
        availConditions(availCount) = "t.avail = 2"
        availCount = availCount + 1
    End If
    
    If availCount > 0 Then
        whereClause = whereClause & " AND (" & Join(availConditions, " OR ") & ")"
    End If
    
    ' Add APP filter when provided
    If appValue <> "" Then
        whereClause = whereClause & " AND t.app = '" & Replace(appValue, "'", "''") & "'"
    End If
    
    sql = "SELECT t.book_id, SUM(t.qty) AS total_qty " & _
          "FROM " & DB_SCHEMA & ".transactions t " & _
          "INNER JOIN " & DB_SCHEMA & ".books b ON t.book_id = b.book_id " & _
          whereClause & " AND b.is_deleted = FALSE " & _
          "GROUP BY t.book_id;"
    
    Set rs = conn.Execute(sql)
    
    Set totals = CreateObject("Scripting.Dictionary")
    
    While Not rs.EOF
        Dim bookId As Long
        Dim qtyValue As Double
        bookId = CLng(rs.Fields("book_id").Value)
        qtyValue = 0
        If Not IsNull(rs.Fields("total_qty").Value) Then
            qtyValue = CDbl(rs.Fields("total_qty").Value)
        End If
        totals(CStr(bookId)) = qtyValue
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
    Set GetBookTotalsForFilters = totals
    Exit Function
    
ErrorHandler:
    Set totals = Nothing
    Resume Cleanup
End Function

' ============================================================================
' SECTOR SUMMARY REPORT INTEGRATION
' ============================================================================

' Generate Sector Summary Report (for form integration)
Public Sub GenerateSectorSummaryReportFromForm(groupSector As String, period As String, selectedBookIds As Collection, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
    On Error GoTo ErrorHandler
    
    ' Get sector summary data using the function from SectorSummaryReport.bas
    Dim summaryData As Collection
    Dim orgName As String
    If organizationName = "" Then
        orgName = "VIMARSH"
    Else
        orgName = organizationName
    End If
    If selectedBookIds Is Nothing Then
        Set selectedBookIds = New Collection
    End If
    Set summaryData = GetSectorSummaryDataForPDF(period, groupSector, selectedBookIds, orgName)
    
    If summaryData.count = 0 Then
        MsgBox "No data found for the selected criteria." & vbCrLf & _
               "Please check your sector and period selections.", vbInformation, "No Data"
        Exit Sub
    End If
    
    ' Generate PDF using the function from SectorSummaryReport.bas
    GenerateSectorSummaryPDFWrapper summaryData, period, groupSector, customFolder, organizationName
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating sector summary report: " & Err.description, vbCritical
End Sub

' ============================================================================
' ALL SUMMARY REPORT INTEGRATION
' ============================================================================

' Generate All Summary Report (for form integration)
Public Sub GenerateAllSummaryReportFromForm(period As String, selectedBookIds As Collection, Optional customFolder As String = "", Optional organizationName As String = "VIMARSH")
    On Error GoTo ErrorHandler
    
    ' Get all summary data using the function from AllSummaryReport.bas
    Dim summaryData As Collection
    Dim orgName As String
    If organizationName = "" Then
        orgName = "VIMARSH"
    Else
        orgName = organizationName
    End If
    If selectedBookIds Is Nothing Then
        Set selectedBookIds = New Collection
    End If
    Set summaryData = GetAllSummaryDataForPDF(period, selectedBookIds, orgName)
    
    If summaryData.count = 0 Then
        MsgBox "No data found for the selected criteria." & vbCrLf & _
               "Please check your period selection.", vbInformation, "No Data"
        Exit Sub
    End If
    
    ' Generate PDF using the function from AllSummaryReport.bas
    GenerateAllSummaryPDFWrapper summaryData, period, customFolder, organizationName
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating all summary report: " & Err.description, vbCritical
End Sub

' Function to convert language name to 2-char short code
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

' Function to convert availability text to short code
Private Function GetAvailabilityShortCode(ByVal availText As String) As String
    Select Case LCase(Trim(availText))
        Case "not in stock"
            GetAvailabilityShortCode = "NS"
        Case "yes"
            GetAvailabilityShortCode = "Yes"
        Case "b-yes"
            GetAvailabilityShortCode = "B-Yes"
        Case Else
            GetAvailabilityShortCode = availText
    End Select
End Function

' Sub to format availability cell - light color and small font for NS
Private Sub FormatAvailabilityCell(cell As Range, availText As String)
    cell.Value = GetAvailabilityShortCode(availText)
    If LCase(Trim(availText)) = "not in stock" Then
        cell.Font.Size = 8
        cell.Font.Color = RGB(150, 150, 150)  ' Light gray color
    End If
End Sub
