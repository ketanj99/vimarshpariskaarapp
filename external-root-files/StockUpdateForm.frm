Option Explicit

' IMPORTANT: For this form to compile, you MUST create these controls in the VBA Form Designer:
' Steps:
' 1. Open StockUpdateForm in VBA Editor (Alt+F11)
' 2. Right-click on the form -> View Object
' 3. Add these controls from Toolbox:
'    - TextBox: Name = "txtFilterBookName" (for filtering by book name)
'    - TextBox: Name = "txtFilterPeriod" (for filtering by period)
'    - TextBox: Name = "txtFilterPNo" (for filtering by P No. text)
'    - ToggleButton: Name = "tglClear", Caption = "Clear"
'    - ToggleButton: Name = "tglBlank", Caption = "Blank"
'    - ToggleButton: Name = "tglAll", Caption = "All"
'    - ListBox: Name = "lstNotInStockBooks"
'    - TextBox: Name = "txtPNo"  
'    - CommandButton: Name = "btnUpdate", Caption = "Update"
'    - CommandButton: Name = "btnClose", Caption = "Close"
'    - Label: Name = "lblStatus"
'    - CommandButton: Name = "btnSelectAll", Caption = "Select All"
'    - CommandButton: Name = "btnGeneratePDF", Caption = "Generate PDF"
'    - CommandButton: Name = "btnGenerateSectorPDF", Caption = "Generate Sector PDF"
' 4. Controls can be small/hidden - code will resize/reposition them
' 5. Save the form
'
' Form controls are automatically declared by VBA when created in form designer
' Controls needed:
' - lstNotInStockBooks (ListBox) - Name: lstNotInStockBooks
' - txtFilterBookName (TextBox) - Name: txtFilterBookName (for filtering by book name)
' - txtFilterPeriod (TextBox) - Name: txtFilterPeriod (for filtering by period)
' - txtFilterPNo (TextBox) - Name: txtFilterPNo (for filtering by P No. text)
' - tglClear (ToggleButton) - Name: tglClear
' - tglBlank (ToggleButton) - Name: tglBlank
' - tglAll (ToggleButton) - Name: tglAll
' - txtPNo (TextBox) - Name: txtPNo
' - btnUpdate (CommandButton) - Name: btnUpdate
' - btnClose (CommandButton) - Name: btnClose
' - lblStatus (Label) - Name: lblStatus
' - btnSelectAll (CommandButton) - Name: btnSelectAll
' - btnGenerateSectorPDF (CommandButton) - Name: btnGenerateSectorPDF

' Note: When controls exist in form designer, VBA automatically creates variables for them
' So we don't need to declare: lstNotInStockBooks, txtPNo, btnUpdate, btnClose, lblStatus
' They are automatically available as form controls

' Only declare variables that are NOT form controls
Private allNotInStockBooks As Collection
Private bookDataMap As Object
Private allBooksData As Collection  ' Store all books for filtering
Private filterMode As String  ' "Clear", "Blank", or "All"

Private Sub UserForm_Initialize()
    InitializeControls
    LoadNotInStockBooks
End Sub

Private Sub InitializeControls()
    On Error Resume Next
    
    ' Controls already exist in form designer, so VBA automatically creates variables for them
    ' We just need to configure their properties
    
    ' Configure Filter TextBoxes - 3 separate filters
    ' Book Name Filter
    With txtFilterBookName
        .Left = 18
        .Top = 18
        .Width = 220
        .Height = 20
        .Text = ""
    End With
    
    ' Period Filter
    With txtFilterPeriod
        .Left = 250
        .Top = 18
        .Width = 220
        .Height = 20
        .Text = ""
    End With
    
    ' P No. Filter
    With txtFilterPNo
        .Left = 482
        .Top = 18
        .Width = 150
        .Height = 20
        .Text = ""
    End With
    
    ' Configure Toggle Buttons (for P No. status filter)
    With tglClear
        .Left = 642
        .Top = 18
        .Width = 60
        .Height = 20
        .Caption = "Clear"
        .Value = False
    End With
    
    With tglBlank
        .Left = 712
        .Top = 18
        .Width = 60
        .Height = 20
        .Caption = "Blank"
        .Value = True  ' Default to "Blank"
    End With
    
    With tglAll
        .Left = 782
        .Top = 18
        .Width = 50
        .Height = 20
        .Caption = "All"
        .Value = False
    End With
    
    ' Configure Select All Button
    With btnSelectAll
        .Left = 842
        .Top = 18
        .Width = 80
        .Height = 20
        .Caption = "Select All"
        .Visible = True  ' Always visible
    End With
    
    ' Configure ListBox (below filter) - Enable multi-select
    ' Mouse wheel scrolling works automatically when listbox has focus
    With lstNotInStockBooks
        .Left = 18
        .Top = 42
        .Width = 700
        .Height = 380
        .ColumnCount = 6
        .ColumnWidths = "250 pt;100 pt;150 pt;80 pt;100 pt;0 pt"
        .MultiSelect = fmMultiSelectMulti  ' Enable multiple selection
    End With
    
    ' Configure P No TextBox
    With txtPNo
        .Left = 18
        .Top = 430
        .Width = 200
        .Height = 20
        .Text = ""
    End With
    
    ' Configure Update Button
    With btnUpdate
        .Left = 230
        .Top = 428
        .Width = 100
        .Height = 24
        .Caption = "Update Selected"
    End With
    
    ' Configure Generate PDF Button
    With btnGeneratePDF
        .Left = 340
        .Top = 428
        .Width = 100
        .Height = 24
        .Caption = "Generate PDF"
    End With
    
    ' Configure Generate Sector PDF Button - REMOVED (now in separate module)
    ' Sector-wise PDF generation moved to SectorWiseStockPDF.bas module
    ' Use GenerateSectorWiseStockPDF() function from that module
    
    ' Configure Close Button
    With btnClose
        .Left = 580
        .Top = 428
        .Width = 80
        .Height = 24
        .Caption = "Close"
    End With
    
    ' Configure Status Label
    With lblStatus
        .Left = 18
        .Top = 460
        .Width = 600
        .Height = 20
        .Caption = "Loading..."
    End With
    
    ' Initialize filter mode
    filterMode = "Blank"
    
    On Error GoTo 0
End Sub

' Filter function to filter books by Book Name, Period, and P No. separately
Private Sub ApplyFilter()
    On Error Resume Next
    
    Dim filterBookName As String
    Dim filterPeriod As String
    Dim filterPNo As String
    
    filterBookName = Trim(txtFilterBookName.Text)
    filterBookName = LCase(filterBookName)
    
    filterPeriod = Trim(txtFilterPeriod.Text)
    filterPeriod = LCase(filterPeriod)
    
    filterPNo = Trim(txtFilterPNo.Text)
    filterPNo = LCase(filterPNo)
    
    ' Clear listbox
    lstNotInStockBooks.Clear
    
    ' Add header row
    lstNotInStockBooks.AddItem "Book Name"
    lstNotInStockBooks.List(0, 0) = "Book Name"
    lstNotInStockBooks.List(0, 1) = "Language"
    lstNotInStockBooks.List(0, 2) = "Period"
    lstNotInStockBooks.List(0, 3) = "Qty"
    lstNotInStockBooks.List(0, 4) = "P No."
    lstNotInStockBooks.List(0, 5) = ""
    
    ' Filter and add matching books
    Dim i As Long
    Dim matchCount As Long
    matchCount = 0
    
    For i = 1 To allBooksData.Count
        Dim bookData As Variant
        bookData = allBooksData.Item(i)
        
        Dim bookName As String
        Dim language As String
        Dim periodDisplay As String
        Dim qty As Long
        Dim stockNumber As String
        Dim importId As Long
        Dim displayName As String
        
        bookName = CStr(bookData(0))
        language = CStr(bookData(1))
        periodDisplay = CStr(bookData(2))
        qty = CLng(bookData(3))
        stockNumber = CStr(bookData(4))
        importId = CLng(bookData(5))
        displayName = CStr(bookData(6))
        
        ' Check if matches Book Name filter
        Dim bookNameMatches As Boolean
        bookNameMatches = False
        If filterBookName = "" Then
            bookNameMatches = True
        Else
            If InStr(LCase(bookName), filterBookName) > 0 Or _
               InStr(LCase(displayName), filterBookName) > 0 Then
                bookNameMatches = True
            End If
        End If
        
        ' Check if matches Period filter
        Dim periodMatches As Boolean
        periodMatches = False
        If filterPeriod = "" Then
            periodMatches = True
        Else
            If InStr(LCase(periodDisplay), filterPeriod) > 0 Then
                periodMatches = True
            End If
        End If
        
        ' Check if matches P No. filter (both text filter and status filter)
        Dim pNoMatches As Boolean
        pNoMatches = False
        Dim stockNumberTrimmed As String
        stockNumberTrimmed = Trim(stockNumber)
        
        ' First check toggle button filter mode (Clear/Blank/All)
        Dim pNoStatusMatches As Boolean
        pNoStatusMatches = False
        
        If filterMode = "Clear" Then
            ' Show only items with P No. (not blank)
            pNoStatusMatches = (stockNumberTrimmed <> "")
        ElseIf filterMode = "Blank" Then
            ' Show only items without P No. (blank)
            pNoStatusMatches = (stockNumberTrimmed = "")
        Else ' filterMode = "All"
            ' Show all items
            pNoStatusMatches = True
        End If
        
        ' Then check text filter (if provided)
        Dim pNoTextMatches As Boolean
        pNoTextMatches = False
        
        If filterPNo = "" Then
            ' No text filter, so it matches
            pNoTextMatches = True
        Else
            ' Text filter provided, check if P No. contains the text
            If InStr(LCase(stockNumberTrimmed), filterPNo) > 0 Then
                pNoTextMatches = True
            End If
        End If
        
        ' Both P No. filters must match
        pNoMatches = pNoStatusMatches And pNoTextMatches
        
        ' All filters must match (Book Name, Period, and P No.)
        If bookNameMatches And periodMatches And pNoMatches Then
            lstNotInStockBooks.AddItem displayName
            lstNotInStockBooks.List(lstNotInStockBooks.ListCount - 1, 0) = bookName
            lstNotInStockBooks.List(lstNotInStockBooks.ListCount - 1, 1) = language
            lstNotInStockBooks.List(lstNotInStockBooks.ListCount - 1, 2) = periodDisplay
            lstNotInStockBooks.List(lstNotInStockBooks.ListCount - 1, 3) = qty
            lstNotInStockBooks.List(lstNotInStockBooks.ListCount - 1, 4) = stockNumber
            lstNotInStockBooks.List(lstNotInStockBooks.ListCount - 1, 5) = importId
            matchCount = matchCount + 1
        End If
    Next i
    
    ' Keep Select All button visible (always available)
    btnSelectAll.Visible = True
    
    ' Clear P No. field and selection when filter changes
    txtPNo.Text = ""
    
    ' Update status
    Dim statusText As String
    statusText = "Showing " & matchCount & " of " & allBooksData.Count & " books"
    If filterBookName <> "" Or filterPeriod <> "" Or filterPNo <> "" Or filterMode <> "All" Then
        statusText = statusText & " (Filtered)"
    End If
    lblStatus.Caption = statusText
    
    On Error GoTo 0
End Sub

' Handle Filter TextBox Change events
Private Sub txtFilterBookName_Change()
    ApplyFilter
End Sub

Private Sub txtFilterPeriod_Change()
    ApplyFilter
End Sub

Private Sub txtFilterPNo_Change()
    ApplyFilter
End Sub

' Handle Toggle Button events - Radio button behavior (only one active at a time)
Private Sub tglClear_Click()
    ' Always activate this toggle and turn off others
    tglClear.Value = True
    tglBlank.Value = False
    tglAll.Value = False
    filterMode = "Clear"
    ApplyFilter
End Sub

Private Sub tglBlank_Click()
    ' Always activate this toggle and turn off others
    tglBlank.Value = True
    tglClear.Value = False
    tglAll.Value = False
    filterMode = "Blank"
    ApplyFilter
End Sub

Private Sub tglAll_Click()
    ' Always activate this toggle and turn off others
    tglAll.Value = True
    tglClear.Value = False
    tglBlank.Value = False
    filterMode = "All"
    ApplyFilter
End Sub

' Handle Select All button click
Private Sub btnSelectAll_Click()
    On Error Resume Next
    Dim i As Long
    ' Select all items except header (index 0)
    For i = 1 To lstNotInStockBooks.ListCount - 1
        lstNotInStockBooks.Selected(i) = True
    Next i
    On Error GoTo 0
End Sub

' Handle ListBox click event (called when lstNotInStockBooks is clicked)
' This will work if control name matches
Private Sub lstNotInStockBooks_Click()
    On Error Resume Next
    ' Skip header row (index 0)
    ' If single item selected, fill P No. field
    If lstNotInStockBooks.ListIndex > 0 Then
        Dim selectedCount As Long
        selectedCount = 0
        Dim i As Long
        For i = 1 To lstNotInStockBooks.ListCount - 1
            If lstNotInStockBooks.Selected(i) Then
                selectedCount = selectedCount + 1
            End If
        Next i
        
        ' Update status to show selection info
        If selectedCount = 1 Then
            Dim currentStockNumber As String
            currentStockNumber = lstNotInStockBooks.List(lstNotInStockBooks.ListIndex, 4) & ""  ' P No. is now column 4
            
            ' Show current P No. in txtPNo field for editing
            txtPNo.Text = currentStockNumber
            
            If Trim(currentStockNumber) = "" Then
                lblStatus.Caption = "1 book selected. Enter P No. to update or leave blank."
            Else
                lblStatus.Caption = "1 book selected. Current P No: " & currentStockNumber & ". Edit or clear to update."
            End If
        ElseIf selectedCount > 1 Then
            lblStatus.Caption = selectedCount & " books selected. Enter P No. to update or leave blank to clear."
            txtPNo.Text = ""  ' Clear when multiple selected
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub LoadNotInStockBooks()
    On Error GoTo ErrorHandler
    
    ' Clear the listbox
    lstNotInStockBooks.Clear
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"
    
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        lblStatus.Caption = "Connection failed: " & Err.Description
        lblStatus.ForeColor = RGB(192, 0, 0)
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Get all "Not In Stock" books with period information
    ' Group by book_id and period only (no sector grouping in form)
    ' Each book_id appears once per period, regardless of sector
    Dim sql As String
    sql = "SELECT " & _
          "MIN(t.importid) as importid, " & _
          "b.book_id, " & _
          "b.book_name, " & _
          "b.language, " & _
          "t.year, " & _
          "t.month, " & _
          "t.half, " & _
          "MAX(t.group_sector) as group_sector, " & _
          "SUM(t.qty) as total_qty, " & _
          "MAX(COALESCE(t.stock_number, '')) as stock_number, " & _
          "MAX(COALESCE(t.stock_date::text, '')) as stock_date " & _
          "FROM vimars.transactions t " & _
          "INNER JOIN vimars.books b ON t.book_id = b.book_id " & _
          "WHERE t.is_deleted = FALSE AND b.is_deleted = FALSE " & _
          "AND t.avail = 0 " & _
          "GROUP BY b.book_id, b.book_name, b.language, t.year, t.month, t.half " & _
          "ORDER BY b.language ASC, b.book_name ASC, t.year DESC, t.month DESC, t.half DESC;"
    
    Dim rs As Object
    Set rs = conn.Execute(sql)
    
    Set allNotInStockBooks = New Collection
    Set bookDataMap = CreateObject("Scripting.Dictionary")
    Set allBooksData = New Collection
    
    While Not rs.EOF
        Dim bookName As String
        Dim period As String
        Dim periodDisplay As String
        Dim qty As Long
        Dim importId As Long
        
        bookName = rs.Fields("book_name").Value & ""
        Dim language As String
        language = rs.Fields("language").Value & ""
        Dim year As Integer
        Dim monthNum As Integer
        Dim half As String
        year = CInt(rs.Fields("year").Value)
        monthNum = CInt(rs.Fields("month").Value)
        half = CStr(rs.Fields("half").Value)
        
        ' Format period: "2025-09-2" -> "2025-Sep (1 To 15)" or "2025-Oct (16 To 30)"
        period = year & "-" & Right("0" & monthNum, 2) & "-" & half
        
        ' Calculate period display with month name
        Dim monthNameStr As String
        monthNameStr = Left(MonthName(monthNum), 3)  ' Get first 3 letters: Sep, Oct, etc.
        
        ' Handle "1 & 2" for full month - normalize by removing spaces
        Dim halfNormalized As String
        halfNormalized = Replace(Trim(LCase(half)), " ", "")  ' Remove all spaces
        
        If halfNormalized = "1&2" Then
            ' Full month range
            Dim lastDay As Integer
            lastDay = Day(DateSerial(year, monthNum + 1, 0))
            periodDisplay = year & "-" & monthNameStr & " (1 To " & lastDay & ")"
        ElseIf Trim(half) = "1" Then
            periodDisplay = year & "-" & monthNameStr & " (1 To 15)"
        ElseIf Trim(half) = "2" Then
            ' Get last day of month
            Dim lastDay2 As Integer
            lastDay2 = Day(DateSerial(year, monthNum + 1, 0))
            periodDisplay = year & "-" & monthNameStr & " (16 To " & lastDay2 & ")"
        Else
            periodDisplay = year & "-" & monthNameStr & " (" & period & ")"
        End If
        
        qty = CLng(rs.Fields("total_qty").Value)
        importId = CLng(rs.Fields("importid").Value)
        
        Dim stockNumber As String
        Dim stockDate As String
        Dim groupSector As String
        stockNumber = rs.Fields("stock_number").Value & ""
        stockDate = rs.Fields("stock_date").Value & ""
        groupSector = rs.Fields("group_sector").Value & ""
        
        ' Add to list (columns: Book Name, Language, Period, Qty, P No., importId)
        ' Display format: Book Name - Language
        Dim displayName As String
        If language <> "" Then
            displayName = bookName & " - " & language
        Else
            displayName = bookName
        End If
        
        ' Store data for filtering (now includes sector)
        Dim bookData(7) As Variant
        bookData(0) = bookName
        bookData(1) = language
        bookData(2) = periodDisplay
        bookData(3) = qty
        bookData(4) = stockNumber
        bookData(5) = importId
        bookData(6) = displayName
        bookData(7) = groupSector
        allBooksData.Add bookData
        
        ' Store data for update (using book+period as key since we group by them)
        Dim bookKey As String
        bookKey = bookName & "|" & period
        Dim bookInfo(7) As Variant
        bookInfo(0) = importId
        bookInfo(1) = bookName
        bookInfo(2) = language
        bookInfo(3) = period
        bookInfo(4) = qty
        bookInfo(5) = year
        bookInfo(6) = monthNum
        bookInfo(7) = half
        If Not bookDataMap.Exists(bookKey) Then
            bookDataMap.Add bookKey, bookInfo
        End If
        
        allNotInStockBooks.Add bookInfo
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    ' Apply filter to display books
    ApplyFilter
    
    lblStatus.Caption = "Loaded " & allBooksData.Count & " Not In Stock books"
    lblStatus.ForeColor = RGB(0, 176, 80)
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error loading books: " & Err.Description
    lblStatus.ForeColor = RGB(192, 0, 0)
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

Private Sub btnClearPNo_Click()
    ' Clear P No. textbox (if button is added to form)
    txtPNo.Text = ""
    txtPNo.SetFocus
    lblStatus.Caption = "P No. cleared. Leave blank and click Update to remove P No. from database."
    lblStatus.ForeColor = RGB(0, 112, 192)
End Sub

Private Sub txtPNo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Press Escape to clear P No. textbox quickly
    If KeyCode = 27 Then  ' Escape key
        txtPNo.Text = ""
        lblStatus.Caption = "P No. cleared. Leave blank and click Update to remove P No. from database."
        lblStatus.ForeColor = RGB(0, 112, 192)
    End If
End Sub

Private Sub btnUpdate_Click()
    On Error GoTo ErrorHandler
    
    ' Get P No input
    Dim pNo As String
    pNo = Trim(txtPNo.Text)
    
    Debug.Print "=== btnUpdate_Click Started ==="
    Debug.Print "P No. input: '" & pNo & "' (isClearing: " & (pNo = "") & ")"
    
    ' Get all selected items (skip header row index 0)
    Dim selectedIndices As Collection
    Set selectedIndices = New Collection
    Dim i As Long
    For i = 1 To lstNotInStockBooks.ListCount - 1
        If lstNotInStockBooks.Selected(i) Then
            selectedIndices.Add i
        End If
    Next i
    
    ' Check if any item is selected
    If selectedIndices.Count = 0 Then
        MsgBox "Please select at least one book from the list.", vbExclamation
        Exit Sub
    End If
    
    ' Determine if clearing P No. (blank input)
    Dim isClearing As Boolean
    isClearing = (pNo = "")
    
    ' Confirm update
    Dim confirmMsg As String
    If isClearing Then
        If selectedIndices.Count > 1 Then
            confirmMsg = "Clear P No. for " & selectedIndices.Count & " selected books?" & vbCrLf & vbCrLf & _
                        "P No. will be removed from database."
        Else
            confirmMsg = "Clear P No. for selected book?" & vbCrLf & vbCrLf & _
                        "P No. will be removed from database."
        End If
    Else
        If selectedIndices.Count > 1 Then
            confirmMsg = "Update P No. to '" & pNo & "' for " & selectedIndices.Count & " selected books?"
        Else
            confirmMsg = "Update P No. to '" & pNo & "' for selected book?"
        End If
    End If
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Update") <> vbYes Then
        Exit Sub
    End If
    
    ' Update database
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"
    
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        MsgBox "Connection failed: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Dim todayDate As String
    todayDate = Format(Date, "yyyy-mm-dd")
    
    Dim updatedCount As Long
    updatedCount = 0
    Dim failedCount As Long
    failedCount = 0
    
    ' Process each selected item
    Dim j As Long
    For j = 1 To selectedIndices.Count
        Dim selectedIndex As Long
        selectedIndex = selectedIndices.Item(j)
        
        ' Get book info from selected row
        Dim bookName As String
        Dim language As String
        Dim periodDisplay As String
        bookName = lstNotInStockBooks.List(selectedIndex, 0)
        language = lstNotInStockBooks.List(selectedIndex, 1)  ' Get language from listbox
        periodDisplay = lstNotInStockBooks.List(selectedIndex, 2)
        
        ' Find matching book info
        Dim bookInfo As Variant
        Dim foundInfo As Boolean
        foundInfo = False
        Dim k As Long
        For k = 1 To allNotInStockBooks.Count
            bookInfo = allNotInStockBooks.Item(k)
            
            Dim bookNameFromInfo As String
            Dim languageFromInfo As String
            Dim infoYear As Integer
            Dim infoMonth As Integer
            Dim infoHalf As String
            
            bookNameFromInfo = CStr(bookInfo(1))
            languageFromInfo = CStr(bookInfo(2))  ' Get language from collection
            infoYear = CInt(bookInfo(5))
            infoMonth = CInt(bookInfo(6))
            infoHalf = CStr(bookInfo(7))
            
            ' Match both book name AND language
            If bookNameFromInfo = bookName And languageFromInfo = language Then
                ' Check if period matches (new format: "2025-Sep (1 To 15)")
                Dim infoPeriodDisplay As String
                Dim infoMonthNameStr As String
                
                infoMonthNameStr = Left(MonthName(infoMonth), 3)
                
                ' Handle "1 & 2" for full month - normalize by removing spaces
                Dim infoHalfNormalized As String
                infoHalfNormalized = Replace(Trim(LCase(infoHalf)), " ", "")  ' Remove all spaces
                
                If infoHalfNormalized = "1&2" Then
                    ' Full month range
                    Dim lastDayFull As Integer
                    lastDayFull = Day(DateSerial(infoYear, infoMonth + 1, 0))
                    infoPeriodDisplay = infoYear & "-" & infoMonthNameStr & " (1 To " & lastDayFull & ")"
                ElseIf Trim(infoHalf) = "1" Then
                    infoPeriodDisplay = infoYear & "-" & infoMonthNameStr & " (1 To 15)"
                ElseIf Trim(infoHalf) = "2" Then
                    Dim lastDay As Integer
                    lastDay = Day(DateSerial(infoYear, infoMonth + 1, 0))
                    infoPeriodDisplay = infoYear & "-" & infoMonthNameStr & " (16 To " & lastDay & ")"
                End If
                
                Debug.Print "Comparing periods: '" & infoPeriodDisplay & "' vs '" & periodDisplay & "'"
                Debug.Print "infoHalf = '" & infoHalf & "', infoHalfNormalized = '" & infoHalfNormalized & "'"
                
                If infoPeriodDisplay = periodDisplay Then
                    Debug.Print "Period matched! Executing update for: " & bookName & " - " & language
                    ' Build SQL based on whether clearing or updating
                    Dim sql As String
                    If isClearing Then
                        ' Clear stock_number and stock_date
                        sql = "UPDATE vimars.transactions t " & _
                              "SET stock_number = NULL, " & _
                              "stock_date = NULL, " & _
                              "updated_date = CURRENT_TIMESTAMP " & _
                              "FROM vimars.books b " & _
                              "WHERE t.book_id = b.book_id " & _
                              "AND b.book_name = '" & Replace(bookName, "'", "''") & "' " & _
                              "AND b.language = '" & Replace(language, "'", "''") & "' " & _
                              "AND t.year = " & infoYear & " " & _
                              "AND t.month = " & infoMonth & " " & _
                              "AND t.half = '" & Replace(infoHalf, "'", "''") & "' " & _
                              "AND t.avail = 0 " & _
                              "AND t.is_deleted = FALSE " & _
                              "AND b.is_deleted = FALSE;"
                    Else
                        ' Update stock_number and stock_date
                        sql = "UPDATE vimars.transactions t " & _
                              "SET stock_number = '" & Replace(pNo, "'", "''") & "', " & _
                              "stock_date = '" & todayDate & "', " & _
                              "updated_date = CURRENT_TIMESTAMP " & _
                              "FROM vimars.books b " & _
                              "WHERE t.book_id = b.book_id " & _
                              "AND b.book_name = '" & Replace(bookName, "'", "''") & "' " & _
                              "AND b.language = '" & Replace(language, "'", "''") & "' " & _
                              "AND t.year = " & infoYear & " " & _
                              "AND t.month = " & infoMonth & " " & _
                              "AND t.half = '" & Replace(infoHalf, "'", "''") & "' " & _
                              "AND t.avail = 0 " & _
                              "AND t.is_deleted = FALSE " & _
                              "AND b.is_deleted = FALSE;"
                    End If
                    
                    Debug.Print "Executing SQL: " & sql
                    conn.Execute sql
                    updatedCount = updatedCount + 1
                    foundInfo = True
                    Debug.Print "Update successful for: " & bookName
                    Exit For
                Else
                    Debug.Print "Period did NOT match"
                End If
            End If
        Next k
        
        If Not foundInfo Then
            failedCount = failedCount + 1
            Debug.Print "Failed to find match for: " & bookName & " - " & language & " - " & periodDisplay
        End If
    Next j
    
    Debug.Print "=== Update Complete: Success=" & updatedCount & ", Failed=" & failedCount & " ==="
    
    conn.Close
    Set conn = Nothing
    
    ' Refresh list
    LoadNotInStockBooks
    
    ' Clear P No field after successful update
    txtPNo.Text = ""
    
    ' Show success message
    Dim successMsg As String
    If updatedCount > 0 Then
        If isClearing Then
            successMsg = "P No. cleared successfully!" & vbCrLf & _
                         "Updated: " & updatedCount & " book(s)"
            lblStatus.Caption = "Cleared P No. for " & updatedCount & " book(s)"
        Else
            successMsg = "P No. updated successfully!" & vbCrLf & _
                         "P No: " & pNo & vbCrLf & _
                         "Updated: " & updatedCount & " book(s)"
            lblStatus.Caption = "Updated " & updatedCount & " book(s) with P No: " & pNo
        End If
        
        If failedCount > 0 Then
            successMsg = successMsg & vbCrLf & "Failed: " & failedCount & " book(s)"
        End If
        lblStatus.ForeColor = RGB(0, 176, 80)
        MsgBox successMsg, vbInformation
    Else
        MsgBox "No books were updated. Please check your selection.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error updating stock: " & Err.Description
    lblStatus.ForeColor = RGB(192, 0, 0)
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

Private Sub btnGeneratePDF_Click()
    On Error GoTo ErrorHandler
    
    ' Check if there are any items to generate PDF
    If lstNotInStockBooks.ListCount <= 1 Then
        MsgBox "No data available to generate PDF.", vbExclamation
        Exit Sub
    End If
    
    ' Get save file path
    Dim savePath As String
    Dim defaultFileName As String
    defaultFileName = "StockReport_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
    
    ' Use Excel to create PDF
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    ' Add header row with "Pending : Surat City" on left and date/time on right
    xlWorksheet.Cells(1, 1).Value = "Pending : Surat City"
    xlWorksheet.Cells(1, 1).Font.Bold = True
    xlWorksheet.Cells(1, 1).HorizontalAlignment = -4131  ' xlLeft
    
    ' Format date/time: "Dt : 18/11/2025 10:53 AM"
    Dim dateTimeStr As String
    dateTimeStr = "Dt : " & Format(Now, "dd/mm/yyyy hh:mm AM/PM")
    xlWorksheet.Cells(1, 5).Value = dateTimeStr
    xlWorksheet.Cells(1, 5).Font.Bold = True
    xlWorksheet.Cells(1, 5).HorizontalAlignment = -4152  ' xlRight
    
    ' Merge cells for header row if needed (optional - for better appearance)
    ' xlWorksheet.Range("A1:E1").Merge
    
    ' Set column headers (5 columns) - moved to row 2
    xlWorksheet.Cells(2, 1).Value = "Book Name"
    xlWorksheet.Cells(2, 2).Value = "Language"
    xlWorksheet.Cells(2, 3).Value = "Period"
    xlWorksheet.Cells(2, 4).Value = "Qty"
    xlWorksheet.Cells(2, 5).Value = "P No."
    
    ' Format header row
    With xlWorksheet.Range("A2:E2")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .VerticalAlignment = -4107  ' xlCenter - Middle align text in cells
    End With
    
    ' Set header alignments: Language and Qty - Center, others - Left
    xlWorksheet.Cells(2, 1).HorizontalAlignment = -4131  ' xlLeft
    xlWorksheet.Cells(2, 2).HorizontalAlignment = -4108  ' xlCenter
    xlWorksheet.Cells(2, 3).HorizontalAlignment = -4131  ' xlLeft
    xlWorksheet.Cells(2, 4).HorizontalAlignment = -4108  ' xlCenter
    xlWorksheet.Cells(2, 5).HorizontalAlignment = -4131  ' xlLeft
    
    ' Collect data from listbox (skip header row)
    Dim rowNum As Long
    rowNum = 3  ' Start from row 3 (row 1 is header, row 2 is table header)
    Dim i As Long
    Dim currentBookName As String
    currentBookName = ""
    Dim rowsPerPage As Long
    rowsPerPage = 50  ' Approximate rows per A4 page (with 21 pixels row height and 0 margins, same as SectorSummaryReport.bas)
    Dim firstRowOfCurrentBook As Long
    firstRowOfCurrentBook = 0
    Dim bookEntryCount As Long
    bookEntryCount = 0
    Dim bookTotalQty As Long
    bookTotalQty = 0
    Dim grandTotalQty As Long  ' Track total of all quantities
    grandTotalQty = 0
    
    For i = 1 To lstNotInStockBooks.ListCount - 1
        Dim bookName As String
        Dim language As String
        Dim period As String
        Dim qty As String
        Dim pNo As String
        Dim nextBookName As String
        
        bookName = lstNotInStockBooks.List(i, 0) & ""
        language = lstNotInStockBooks.List(i, 1) & ""
        period = lstNotInStockBooks.List(i, 2) & ""
        qty = lstNotInStockBooks.List(i, 3) & ""
        pNo = lstNotInStockBooks.List(i, 4) & ""
        
        ' Check next item's book name (if exists)
        If i < lstNotInStockBooks.ListCount - 1 Then
            nextBookName = lstNotInStockBooks.List(i + 1, 0) & ""
        Else
            nextBookName = ""  ' Last item
        End If
        
        ' Check if this is a new book
        If bookName <> currentBookName Then
            ' Calculate how many rows this book will take (count entries + 1 total row)
            Dim bookRowCount As Long
            bookRowCount = 0
            Dim k As Long
            For k = i To lstNotInStockBooks.ListCount - 1
                If lstNotInStockBooks.List(k, 0) & "" = bookName Then
                    bookRowCount = bookRowCount + 1
                Else
                    Exit For
                End If
            Next k
            bookRowCount = bookRowCount + 1  ' Add 1 total row
            
            ' Check if book would split across pages
            ' Account for header row (row 1) and table header (row 2) on first page
            ' On subsequent pages, we add header row (row 1) and table header (row 2)
            Dim rowsFromPageStart As Long
            Dim firstPageHeaderRows As Long
            firstPageHeaderRows = 2  ' Row 1 (header) + Row 2 (table header)
            rowsFromPageStart = ((rowNum - firstPageHeaderRows - 1) Mod rowsPerPage) + 1  ' Rows from start of current page (excluding header rows)
            
            ' If book would split, add page break and headers
            If rowsFromPageStart + bookRowCount > rowsPerPage And rowsFromPageStart > 1 Then
                ' Add page break before this book
                xlWorksheet.HPageBreaks.Add xlWorksheet.Cells(rowNum, 1)
                
                ' Add header row on new page (row 1)
                xlWorksheet.Cells(rowNum, 1).Value = "Pending : Surat City"
                xlWorksheet.Cells(rowNum, 1).Font.Bold = True
                xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
                
                dateTimeStr = "Dt : " & Format(Now, "dd/mm/yyyy hh:mm AM/PM")
                xlWorksheet.Cells(rowNum, 5).Value = dateTimeStr
                xlWorksheet.Cells(rowNum, 5).Font.Bold = True
                xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4152  ' xlRight
                
                rowNum = rowNum + 1
                
                ' Add table header row on new page (row 2)
                xlWorksheet.Cells(rowNum, 1).Value = "Book Name"
                xlWorksheet.Cells(rowNum, 2).Value = "Language"
                xlWorksheet.Cells(rowNum, 3).Value = "Period"
                xlWorksheet.Cells(rowNum, 4).Value = "Qty"
                xlWorksheet.Cells(rowNum, 5).Value = "P No."
                
                ' Format header row
                With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 5))
                    .Font.Bold = True
                    .Interior.Color = RGB(200, 200, 200)
                    .Borders.LineStyle = 1  ' xlContinuous
                    .Borders.Weight = 2  ' xlMedium
                    .VerticalAlignment = -4107  ' xlCenter - Middle align text in cells
                End With
                
                ' Set header alignments
                xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
                xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4108  ' xlCenter
                xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4131  ' xlLeft
                xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4108  ' xlCenter
                xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4131  ' xlLeft
                
                rowNum = rowNum + 1
            End If
            
            ' Reset for new book
            currentBookName = bookName
            firstRowOfCurrentBook = rowNum
            bookEntryCount = 0
            bookTotalQty = 0
        End If
        
        bookEntryCount = bookEntryCount + 1
        ' Add to book total Qty and grand total Qty
        If IsNumeric(qty) Then
            Dim qtyValue As Long
            qtyValue = CLng(qty)
            bookTotalQty = bookTotalQty + qtyValue
            grandTotalQty = grandTotalQty + qtyValue  ' Add to grand total
        End If
        
        ' Write data to worksheet
        xlWorksheet.Cells(rowNum, 1).Value = bookName
        xlWorksheet.Cells(rowNum, 2).Value = language
        xlWorksheet.Cells(rowNum, 3).Value = period
        xlWorksheet.Cells(rowNum, 4).Value = qty
        xlWorksheet.Cells(rowNum, 5).Value = pNo
        
        ' Set alignments: Language and Qty - Center, others - Left
        xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
        xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4108  ' xlCenter
        xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4131  ' xlLeft
        xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4108  ' xlCenter
        xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4131  ' xlLeft
        
        ' Make Qty column bold
        xlWorksheet.Cells(rowNum, 4).Font.Bold = True
        
        ' Apply borders and vertical alignment to this row
        With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 5))
            .Borders.LineStyle = 1  ' xlContinuous
            .Borders.Weight = 1  ' xlThin
            .VerticalAlignment = -4107  ' xlCenter - Middle align text in cells
        End With
        
        rowNum = rowNum + 1
        
        ' Check if next item has different book name or this is the last item
        ' If yes, add total Qty row after current book's entries
        If nextBookName <> bookName Then
            ' Add total Qty row for current book
            xlWorksheet.Cells(rowNum, 1).Value = ""
            xlWorksheet.Cells(rowNum, 2).Value = ""
            xlWorksheet.Cells(rowNum, 3).Value = "Total : " & bookTotalQty  ' Total : total qty in Period column
            xlWorksheet.Cells(rowNum, 4).Value = ""  ' Qty column blank
            xlWorksheet.Cells(rowNum, 5).Value = ""
            
            ' Format total row - only "Total : total qty" text bold
            With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 5))
                .Font.Bold = False  ' Default: not bold
                .Interior.Color = RGB(240, 240, 240)
                .Borders.LineStyle = 1  ' xlContinuous
                .Borders.Weight = 1  ' xlThin
                .VerticalAlignment = -4107  ' xlCenter - Middle align text in cells
            End With
            
            ' Make only "Total : total qty" text bold (Period column)
            xlWorksheet.Cells(rowNum, 3).Font.Bold = True
            
            ' Set alignments for total row
            xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
            xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4108  ' xlCenter
            xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4131  ' xlLeft
            xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4108  ' xlCenter
            xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4131  ' xlLeft
            
            rowNum = rowNum + 1
        End If
    Next i
    
    ' Add final grand total row at the end
    xlWorksheet.Cells(rowNum, 1).Value = ""
    xlWorksheet.Cells(rowNum, 2).Value = ""
    xlWorksheet.Cells(rowNum, 3).Value = "Total All Qty : " & grandTotalQty  ' Total All Qty in Period column
    xlWorksheet.Cells(rowNum, 4).Value = ""  ' Qty column blank
    xlWorksheet.Cells(rowNum, 5).Value = ""
    
    ' Format grand total row - make it stand out
    With xlWorksheet.Range(xlWorksheet.Cells(rowNum, 1), xlWorksheet.Cells(rowNum, 5))
        .Font.Bold = False  ' Default: not bold
        .Interior.Color = RGB(220, 220, 220)  ' Slightly darker background
        .Borders.LineStyle = 1  ' xlContinuous
        .Borders.Weight = 2  ' xlMedium - thicker border
        .VerticalAlignment = -4107  ' xlCenter - Middle align text in cells
    End With
    
    ' Make only "Total All Qty : total" text bold (Period column)
    xlWorksheet.Cells(rowNum, 3).Font.Bold = True
    xlWorksheet.Cells(rowNum, 3).Font.Size = 11  ' Slightly larger font (default is usually 10)
    
    ' Set alignments for grand total row
    xlWorksheet.Cells(rowNum, 1).HorizontalAlignment = -4131  ' xlLeft
    xlWorksheet.Cells(rowNum, 2).HorizontalAlignment = -4108  ' xlCenter
    xlWorksheet.Cells(rowNum, 3).HorizontalAlignment = -4131  ' xlLeft
    xlWorksheet.Cells(rowNum, 4).HorizontalAlignment = -4108  ' xlCenter
    xlWorksheet.Cells(rowNum, 5).HorizontalAlignment = -4131  ' xlLeft
    
    rowNum = rowNum + 1
    
    ' Apply borders to table header row (row 2)
    With xlWorksheet.Range("A2:E2")
        .Borders.LineStyle = 1  ' xlContinuous
        .Borders.Weight = 2  ' xlMedium
    End With
    
    ' Set column widths (P No. column wider for future entries)
    xlWorksheet.Columns("A").ColumnWidth = 30   ' Book Name
    xlWorksheet.Columns("B").ColumnWidth = 15   ' Language
    xlWorksheet.Columns("C").ColumnWidth = 25    ' Period
    xlWorksheet.Columns("D").ColumnWidth = 10   ' Qty
    xlWorksheet.Columns("E").ColumnWidth = 40   ' P No. - Wider for future entries
    
    ' Set row heights to match SectorSummaryReport.bas (21 pixels)
    Dim lastDataRow As Long
    lastDataRow = rowNum - 1
    If lastDataRow > 1 Then
        ' Set all rows to 21 pixels height (same as SectorSummaryReport.bas)
        xlWorksheet.Rows("1:" & lastDataRow).RowHeight = 21
        ' Header row (row 1) can have slightly more height if needed
        ' xlWorksheet.Rows("1:1").RowHeight = 25  ' Optional: make header row taller
    End If
    
    ' Set page format to A4 Portrait with minimal margins to maximize content
    With xlWorksheet.PageSetup
        .PaperSize = 9  ' xlPaperA4
        .Orientation = 1  ' xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .TopMargin = 0    ' No top margin
        .BottomMargin = 0 ' No bottom margin
        .LeftMargin = 36  ' 0.5 inches = 36 points
        .RightMargin = 36 ' 0.5 inches = 36 points
        .HeaderMargin = 0
        .FooterMargin = 0
        .PrintGridlines = False
        .PrintHeadings = False
        .CenterHorizontally = False
        .CenterVertically = True  ' Center content vertically on page
    End With
    
    ' Set print area
    xlWorksheet.PageSetup.PrintArea = "A1:E" & (rowNum - 1)
    
    ' Get save path from user
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFileName, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save PDF As")
    
    If savePath = "False" Then
        ' User cancelled
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
    
    lblStatus.Caption = "PDF generated successfully: " & savePath
    lblStatus.ForeColor = RGB(0, 176, 80)
    MsgBox "PDF generated successfully!" & vbCrLf & savePath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error generating PDF: " & Err.Description
    lblStatus.ForeColor = RGB(192, 0, 0)
    MsgBox "Error generating PDF: " & Err.Description, vbCritical
    
    ' Cleanup on error
    On Error Resume Next
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

' Sector-wise PDF generation removed from form
' Use GenerateSectorWiseStockPDF() function from SectorWiseStockPDF.bas module instead

Private Sub btnClose_Click()
    Unload Me
End Sub

