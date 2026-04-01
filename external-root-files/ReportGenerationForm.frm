
Option Explicit

' Form controls are automatically declared by VBA when created in form designer
' No need to declare them here - VBA handles this automatically

Private WithEvents bookList As MSForms.ListBox
Private WithEvents txtBookFilter As MSForms.TextBox
Private WithEvents cboOrganization As MSForms.ComboBox
Private allBooks As Collection
Private bookSelectionMap As Object
Private bookTotals As Object

Private Sub UserForm_Initialize()
    InitializeBookListControl
    
    ' Initialize organization dropdown
    InitializeOrganizationControl
    
    ' Add 5th checkbox - chkAllReport (SummaryReportPdf.bas)
    EnsureChkAllReportCheckbox
    ' Add toggle button (Select All / Deselect All) for report type checkboxes
    EnsureReportTypeToggleButton
    
    ' Load filter data
    LoadFilterData
    
    ' Load book catalog for selection
    LoadBookListData
    
    ' Set default values
    SetDefaultValues
End Sub

Private Sub InitializeBookListControl()
    On Error Resume Next
    Set bookList = Me.Controls("lstBooks")
    On Error GoTo 0
    
    If bookList Is Nothing Then
        Set bookList = Me.Controls.Add("Forms.ListBox.1", "lstBooks", True)
        With bookList
            .Left = 18
            .Top = 210
            .Width = 320
            .Height = 200
        End With
    End If
    
    If TypeName(bookList) <> "ListBox" Then
        Set bookList = Nothing
        Exit Sub
    End If
    
    InitializeBookFilterControl
    
    On Error Resume Next
    With bookList
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
        .IntegralHeight = False
        .ColumnCount = 3
        .ColumnWidths = "220 pt;60 pt;0 pt"
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
    End With
    On Error GoTo 0
End Sub

Private Sub InitializeBookFilterControl()
    On Error Resume Next
    Set txtBookFilter = Me.Controls("txtBookFilter")
    On Error GoTo 0
    
    If txtBookFilter Is Nothing Then
        Set txtBookFilter = Me.Controls.Add("Forms.TextBox.1", "txtBookFilter", True)
        With txtBookFilter
            .Left = 18
            .Top = bookList.Top - 28    
            .Width = bookList.Width
            .Height = 18
        End With
    End If
    
    If Not txtBookFilter Is Nothing Then
        txtBookFilter.Text = ""
    End If
End Sub

Private Sub LoadBookListData()
    On Error GoTo ErrorHandler
    
    If bookList Is Nothing Then
        InitializeBookListControl
        If bookList Is Nothing Then
            lblStatus.Caption = "Unable to initialize book selector."
            lblStatus.ForeColor = RGB(192, 0, 0)
            Exit Sub
        End If
    End If
    
    lblStatus.Caption = "Loading available books..."
    lblStatus.ForeColor = RGB(0, 112, 192)
    DoEvents
    
    bookList.Clear
    
    Dim books As Collection
    Set books = SectorVillageDetailsReport.LoadAvailableBooks()
    
    If books Is Nothing Or books.count = 0 Then
        lblStatus.Caption = "No books available for selection."
        lblStatus.ForeColor = RGB(255, 140, 0)
        Exit Sub
    End If
    Set allBooks = books
    
    EnsureBookSelectionCache allBooks
    RecalculateBookTotals
    
    lblStatus.Caption = "Filters and book list loaded. Ready to generate report..."
    lblStatus.ForeColor = RGB(0, 176, 80)
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error loading books: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub RecalculateBookTotals()
    If allBooks Is Nothing Then Exit Sub
    
    Dim currentFilter As String
    currentFilter = ""
    If Not txtBookFilter Is Nothing Then
        currentFilter = txtBookFilter.Text
    End If
    
    PersistDisplayedSelections
    
    ' Always use YES and Not in Stock (B-Yes not used) regardless of checkbox values
    Dim orgName As String
    orgName = GetOrganizationName()
    Set bookTotals = SectorVillageDetailsReport.GetBookTotalsForFilters( _
        cboGroupSector.Text, cboPeriod.Text, True, True, False, orgName)
    
    If bookTotals Is Nothing Then
        Set bookTotals = CreateObject("Scripting.Dictionary")
    End If
    
    PopulateBookList currentFilter
End Sub

Private Sub PopulateBookList(Optional filterText As String = "")
    If bookList Is Nothing Then Exit Sub
    If allBooks Is Nothing Then Exit Sub
    
    Dim normalizedFilter As String
    normalizedFilter = LCase$(Trim$(filterText))
    
    bookList.Clear
    
    Dim totals As Object
    Set totals = bookTotals
    If totals Is Nothing Then
        Set totals = CreateObject("Scripting.Dictionary")
    End If
    
    ' Collect books with qty > 0 and apply filter
    Dim filteredBooks As Collection
    Set filteredBooks = New Collection
    
    Dim i As Long
    For i = 1 To allBooks.count
        Dim bookInfo As Variant
        bookInfo = allBooks.Item(i)
        
        Dim languageLabel As String
        Dim bookLabel As String
        languageLabel = Trim$(CStr(bookInfo(2)))
        bookLabel = Trim$(CStr(bookInfo(1)))
        
        ' Change format: Book Name - Language (instead of Language - Book Name)
        Dim displayLabel As String
        If languageLabel <> "" Then
            displayLabel = bookLabel & " - " & languageLabel
        Else
            displayLabel = bookLabel
        End If
        
        ' Only show books that have data for selected APP and period
        ' If book doesn't exist in totals, it means no data for selected APP/period
        Dim totalValue As Double
        totalValue = 0
        If totals.Exists(CStr(bookInfo(0))) Then
            totalValue = CDbl(totals(CStr(bookInfo(0))))
        Else
            ' Book doesn't exist in filtered totals (no data for selected APP/period)
            ' Skip this book - don't show it in the list
            GoTo NextBook
        End If
        
        ' Skip books with qty = 0 for selected period and APP
        If totalValue = 0 Then
            GoTo NextBook
        End If
        
        Dim includeBook As Boolean
        includeBook = (normalizedFilter = "")
        If Not includeBook Then
            Dim searchableText As String
            searchableText = LCase$(displayLabel & " " & languageLabel & " " & bookLabel & " " & CStr(totalValue))
            includeBook = InStr(searchableText, normalizedFilter) > 0
        End If
        
        If includeBook Then
            ' Store book info with display label for sorting
            Dim bookData(4) As Variant
            bookData(0) = bookInfo(0)  ' book_id
            bookData(1) = bookInfo(1)  ' book_name (for sorting)
            bookData(2) = bookInfo(2)  ' language
            bookData(3) = totalValue   ' qty
            bookData(4) = displayLabel ' display label
            filteredBooks.Add bookData
        End If
        
NextBook:
    Next i
    
    ' Sort books alphabetically by book name (A to Z)
    Call SortBooksByName(filteredBooks)
    
    ' Add sorted books to list
    For i = 1 To filteredBooks.count
        Dim sortedBookData As Variant
        sortedBookData = filteredBooks.Item(i)
        
        bookList.AddItem sortedBookData(4)  ' displayLabel
        
        bookList.List(bookList.ListCount - 1, 1) = Format$(sortedBookData(3), "#,##0")  ' qty
        bookList.List(bookList.ListCount - 1, 2) = CLng(sortedBookData(0))  ' book_id
        bookList.Selected(bookList.ListCount - 1) = GetBookSelectionState(CLng(sortedBookData(0)), True)
    Next i
End Sub

' Helper function to sort books by name (A to Z)
Private Sub SortBooksByName(books As Collection)
    If books.count <= 1 Then Exit Sub
    
    ' Convert collection to array for easier sorting
    Dim bookArray() As Variant
    ReDim bookArray(1 To books.count)
    
    Dim i As Long
    For i = 1 To books.count
        bookArray(i) = books.Item(i)
    Next i
    
    ' Simple bubble sort by book name (index 1)
    Dim j As Long
    Dim tempBook As Variant
    
    For i = 1 To books.count - 1
        For j = i + 1 To books.count
            If UCase$(CStr(bookArray(i)(1))) > UCase$(CStr(bookArray(j)(1))) Then
                ' Swap books
                tempBook = bookArray(i)
                bookArray(i) = bookArray(j)
                bookArray(j) = tempBook
            End If
        Next j
    Next i
    
    ' Clear collection and add sorted books back
    While books.count > 0
        books.Remove 1
    Wend
    
    For i = 1 To UBound(bookArray)
        books.Add bookArray(i)
    Next i
End Sub

Private Sub EnsureBookSelectionCache(books As Collection)
    If bookSelectionMap Is Nothing Then
        Set bookSelectionMap = CreateObject("Scripting.Dictionary")
    End If
    
    Dim workingMap As Object
    Set workingMap = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To books.count
        Dim bookInfo As Variant
        bookInfo = books.Item(i)
        Dim bookId As Long
        bookId = CLng(bookInfo(0))
        Dim key As String
        key = CStr(bookId)
        If bookSelectionMap.Exists(key) Then
            workingMap.Add key, bookSelectionMap(key)
        Else
            workingMap.Add key, True
        End If
    Next i
    
    Set bookSelectionMap = workingMap
End Sub

Private Function GetBookSelectionState(bookId As Long, Optional defaultValue As Boolean = True) As Boolean
    If bookSelectionMap Is Nothing Then
        Set bookSelectionMap = CreateObject("Scripting.Dictionary")
    End If
    
    Dim key As String
    key = CStr(bookId)
    
    If Not bookSelectionMap.Exists(key) Then
        bookSelectionMap.Add key, defaultValue
    End If
    
    GetBookSelectionState = CBool(bookSelectionMap(key))
End Function

Private Sub PersistDisplayedSelections()
    If bookSelectionMap Is Nothing Then Exit Sub
    If bookList Is Nothing Then Exit Sub
    
    Dim i As Long
    For i = 0 To bookList.ListCount - 1
        Dim bookId As Long
        bookId = CLng(bookList.List(i, 2))
        bookSelectionMap(CStr(bookId)) = bookList.Selected(i)
    Next i
End Sub

Private Sub bookList_Click()
    PersistDisplayedSelections
End Sub

Private Sub txtBookFilter_Change()
    PersistDisplayedSelections
    PopulateBookList txtBookFilter.Text
End Sub

Private Sub LoadFilterData()
    On Error GoTo ErrorHandler
    
    ' Load Organization data
    LoadOrganizationData
    
    ' Load Group Sector data
    LoadGroupSectorData
    
    ' Load Period data
    LoadPeriodData
    
    lblStatus.Caption = "Filters loaded successfully. Ready to generate report..."
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error loading filter data: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub LoadGroupSectorData()
    On Error GoTo ErrorHandler
    
    Dim groupSectors As Collection
    Dim sector As Variant
    
    ' Clear existing items
    cboGroupSector.Clear
    
    ' Add "All" option
    cboGroupSector.AddItem "All"
    
    ' Get group sector data from SectorVillageDetailsReport
    Set groupSectors = SectorVillageDetailsReport.LoadGroupSectorData()
    
    
    ' Add group sectors to combo box
    For Each sector In groupSectors
        cboGroupSector.AddItem sector
    Next sector
    
    
    ' Set default selection
    cboGroupSector.ListIndex = 0
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub LoadPeriodData()
    On Error GoTo ErrorHandler
    
    Dim periods As Collection
    Dim period As Variant
    Dim orgName As String
    
    ' Get selected APP value
    orgName = GetOrganizationName()
    
    ' Clear existing items
    cboPeriod.Clear
    
    ' Get period data from SectorVillageDetailsReport filtered by selected APP
    Set periods = SectorVillageDetailsReport.LoadPeriodData(orgName)
    
    
    ' Add periods to combo box in ascending order
    For Each period In periods
        cboPeriod.AddItem period
    Next period
    
    
    ' Set default selection to first period (if any exist)
    If cboPeriod.ListCount > 0 Then
        cboPeriod.ListIndex = 0
    End If
    
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub InitializeOrganizationControl()
    On Error Resume Next
    Set cboOrganization = Me.Controls("cboOrganization")
    On Error GoTo 0
    
    If cboOrganization Is Nothing Then
        ' Try to add the control programmatically
        On Error Resume Next
        Set cboOrganization = Me.Controls.Add("Forms.ComboBox.1", "cboOrganization", True)
        If Not cboOrganization Is Nothing Then
            ' Position it near other dropdowns (adjust position as needed)
            ' You may need to adjust Left, Top, Width, Height based on your form layout
            With cboOrganization
                .Left = 18
                .Top = 60  ' Adjust based on your form layout
                .Width = 120
                .Height = 20
            End With
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub EnsureChkAllReportCheckbox()
    On Error Resume Next
    Dim chk As Object
    Set chk = Me.Controls("chkAllReport")
    On Error GoTo 0
    If chk Is Nothing Then
        Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkAllReport", True)
        With chk
            .Caption = "All Report (SummaryReportPdf)"
            .Left = 18
            .Top = 168
            .Width = 180
            .Height = 18
            .Value = True
        End With
    End If
End Sub

Private Sub EnsureReportTypeToggleButton()
    On Error Resume Next
    Dim btn As Object
    Set btn = Me.Controls("btnChkselectUnselect")
    On Error GoTo 0
    If btn Is Nothing Then
        Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnChkselectUnselect", True)
        With btn
            .Caption = "Select All / Deselect All"
            .Left = 208
            .Top = 168
            .Width = 130
            .Height = 22
        End With
    End If
End Sub

Private Sub btnChkselectUnselect_Click()
    On Error Resume Next
    Dim anyChecked As Boolean
    anyChecked = chkDetailsReport.Value Or chkSummaryReport.Value Or chkSecoreSummary.Value Or chkAllSummary.Value
    On Error Resume Next
    anyChecked = anyChecked Or Me.Controls("chkAllReport").Value
    On Error GoTo 0
    Dim newVal As Boolean
    newVal = Not anyChecked
    chkDetailsReport.Value = newVal
    chkSummaryReport.Value = newVal
    chkSecoreSummary.Value = newVal
    chkAllSummary.Value = newVal
    Me.Controls("chkAllReport").Value = newVal
End Sub

Private Sub LoadOrganizationData()
    On Error GoTo ErrorHandler
    
    ' Initialize control if needed
    InitializeOrganizationControl
    
    ' Clear existing items
    On Error Resume Next
    If Not cboOrganization Is Nothing Then
        cboOrganization.Clear
    Else
        Me.Controls("cboOrganization").Clear
    End If
    On Error GoTo 0
    
    ' Load unique APP values from database
    Dim appValues As Collection
    Set appValues = GetUniqueAppValuesFromDatabase()
    
    ' Add organization options from database (VIMARSH first if available)
    If Not appValues Is Nothing And appValues.count > 0 Then
        Dim appValue As Variant
        Dim hasVimarsh As Boolean
        hasVimarsh = False
        
        ' First, add VIMARSH if it exists
        For Each appValue In appValues
            If UCase(Trim(CStr(appValue))) = "VIMARSH" Then
                On Error Resume Next
                If Not cboOrganization Is Nothing Then
                    cboOrganization.AddItem CStr(appValue)
                Else
                    Me.Controls("cboOrganization").AddItem CStr(appValue)
                End If
                On Error GoTo 0
                hasVimarsh = True
                Exit For
            End If
        Next appValue
        
        ' Then, add all other APP values
        For Each appValue In appValues
            If UCase(Trim(CStr(appValue))) <> "VIMARSH" Then
                On Error Resume Next
                If Not cboOrganization Is Nothing Then
                    cboOrganization.AddItem CStr(appValue)
                Else
                    Me.Controls("cboOrganization").AddItem CStr(appValue)
                End If
                On Error GoTo 0
            End If
        Next appValue
    Else
        ' Fallback to default values if database query fails
        On Error Resume Next
        If Not cboOrganization Is Nothing Then
            cboOrganization.AddItem "VIMARSH"
            cboOrganization.AddItem "PARISHKAAR"
        Else
            Me.Controls("cboOrganization").AddItem "VIMARSH"
            Me.Controls("cboOrganization").AddItem "PARISHKAAR"
        End If
        On Error GoTo 0
    End If
    
    ' Set default selection to VIMARSH (find and select it)
    On Error Resume Next
    Dim i As Long
    Dim foundVimarsh As Boolean
    foundVimarsh = False
    
    If Not cboOrganization Is Nothing Then
        If cboOrganization.ListCount > 0 Then
            ' Search for VIMARSH in the list
            For i = 0 To cboOrganization.ListCount - 1
                If UCase(Trim(cboOrganization.List(i))) = "VIMARSH" Then
                    cboOrganization.ListIndex = i
                    foundVimarsh = True
                    Exit For
                End If
            Next i
            ' If VIMARSH not found, select first item
            If Not foundVimarsh Then
                cboOrganization.ListIndex = 0
            End If
        End If
    Else
        Dim orgControl As Object
        Set orgControl = Me.Controls("cboOrganization")
        If orgControl.ListCount > 0 Then
            ' Search for VIMARSH in the list
            For i = 0 To orgControl.ListCount - 1
                If UCase(Trim(orgControl.List(i))) = "VIMARSH" Then
                    orgControl.ListIndex = i
                    foundVimarsh = True
                    Exit For
                End If
            Next i
            ' If VIMARSH not found, select first item
            If Not foundVimarsh Then
                orgControl.ListIndex = 0
            End If
        End If
    End If
    On Error GoTo 0
    
    Exit Sub
    
ErrorHandler:
    ' On error, use default values
    On Error Resume Next
    If Not cboOrganization Is Nothing Then
        cboOrganization.AddItem "VIMARSH"
        cboOrganization.AddItem "PARISHKAAR"
        If cboOrganization.ListCount > 0 Then
            ' Select VIMARSH (should be at index 0 since we added it first)
            cboOrganization.ListIndex = 0
        End If
    Else
        Me.Controls("cboOrganization").AddItem "VIMARSH"
        Me.Controls("cboOrganization").AddItem "PARISHKAAR"
        If Me.Controls("cboOrganization").ListCount > 0 Then
            ' Select VIMARSH (should be at index 0 since we added it first)
            Me.Controls("cboOrganization").ListIndex = 0
        End If
    End If
    On Error GoTo 0
End Sub

' Get unique APP values from database
Private Function GetUniqueAppValuesFromDatabase() As Collection
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim appValues As Collection
    Set appValues = New Collection
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connect to PostgreSQL
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=localhost;Port=5432;Database=vimarshbooks;Uid=postgres;Pwd=Ketan@757399;"
    
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        Set GetUniqueAppValuesFromDatabase = appValues
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Get unique APP values from transactions table
    sql = "SELECT DISTINCT app " & _
          "FROM vimars.transactions " & _
          "WHERE app IS NOT NULL " & _
          "AND app <> '' " & _
          "AND is_deleted = FALSE " & _
          "ORDER BY app;"
    
    Set rs = conn.Execute(sql)
    
    ' Add APP values to collection
    While Not rs.EOF
        Dim appValue As String
        appValue = Trim(CStr(rs.Fields("app").Value))
        If appValue <> "" Then
            ' Check if already added (to avoid duplicates)
            Dim found As Boolean
            found = False
            Dim existingApp As Variant
            For Each existingApp In appValues
                If UCase(CStr(existingApp)) = UCase(appValue) Then
                    found = True
                    Exit For
                End If
            Next existingApp
            
            If Not found Then
                appValues.Add appValue
            End If
        End If
        rs.MoveNext
    Wend
    
    ' Close connection
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Set GetUniqueAppValuesFromDatabase = appValues
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
    Set GetUniqueAppValuesFromDatabase = appValues
End Function

Private Sub cboGroupSector_Change()
    RecalculateBookTotals
End Sub

Private Sub cboGroupSector_Click()
    RecalculateBookTotals
End Sub

Private Sub cboPeriod_Change()
    RecalculateBookTotals
End Sub

Private Sub cboPeriod_Click()
    RecalculateBookTotals
End Sub

Private Sub cboOrganization_Change()
    ' Reload periods filtered by selected APP
    LoadPeriodData
    ' Recalculate book totals for new APP and current period
    RecalculateBookTotals
End Sub

Private Sub cboOrganization_Click()
    ' Reload periods filtered by selected APP
    LoadPeriodData
    ' Recalculate book totals for new APP and current period
    RecalculateBookTotals
End Sub

Private Sub chkYes_Click()
    ' Availability checkboxes don't affect reports - always use YES and NS
    RecalculateBookTotals
End Sub

Private Sub chkNotInStock_Click()
    ' Availability checkboxes don't affect reports - always use YES and NS
    RecalculateBookTotals
End Sub

Private Sub chkBYes_Click()
    ' Availability checkboxes don't affect reports - always use YES and NS
    RecalculateBookTotals
End Sub

Private Sub SetDefaultValues()
    ' Set default values for dropdowns
    On Error Resume Next
    If cboGroupSector.ListCount > 0 Then cboGroupSector.ListIndex = 0  ' "All" is at index 0 for group sector
    If cboPeriod.ListCount > 0 Then cboPeriod.ListIndex = 0  ' First period is at index 0
    
    ' Ensure VIMARSH is selected in organization dropdown
    Dim i As Long
    Dim foundVimarsh As Boolean
    foundVimarsh = False
    
    If Not cboOrganization Is Nothing Then
        If cboOrganization.ListCount > 0 Then
            ' Search for VIMARSH in the list
            For i = 0 To cboOrganization.ListCount - 1
                If UCase(Trim(cboOrganization.List(i))) = "VIMARSH" Then
                    cboOrganization.ListIndex = i
                    foundVimarsh = True
                    Exit For
                End If
            Next i
            ' If VIMARSH not found, select first item
            If Not foundVimarsh Then
                cboOrganization.ListIndex = 0
            End If
        End If
    Else
        Dim orgControl As Object
        Set orgControl = Me.Controls("cboOrganization")
        If orgControl.ListCount > 0 Then
            ' Search for VIMARSH in the list
            For i = 0 To orgControl.ListCount - 1
                If UCase(Trim(orgControl.List(i))) = "VIMARSH" Then
                    orgControl.ListIndex = i
                    foundVimarsh = True
                    Exit For
                End If
            Next i
            ' If VIMARSH not found, select first item
            If Not foundVimarsh Then
                orgControl.ListIndex = 0
            End If
        End If
    End If
    On Error GoTo 0
    
    ' Set default availability checkboxes - YES and Not in Stock (B-Yes not used)
    On Error Resume Next
    Me.Controls("chkYes").Value = True
    Me.Controls("chkNotInStock").Value = True
    Me.Controls("chkBYes").Value = False
    On Error GoTo 0
    
    ' Set default report type checkboxes - all 5 checked
    chkDetailsReport.Value = True      ' Sector-Village Details (default checked)
    chkSummaryReport.Value = True     ' Sector-Village Summary (default checked)
    chkSecoreSummary.Value = True     ' Sector Summary (default checked)
    chkAllSummary.Value = True        ' All Summary Selected Period (default checked)
    On Error Resume Next
    Me.Controls("chkAllReport").Value = True
    On Error GoTo 0
    
    RecalculateBookTotals
End Sub

Private Sub btnPreview_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedBookIds As Collection
    Set selectedBookIds = GetSelectedBookIds()
    
    If selectedBookIds.count = 0 Then
        MsgBox "Please keep at least one book selected to generate reports.", vbExclamation
        lblStatus.Caption = "Select at least one book to generate reports."
        lblStatus.ForeColor = RGB(192, 0, 0)
        Exit Sub
    End If
    
    lblStatus.Caption = "Generating preview..."
    DoEvents
    
    ' Generate and show preview
    Dim reportData As Collection
    Set reportData = GetFilteredData(selectedBookIds)
    
            If reportData.count > 0 Then
            lblStatus.Caption = "Data found successfully. Found " & reportData.count & " records."
        Else
        lblStatus.Caption = "No data found with selected filters."
        lblStatus.ForeColor = RGB(255, 140, 0)
    End If
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error generating preview: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub btnGenerateReport_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedBookIds As Collection
    Set selectedBookIds = GetSelectedBookIds()
    
    If selectedBookIds.count = 0 Then
        MsgBox "Please keep at least one book selected to generate PDF reports.", vbExclamation
        lblStatus.Caption = "Select at least one book to generate PDF reports."
        lblStatus.ForeColor = RGB(192, 0, 0)
        Exit Sub
    End If
    
    ' Check if at least one report type is selected
    Dim chkAllReportVal As Boolean
    chkAllReportVal = False
    On Error Resume Next
    chkAllReportVal = Me.Controls("chkAllReport").Value
    On Error GoTo ErrorHandler
    If Not (chkDetailsReport.Value Or chkSummaryReport.Value Or chkSecoreSummary.Value Or chkAllSummary.Value Or chkAllReportVal) Then
        MsgBox "Please select at least one report type to generate.", vbExclamation
        Exit Sub
    End If
    
    ' Ask user to select folder for saving reports
    Dim selectedFolder As String
    selectedFolder = BrowseForFolder("Select folder to save reports")
    
    If selectedFolder = "" Then
        MsgBox "No folder selected. Report generation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Create folder name based on selected APP (organization), group sector and period
    Dim folderName As String
    Dim orgName As String
    orgName = GetOrganizationName()
    folderName = orgName & "_" & cboGroupSector.Text & "_" & cboPeriod.Text
    
    ' Create the full folder path
    Dim fullFolderPath As String
    fullFolderPath = selectedFolder & "\" & folderName
    
    ' Create the folder if it doesn't exist
    If Dir(fullFolderPath, vbDirectory) = "" Then
        MkDir fullFolderPath
    End If
    
    lblStatus.Caption = "Generating PDF reports in folder: " & fullFolderPath
    DoEvents
    
    Dim totalReportsGenerated As Integer
    totalReportsGenerated = 0
    
    ' Get organization name once for all reports (same as used in folder name)
    
    ' Generate Details Report (Secor-Village Details)
    If chkDetailsReport.Value Then
        lblStatus.Caption = "Generating Details Report..."
        DoEvents
        
        Dim reportData As Collection
        Set reportData = GetFilteredData(selectedBookIds)
        
        If reportData.count > 0 Then
            SectorVillageDetailsReport.GeneratePDFReport reportData, cboGroupSector.Text, cboPeriod.Text, GetAvailabilityText(), fullFolderPath, orgName
            totalReportsGenerated = totalReportsGenerated + 1
            lblStatus.Caption = "Details Report generated. Generating next report..."
            DoEvents
        Else
            lblStatus.Caption = "No data found for Details Report."
            lblStatus.ForeColor = RGB(255, 140, 0)
        End If
    End If
    
    ' Generate Summary Report (Secor-Village Summary)
    If chkSummaryReport.Value Then
        lblStatus.Caption = "Generating Summary Report..."
        DoEvents
        
        Dim summaryData As Collection
        Set summaryData = SectorVillageSummaryReport.GetSummaryDataForPDF(cboPeriod.Text, cboGroupSector.Text, selectedBookIds, orgName)
        
        If summaryData.count > 0 Then
            SectorVillageSummaryReport.GenerateSummaryPDF summaryData, cboPeriod.Text, cboGroupSector.Text, fullFolderPath, orgName
            totalReportsGenerated = totalReportsGenerated + 1
            lblStatus.Caption = "Summary Report generated. Generating next report..."
            DoEvents
        Else
            lblStatus.Caption = "No data found for Summary Report."
            lblStatus.ForeColor = RGB(255, 140, 0)
        End If
    End If
    
    ' Generate Secore Summary Report
    If chkSecoreSummary.Value Then
        lblStatus.Caption = "Generating Secore Summary Report..."
        DoEvents
        
        ' Get selected values from form controls
        Dim selectedSector As String
        Dim selectedPeriod As String
        
        selectedSector = cboGroupSector.Text
        selectedPeriod = cboPeriod.Text
        
        ' Generate the sector summary report
        Call SectorVillageDetailsReport.GenerateSectorSummaryReportFromForm(selectedSector, selectedPeriod, selectedBookIds, fullFolderPath, orgName)
        totalReportsGenerated = totalReportsGenerated + 1
        lblStatus.Caption = "Secore Summary Report generated. Generating next report..."
        DoEvents
    End If
    
    ' Generate All Summary Selected Period Report
    If chkAllSummary.Value Then
        lblStatus.Caption = "Generating All Summary Report..."
        DoEvents
        
        ' Get selected period from form control (reuse the variable declared above)
        selectedPeriod = cboPeriod.Text
        
        ' Generate the all summary report
        Call SectorVillageDetailsReport.GenerateAllSummaryReportFromForm(selectedPeriod, selectedBookIds, fullFolderPath, orgName)
        totalReportsGenerated = totalReportsGenerated + 1
        lblStatus.Caption = "All Summary Report generated. Generating next report..."
        DoEvents
    End If
    
    ' Generate All Report (5th - SummaryReportPdf.bas)
    On Error Resume Next
    chkAllReportVal = Me.Controls("chkAllReport").Value
    On Error GoTo ErrorHandler
    If chkAllReportVal Then
        lblStatus.Caption = "Generating All Report (SummaryReportPdf)..."
        DoEvents
        Call SummaryReportPdf.GenerateSummaryReportPdfFromForm(cboPeriod.Text, cboGroupSector.Text, selectedBookIds, fullFolderPath, orgName)
        totalReportsGenerated = totalReportsGenerated + 1
        lblStatus.Caption = "All Report generated. Generating next report..."
        DoEvents
    End If
    
    ' Final status message
    If totalReportsGenerated > 0 Then
        lblStatus.Caption = "Successfully generated " & totalReportsGenerated & " PDF report(s)!"
        lblStatus.ForeColor = RGB(0, 176, 80)
        
        ' Open the folder containing the generated reports
        Shell "explorer.exe """ & fullFolderPath & """", vbNormalFocus
        
        MsgBox "Reports generated successfully!" & vbCrLf & vbCrLf & _
               "Total reports: " & totalReportsGenerated & vbCrLf & _
               "Location: " & fullFolderPath & vbCrLf & vbCrLf & _
               "The folder will open automatically.", vbInformation, "Reports Generated"
        
        ' Close the form after user clicks OK
        Unload Me
    Else
        lblStatus.Caption = "No reports were generated. Please check your selections."
        lblStatus.ForeColor = RGB(255, 140, 0)
    End If
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error generating PDF: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub btnCancel_Click()
    ' Close form
    Unload Me
End Sub

Private Sub btnFixBookLanguage_Click()
    On Error GoTo ErrorHandler
    
    lblStatus.Caption = "Fixing book language in database..."
    DoEvents
    
    ' Call the function to fix book language
    SectorVillageSummaryReport.FixBookLanguage
    
    lblStatus.Caption = "Book language fix completed. Check the message box for details."
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error fixing book language: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub btnCheckBookLanguages_Click()
    On Error GoTo ErrorHandler
    
    lblStatus.Caption = "Checking book languages in database..."
    DoEvents
    
    ' Call the function to check book languages
    SectorVillageSummaryReport.CheckBookLanguages
    
    lblStatus.Caption = "Book language check completed. Check the message box for details."
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error checking book languages: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Sub btnCheckE10Data_Click()
    On Error GoTo ErrorHandler
    
    lblStatus.Caption = "Checking E10 sector data..."
    DoEvents
    
    ' Call the function to check E10 sector data
    SectorVillageSummaryReport.CheckE10SectorData
    
    lblStatus.Caption = "E10 sector data check completed. Check the message box for details."
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error checking E10 sector data: " & Err.description
    lblStatus.ForeColor = RGB(192, 0, 0)
End Sub

Private Function GetFilteredData(Optional selectedBookIds As Collection = Nothing) As Collection
    On Error GoTo ErrorHandler
    
    If selectedBookIds Is Nothing Then
        Set selectedBookIds = GetSelectedBookIds()
    End If
    
    If selectedBookIds.count = 0 Then
        lblStatus.Caption = "Select at least one book to continue."
        lblStatus.ForeColor = RGB(192, 0, 0)
        Set GetFilteredData = New Collection
        Exit Function
    End If
    
    ' Always use YES and Not in Stock (B-Yes not used) regardless of checkbox values
    Dim orgName As String
    orgName = GetOrganizationName()
    Set GetFilteredData = SectorVillageDetailsReport.GetFilteredReportData( _
        cboGroupSector.Text, cboPeriod.Text, True, True, False, selectedBookIds, orgName)
    Exit Function
    
ErrorHandler:
    Set GetFilteredData = New Collection
End Function

Private Function GetSelectedBookIds() As Collection
    Dim selectedIds As Collection
    Set selectedIds = New Collection
    
    PersistDisplayedSelections
    
    If bookSelectionMap Is Nothing Then
        Set GetSelectedBookIds = selectedIds
        Exit Function
    End If
    
    Dim key As Variant
    For Each key In bookSelectionMap.Keys
        If CBool(bookSelectionMap(key)) Then
            On Error Resume Next
            selectedIds.Add CLng(key)
            On Error GoTo 0
        End If
    Next key
    
    Set GetSelectedBookIds = selectedIds
End Function

Private Function GetAvailabilityText() As String
    ' Always return YES and Not In Stock (B-Yes not used)
    GetAvailabilityText = "Yes, Not In Stock"
End Function

Private Function GetOrganizationName() As String
    On Error Resume Next
    Dim orgName As String
    If Not cboOrganization Is Nothing Then
        orgName = cboOrganization.Text
    Else
        orgName = Me.Controls("cboOrganization").Text
    End If
    If orgName = "" Then
        orgName = "VIMARSH"  ' Default value
    End If
    GetOrganizationName = orgName
    On Error GoTo 0
End Function

' Function to browse for folder
Private Function BrowseForFolder(Optional prompt As String = "Select a folder") As String
    Dim shellApp As Object
    Dim folder As Object
    
    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, prompt, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
    
    Set folder = Nothing
    Set shellApp = Nothing
End Function



