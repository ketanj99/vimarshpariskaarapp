Option Explicit

' Import multiple Excel files
Public Sub ImportMultipleExcelFiles()
    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    Dim fileName As String
    Dim fileCount As Long
    
    ' Get folder path from user
    folderPath = BrowseForFolder("Select folder containing Excel files")
    If folderPath = "" Then Exit Sub
    
    ' Initialize system if not already done
    Call InitializeSystem
    
    ' Get connection
    Dim conn As Object
    Set conn = GetConnection()
    
    ' Loop through all Excel files in folder
    fileName = Dir(folderPath & "\*.xls*")
    Do While fileName <> ""
        Application.StatusBar = "Importing: " & fileName
        
        ' Import single file
        ImportExcelFile folderPath & "\" & fileName, conn
        
        fileCount = fileCount + 1
        fileName = Dir()
    Loop
    
    Application.StatusBar = False
    MsgBox "Successfully imported " & fileCount & " files!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error importing files: " & Err.Description, vbCritical
End Sub

' Import single Excel file
Private Sub ImportExcelFile(filePath As String, conn As Object)
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim bookName As String, bookID As Long
    Dim language As String, year As Integer, month As Integer
    Dim half As String, qty As Integer, avail As Integer
    Dim village As String, groupSector As String, contactName As String
    Dim transactionDate As Date
    Dim amount As Currency
    Dim transactionType As String
    Dim description As String
    
    ' Open workbook
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Sheets(1) ' Assume first sheet
    
    ' Find last row and column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Validate headers
    If Not ValidateHeaders(ws) Then
        MsgBox "Invalid file format. Please check column headers.", vbExclamation
        wb.Close False
        Set wb = Nothing
        Exit Sub
    End If
    
    ' Process each row (skip header)
    For i = 2 To lastRow
        ' Extract data from columns based on your format
        bookName = Trim(ws.Cells(i, 1).Value)        ' BOOK TITLE
        language = Trim(ws.Cells(i, 2).Value)        ' LANGUAGE
        year = Val(ws.Cells(i, 3).Value)             ' YEAR
        month = Val(ws.Cells(i, 4).Value)            ' MONTH
        half = Trim(ws.Cells(i, 5).Value)            ' HALF
        qty = Val(ws.Cells(i, 6).Value)              ' QTY
        avail = Val(ws.Cells(i, 7).Value)            ' AVAIL
        village = Trim(ws.Cells(i, 8).Value)         ' VILLAGE
        groupSector = Trim(ws.Cells(i, 9).Value)     ' GROUP(SECTOR)
        contactName = Trim(ws.Cells(i, 10).Value)    ' NAME(CONTACT NO.)
        
        ' Validate required data
        If bookName <> "" And year > 0 And month > 0 And month <= 12 And qty > 0 Then
            ' Create transaction date from year and month
            transactionDate = DateSerial(year, month, 1)
            
            ' Calculate amount based on quantity (you can adjust this logic)
            amount = qty * 100 ' Assuming 100 per unit, adjust as needed
            
            ' Set transaction type based on availability
            If avail > 0 Then
                transactionType = "Available"
            Else
                transactionType = "Not Available"
            End If
            
            ' Create description with additional details
            description = "Language: " & language & ", Village: " & village & _
                         ", Group: " & groupSector & ", Contact: " & contactName & _
                         ", Half: " & half
            
            ' Get or create book ID with additional information
            bookID = GetOrCreateBookID(bookName, language, village, groupSector, contactName, conn)
            
            ' Insert transaction
            InsertTransaction bookID, transactionDate, amount, transactionType, description, conn
        End If
    Next i
    
    ' Close workbook
    wb.Close False
    Set wb = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not wb Is Nothing Then
        wb.Close False
        Set wb = Nothing
    End If
    MsgBox "Error importing file " & filePath & ": " & Err.Description, vbCritical
End Sub

' Get or create book ID with additional information
Private Function GetOrCreateBookID(bookName As String, language As String, village As String, _
                                 groupSector As String, contactName As String, conn As Object) As Long
    On Error GoTo ErrorHandler
    
    Dim rs As Object
    Dim sql As String
    
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Check if book exists
    sql = "SELECT BookID FROM Books WHERE BookName = '" & Replace(bookName, "'", "''") & "'"
    rs.Open sql, conn
    
    If Not rs.EOF Then
        GetOrCreateBookID = rs.Fields("BookID").Value
        
        ' Update additional information if it's different
        sql = "UPDATE Books SET Language = '" & Replace(language, "'", "''") & "', " & _
              "Village = '" & Replace(village, "'", "''") & "', " & _
              "GroupSector = '" & Replace(groupSector, "'", "''") & "', " & _
              "ContactName = '" & Replace(contactName, "'", "''") & "' " & _
              "WHERE BookID = " & GetOrCreateBookID
        conn.Execute sql
    Else
        ' Create new book with all information
        rs.Close
        sql = "INSERT INTO Books (BookName, Language, Village, GroupSector, ContactName) " & _
              "VALUES ('" & Replace(bookName, "'", "''") & "', '" & _
              Replace(language, "'", "''") & "', '" & _
              Replace(village, "'", "''") & "', '" & _
              Replace(groupSector, "'", "''") & "', '" & _
              Replace(contactName, "'", "''") & "')"
        conn.Execute sql
        
        ' Get the new book ID
        sql = "SELECT @@IDENTITY AS NewID"
        rs.Open sql, conn
        GetOrCreateBookID = rs.Fields("NewID").Value
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrorHandler:
    GetOrCreateBookID = 0
End Function

' Insert transaction record
Private Sub InsertTransaction(bookID As Long, transactionDate As Date, amount As Currency, _
                           transactionType As String, description As String, conn As Object)
    On Error GoTo ErrorHandler
    
    Dim sql As String
    
    sql = "INSERT INTO Transactions (BookID, TransactionDate, Amount, TransactionType, Description) " & _
          "VALUES (" & bookID & ", #" & Format(transactionDate, "yyyy-mm-dd") & "#, " & _
          amount & ", '" & Replace(transactionType, "'", "''") & "', '" & _
          Replace(description, "'", "''") & "')"
    
    conn.Execute sql
    Exit Sub
    
ErrorHandler:
    ' Continue processing other records
End Sub

' Browse for folder
Private Function BrowseForFolder(Optional prompt As String = "Select Folder") As String
    On Error GoTo ErrorHandler
    
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
    Exit Function
    
ErrorHandler:
    BrowseForFolder = ""
End Function

' Import from specific Excel file
Public Sub ImportFromSpecificFile()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    
    ' Get file path from user
    filePath = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", , "Select Excel file to import")
    If filePath = "False" Then Exit Sub
    
    ' Initialize system
    Call InitializeSystem
    
    ' Import file
    ImportExcelFile filePath, GetConnection()
    
    MsgBox "File imported successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error importing file: " & Err.Description, vbCritical
End Sub

' Validate Excel file format
Public Function ValidateExcelFormat(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headerRow As Range
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Sheets(1)
    
    ' Check if required columns exist for your format
    Set headerRow = ws.Rows(1)
    
    ValidateExcelFormat = (InStr(1, headerRow.Cells(1, 1).Value, "BOOK TITLE", vbTextCompare) > 0) And _
                         (InStr(1, headerRow.Cells(1, 2).Value, "LANGUAGE", vbTextCompare) > 0) And _
                         (InStr(1, headerRow.Cells(1, 3).Value, "YEAR", vbTextCompare) > 0) And _
                         (InStr(1, headerRow.Cells(1, 4).Value, "MONTH", vbTextCompare) > 0) And _
                         (InStr(1, headerRow.Cells(1, 6).Value, "QTY", vbTextCompare) > 0)
    
    wb.Close False
    Set wb = Nothing
    Exit Function
    
ErrorHandler:
    ValidateExcelFormat = False
    If Not wb Is Nothing Then
        wb.Close False
        Set wb = Nothing
    End If
End Function

' Validate headers for your specific format
Private Function ValidateHeaders(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim headerRow As Range
    Set headerRow = ws.Rows(1)
    
    ' Check for your specific column headers
    ValidateHeaders = (InStr(1, headerRow.Cells(1, 1).Value, "BOOK TITLE", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 2).Value, "LANGUAGE", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 3).Value, "YEAR", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 4).Value, "MONTH", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 5).Value, "HALF", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 6).Value, "QTY", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 7).Value, "AVAIL", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 8).Value, "VILLAGE", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 9).Value, "GROUP", vbTextCompare) > 0) And _
                     (InStr(1, headerRow.Cells(1, 10).Value, "NAME", vbTextCompare) > 0)
    
    Exit Function
    
ErrorHandler:
    ValidateHeaders = False
End Function
