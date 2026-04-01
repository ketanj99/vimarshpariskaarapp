Option Explicit

' Database Backup and Restore Module
' This module provides functions to backup and restore PostgreSQL database

' Constants
Private Const DB_HOST As String = "localhost"
Private Const DB_PORT As String = "5432"
Private Const DB_NAME As String = "vimarshbooks"
Private Const DB_USER As String = "postgres"
Private Const DB_PASSWORD As String = "Ketan@757399"
Private Const DB_SCHEMA As String = "vimars"
Private Const BACKUP_FOLDER As String = "F:\VIMARS App Excle\Backups"

' Backup database data to SQL file
Public Sub BackupDatabase()
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim backupFile As String
    Dim fileNum As Integer
    Dim backupDateTime As String
    Dim folderPath As String
    
    ' Create backup folder if it doesn't exist
    folderPath = BACKUP_FOLDER
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Generate backup filename with timestamp
    backupDateTime = Format(Now, "yyyymmdd_hhmmss")
    backupFile = folderPath & "\backup_" & backupDateTime & ".sql"
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        MsgBox "Failed to connect to PostgreSQL database.", vbCritical
        Exit Sub
    End If
    
    ' Open file for writing
    fileNum = FreeFile
    Open backupFile For Output As #fileNum
    
    ' Write backup header
    Print #fileNum, "-- VIMARS Database Backup"
    Print #fileNum, "-- Created: " & Format(Now, "dd-mm-yyyy hh:mm:ss")
    Print #fileNum, "-- Schema: " & DB_SCHEMA
    Print #fileNum, ""
    Print #fileNum, "BEGIN;"
    Print #fileNum, ""
    Print #fileNum, "-- Delete existing data in correct order (considering foreign keys)"
    Print #fileNum, "-- Delete from child tables first (transactions), then parent tables"
    Print #fileNum, "DELETE FROM " & DB_SCHEMA & ".transactions WHERE is_deleted = FALSE OR is_deleted IS NULL;"
    Print #fileNum, "DELETE FROM " & DB_SCHEMA & ".books WHERE is_deleted = FALSE OR is_deleted IS NULL;"
    Print #fileNum, "DELETE FROM " & DB_SCHEMA & ".import_history WHERE is_deleted = FALSE OR is_deleted IS NULL;"
    Print #fileNum, ""
    
    ' Backup tables in correct order (considering foreign keys)
    ' First backup parent tables (books, import_history)
    Print #fileNum, "-- Backup Books Table"
    BackupTableToSQLWithoutDelete conn, "books", fileNum
    
    Print #fileNum, "-- Backup Import History Table"
    BackupTableToSQLWithoutDelete conn, "import_history", fileNum
    
    ' Then backup child tables (transactions)
    Print #fileNum, "-- Backup Transactions Table"
    BackupTableToSQLWithoutDelete conn, "transactions", fileNum
    
    ' Write backup footer
    Print #fileNum, ""
    Print #fileNum, "COMMIT;"
    Print #fileNum, "-- Backup completed successfully"
    
    ' Close file
    Close #fileNum
    
    ' Close connection
    conn.Close
    Set conn = Nothing
    
    MsgBox "Database backup completed successfully!" & vbCrLf & vbCrLf & _
           "Backup file: " & backupFile & vbCrLf & vbCrLf & _
           "File size: " & Format(FileLen(backupFile) / 1024, "#,##0.00") & " KB", vbInformation, "Backup Successful"
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    MsgBox "Error creating database backup: " & Err.Description, vbCritical
End Sub

' Backup a specific table to SQL file (with DELETE statement)
Private Sub BackupTableToSQL(conn As Object, tableName As String, fileNum As Integer)
    On Error GoTo ErrorHandler
    
    ' Delete existing data first
    Print #fileNum, "DELETE FROM " & DB_SCHEMA & "." & tableName & ";"
    Print #fileNum, ""
    
    ' Then backup data
    BackupTableToSQLWithoutDelete conn, tableName, fileNum
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in BackupTableToSQL: " & Err.Description
End Sub

' Backup a specific table to SQL file (without DELETE statement)
Private Sub BackupTableToSQLWithoutDelete(conn As Object, tableName As String, fileNum As Integer)
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    Dim field As Object
    Dim values As String
    Dim columns As String
    Dim rowCount As Long
    
    ' Get all columns for the table
    sql = "SELECT column_name, data_type " & _
          "FROM information_schema.columns " & _
          "WHERE table_schema = '" & DB_SCHEMA & "' " & _
          "AND table_name = '" & tableName & "' " & _
          "ORDER BY ordinal_position;"
    
    Set rs = conn.Execute(sql)
    
    ' Build column list
    columns = ""
    Dim columnNames As Collection
    Set columnNames = New Collection
    Dim columnTypes As Object
    Set columnTypes = CreateObject("Scripting.Dictionary")
    
    While Not rs.EOF
        Dim colName As String
        Dim colType As String
        colName = rs.Fields("column_name").Value
        colType = rs.Fields("data_type").Value
        
        If columns <> "" Then columns = columns & ", "
        columns = columns & colName
        
        columnNames.Add colName
        columnTypes.Add colName, colType
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    ' Get all data from table
    sql = "SELECT * FROM " & DB_SCHEMA & "." & tableName & " " & _
          "WHERE is_deleted = FALSE OR is_deleted IS NULL " & _
          "ORDER BY "
    
    ' Add appropriate ordering column based on table
    If tableName = "books" Then
        sql = sql & "book_id;"
    ElseIf tableName = "transactions" Then
        sql = sql & "importid;"
    ElseIf tableName = "import_history" Then
        sql = sql & "history_id;"
    Else
        sql = sql & "1;"
    End If
    
    Set rs = conn.Execute(sql)
    rowCount = 0
    
    While Not rs.EOF
        values = ""
        Dim colName As Variant
        Dim colValue As Variant
        
        For Each colName In columnNames
            If values <> "" Then values = values & ", "
            
            colValue = rs.Fields(CStr(colName)).Value
            
            ' Handle NULL values
            If IsNull(colValue) Then
                values = values & "NULL"
            Else
                Dim colType As String
                Dim strValue As String
                colType = columnTypes(CStr(colName))
                
                ' Format value based on data type
                If colType = "integer" Or colType = "bigint" Or colType = "smallint" Then
                    values = values & CStr(colValue)
                ElseIf colType = "boolean" Then
                    If CBool(colValue) Then
                        values = values & "TRUE"
                    Else
                        values = values & "FALSE"
                    End If
                ElseIf colType = "timestamp without time zone" Or colType = "timestamp with time zone" Or colType = "date" Or colType = "time" Then
                    ' Date/timestamp formatting for PostgreSQL
                    Dim dateValue As Date
                    If IsDate(colValue) Then
                        dateValue = CDate(colValue)
                        If colType = "date" Then
                            ' Format as date: YYYY-MM-DD
                            strValue = Format(dateValue, "yyyy-mm-dd")
                        Else
                            ' Format as timestamp: YYYY-MM-DD HH:MM:SS
                            strValue = Format(dateValue, "yyyy-mm-dd hh:nn:ss")
                        End If
                    Else
                        strValue = CStr(colValue)
                    End If
                    strValue = Replace(strValue, "'", "''")
                    values = values & "'" & strValue & "'"
                Else
                    ' Text and other types - need quotes and escape
                    strValue = CStr(colValue)
                    strValue = Replace(strValue, "'", "''")
                    strValue = Replace(strValue, "\", "\\")
                    values = values & "'" & strValue & "'"
                End If
            End If
        Next colName
        
        ' Write INSERT statement
        Print #fileNum, "INSERT INTO " & DB_SCHEMA & "." & tableName & " (" & columns & ") VALUES (" & values & ");"
        
        rowCount = rowCount + 1
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Print #fileNum, "-- " & tableName & ": " & rowCount & " records backed up"
    Print #fileNum, ""
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Debug.Print "Error backing up table " & tableName & ": " & Err.Description
End Sub

' Restore database from backup SQL file
Public Sub RestoreDatabase()
    On Error GoTo ErrorHandler
    
    Dim backupFile As String
    Dim conn As Object
    Dim sql As String
    Dim fileNum As Integer
    Dim fileLine As String
    Dim sqlCommand As String
    Dim confirmMsg As String
    Dim fileContent As String
    Dim allCommands As String
    
    ' Ask user to select backup file
    backupFile = Application.GetOpenFilename( _
        FileFilter:="SQL Backup Files (*.sql),*.sql,All Files (*.*),*.*", _
        Title:="Select backup file to restore", _
        MultiSelect:=False)
    
    If backupFile = "False" Or backupFile = "" Then
        Exit Sub
    End If
    
    ' Confirm restore action (this will delete existing data)
    confirmMsg = "WARNING: This will restore data from backup file and may overwrite existing data!" & vbCrLf & vbCrLf & _
                 "Backup file: " & backupFile & vbCrLf & vbCrLf & _
                 "File size: " & Format(FileLen(backupFile) / 1024, "#,##0.00") & " KB" & vbCrLf & vbCrLf & _
                 "Do you want to continue?" & vbCrLf & _
                 "This action cannot be undone!"
    
    If MsgBox(confirmMsg, vbYesNo + vbCritical, "Confirm Restore") <> vbYes Then
        Exit Sub
    End If
    
    ' Create database connection
    Set conn = CreateObject("ADODB.Connection")
    If Not ConnectToPostgreSQL(conn) Then
        MsgBox "Failed to connect to PostgreSQL database.", vbCritical
        Exit Sub
    End If
    
    ' Open backup file for reading
    fileNum = FreeFile
    Open backupFile For Input As #fileNum
    
    ' Read entire file content
    fileContent = Input(LOF(fileNum), #fileNum)
    Close #fileNum
    
    ' Split by semicolon and execute each SQL command
    Dim commands() As String
    Dim i As Long
    Dim executedCount As Long
    Dim failedCount As Long
    executedCount = 0
    failedCount = 0
    
    ' Simple split by semicolon (handle multi-line statements)
    commands = Split(fileContent, ";")
    
    ' Execute each command
    For i = 0 To UBound(commands)
        sqlCommand = Trim(commands(i))
        
        ' Skip empty commands and comments
        If sqlCommand = "" Or Left(sqlCommand, 2) = "--" Then
            GoTo NextCommand
        End If
        
        ' Skip BEGIN/COMMIT TRANSACTION if present
        If UCase(Trim(sqlCommand)) = "BEGIN TRANSACTION" Or UCase(Trim(sqlCommand)) = "COMMIT TRANSACTION" Then
            GoTo NextCommand
        End If
        
        ' Execute SQL command
        On Error Resume Next
        conn.Execute sqlCommand
        If Err.Number = 0 Then
            executedCount = executedCount + 1
        Else
            failedCount = failedCount + 1
            Debug.Print "Error executing SQL command " & (i + 1) & ": " & Err.Description
            Debug.Print "SQL: " & Left(sqlCommand, 150)
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        
NextCommand:
    Next i
    
    ' Close connection
    conn.Close
    Set conn = Nothing
    
    Dim resultMsg As String
    resultMsg = "Database restore completed!" & vbCrLf & vbCrLf & _
                "Backup file: " & backupFile & vbCrLf & _
                "SQL commands executed successfully: " & executedCount
    
    If failedCount > 0 Then
        resultMsg = resultMsg & vbCrLf & "SQL commands failed: " & failedCount & vbCrLf & _
                    "Check Debug window for details."
    End If
    
    MsgBox resultMsg, IIf(failedCount > 0, vbExclamation, vbInformation), "Restore Completed"
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
    MsgBox "Error restoring database: " & Err.Description, vbCritical
End Sub

' Quick backup function (for easy access)
Public Sub QuickBackup()
    BackupDatabase
End Sub

' Quick restore function (for easy access)
Public Sub QuickRestore()
    RestoreDatabase
End Sub

' Connect to PostgreSQL (consistent with other modules)
Private Function ConnectToPostgreSQL(conn As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim connectionString As String
    Dim password As String
    Dim driverOptions() As String
    Dim i As Integer
    
    ' Use constant password
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
        ' Build connection string using constants
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
    Debug.Print "Error connecting to PostgreSQL: " & Err.Description
    MsgBox "Error connecting to PostgreSQL: " & Err.Description, vbCritical
    ConnectToPostgreSQL = False
End Function
