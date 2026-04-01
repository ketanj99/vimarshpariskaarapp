Option Explicit

' Stock Update Module for Not In Stock Books
' This module handles stock number and date updates

Private Const DB_HOST As String = "localhost"
Private Const DB_PORT As String = "5432"
Private Const DB_NAME As String = "vimarshbooks"
Private Const DB_USER As String = "postgres"
Private Const DB_PASSWORD As String = "Ketan@757399"
Private Const DB_SCHEMA As String = "vimars"

' Add stock_number and stock_date columns to transactions table if they don't exist
Public Sub EnsureStockColumnsExist()
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim connectionString As String
    connectionString = "Driver={PostgreSQL UNICODE};Server=" & DB_HOST & ";Port=" & DB_PORT & ";Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASSWORD & ";"
    
    On Error Resume Next
    conn.Open connectionString
    If Err.Number <> 0 Then
        MsgBox "Connection failed: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Dim sql As String
    Dim rs As Object
    
    ' Check if stock_number column exists
    sql = "SELECT column_name FROM information_schema.columns " & _
          "WHERE table_schema = '" & DB_SCHEMA & "' " & _
          "AND table_name = 'transactions' " & _
          "AND column_name = 'stock_number';"
    
    Set rs = conn.Execute(sql)
    
    If rs.EOF Then
        ' Add stock_number column
        sql = "ALTER TABLE " & DB_SCHEMA & ".transactions ADD COLUMN stock_number TEXT;"
        conn.Execute sql
        Debug.Print "Added stock_number column"
    End If
    
    rs.Close
    
    ' Check if stock_date column exists
    sql = "SELECT column_name FROM information_schema.columns " & _
          "WHERE table_schema = '" & DB_SCHEMA & "' " & _
          "AND table_name = 'transactions' " & _
          "AND column_name = 'stock_date';"
    
    Set rs = conn.Execute(sql)
    
    If rs.EOF Then
        ' Add stock_date column
        sql = "ALTER TABLE " & DB_SCHEMA & ".transactions ADD COLUMN stock_date DATE;"
        conn.Execute sql
        Debug.Print "Added stock_date column"
    End If
    
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    Debug.Print "Stock columns verified/created successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error ensuring stock columns: " & Err.Description
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub

' Show Stock Update Form
Public Sub ShowStockUpdateForm()
    On Error GoTo ErrorHandler
    
    ' Ensure columns exist first
    Call EnsureStockColumnsExist
    
    ' Show form
    StockUpdateForm.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error showing stock update form: " & Err.Description, vbCritical
End Sub

' Simple macro name for easy access from Macro list (Alt+F8)
Public Sub OpenStockUpdateForm()
    Call ShowStockUpdateForm
End Sub

' Format period display (e.g., "2025-09-2" -> "16 To 30")
Public Function FormatPeriodDisplay(year As Integer, month As Integer, half As String) As String
    If half = "1" Then
        FormatPeriodDisplay = "1 To 15"
    ElseIf half = "2" Then
        ' Get last day of month
        Dim lastDay As Integer
        lastDay = Day(DateSerial(year, month + 1, 0))
        FormatPeriodDisplay = "16 To " & lastDay
    Else
        FormatPeriodDisplay = year & "-" & Right("0" & month, 2) & "-" & half
    End If
End Function

