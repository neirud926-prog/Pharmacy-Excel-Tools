Attribute VB_Name = "MQueryHandler"
Sub RefreshQuery(queryName As String)
    On Error GoTo ErrorHandler
    Dim conn As WorkbookConnection
    
    ' Check if the connection exists
    On Error Resume Next
    Set conn = ThisWorkbook.Connections(queryName)
    If conn Is Nothing Then
        MsgBox "Connection '" & queryName & "' not found in this workbook.", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Refresh the connection
    conn.Refresh
    'MsgBox "Query '" & queryName & "' refreshed successfully.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Failed to refresh query '" & queryName & "'. Error: " & err.Description, vbExclamation
End Sub

Sub test()
  Call RefreshQuery("Query - Error Handling Query")
End Sub


Sub ListConnections()
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        'Debug.Print conn.Name
    Next conn
End Sub

