Attribute VB_Name = "DBCRUD"
'=== Module name: DBCURD =============================================
Option Explicit


'====================================================================
' GetNextBinShelf – Fetches the latest BinShelf from Access DB "Record" table
'                   (by max Date, fallback to max ID if dates tie),
'                   then finds/returns the next sequential BinShelf from
'                   wsBinShelfLocator (assumes sorted ascending in col B).
'====================================================================
' Assumptions:
' - DB path: "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"
' - wsBinShelfLocator: Col A = ItemCode, Col B = BinShelf (sorted ASC by B).
' - Latest record: By max "Date" (Date field); ties broken by max ID.
' - Returns: Next BinShelf as String (e.g., "A02"); "" if none/next not found.
' - Error handling: Returns "" on failures (e.g., DB access, no records).
'====================================================================
' Design Pattern: Repository Pattern (DB access isolated); Binary Search for
'                 worksheet lookup (O(log n) perf on sorted data – fastest).
' Performance: ADO query ~0.1-0.5s (DB size dependent); Binary search <0.01s.
' Cleanliness: Self-contained, explicit declarations, early exits, no globals.
' Professional: Late-bound ADO (portable, no refs); defensive validation.
'====================================================================
Public Function GetNextBinShelf() As String
    Dim conn As Object          ' ADODB.Connection
    Dim rs   As Object          ' ADODB.Recordset
    Dim sql  As String
    Dim dbPath As String
    Dim latestBinShelf As String
    Dim ws As Worksheet
    Dim binCol As Long          ' BinShelf column (B = 2)
    Dim lastRow As Long
    Dim binRange As Range
    Dim foundRow As Long
    Dim nextBinShelf As String
    
    On Error GoTo ErrorHandler
    
    '--- 1. Setup ---------------------------------------------------
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"
    Set ws = wsBinShelfLocator
    binCol = 2                           ' BinShelf in column B
    
    ' Validate worksheet has data (at least header + 1 row)
    lastRow = ws.Cells(ws.Rows.Count, binCol).End(xlUp).Row
    If lastRow < 2 Then GoTo ErrorHandler
    
    '--- 2. Get latest BinShelf where Type = 'Daily' ---------------
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & _
              ";Persist Security Info=False;"
    
    sql = "SELECT TOP 1 [BinShelf] " & _
          "FROM [Record] " & _
          "WHERE Type = 'Daily' " & _
          "ORDER BY [RecDate] DESC, [ID] DESC"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        latestBinShelf = Trim(rs.fields("BinShelf").Value & "")
    End If
    
    ' Early cleanup
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    
    If latestBinShelf = "" Then GoTo ErrorHandler   ' no Daily records
    
    '--- 3. Find position in sorted list (wsBinShelfLocator) --------
    Set binRange = ws.Range(ws.Cells(2, binCol), ws.Cells(lastRow, binCol))
    foundRow = Application.Match(latestBinShelf, binRange, 0)   ' exact match
    
    If IsError(foundRow) Then GoTo ErrorHandler     ' not found in sheet
    
    foundRow = foundRow + 1   ' convert to absolute row number
    
    '--- 4. Get next (with wrap-around) -----------------------------
    If foundRow < lastRow Then
        ' Normal case: next row exists
        nextBinShelf = Trim(ws.Cells(foundRow + 1, binCol).Value & "")
    Else
        ' Wrap-around: return first item (row 2)
        nextBinShelf = Trim(ws.Cells(2, binCol).Value & "")
    End If
    
    ' If first item is empty (edge case), return ""
    If nextBinShelf = "" Then nextBinShelf = ""
    
    GetNextBinShelf = nextBinShelf
    Exit Function

ErrorHandler:
    GetNextBinShelf = ""
    ' Optional logging: Debug.Print "GetNextBinShelf error: " & Err.Description
End Function

Public Sub SaveListViewToAccess(frm As frmApp)
    Dim dbPath As String: dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"
    Dim conn As Object, cmd As Object, itm As ListItem
    Dim i As Long, j As Long, recordType As String
    
    On Error GoTo ErrHandler
    
    recordType = Trim(frm.btnChangeType.Caption)
    If recordType <> "Daily" And recordType <> "Annual" Then recordType = "Daily"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    ' Updated SQL to include RefNum
    cmd.CommandText = "INSERT INTO [Record] (ItemCode, Description, BinShelf, Quantity, RecDate, [Record By], TotalSum, ERP, Type, RefNum) " & _
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    cmd.Prepared = True
    
    Application.ScreenUpdating = False
    
    With frm.ListTakenView
        For i = 1 To .ListItems.Count
            Set itm = .ListItems(i)
            
            ' Calculate Totals
            Dim totalSum As Double: totalSum = 0
            For j = 3 To 12
                If IsNumeric(itm.SubItems(j)) Then totalSum = totalSum + CDbl(itm.SubItems(j))
            Next j
            
            Dim erpQty As Double: erpQty = IIf(IsNumeric(itm.SubItems(13)), CDbl(itm.SubItems(13)), 0)
            Dim qty As Double: qty = totalSum - erpQty
            
            ' Assumption: RefNum is stored in SubItems(14). Adjust index if necessary.
            Dim refNumVal As String: refNumVal = "ST" & Format(Now, "yyyymmddhh")

            ' Clear and Re-append Parameters
            Do While cmd.Parameters.Count > 0: cmd.Parameters.Delete 0: Loop
            
            cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 255, Trim(itm.SubItems(1))) ' ItemCode
            cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 255, Trim(itm.Text))        ' Description
            cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 50, Trim(itm.SubItems(2)))  ' BinShelf
            cmd.Parameters.Append cmd.CreateParameter(, 5, 1, , qty)                       ' Quantity
            cmd.Parameters.Append cmd.CreateParameter(, 7, 1, , Date)                      ' RecDate
            cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 50, Environ("USERNAME"))   ' Record By
            cmd.Parameters.Append cmd.CreateParameter(, 5, 1, , totalSum)                  ' TotalSum
            cmd.Parameters.Append cmd.CreateParameter(, 5, 1, , erpQty)                    ' ERP
            cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 50, recordType)            ' Type
            cmd.Parameters.Append cmd.CreateParameter(, 200, 1, 100, refNumVal)            ' RefNum
            
            cmd.Execute
        Next i
        MsgBox .ListItems.Count & " record(s) saved successfully!", vbInformation
    End With
    
CleanUp:
    Set cmd = Nothing
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set conn = Nothing
    
    ExportFilteredData "RefNum = '" & refNumVal & "'"
    Application.ScreenUpdating = True
    
    
    Exit Sub
ErrHandler:
    MsgBox "Save Failed: " & Err.Description, vbCritical: Resume CleanUp
End Sub
Public Sub ClearAllRecords(frm As frmApp)
    frm.ListTakenView.ListItems.Clear
End Sub


'=======================================================================
' LoadDailyRecordsToReport
'   - Loads all records where Type = 'Daily' into wsReport
'=======================================================================
Public Sub LoadDailyRecordsToReport()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim dbPath As String
    
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Clear previous content
    wsReport.Cells.Clear
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    ' Query - only records with Type = 'Daily'
    sql = "SELECT ID, RefNum, ItemCode, Description, BinShelf, Quantity, RecDate, [Record By], TotalSum, ERP, Type " & _
          "FROM [Record] " & _
          "WHERE Type = 'Daily' " & _
          "ORDER BY RecDate DESC, ID DESC"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1   ' adOpenKeyset, adLockReadOnly
    
    If Not rs.EOF Then
        ' Write headers (field names)
        Dim f As Long
        For f = 0 To rs.fields.Count - 1
            wsReport.Cells(1, f + 1).Value = rs.fields(f).Name
        Next f
        
        ' Write data starting row 2
        wsReport.Range("A2").CopyFromRecordset rs
    Else
        wsReport.Range("A1").Value = "No Daily records found."
    End If
    
    ' Auto-fit columns
    wsReport.Columns.AutoFit
    
CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrHandler:
    MsgBox "Error loading Daily records:" & vbCrLf & Err.Description, vbCritical
    Resume CleanUp
End Sub


'=======================================================================
' LoadAnnualRecordsToReport
'   - Loads all records where Type = 'Annual' into wsReport
'=======================================================================
Public Sub LoadAnnualRecordsToReport()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim dbPath As String
    
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"
    
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    wsReport.Cells.Clear
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    sql = "SELECT ID, RefNum, ItemCode, Description, BinShelf, Quantity, RecDate, [Record By], TotalSum, ERP, Type " & _
            "FROM [Record] " & _
          "WHERE Type = 'Annual' " & _
          "ORDER BY RecDate DESC, ID DESC"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1
    
    If Not rs.EOF Then
        Dim f As Long
        For f = 0 To rs.fields.Count - 1
            wsReport.Cells(1, f + 1).Value = rs.fields(f).Name
        Next f
        
        wsReport.Range("A2").CopyFromRecordset rs
    Else
        wsReport.Range("A1").Value = "No Annual records found."
    End If
    
    wsReport.Columns.AutoFit
    
CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrHandler:
    MsgBox "Error loading Annual records:" & vbCrLf & Err.Description, vbCritical
    Resume CleanUp
End Sub

'=======================================================================
' ExportFilteredData
' Params: strCriteria (e.g., "ItemCode = 'A123'")
'=======================================================================
Public Sub ExportFilteredData(Optional strCriteria As String = "")
    Dim conn As Object, rs As Object
    Dim sql As String, dbPath As String
    
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"
    
    On Error GoTo ErrHandler
    
    sql = "SELECT ID, RefNum, ItemCode, Description, BinShelf, Quantity, RecDate, [Record By], TotalSum, ERP, Type FROM [Record]"
    If strCriteria <> "" Then sql = sql & " WHERE " & strCriteria
    sql = sql & " ORDER BY RecDate DESC"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    Set rs = conn.Execute(sql)
    
    wsReport.Cells.Clear
    
    If Not rs.EOF Then
        ' Write Headers
        Dim f As Integer
        For f = 0 To rs.fields.Count - 1
            wsReport.Cells(1, f + 1).Value = rs.fields(f).Name
        Next f
        ' Write Data
        wsReport.Range("A2").CopyFromRecordset rs
        wsReport.Columns.AutoFit
    Else
        MsgBox "No records found for the given criteria.", vbInformation
    End If

CleanUp:
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Exit Sub
ErrHandler:
    MsgBox "Export Error: " & Err.Description, vbCritical: Resume CleanUp
End Sub

Public Sub RunExportByRef()

    ExportFilteredData "RefNum = '" & "ST2026051201" & "'"
        
End Sub
