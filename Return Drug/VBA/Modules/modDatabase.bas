Attribute VB_Name = "modDatabase"
Option Explicit

' --- Private constant to hold the database path for easy updating ---
Private Const DB_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Database\ReturnDrugRecord.accdb"

' --- Private function to get the Windows username ---
Private Function GetUsername() As String
    GetUsername = Environ("USERNAME")
End Function

Public Function SaveReturnRecord(ByVal returnType As String, ByVal itemCode As String, ByVal quantity As Long, ByVal refNum As String, ByVal locationCode As String, ByVal entryMethod As String) As Long
    On Error GoTo ErrorHandler
    Dim conn As Object
    Dim existingRecordId As Long, existingQuantity As Long, existingCount As Long, recordId As Long

    itemCode = UCase(itemCode) ' Safety net

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"

    ' Check if we already have this item in this Reference Number
    FindExistingRecord conn, itemCode, refNum, existingRecordId, existingQuantity, existingCount

    If existingRecordId > 0 Then
        ' ADD the new quantity to the old quantity, and ADD +1 to the scan count
        UpdateReturnQuantity conn, existingRecordId, existingQuantity + quantity, existingCount + 1
        recordId = existingRecordId
    Else
        ' New item! Set EntryCount to 1
        recordId = InsertNewReturn(conn, Now(), returnType, itemCode, quantity, refNum, locationCode, GetUsername(), "Processing", entryMethod, 1)
    End If
    
    SaveReturnRecord = recordId
    conn.Close
    Exit Function
    
ErrorHandler:
    SaveReturnRecord = 0
End Function
' --- CRUD: CREATE (Insert) a new record ---
Private Function InsertNewReturn(ByVal conn As Object, ByVal returnDateTime As Date, ByVal returnType As String, ByVal itemCode As String, ByVal quantity As Long, ByVal refNum As String, ByVal locationCode As String, ByVal dataEntryBy As String, ByVal recordStatus As String, ByVal entryMethod As String, ByVal entryCount As Long) As Long
    Dim cmd As Object, rs As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    cmd.CommandText = "INSERT INTO DrugReturns (ReturnDateTime, ReturnType, ItemCode, Quantity, ReferenceNum, LocationCode, DataEntryBy, RecordStatus, EntryMethod, EntryCount) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    cmd.Parameters.Append cmd.CreateParameter("p1", 7, 1, , returnDateTime)
    cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 50, returnType)
    cmd.Parameters.Append cmd.CreateParameter("p3", 200, 1, 20, itemCode)
    cmd.Parameters.Append cmd.CreateParameter("p4", 3, 1, , quantity)
    cmd.Parameters.Append cmd.CreateParameter("p5", 200, 1, 50, refNum)
    cmd.Parameters.Append cmd.CreateParameter("p6", 200, 1, 50, locationCode)
    cmd.Parameters.Append cmd.CreateParameter("p7", 200, 1, 100, dataEntryBy)
    cmd.Parameters.Append cmd.CreateParameter("p8", 200, 1, 50, recordStatus)
    cmd.Parameters.Append cmd.CreateParameter("p9", 200, 1, 50, entryMethod)
    cmd.Parameters.Append cmd.CreateParameter("p10", 3, 1, , entryCount) '<-- NEW!
    
    cmd.Execute
    Set rs = conn.Execute("SELECT @@IDENTITY")
    InsertNewReturn = rs(0).Value
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
End Function

Public Sub LogVerificationAttempt(ByVal itemCode As String, ByVal shelfHash As String, ByVal result As String, ByVal targetBin As String, ByVal scannedBin As String, Optional ByVal fk_ReturnID As Long = 0, Optional ByVal scanMethod As String = "Scan")
    On Error Resume Next ' If logging fails, don't stop the user's work
    
    Dim conn As Object, cmd As Object
    Set conn = CreateObject("ADODB.Connection")
    
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    cmd.CommandText = "INSERT INTO VerificationLog (ScanDateTime, ScannedBy, ScannedItemCode, ScannedShelfHash, VerificationResult, ScanMethod, TargetBin, ScannedBin, fk_ReturnID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
    
    ' Append parameters matching your schema
    cmd.Parameters.Append cmd.CreateParameter("p1", 7, 1, , Now()) ' adDate
    cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 100, GetUsername()) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p3", 200, 1, 20, itemCode) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p4", 200, 1, 50, shelfHash) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p5", 200, 1, 50, result) ' adVarChar
    
    ' ---> UPDATED: Uses the new scanMethod variable <---
    cmd.Parameters.Append cmd.CreateParameter("p6", 200, 1, 20, scanMethod) ' adVarChar
    
    cmd.Parameters.Append cmd.CreateParameter("p7", 200, 1, 50, targetBin) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p8", 200, 1, 50, scannedBin) ' adVarChar
    
    If fk_ReturnID > 0 Then
        cmd.Parameters.Append cmd.CreateParameter("p9", 3, 1, , fk_ReturnID) ' adInteger
    Else
        cmd.Parameters.Append cmd.CreateParameter("p9", 3, 1, , Null)
    End If
    
    cmd.Execute
    
    If conn.State = 1 Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub

' --- Finds an existing record based on ItemCode and ReferenceNum ---
Private Sub FindExistingRecord(ByVal conn As Object, ByVal itemCode As String, ByVal refNum As String, ByRef outRecordId As Long, ByRef outQuantity As Long, ByRef outCount As Long)
    Dim rs As Object
    Set rs = conn.Execute("SELECT ReturnID, Quantity, EntryCount FROM DrugReturns WHERE ItemCode = '" & itemCode & "' AND ReferenceNum = '" & refNum & "'")
    If Not rs.EOF Then
        outRecordId = rs.fields("ReturnID").Value
        outQuantity = rs.fields("Quantity").Value
        outCount = IIf(IsNull(rs.fields("EntryCount").Value), 1, rs.fields("EntryCount").Value)
    Else
        outRecordId = 0
        outQuantity = 0
        outCount = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub






' --- CRUD: UPDATE a record's quantity ---
Private Sub UpdateReturnQuantity(ByVal conn As Object, ByVal recordId As Long, ByVal newQuantity As Long, ByVal newCount As Long)
    conn.Execute "UPDATE DrugReturns SET Quantity = " & newQuantity & ", EntryCount = " & newCount & " WHERE ReturnID = " & recordId
End Sub



' --- CRUD: DELETE a record ---
' You can use this for a "Delete" button later
Public Sub DeleteReturnRecord(ByVal returnId As Long)
    On Error GoTo ErrorHandler
    
    Dim conn As Object 'ADODB.Connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Dim cmd As Object 'ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    cmd.CommandText = "DELETE FROM DrugReturns WHERE ReturnID = ?"
    cmd.Parameters.Append cmd.CreateParameter("p1", 3, 1, , returnId) ' adInteger
    
    cmd.Execute
    
    MsgBox "Record " & returnId & " has been deleted.", vbInformation

ExitSub:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while deleting the record:" & vbCrLf & Err.Description, vbCritical, "Database Error"
    Resume ExitSub
End Sub

Public Function GetNextReferenceNumber() As String
    On Error GoTo ErrorHandler

    Dim conn As Object 'ADODB.Connection
    Dim rs As Object   'ADODB.Recordset
    Dim sql As String
    Dim prefix As String
    Dim maxRef As String
    Dim currentNum As Long
    
    ' 1. Define the 11-character prefix for today's date (e.g., "NLTH-260323")
    prefix = "NLTH-" & Format(Date, "yymmdd")
    
    ' 2. Establish connection to the database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"

    ' 3. Query 1: Find the absolute highest ReferenceNum for today
    ' Using MAX() in SQL is much faster than looping in VBA
    sql = "SELECT MAX(ReferenceNum) AS MaxRef FROM DrugReturns WHERE ReferenceNum LIKE '" & prefix & "%'"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1 ' adOpenKeyset, adLockReadOnly

    ' Check if we found ANY records for today
    If Not rs.EOF And Not IsNull(rs.fields("MaxRef").Value) Then
        maxRef = rs.fields("MaxRef").Value
        rs.Close ' Close the recordset so we can reuse the 'rs' variable
        
        ' 4. Query 2: Check if this highest ReferenceNum is still "Processing"
        ' We count how many records with this ref number are processing
        sql = "SELECT COUNT(*) AS ProcCount FROM DrugReturns WHERE ReferenceNum = '" & maxRef & "' AND RecordStatus = 'Processing'"
        rs.Open sql, conn, 1, 1
        
        If rs.fields("ProcCount").Value > 0 Then
            ' STATUS IS PROCESSING: We return the existing Reference Number
            GetNextReferenceNumber = maxRef
        Else
            ' STATUS IS NOT PROCESSING (e.g. Printed): We need a brand new number
            ' We extract the 3-digit sequence starting at character 12
            currentNum = val(Mid(maxRef, 12))
            GetNextReferenceNumber = prefix & Format(currentNum + 1, "000")
        End If
        
    Else
        ' NO RECORDS FOR TODAY: This is the very first scan of the day!
        GetNextReferenceNumber = prefix & "001"
    End If

ExitSub:
    ' 5. Clean up objects safely
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Function

ErrorHandler:
    MsgBox "An error occurred while generating the Reference Number:" & vbCrLf & Err.Description, vbCritical, "Database Error"
    ' In case of an error, return a fallback so the app doesn't totally crash
    GetNextReferenceNumber = prefix & "ERR"
    Resume ExitSub
End Function

Public Function GetRecordsForRefNum(ByVal refNum As String) As Object
    On Error GoTo ErrorHandler
    
    Dim conn As Object 'ADODB.Connection
    Dim rs As Object   'ADODB.Recordset
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Set rs = CreateObject("ADODB.Recordset")
    ' We use a client-side cursor so we can safely disconnect the data from the database
    rs.CursorLocation = 3 ' adUseClient
    
    ' Select the columns we want to show, ordered by newest first (descending)
    sql = "SELECT ReturnID, ReturnDateTime, ItemCode, Quantity, LocationCode, RecordStatus FROM DrugReturns WHERE ReferenceNum = '" & refNum & "' ORDER BY ReturnDateTime DESC"
    
    ' Open the recordset
    rs.Open sql, conn, 3, 1 ' adOpenStatic, adLockReadOnly
    
    ' Disconnect the recordset so we can safely close the connection but keep the data in memory!
    Set rs.ActiveConnection = Nothing
    
    ' Close connection
    conn.Close
    Set conn = Nothing
    
    ' Return the data
    Set GetRecordsForRefNum = rs
    Exit Function

ErrorHandler:
    MsgBox "Error loading ListView data: " & Err.Description
    Set GetRecordsForRefNum = Nothing
End Function
Public Sub GeneratePrintReport(ByVal refNum As String)
    On Error GoTo ErrorHandler
    
    Dim conn As Object 'ADODB.Connection
    Dim rs As Object   'ADODB.Recordset
    Dim sql As String
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim itemCode As String
    Dim desc As Variant
    Dim eMethod As String, sMethod As String, vResult As String
    
    ' Trackers for the Audit Trail logic
    Dim currentReturnID As Long
    Dim previousReturnID As Long
    Dim previousRowWritten As Long
    Dim previousResult As String
    
    ' 1. Set the target worksheet
    Set ws = wsReport
    
    ' 2. Clear old data from row 8 downwards (Columns A to H)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 8 Then
        ws.Range("A8:H" & lastRow).ClearContents
        ws.Range("A8:H" & lastRow).Style = "Normal" ' Reset any old highlighting
        ws.Range("A8:H" & lastRow).Font.ColorIndex = xlAutomatic
        ws.Range("A8:H" & lastRow).Font.Italic = False
        ws.Range("A8:H" & lastRow).HorizontalAlignment = xlLeft
    End If
    
    ' 3. Connect to Database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    ' 4. Fetch records
    ' ---> UPDATED SQL: Now selecting VL.ScannedBin <---
    sql = "SELECT DR.ReturnID, DR.ReturnDateTime, DR.ItemCode, DR.Quantity, DR.LocationCode, DR.DataEntryBy, DR.EntryMethod, DR.CheckBy, " & _
          "VL.TargetBin, VL.ScannedBin, VL.VerificationResult, VL.ScanMethod " & _
          "FROM DrugReturns AS DR " & _
          "LEFT JOIN VerificationLog AS VL ON DR.ReturnID = VL.fk_ReturnID " & _
          "WHERE DR.ReferenceNum = '" & refNum & "' " & _
          "ORDER BY DR.ReturnDateTime ASC, VL.ScanDateTime ASC"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1 ' adOpenKeyset, adLockReadOnly
    
    ' 5. Loop and fill the worksheet
    If Not rs.EOF Then
        currentRow = 8 ' Starting Row
        previousReturnID = 0
        previousRowWritten = 0
        previousResult = ""
        
        Do While Not rs.EOF
            currentReturnID = rs.fields("ReturnID").Value
            itemCode = CStr(rs.fields("ItemCode").Value)
            
            ' Lookup Description
            desc = General.iLookup(itemCode, wsNLTItem.Range("A:B"), 2)
            If IsEmpty(desc) Or IsError(desc) Then desc = "Unknown Item"
            
            ' Safely extract the methods and results to check for highlighting
            eMethod = "" & rs.fields("EntryMethod").Value
            sMethod = "" & rs.fields("ScanMethod").Value
            vResult = "" & rs.fields("VerificationResult").Value
            
            ' --- AUDIT TRAIL DISPLAY LOGIC ---
            If currentReturnID = previousReturnID And previousRowWritten >= 8 Then
                
                ' 1. Tag with (Retry) ONLY if the previous scan was an actual Mismatch!
                If UCase(previousResult) = "MISMATCH" Then
                    ws.Cells(previousRowWritten, 6).Value = ws.Cells(previousRowWritten, 6).Value & " (Retry)"
                End If
                
                ' 2. Blank out the quantity to prevent accidental double-summing
                ws.Cells(previousRowWritten, 5).Value = "-"
                ws.Cells(previousRowWritten, 5).HorizontalAlignment = xlCenter
                
                ' 3. Grey out the entire row so the Final scan stands out
                ws.Range("A" & previousRowWritten & ":H" & previousRowWritten).Font.Color = RGB(150, 150, 150)
                ws.Range("A" & previousRowWritten & ":H" & previousRowWritten).Font.Italic = True
                
            End If
            
            ' Fill Columns A to H
            ws.Cells(currentRow, 1).Value = Format(rs.fields("ReturnDateTime").Value, "yyyy-mm-dd")
            ws.Cells(currentRow, 2).Value = Format(rs.fields("ReturnDateTime").Value, "hh:mm:ss")
            ws.Cells(currentRow, 3).Value = itemCode & vbLf & "(" & rs.fields("LocationCode").Value & ")"
            ws.Cells(currentRow, 4).Value = desc
            ws.Cells(currentRow, 5).Value = rs.fields("Quantity").Value
            
            ' ---> UPDATED: Fill Col F with ScannedBin <---
            ws.Cells(currentRow, 6).Value = "" & rs.fields("ScannedBin").Value      ' Col F
            ws.Cells(currentRow, 7).Value = "" & rs.fields("DataEntryBy").Value    ' Col G
            ws.Cells(currentRow, 8).Value = "" & rs.fields("CheckBy").Value        ' Col H
            
            ' --- HIGHLIGHTING LOGIC ---
            If UCase(eMethod) = "MANUAL" Or UCase(sMethod) = "MANUAL" Then
                ws.Range("A" & currentRow & ":H" & currentRow).Style = "Neutral"
            End If
            
            If UCase(vResult) = "MISMATCH" Then
                ws.Range("A" & currentRow & ":H" & currentRow).Style = "Bad"
            End If
            
            ' Update trackers before moving to the next row
            previousReturnID = currentReturnID
            previousRowWritten = currentRow
            previousResult = vResult
            
            currentRow = currentRow + 1
            rs.MoveNext
        Loop
        
        ' 6. UPDATE STATUS: Mark records as "Printed"
        sql = "UPDATE DrugReturns SET RecordStatus = 'Printed' WHERE ReferenceNum = '" & refNum & "'"
        conn.Execute sql
        
        ws.Range("F6").Value = "Reference No: " & refNum
        
        ws.Visible = xlSheetVisible
        ws.Activate
        
        MsgBox "Report generated and " & (currentRow - 8) & " records marked as 'Printed'.", vbInformation
    Else
        MsgBox "No records found for Reference Number: " & refNum, vbExclamation
    End If

ExitSub:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Report Error: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

'------------
Public Function GetItemCodeFromGTIN(ByVal barcodeVal As String) As String
    On Error GoTo ErrorHandler
    Dim conn As Object, rs As Object, repDbPath As String
    
    repDbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & repDbPath & ";"
    
    ' SQL: Searching [Hash] column, returning [Item Code] column
    Dim sql As String
    
    sql = "SELECT [Item Code] FROM GTIN WHERE [Hash] = '" & barcodeVal & "'"
    
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        ' Accessing field with space in name
        GetItemCodeFromGTIN = CStr(rs.fields("Item Code").Value)
    Else
        GetItemCodeFromGTIN = ""
    End If
    
    rs.Close: conn.Close
    Exit Function
ErrorHandler:
    GetItemCodeFromGTIN = ""
End Function

'--------------------------------Page Scan
' --- Load Reference Numbers from the last 14 days into a ComboBox ---
Public Sub LoadRecentRefNumbers(ByRef cb As Object)
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim cutoffDate As Date
    
    ' 1. Set up the ComboBox defaults
    cb.Clear
    cb.AddItem "Please Select Ref. No."
    cb.Value = "Please Select Ref. No."
    
    ' 2. Calculate the date for 2 weeks ago
    cutoffDate = Date - 14
    
    ' 3. Connect to the Database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    ' 4. SQL Query:
    ' - Get ReferenceNum
    ' - Filter for the last 14 days (Access uses # around dates)
    ' - GROUP BY groups duplicates into a single line
    ' - ORDER BY MAX(ReturnDateTime) DESC puts the newest ones at the top
    sql = "SELECT ReferenceNum FROM DrugReturns " & _
          "WHERE ReturnDateTime >= #" & Format(cutoffDate, "yyyy-mm-dd") & "# " & _
          "GROUP BY ReferenceNum " & _
          "ORDER BY MAX(ReturnDateTime) DESC"
          
    Set rs = conn.Execute(sql)
    
    ' 5. Loop through the results and add them to the dropdown
    Do While Not rs.EOF
        If Not IsNull(rs.fields("ReferenceNum").Value) Then
            cb.AddItem rs.fields("ReferenceNum").Value
        End If
        rs.MoveNext
    Loop
    
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
    Exit Sub
    
ErrorHandler:
    MsgBox "Failed to load Reference Numbers: " & Err.Description, vbCritical, "Database Error"
    Resume CleanUp
End Sub

' --- NEW: Find ReturnID based on Reference Number and Item Code ---
Public Function GetReturnID(ByVal refNum As String, ByVal itemCode As String) As Long
    On Error GoTo ErrorHandler
    
    Dim conn As Object, cmd As Object, rs As Object
    
    ' Default to 0 if we can't find a match
    GetReturnID = 0
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    ' Use parameters to prevent errors and find the exact ReturnID
    cmd.CommandText = "SELECT TOP 1 ReturnID FROM DrugReturns WHERE ReferenceNum = ? AND ItemCode = ?"
    cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 50, refNum)   ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 20, itemCode) ' adVarChar
    
    Set rs = cmd.Execute
    
    ' If a record is found, grab the ID!
    If Not rs.EOF Then
        GetReturnID = rs.fields("ReturnID").Value
    End If
    
    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Exit Function

ErrorHandler:
    ' If something goes wrong, quietly return 0 so the scan process doesn't crash
    GetReturnID = 0
End Function

' --- NEW: Update the ReturnBy field in the DrugReturns table ---
Public Sub UpdateReturnBy(ByVal returnId As Long)
    On Error GoTo ErrorHandler
    
    ' Safety Check: Only update if we have a valid ReturnID
    If returnId <= 0 Then Exit Sub
    
    Dim conn As Object, cmd As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    ' Update the ReturnBy column where the ID matches
    cmd.CommandText = "UPDATE DrugReturns SET ReturnBy = ? WHERE ReturnID = ?"
    
    ' p1: The username (Using your existing GetUsername function)
    cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 100, GetUsername()) ' adVarChar
    ' p2: The ReturnID
    cmd.Parameters.Append cmd.CreateParameter("p2", 3, 1, , returnId)           ' adInteger
    
    cmd.Execute
    
CleanUp:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    ' We silently resume cleanup here so a minor update error doesn't interrupt the user's scanning workflow
    Resume CleanUp
End Sub

'----------------------------------------------------Print Return Sheet
Sub testprint()
    
End Sub

' --- NEW: Generate a Shelving List Grouped by Station ---
' --- Generate a Shelving List Grouped by Station (With Tick Box & Highlights) ---
Public Sub GenerateShelvingList(ByVal refNum As String)
    On Error GoTo ErrorHandler
    
    Dim conn As Object 'ADODB.Connection
    Dim rs As Object   'ADODB.Recordset
    Dim sql As String
    Dim ws As Worksheet
    
    Dim currentRow As Long, tableStartRow As Long
    Dim currentStation As String, thisStation As String
    Dim itemCode As String, locCode As String
    Dim eMethod As String ' <--- NEW: Variable for Entry Method
    Dim desc As Variant
    
    ' 1. Target the worksheet you specified
    Set ws = wsToShelveItem
    
    ' Clear old data and formatting completely
    ws.Cells.Clear
    ws.Cells.Style = "Normal" ' Safety reset to ensure old highlights are gone
    
    ' 2. Set up the Top Header (Extended to E)
    ws.Range("A1").Value = "Reference No.:"
    ws.Range("B1").Value = refNum
    ws.Range("C1").Value = "Date"
    ws.Range("D1").Value = Format(Date, "yyyy-mm-dd")
    ws.Range("A1:E1").Font.Bold = True
    
    ' 3. Connect to Database
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    ' 4. Fetch the records
    ' ---> UPDATED SQL: Now selecting EntryMethod as well <---
    sql = "SELECT ItemCode, Quantity, LocationCode, EntryMethod FROM DrugReturns WHERE ReferenceNum = '" & refNum & "' ORDER BY LocationCode ASC"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1 ' adOpenKeyset, adLockReadOnly
    
    currentRow = 3
    currentStation = ""
    
    ' 5. Loop and build the tables
    If Not rs.EOF Then
        Do While Not rs.EOF
            locCode = "" & rs.fields("LocationCode").Value
            
            ' Extract the first character for the Station
            If locCode = "" Or locCode = "UNKNOWN" Then
                thisStation = "Unknown"
            Else
                thisStation = Left(locCode, 1)
            End If
            
            ' --- DETECT A NEW STATION GROUP ---
            If thisStation <> currentStation Then
                ' If this isn't the first table, draw borders for the PREVIOUS table
                If currentStation <> "" Then
                    ws.Range("A" & tableStartRow & ":E" & (currentRow - 1)).Borders.LineStyle = xlContinuous
                    currentRow = currentRow + 1 ' Leave one blank line
                End If
                
                ' Print the "Station X" Title
                ws.Cells(currentRow, 1).Value = "Station " & thisStation
                ws.Cells(currentRow, 1).Font.Bold = True
                currentRow = currentRow + 1
                
                ' Print Table Headers
                tableStartRow = currentRow
                ws.Cells(currentRow, 1).Value = "Item Code"
                ws.Cells(currentRow, 2).Value = "Description"
                ws.Cells(currentRow, 3).Value = "To Location"
                ws.Cells(currentRow, 4).Value = "Quantity"
                ws.Cells(currentRow, 5).Value = "Tick"
                
                ' Bold the headers
                ws.Range("A" & currentRow & ":E" & currentRow).Font.Bold = True
                ws.Cells(currentRow, 4).HorizontalAlignment = xlCenter
                ws.Cells(currentRow, 5).HorizontalAlignment = xlCenter
                
                currentRow = currentRow + 1
                currentStation = thisStation ' Update our tracker
            End If
            
            ' --- PRINT THE DATA ROW ---
            itemCode = CStr(rs.fields("ItemCode").Value)
            eMethod = "" & rs.fields("EntryMethod").Value ' <--- NEW: Safely get Entry Method
            
            desc = General.iLookup(itemCode, wsNLTItem.Range("A:B"), 2)
            If IsEmpty(desc) Or IsError(desc) Then desc = "Unknown Item"
            
            ws.Cells(currentRow, 1).Value = itemCode
            ws.Cells(currentRow, 2).Value = desc
            ws.Cells(currentRow, 3).Value = locCode
            ws.Cells(currentRow, 4).Value = rs.fields("Quantity").Value
            ws.Cells(currentRow, 5).Value = ""
            ws.Cells(currentRow, 4).HorizontalAlignment = xlCenter
            
            ' ---> NEW: HIGHLIGHTING LOGIC <---
            If UCase(eMethod) = "MANUAL" Then
                ws.Range("A" & currentRow & ":E" & currentRow).Style = "Neutral"
            End If
            
            currentRow = currentRow + 1
            rs.MoveNext
        Loop
        
        ' --- CLEANUP POST-LOOP ---
        If currentStation <> "" Then
            ws.Range("A" & tableStartRow & ":E" & (currentRow - 1)).Borders.LineStyle = xlContinuous
        End If
        
        ws.Columns("A:E").AutoFit
        ws.Columns("E").ColumnWidth = 8
        
        ws.Visible = xlSheetVisible
        ws.Activate
        
        MsgBox "Shelving list generated successfully!", vbInformation
    Else
        MsgBox "No records found for Reference Number: " & refNum, vbExclamation
    End If

ExitSub:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Report Error: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

'----------------------------------------Shelving Status
' --- Fetch Joined Status Data for the Check Page (Latest Scan Only) ---
Public Function GetCheckStatusData(ByVal refNum As String) As Object
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    
    ' ---> UPDATED SQL: Added DR.ReturnID so we can check progress <---
    sql = "SELECT DR.ReturnID, DR.ItemCode, DR.Quantity, DR.LocationCode, DR.DataEntryBy, DR.EntryMethod, " & _
          "VL.TargetBin, VL.VerificationResult, VL.ScannedBy, VL.ScanDateTime " & _
          "FROM DrugReturns AS DR " & _
          "LEFT JOIN (" & _
          "    SELECT * FROM VerificationLog WHERE LogID IN (" & _
          "        SELECT MAX(LogID) FROM VerificationLog GROUP BY fk_ReturnID" & _
          "    )" & _
          ") AS VL ON DR.ReturnID = VL.fk_ReturnID " & _
          "WHERE DR.ReferenceNum = '" & refNum & "' " & _
          "ORDER BY DR.ReturnDateTime ASC"
          
    rs.Open sql, conn, 3, 1 ' adOpenStatic, adLockReadOnly
    
    Set rs.ActiveConnection = Nothing
    conn.Close
    Set conn = Nothing
    Set GetCheckStatusData = rs
    Exit Function

ErrorHandler:
    MsgBox "Error loading Status Data: " & Err.Description, vbCritical, "Database Error"
    Set GetCheckStatusData = Nothing
End Function
' --- NEW: Security Check for Separation of Duties ---
Public Function IsUserAlsoReturner(ByVal refNum As String, ByVal corpId As String) As Boolean
    On Error GoTo ErrorHandler
    Dim conn As Object, rs As Object, sql As String
    
    IsUserAlsoReturner = False
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    ' Count how many items in this lot were returned by this specific CorpID
    sql = "SELECT COUNT(*) FROM DrugReturns WHERE ReferenceNum = '" & refNum & "' AND ReturnBy = '" & corpId & "'"
    
    Set rs = conn.Execute(sql)
    
    ' If the count is greater than 0, they shelved at least one item here and CANNOT counter-sign!
    If rs(0).Value > 0 Then IsUserAlsoReturner = True
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function
    
ErrorHandler:
    IsUserAlsoReturner = False
End Function
' --- NEW: Update the CheckBy field for a specific Reference Number ---
Public Sub UpdateCounterCheck(ByVal refNum As String, ByVal corpId As String)
    On Error GoTo ErrorHandler
    
    Dim conn As Object, cmd As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    ' Update all records in this lot with the Counter Sign ID
    cmd.CommandText = "UPDATE DrugReturns SET CheckBy = ? WHERE ReferenceNum = ?"
    
    ' Append parameters (p1 = CheckBy, p2 = ReferenceNum)
    cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 100, corpId) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 50, refNum)  ' adVarChar
    
    cmd.Execute
    
CleanUp:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set cmd = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving signature to database: " & Err.Description, vbCritical, "Database Error"
    Resume CleanUp
End Sub


Public Sub CheckScanProgress(ByVal returnId As Long, ByRef outRequired As Long, ByRef outDone As Long)
    On Error Resume Next
    outRequired = 1
    outDone = 0
    If returnId <= 0 Then Exit Sub
    
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    
    ' 1. Get how many times they NEED to scan it (EntryCount)
    Set rs = conn.Execute("SELECT EntryCount FROM DrugReturns WHERE ReturnID = " & returnId)
    If Not rs.EOF Then outRequired = IIf(IsNull(rs.fields("EntryCount").Value), 1, rs.fields("EntryCount").Value)
    rs.Close
    
    ' 2. Get how many times they HAVE successfully scanned it (Count of 'Match')
    Set rs = conn.Execute("SELECT COUNT(*) FROM VerificationLog WHERE fk_ReturnID = " & returnId & " AND VerificationResult = 'Match'")
    If Not rs.EOF Then outDone = rs(0).Value
    rs.Close
    
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
