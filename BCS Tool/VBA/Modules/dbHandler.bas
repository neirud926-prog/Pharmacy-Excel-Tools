Attribute VB_Name = "dbHandler"

'==================== Shared DB connection (performance) ====================
' Opening an ADODB connection to the Access DB on the network share is by far
' the most expensive step of every query. Instead of opening and closing a
' brand-new connection per call, we keep ONE connection open for the whole
' session and reuse it. This removes the repeated network handshake that made
' scanning slow. If anything goes wrong the connection is dropped and a fresh
' one is opened on the next call.
Private mSharedConn As Object

Public Function GetDBPath() As String
    GetDBPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
End Function

Public Function GetSharedConn() As Object
    On Error GoTo Reopen
    If Not mSharedConn Is Nothing Then
        If mSharedConn.State = 1 Then
            Set GetSharedConn = mSharedConn
            Exit Function
        End If
    End If
Reopen:
    On Error GoTo 0
    Set mSharedConn = CreateObject("ADODB.Connection")
    mSharedConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & GetDBPath() & ";Persist Security Info=False;"
    Set GetSharedConn = mSharedConn
End Function

Public Sub CloseSharedConn()
    On Error Resume Next
    If Not mSharedConn Is Nothing Then
        If mSharedConn.State = 1 Then mSharedConn.Close
    End If
    Set mSharedConn = Nothing
    On Error GoTo 0
End Sub

Sub InsertRecord(ByVal Dt As String, ByVal GTIN As String, ByVal BinShelfCode As String, ByVal loc As String, ByVal result As String, ByVal RefillType As String, ByVal RefNo As String, ByVal status As String, ByVal by As String, ByVal Lot As String, ByVal Exp As String, ByVal PackSize As String, ByVal LastLot As String, ByVal LastExp As String)
    Dim conn As Object
    Dim rs As Object
    Dim dbPath As String
    Dim sql As String
    Dim errSql As String
    Dim lastID As Long
    ' Path to your Access database
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
   
    ' Build the SQL statement for inserting into Record
    sql = "INSERT INTO [Record] ([Date and Time], [GTIN on Drug], [QR code on Binshelf], [Location], [Result], [Refill Type], [Ref No], [Status], [By], [Lot], [Exp], [Pack Size], [LastLot], [LastExp]) " & _
          "VALUES (#" & Dt & "#, '" & GTIN & "', '" & BinShelfCode & "', '" & loc & "', '" & result & "', '" & RefillType & "', '" & RefNo & "', '" & status & "', '" & by & "', '" & Lot & "', '" & Exp & "', '" & PackSize & "', '" & LastLot & "', '" & LastExp & "');"
    On Error GoTo ErrHandler
    ' Reuse the shared (persistent) connection
    Set conn = GetSharedConn()

    ' Execute the insert into Record
    conn.Execute sql
   
    ' Retrieve the last inserted ID
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT @@IDENTITY AS LastID;", conn, 3, 1 ' 3=adOpenStatic, 1=adLockReadOnly
    If Not rs.EOF Then
        lastID = rs("LastID")
    Else
        MsgBox "Error: Could not retrieve the last inserted ID.", vbCritical, "Error"
        GoTo Cleanup
    End If
    rs.Close
    Set rs = Nothing
    ' If result is "Warning: Not Match", insert into ErrHandling
    If result = "Warning: Not Match" Then
        errSql = "INSERT INTO [ErrHandling] ([FUID]) VALUES (" & lastID & ");"
        conn.Execute errSql
    End If
Cleanup:
    ' Keep the shared connection open for reuse
    Set conn = Nothing
    Exit Sub
ErrHandler:
    MsgBox "Error: " & err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    ' Drop the shared connection so a fresh one is opened next time
    Call CloseSharedConn
    Set conn = Nothing
End Sub

Sub Sample_Insertrecord()
    Call InsertRecord("6/12/2024 12:20", _
                      "THYR02 - 05012617009258", _
                      "THYR02", _
                      "2A4A", _
                      "Matched", _
                      "Out-Partient Daily Replenishment", _
                      "NLTH-241206002", _
                      "Processing", _
                      "wfw325", _
                      "na----", _
                      "na----", _
                      "test")
End Sub


Sub GetRecordDetail(ByVal RefNo As String, Optional ByVal fileSelected As Boolean)
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object   ' ADODB.Recordset
    Dim dbPath As String
    Dim sql As String
    
    
    'Dim ws As Worksheet
    'Dim nextRow As Long
    
    ' Set your target worksheet
    'Set ws = ThisWorkbook.Sheets("Sheet1") ' Change as needed
    
    ' Path to your Access database
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
    
    ' Build the SQL statement with proper quoting
    'sql = "SELECT [Date and Time], [GTIN on Drug], [QR code on Binshelf], [Location],[Result],[Refill Type], [Ref No], [Pack Size] " & _
     '     "FROM [Record] WHERE [Ref No] = '" & Replace(RefNo, "'", "''") & "';"
          
     sql = "SELECT [Date and Time], [GTIN on Drug], [QR code on Binshelf], [Location],[Result],[Refill Type], [Ref No], [Pack Size] " & _
          "FROM [Record] WHERE [Ref No] = '" & Replace(RefNo, "'", "''") & "' ORDER BY [Date and Time] ASC;"
    
    On Error GoTo ErrHandler
    
    ' Create ADO connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open connection to Access database
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & dbPath & ";" & _
              "Persist Security Info=False;"
    
    ' Create ADO Recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Open the Recordset
    rs.Open sql, conn, 3, 1 ' 3=adOpenStatic, 1=adLockReadOnly
    
    ' Check if any records were returned
    If Not rs.EOF Then
        ' Find the next empty row in the worksheet
        nextRow = 1
        
        
        ' Loop through the records and write to the worksheet
        wsReport.Range("A9:G282").value = ""
        wsReport.Range("A9:G282").Style = "Normal2"
        
        
        wsReport.Range("C6").value = rs.Fields("Refill Type").value
        
        If rs.Fields("Refill Type").value = "Replenishment from Main Store (TUE & THU)" Then
                    
                fileSelected = importTransaction
                Dim qt As QueryTable
                Set qt = wsTranscationQty.ListObjects("Transactions_Qty").QueryTable
                            
                qt.BackgroundQuery = False
                qt.Refresh
                
        End If
        
        i = 1
        Do While Not rs.EOF
            
            wsReport.Cells(8 + i, 1).value = rs.Fields("Date and Time").value
            
            If rs.Fields("Result").value = "Warning: Not Match" Then
                wsReport.Range("A" & 8 + i & ":G" & 8 + i).Style = "Bad"
                ItemCode = Left(rs.Fields("GTIN on Drug").value, 6)
            ElseIf rs.Fields("Result").value = "Manual" Then
                wsReport.Range("A" & 8 + i & ":G" & 8 + i).Style = "Neutral"
                ItemCode = rs.Fields("QR code on Binshelf").value
            Else
                ItemCode = Left(rs.Fields("GTIN on Drug").value, 6)
            End If
            
            Dim GTIN As String
            GTIN = Right(rs.Fields("GTIN on Drug").value, 14)
            
            wsReport.Cells(8 + i, 2).value = ItemCode
            Dim PackSize As String
            
            On Error Resume Next
            PackSize = GetSolPackSize(GTIN)
            On Error GoTo 0
            
            
            wsReport.Cells(8 + i, 3).value = Application.WorksheetFunction.VLookup(ItemCode, wsItemLoc.Range("A:C"), 2, 0) & " " & PackSize
            wsReport.Cells(8 + i, 7).value = rs.Fields("Location").value
            
            If fileSelected = True And rs.Fields("Refill Type").value = "Replenishment from Main Store (TUE & THU)" Then
                On Error Resume Next
                wsReport.Cells(8 + i, 6).value = Application.WorksheetFunction.VLookup(ItemCode, wsTranscationQty.Range("A:B"), 2, 0)
                On Error GoTo 0
            End If

            
            If rs.Fields("Result").value = "Warning: Not Match" Then
                wsReport.Range("A" & 8 + i & ":G" & 8 + i).Style = "Bad"
            ElseIf rs.Fields("Result").value = "Manual" Then
                wsReport.Range("A" & 8 + i & ":G" & 8 + i).Style = "Neutral"
                wsReport.Cells(8 + i, 2).value = rs.Fields("QR code on Binshelf").value
            End If
            
            'If Len(wsRecord.Cells(iRow + i - 1, 10).Value) > 0 Then
                'If Not wsRecord.Cells(iRow + i - 1, 10).Value = Application.WorksheetFunction.VLookup(itemCode, wsCurrentLot.Range("A:C"), 2, 0) Then
                    'wsReport.Range("G" & 7 + i).Style = "Good"
                'End If
            'End If
            'RESET
             PackSize = ""
            'I++
            i = i + 1
            rs.MoveNext
        Loop
        
         sqlUpdate = "UPDATE [Record] SET [Status] = 'Printed' WHERE [Ref No] = '" & Replace(RefNo, "'", "''") & "';"
       
        ' Execute the UPDATE statement
        conn.Execute sqlUpdate, , 128 ' 128=adExecuteNoRecords
        'MsgBox "Records retrieved and written to " & ws.Name, vbInformation, "Success"
    Else
        MsgBox "No records found with Ref No: " & RefNo, vbInformation, "No Records"
    End If
    
    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set ws = Nothing
End Sub

Sub TT123()

End Sub
    

Function GetProcessingRefno() As Variant
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object   ' ADODB.Recordset
    Dim dbPath As String
    Dim sql As String
    'Dim ws As Worksheet
    'Dim nextRow As Long
    Dim status As String
    status = "Processing"
    Dim colls As New Collection
    
    ' Path to your Access database
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
    
    ' Build the SQL statement with proper quoting
    sql = "SELECT [Ref No] " & _
          "FROM [Record] WHERE [Status] = '" & Replace(status, "'", "''") & "';"
    
    On Error GoTo ErrHandler
    
    ' Create ADO connection object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Open connection to Access database
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & dbPath & ";" & _
              "Persist Security Info=False;"
    
    ' Create ADO Recordset object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Open the Recordset
    rs.Open sql, conn, 3, 1 ' 3=adOpenStatic, 1=adLockReadOnly
    
    ' Check if any records were returned
    If Not rs.EOF Then
        Do While Not rs.EOF
            ' Find the next empty row in the worksheet
            On Error Resume Next
            colls.Add rs.Fields("Ref No").value, rs.Fields("Ref No").value
            On Error Resume Next
            rs.MoveNext
        Loop
        'MsgBox "Records retrieved and written to " & ws.Name, vbInformation, "Success"
    Else
        MsgBox "No records found with Ref No: " & RefNo, vbInformation, "No Records"
    End If
    
    Dim arr As Variant
    ReDim arr(colls.Count - 1) As Variant
    For i = 1 To colls.Count
        arr(i - 1) = colls(i)
    Next
    
    GetProcessingRefno = arr
    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Set ws = Nothing
    Exit Function
    
ErrHandler:
    MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set ws = Nothing
End Function



'------------------------------------------------------DB SQL-------------------------------

Function DatabaseSQL(ByVal sql As String) As Variant
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object   ' ADODB.Recordset
    Dim dbPath As String
    Dim isSelect As Boolean
    Dim result As Variant
    Dim rowsAffected As Long
    
    On Error GoTo ErrHandler
    
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
    
    ' Check if it's a SELECT query (simple check: look at first word)
    isSelect = (UCase(Left(LTrim(sql), 6)) = "SELECT")

    ' Reuse the shared (persistent) connection
    Set conn = GetSharedConn()

    If isSelect Then
        ' For SELECT queries, return a 2D array of results
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open sql, conn, 3, 1 ' adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            ' Copy recordset to a 2D variant array
            result = RecordsetToArray(rs)
        Else
            ' Return empty array if no records
            result = VBA.Array()
        End If
        
        rs.Close
        Set rs = Nothing
    Else
        ' For action queries (INSERT, UPDATE, DELETE), return rowsAffected
        conn.Execute sql, rowsAffected, 128 ' adExecuteNoRecords
        result = rowsAffected
    End If

    ' Keep the shared connection open for reuse
    Set conn = Nothing

    DatabaseSQL = result
    Exit Function

ErrHandler:
    MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "Error in DatabaseSQL"
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    ' Drop the shared connection so a fresh one is opened next time
    Call CloseSharedConn
    Set conn = Nothing
    ' Return empty or error indicator
    DatabaseSQL = CVErr(xlErrValue)
End Function

'-----------------------------------------------------------db helper Function
' Helper function to convert a recordset to a 2D array
Function RecordsetToArray(rs As Object) As Variant
    Dim fldCount As Long
    Dim recCount As Long
    Dim dataArr As Variant
    Dim i As Long, j As Long
    
    If rs.EOF And rs.BOF Then
        ' No data
        RecordsetToArray = VBA.Array()
        Exit Function
    End If
    
    ' Move to the last record to get record count
    rs.MoveLast
    recCount = rs.RecordCount
    rs.MoveFirst
    
    fldCount = rs.Fields.Count
    ReDim dataArr(1 To recCount, 1 To fldCount)
    
    For i = 1 To recCount
        For j = 1 To fldCount
            dataArr(i, j) = rs.Fields(j - 1).value
            'Debug.Print dataArr(i, j)
        Next j
        rs.MoveNext
    Next i
    
    RecordsetToArray = dataArr
End Function
'----------------------------------------------------InsertSQL---------------------
Function PrepareInsertSQL(ByVal TableName As String, ByVal ItemCode As String, ByVal GTIN As String, ByVal HASH As String, ByVal PackSize As String, Optional ByVal Algo As String, Optional ByVal CustomAlgo As String) As Long
    Dim sql As String
    
    ' Escape single quotes for text fields
    ItemCode = Replace(ItemCode, "'", "''")
    GTIN = Replace(GTIN, "'", "''")
    HASH = Replace(HASH, "'", "''")
    PackSize = Replace(PackSize, "'", "''")
    Algo = Replace(Algo, "'", "''")
    CustomAlgo = Replace(CustomAlgo, "'", "''")
    
    sql = "INSERT INTO [" & TableName & "] ([Item Code], [GTIN], [HASH], [Pack Size], [Algo], [CustomAlgo]) " & _
          "VALUES ('" & ItemCode & "', '" & GTIN & "', '" & HASH & "', '" & PackSize & "','" & Algo & "','" & CustomAlgo & "');"
    
    ' Execute and return rows affected
    PrepareInsertSQL = DatabaseSQL(sql)
End Function


'-----------------------------------PrepareUpdateSQL-------------------------------------------

Function PrepareUpdateSQL(ByVal TableName As String, ByVal GTIN As String, ByVal ItemCode As String, ByVal HASH As String, Optional ByVal PackSize As String, Optional ByVal Algo As String, Optional ByVal CustomAlgo As String) As Long
    Dim sql As String
    
    ItemCode = Replace(ItemCode, "'", "''")
    GTIN = Replace(GTIN, "'", "''")
    HASH = Replace(HASH, "'", "''")
    If Len(PackSize) > 0 Then PackSize = Replace(PackSize, "'", "''")
    Algo = Replace(Algo, "'", "''")
    CustomAlgo = Replace(CustomAlgo, "'", "''")
    
    If Len(PackSize) > 0 Then
    sql = "UPDATE [" & TableName & "] SET " & _
          "[Item Code] = '" & ItemCode & "', " & _
          "[HASH] = '" & HASH & "', " & _
          "[Pack Size] = '" & PackSize & "', " & _
          "[Algo] = '" & Algo & "', " & _
          "[CustomAlgo] = '" & CustomAlgo & "' " & _
          "WHERE [GTIN] = '" & GTIN & "';"
    Else
        sql = "UPDATE [" & TableName & "] SET " & _
          "[Item Code] = '" & ItemCode & "', " & _
          "[HASH] = '" & HASH & "', " & _
          "[Algo] = '" & Algo & "', " & _
          "[CustomAlgo] = '" & CustomAlgo & "' " & _
          "WHERE [GTIN] = '" & GTIN & "';"

    End If
    
    ' Execute and return rows affected
    PrepareUpdateSQL = DatabaseSQL(sql)
End Function

'-----------------------------------PrepareDeleteSQL-------------------------------------------

Function PrepareDeleteSQL(ByVal TableName As String, ByVal GTIN As String) As Long
    Dim sql As String
    
    GTIN = Replace(GTIN, "'", "''")
    
    sql = "DELETE FROM [" & TableName & "] WHERE [GTIN] = '" & GTIN & "';"
    
    ' Execute and return rows affected
    PrepareDeleteSQL = DatabaseSQL(sql)
End Function

'---------------------------------Sample for use-------------------------------
'**Insert Example:**

Sub TestInsert()
    Dim rowsAffected As Long
    rowsAffected = PrepareInsertSQL("GTIN", "123456", "00012345678901", "ABCDEF1234", "Large")
    MsgBox rowsAffected & " rows inserted."
End Sub
```

'**Update Example:**


Sub testupdate()
    Dim rowsAffected As Long
    rowsAffected = PrepareUpdateSQL("GTIN", "00012345678901", "654321", "ZZZZZZ9999", "Small")
    MsgBox rowsAffected & " rows updated."
End Sub


'**Delete Example:**


Sub TestDelete()
    Dim rowsAffected As Long
    rowsAffected = PrepareDeleteSQL("GTIN", "00012345678901")
    MsgBox rowsAffected & " rows deleted."
End Sub



'If you want to select rows and display them, just call `DatabaseSQL` directly with a SELECT statement:

Function ListViewShowGTINDetails(ByVal ItemCode As String)
    Dim results As Variant
    Dim i As Long, j As Long
    
    results = DatabaseSQL("SELECT [Item Code], [GTIN], [HASH], [Pack Size], [Algo], [CustomAlgo] FROM [GTIN] WHERE [Item Code] = '" & ItemCode & "';")
    
    If IsArray(results) Then
        ' Display results in Immediate Window
        'For i = LBound(results, 1) To UBound(results, 1)
        '    For j = LBound(results, 2) To UBound(results, 2)
        '        Debug.Print results(i, j),
        '    Next j
        '    Debug.Print
        'Next i
        
        ListViewShowGTINDetails = results
    Else
        MsgBox "No records found or an error occurred."
    End If
End Function

Sub testsmall()
    MsgBox ShowGTINDetailsByGTIN("06920425210164")(1, 1)
    
End Sub

Function ShowGTINDetailsByGTIN(ByVal GTIN As String)
    Dim results As Variant
    Dim i As Long, j As Long
    
    results = DatabaseSQL("SELECT [Item Code], [GTIN], [HASH], [Pack Size], [Algo], [CustomAlgo] FROM [GTIN] WHERE [GTIN] = '" & GTIN & "';")
    
    On Error GoTo err
    If IsArray(results) Then
        ' Display results in Immediate Window
        
            For j = LBound(results, 2) To UBound(results, 2)
                'Debug.Print results(1, j)
            Next j
            
        
        
        ShowGTINDetailsByGTIN = results
    Else
        MsgBox "No records found or an error occurred."
    End If
    
err:
    If err.Number = 9 Then
        'MsgBox "No records found or an error occurred."
    End If
End Function

Sub Update_wsGTIN()
    
    Dim results As Variant
    Dim i As Long, j As Long
    
    results = DatabaseSQL("SELECT * FROM [GTIN];")
    
    If IsArray(results) Then
        wsGTIN.Range("A1:F20000").value = ""
        wsGTIN.Range("A1:F" & UBound(results)) = results
        
    Else
        MsgBox "No records found or an error occurred."
    End If

End Sub

Sub Update_wsRecord()
    
    Dim results As Variant
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim isNARow As Boolean
    
    ' Set reference to the worksheet
    Set ws = wsRecord
    
    ' Get data from database
    results = DatabaseSQL("SELECT [Date and Time], [GTIN on Drug], [QR code on Binshelf], [Location], [Result], [Refill Type], [Ref No], [Status], [By], [Lot], [Exp] FROM [Record];")
    
    If IsArray(results) Then
        ' Clear existing data
        ws.Range("A1:K100000").value = ""
        ws.Range("A1:K100000").Clear
        
        ' Write data to worksheet
        ws.Range("A1:K" & UBound(results) + 1).value = results
        
        ' Sort data by column G (6th column)
        Set rng = ws.Range("A1:K" & UBound(results) + 1)
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("G1:G" & UBound(results) + 1), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal
            .SetRange rng
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' Find and remove row with all #NA values
        lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        For i = lastRow To 1 Step -1
            isNARow = True
            For j = 1 To 11 ' Columns A to K
                If ws.Cells(i, j).text <> "#N/A" Then
                    isNARow = False
                    Exit For
                End If
            Next j
            If isNARow Then
                ws.Rows(i).Delete
                Exit For ' Exit after deleting the first matching row
            End If
        Next i
        
    Else
        MsgBox "No records found or an error occurred."
    End If

End Sub


Sub GetRecordByRefNo(ByVal RefNo As String)
    Dim results As Variant
    Dim i As Long, j As Long
    
    results = DatabaseSQL("SELECT [Date and Time], [GTIN on Drug], [QR code on Binshelf], [Location], [Result], [Refill Type], [Ref No], [Status], [By], [Lot], [Exp] FROM [Record] WHERE [Ref No] = '" & RefNo & "';")
    
    If IsArray(results) Then
        wsRecordRetrived.Range("A2:K100000").value = ""
        wsRecordRetrived.Range("A2:K100000").Clear
        wsRecordRetrived.Range("A2:K" & UBound(results) + 1) = results
        
    Else
        MsgBox "No records found or an error occurred."
    End If
    
End Sub

Sub GetRecordByTodayReplenType(Optional ByVal tgrdate As String)
    Dim results As Variant
    Dim sqlStr As String
    Dim todayStr As String
    
    ' Format today's date for SQL (e.g., #2026-05-28#)
    todayStr = "#" & Format(Date, "yyyy-mm-dd") & "#"
    If tgrdate <> "" Then todayStr = "#" & Format(tgrdate, "yyyy-mm-dd") & "#"
    
    ' Construct the SQL query focusing only on Refill Type and Today's Date
    sqlStr = "SELECT [Date and Time], [GTIN on Drug], [QR code on Binshelf], [Location], [Result], [Refill Type], [Ref No], [Status], [By], [Lot], [Exp] " & _
             "FROM [Record] " & _
             "WHERE [Refill Type] = 'Replenishment from Main Store (TUE & THU)' " & _
             "AND DateValue([Date and Time]) = " & todayStr & ";"
    
    results = DatabaseSQL(sqlStr)
    
    If IsArray(results) Then
        wsRecordRetrived.Range("A2:K100000").Clear
        wsRecordRetrived.Range("A2:K" & UBound(results, 1) + 1).value = results
    Else
        MsgBox "No records found or an error occurred."
    End If
End Sub



Function GetRecordTypeByRefNo(ByVal RefNo As String)
    On Error GoTo err:
    Dim results As Variant
    Dim i As Long, j As Long
    
    results = DatabaseSQL("SELECT [Refill Type] FROM [Record] WHERE [Ref No] = '" & RefNo & "';")
    
    If IsArray(results) Then
        GetRecordTypeByRefNo = results(1, 1)
    Else
        MsgBox "No records found or an error occurred."
    End If
    
err:
    'MsgBox "Incorrect Ref Num or this Ref Num not exist"
    
End Function

Function PrepareUpdateSQL_ForUpdateType(ByVal TableName As String, ByVal RefNo As String, ByVal RefillType As String) As Long
    Dim sql As String
    
    RefNo = Replace(RefNo, "'", "''")
    RefillType = Replace(RefillType, "'", "''")
    
    sql = "UPDATE [" & TableName & "] SET " & _
          "[Refill Type] = '" & RefillType & "' " & _
          "WHERE [Ref No] = '" & RefNo & "';"
    
    
    ' Execute and return rows affected
    PrepareUpdateSQL_ForUpdateType = DatabaseSQL(sql)
End Function


'------------------------------------------------------------------------------------------
Function PrepareUpdateSQL_ForUpdateAlgo(ByVal TableName As String, ByVal GTIN As String, ByVal AlgoNum As String) As Long
    Dim sql As String
    
    AlgoNum = Replace(AlgoNum, "'", "''")
    GTIN = Replace(GTIN, "'", "''")
    sql = "UPDATE [" & TableName & "] SET " & _
          "[Algo] = '" & AlgoNum & "' " & _
          "WHERE [GTIN] = '" & GTIN & "';"
    
    
    ' Execute and return rows affected
    PrepareUpdateSQL_ForUpdateAlgo = DatabaseSQL(sql)
End Function



'------------------------------------------------------------------------------------------
'Debug or Fix value Area
Function PrepareFixUpdate()
    Dim sql As String
    
    TableName = Replace("Record", "'", "''")
    
    Dt = Replace("16/12/2024 13:00:00", "'", "''")
    RefNo = Replace("NLTH-241216002", "'", "''")
    
    
    sql = "UPDATE [" & TableName & "] SET " & _
          "[Ref No] = '" & RefNo & "', " & _
          "WHERE [Date and Time] => '" & Dt & "';"
    
    ' Execute and return rows affected
    PrepareFixUpdate = DatabaseSQL(sql)
End Function


Function LookUpItemCode(ByVal vgtin As String)

    GTIN = Replace(vgtin, "'", "''")

    Dim results As Variant
    Dim i As Long, j As Long
    
    results = DatabaseSQL("SELECT [GTIN], [Item Code] FROM [GTIN] WHERE [HASH] = '" & GTIN & "';")
    
    If IsArray(results) Then
        
        'Debug.Print (LBound(results))
        
       LookUpItemCode = results(1, 2)
 
    Else
        MsgBox "No records found or an error occurred."
    End If

End Function

Function LoadTodayData(Optional ByVal newdate As String) As Variant
    Dim sql As String
    Dim today, thedate, nextDate As String
    today = Format(Now, "dd/mm/yyyy")
   
    If newdate = "" Then
        thedate = today
    Else
        thedate = Format(CDate(newdate), "dd/mm/yyyy")
    End If
   
    ' Convert to MM/DD/YYYY for Access date format
    idate = Format(CDate(thedate), "mm/dd/yyyy")
    jDate = Format(DateAdd("d", 1, thedate), "mm/dd/yyyy")
   
    sql = "SELECT [Ref No], [Status], [By], [Refill Type], SUM(1) AS [Item Counted] " & _
          "FROM [Record] WHERE [Date and Time] >= #" & idate & "# AND [Date and Time] < #" & jDate & "# " & _
          "GROUP BY [Ref No], [Status], [By], [Refill Type];"
    LoadTodayData = DatabaseSQL(sql)
End Function


'------------------For Home page---------
Function CountUniqueRefNoToday() As Long
    Dim sql As String
    Dim today As String
    Dim nextDate As String
    Dim result As Variant
    today = Format(Date, "mm/dd/yyyy") ' 10/08/2025
    nextDate = Format(DateAdd("d", 1, Date), "mm/dd/yyyy") ' 10/09/2025
    sql = "SELECT COUNT(*) FROM (SELECT DISTINCT [Ref No] FROM [Record] WHERE [Date and Time] >= #" & today & "# AND [Date and Time] < #" & nextDate & "#);"
    result = DatabaseSQL(sql)
    If IsArray(result) Then
        CountUniqueRefNoToday = result(1, 1)
    Else
        CountUniqueRefNoToday = 0 ' Default to 0 on error or no data
    End If
End Function

Function TotalLinesRecordToday() As Long
    Dim sql As String
    Dim today As String
    Dim nextDate As String
    Dim result As Variant
    today = Format(Date, "mm/dd/yyyy") ' 10/08/2025
    nextDate = Format(DateAdd("d", 1, Date), "mm/dd/yyyy") ' 10/09/2025
    sql = "SELECT COUNT(*) FROM [Record] WHERE [Date and Time] >= #" & today & "# AND [Date and Time] < #" & nextDate & "#;"
    result = DatabaseSQL(sql)
    If IsArray(result) Then
        TotalLinesRecordToday = result(1, 1)
    Else
        TotalLinesRecordToday = 0 ' Default to 0 on error or no data
    End If
End Function

Function CountLinesErrHandlingCheckByNull() As Long
    Dim sql As String
    Dim today As String
    Dim nextDate As String
    Dim result As Variant
    today = Format(Date, "mm/dd/yyyy") ' 10/08/2025
    nextDate = Format(DateAdd("d", 1, Date), "mm/dd/yyyy") ' 10/09/2025
    sql = "SELECT COUNT(*) FROM [ErrHandling] E INNER JOIN [Record] R ON E.FUID = R.ID WHERE R.[Date and Time] >= #" & today & "# AND R.[Date and Time] < #" & nextDate & "# AND E.[CheckBy] IS NULL;"
    result = DatabaseSQL(sql)
    If IsArray(result) Then
        CountLinesErrHandlingCheckByNull = result(1, 1)
    Else
        CountLinesErrHandlingCheckByNull = 0 ' Default to 0 on error or no data
    End If
End Function

Function CountLinesErrHandlingCheckByNotNull() As Long
    Dim sql As String
    Dim today As String
    Dim nextDate As String
    Dim result As Variant
    today = Format(Date, "mm/dd/yyyy") ' 10/08/2025
    nextDate = Format(DateAdd("d", 1, Date), "mm/dd/yyyy") ' 10/09/2025
    sql = "SELECT COUNT(*) FROM [ErrHandling] E INNER JOIN [Record] R ON E.FUID = R.ID WHERE R.[Date and Time] >= #" & today & "# AND R.[Date and Time] < #" & nextDate & "# AND E.[CheckBy] IS NOT NULL;"
    result = DatabaseSQL(sql)
    If IsArray(result) Then
        CountLinesErrHandlingCheckByNotNull = result(1, 1)
    Else
        CountLinesErrHandlingCheckByNotNull = 0 ' Default to 0 on error or no data
    End If
End Function
Function AllCountLinesErrHandlingCheckByNull() As Long
    Dim sql As String
    Dim result As Variant
    sql = "SELECT COUNT(*) FROM [ErrHandling] E INNER JOIN [Record] R ON E.FUID = R.ID WHERE E.[CheckBy] IS NULL;"
    result = DatabaseSQL(sql)
    If IsArray(result) Then
        AllCountLinesErrHandlingCheckByNull = result(1, 1)
    Else
        CAllCountLinesErrHandlingCheckByNull = 0 ' Default to 0 on error or no data
    End If
End Function

'Report-------------------------------------

Function GetShelvingFrequencyLast30Days() As Variant
    Dim sql As String
    Dim today As String
    Dim nextDate As String
    today = Format(Date, "mm/dd/yyyy") ' 10/09/2025
    nextDate = Format(DateAdd("d", 1, Date), "mm/dd/yyyy") ' 10/10/2025
    sql = "SELECT Record.[QR code on Binshelf], Count(Record.[QR code on Binshelf]) AS ShelvingCount " & _
          "FROM Record " & _
          "WHERE (((Record.[Date and Time])>=DateAdd('d',-30,Date()) And (Record.[Date and Time])<DateAdd('d',1,Date())) " & _
          "AND ((Record.Result)='Matched' Or (Record.Result)='Manual')) " & _
          "GROUP BY Record.[QR code on Binshelf] " & _
          "ORDER BY Count(Record.[QR code on Binshelf]) DESC;"
    GetShelvingFrequencyLast30Days = DatabaseSQL(sql)
End Function

Sub LoadShelvingReportToWorksheet()
    Dim results As Variant
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
   
    Set ws = wsShelvingReport

    results = GetShelvingFrequencyLast30Days()
   
    ws.Range("A1:B" & ws.Rows.Count).ClearContents
   
    If IsArray(results) Then
        ws.Range("A1").value = "QR Code on Binshelf"
        ws.Range("B1").value = "Shelving Count"
        ws.Range("C1").value = "Gen. on: " & Date
        
        lastRow = UBound(results, 1)
        For i = 1 To lastRow
            ws.Cells(i + 1, 1).value = results(i, 1) ' [QR code on Binshelf]
            ws.Cells(i + 1, 2).value = results(i, 2) ' ShelvingCount
        Next i
    Else
        MsgBox "No data found or an error occurred."
    End If
End Sub

Function GetShelvingLastDetails() As Variant
    Dim sql As String
    Dim today As String
    Dim thirtyDaysAgo As String
    
        sql = "SELECT Record.[QR code on Binshelf], Last(Record.Lot) AS LotOfLast, Last(Record.Exp) AS ExpOfLast, Last(Record.[Date and Time]) AS [Date and TimeOfLast] " & _
              "FROM Record " & _
              "GROUP BY Record.[QR code on Binshelf] " & _
              "HAVING Last(Record.Lot) IS NOT NULL AND Last(Record.Lot) <> '' AND Last(Record.Lot) <> 'na----';"
   
    
    GetShelvingLastDetails = DatabaseSQL(sql)
End Function

Function GetLastLotAndExp(ByVal binshelf As String) As Variant
    Dim sql As String
    Dim result As Variant
    binshelf = Replace(binshelf, "'", "''") ' Escape single quotes
    sql = "SELECT TOP 1 Lot, Exp FROM [Record] WHERE [QR code on Binshelf] = '" & binshelf & "' ORDER BY [Date and Time] DESC;"
    result = DatabaseSQL(sql)
    On Error Resume Next
    If IsArray(result) Then
        GetLastLotAndExp = Array(result(1, 1), result(1, 2))
    Else
        GetLastLotAndExp = Array("", "") ' Default to empty if no data
    End If
    On Error GoTo 0
End Function


