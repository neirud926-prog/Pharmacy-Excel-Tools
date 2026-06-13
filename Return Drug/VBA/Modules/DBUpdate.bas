Attribute VB_Name = "DBUpdate"


Public Sub LoadFileToWorksheet(filePath As String, Optional ByVal Tgrws As Worksheet, _
                               Optional ByVal importWsName As String = "Sheet1", _
                               Optional ByVal customSQL As String = "")
    Dim ws As Worksheet
    Dim conn As Object  ' ADO Connection
    Dim rs As Object    ' ADO Recordset
    Dim sql As String
    Dim fileExt As String
    Dim effectiveWsName As String
    Dim excelProps As String  ' Format-specific connection properties
    Dim i As Integer          ' Loop counter for headers
    Dim usedCustomSQL As Boolean  ' Flag for messaging
    
    ' Performance optimization: Disable screen updating and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbCritical, "Error"
        GoTo CleanUp
    End If
    
    ' Set effective import worksheet name (validate/trim; used only for default SQL)
    effectiveWsName = importWsName
    If effectiveWsName = "" Then effectiveWsName = "Sheet1"
    
    ' Set target worksheet: Use provided Tgrws or fallback to ActiveSheet
    If Tgrws Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = Tgrws
    End If
    ws.Cells.Clear  ' Clear existing data in the target worksheet
    
    fileExt = UCase(Right(filePath, 4))
    
    ' Validate supported Excel format
    If fileExt <> "XLSX" And fileExt <> ".XLS" And fileExt <> "XLSM" Then
        MsgBox "Unsupported file type: " & fileExt & ". Use .XLSX or .XLS.", vbExclamation, "Error"
        GoTo CleanUp
    End If
    
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Set format-specific properties
    If fileExt = ".XLSX" Then
        excelProps = "Excel 12.0 Xml;HDR=Yes;IMEX=1"  ' Modern XML format
    Else  ' .XLS
        excelProps = "Excel 8.0;HDR=Yes;IMEX=1"  ' Legacy binary format
    End If
    
    ' Open connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & _
              ";Extended Properties='" & excelProps & "';"
    
    ' Set SQL: Use custom if provided, else default
    usedCustomSQL = (customSQL <> "")
    If usedCustomSQL Then
        sql = customSQL
    Else
        sql = "SELECT * FROM [" & effectiveWsName & "$]"
    End If
    
    ' Execute query
    rs.Open sql, conn
    
    ' Output headers to row 1 (column names from recordset fields; adapts to selected columns)
    If rs.fields.Count > 0 Then
        For i = 0 To rs.fields.Count - 1
            ws.Cells(1, i + 1).Value = rs.fields(i).Name
        Next i
    End If
    
    ' Output data rows starting from row 2 (bulk copy for performance)
    If Not rs.EOF Then
        ws.Range("A2").CopyFromRecordset rs  ' Efficient bulk copy
    End If
    
    ' Success message (RecordCount excludes headers; notes custom SQL if used)
    'MsgBox "Data loaded successfully from " & filePath & " (" & rs.RecordCount & " data rows + headers)." & vbCrLf & _
    '       "Target Worksheet: " & ws.Name & vbCrLf & _
    '       IIf(usedCustomSQL, "Custom SQL Used", "Source Sheet: " & effectiveWsName) & " (" & rs.Fields.Count & " columns)", _
    '       vbInformation, "Success"
    
CleanUp:
    ' Clean up objects
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading file: " & Err.Description & " (Code: " & Err.Number & ")" & vbCrLf & _
           IIf(usedCustomSQL, "Check your custom SQL statement.", ""), vbCritical, "VBA Error"
    GoTo CleanUp
End Sub




Sub LoadBinshelfLocator()
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\ItemLocationXls.xls"
    Dim vfilePath As String, sql As String  ' Combined Dim for brevity
    
    vfilePath = FILE_PATH
    
    sql = "SELECT [Item Code], [Bin No#] FROM [itemLocationXls$] " & _
          "WHERE [Dispensing Method] = 'Manual' AND LEFT([Bin No#], 1) NOT IN ('C', 'D') " & _
          "ORDER BY [Bin No#] ASC"
    
    Call LoadFileToWorksheet(vfilePath, wsBinShelfLocator, , sql)
End Sub

Sub LoadEV54Locator()
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\EV54.xls"
    Dim vfilePath As String, sql As String  ' Combined Dim for brevity
    
    vfilePath = FILE_PATH
    
    sql = "SELECT [Item Code], [Bin No#] FROM [itemLocationXls$]"
    Call LoadFileToWorksheet(vfilePath, wsEV54, , sql)
End Sub

Sub LoadEV180Locator()
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\EV180.xls"
    Dim vfilePath As String, sql As String  ' Combined Dim for brevity
    
    vfilePath = FILE_PATH
    
    sql = "SELECT [Item Code], [Bin No#] FROM [itemLocationXls$]"
    Call LoadFileToWorksheet(vfilePath, wsEV180, , sql)
End Sub

Sub LoadTOSHOLocator()
    ' ADO Constants (numeric for late-binding/portability; avoids ref dependency)
    Const adUseServer As Long = 2      ' Cursor location (server-side for efficiency)
    Const adLockReadOnly As Long = 1   ' Lock type (read-only to minimize NAS conflicts)
    Const adCmdText As Long = 1        ' Command type (text SQL)
    
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\ItemLocationXls.xls"
    Dim vfilePath As String, localWbPath As String
    Dim sqlExternal As String, sqlLocal As String
    Dim rsExternal As Object, rsLocal As Object
    Dim dictLocal As Object  ' Scripting.Dictionary for fast lookups (Cas by Item Code)
    Dim ws As Worksheet
    Dim key As Variant, rowOutput As Long
    Dim connExternal As Object, connLocal As Object
    
    ' Performance: Disable updates/calc
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler
    
    ' Inputs
    vfilePath = FILE_PATH
    localWbPath = ThisWorkbook.FullName  ' Path to current workbook for ADO (works even if unsaved)
    
    ' External SQL: ATDPS filter (TOSHO alias implicit)
    sqlExternal = "SELECT [Item Code], [Bin No#] FROM [itemLocationXls$] " & _
                  "WHERE [Dispensing Method] = 'ATDPS'"
    
    ' Local SQL: All from TOSHODetail (no filter; join handles INNER)
    sqlLocal = "SELECT [Item Code], [Cas] FROM [TOSHODetail$]"
    
    ' Step 1: Query external data (Full file path in Data Source for .XLS)
    Set rsExternal = CreateObject("ADODB.Recordset")
    Set connExternal = CreateObject("ADODB.Connection")
    connExternal.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & vfilePath & _
                      ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';"
    If connExternal.State <> 1 Then Err.Raise 1001, , "Failed to open external connection."
    rsExternal.Open sqlExternal, connExternal, adUseServer, adLockReadOnly, adCmdText
    
    If rsExternal.EOF Then
        MsgBox "No external data found (ATDPS filter).", vbExclamation
        GoTo CleanUp
    End If
    
    ' Step 2: Query local data
    Set rsLocal = CreateObject("ADODB.Recordset")
    Set connLocal = CreateObject("ADODB.Connection")
    connLocal.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & localWbPath & _
                   ";Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1';"
    If connLocal.State <> 1 Then Err.Raise 1002, , "Failed to open local connection."
    rsLocal.Open sqlLocal, connLocal, adUseServer, adLockReadOnly, adCmdText
    
    If rsLocal.EOF Then
        MsgBox "No local data in [TOSHODetail].", vbExclamation
        GoTo CleanUp
    End If
    
    ' Step 3: Build Dictionary for local lookups (key: Item Code, value: Cas)
    Set dictLocal = CreateObject("Scripting.Dictionary")
    Do While Not rsLocal.EOF
        key = rsLocal.fields("Item Code").Value
        If Not IsNull(key) And key <> "" And Not dictLocal.Exists(key) Then  ' Skip nulls/duplicates (first match)
            dictLocal(key) = rsLocal.fields("Cas").Value
        End If
        rsLocal.MoveNext
    Loop
    
    ' Step 4: INNER JOIN & Output to wsTOSHO
    Set ws = wsTOSHO
    ws.Cells.Clear
    rowOutput = 2  ' Start data after headers
    
    ' Headers (fixed for three columns)
    ws.Range("A1").Value = "Item Code"
    ws.Range("B1").Value = "Cas"
    
    ' Join: Loop external, lookup local (INNER: only if exists)
    rsExternal.MoveFirst
    Do While Not rsExternal.EOF
        key = rsExternal.fields("Item Code").Value
        If Not IsNull(key) And key <> "" And dictLocal.Exists(key) Then
            ws.Cells(rowOutput, 1).Value = key  ' Item Code
            'ws.Cells(rowOutput, 2).Value = rsExternal.Fields("Bin No#").Value  ' Bin No#
            ws.Cells(rowOutput, 2).Value = dictLocal(key)  ' Cas
            rowOutput = rowOutput + 1
        End If
        rsExternal.MoveNext
    Loop
    
CleanUp:
    ' Cleanup ADO objects
    If Not rsExternal Is Nothing Then If rsExternal.State = 1 Then rsExternal.Close: Set rsExternal = Nothing
    If Not rsLocal Is Nothing Then If rsLocal.State = 1 Then rsLocal.Close: Set rsLocal = Nothing
    If Not connExternal Is Nothing Then If connExternal.State = 1 Then connExternal.Close: Set connExternal = Nothing
    If Not connLocal Is Nothing Then If connLocal.State = 1 Then connLocal.Close: Set connLocal = Nothing
    Set dictLocal = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in LoadTOSHOLocator: " & Err.Description & " (Code: " & Err.Number & ")" & vbCrLf & _
           "Tip: Ensure ItemLocationXls.xls is closed elsewhere and accessible.", vbCritical
    GoTo CleanUp
End Sub

Sub LoadTopupSMItemLocator()
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Topup SM Item Locator.XLSX"
    Dim vfilePath As String  ' Combined Dim for brevity
    
    vfilePath = FILE_PATH
    
    Call LoadFileToWorksheet(vfilePath, wsTopupSMLoc, , "")
End Sub


Sub LoadwsPPLocator()
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Prepack Item Locator.XLSX"
    Dim vfilePath As String  ' Combined Dim for brevity
    
    vfilePath = FILE_PATH
    
    Call LoadFileToWorksheet(vfilePath, wsPPLoc, , "")
End Sub

Sub LoadMSData()
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Cabinet & Mobile Shelf Label.XLSM"
    Dim vfilePath As String, sql As String  ' Combined Dim for brevity
    
    vfilePath = FILE_PATH
    
    sql = "SELECT [Item Code], [Location] FROM [Summary$]"
    Call LoadFileToWorksheet(vfilePath, wsMobileShelf, , sql)
End Sub

Sub LoadNameList()
    ' Define the file path
    Const FILE_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Namelist.xlsx"
    Dim vfilePath As String
    
    vfilePath = FILE_PATH
    
    ' Call the helper function:
    ' Param 1: File Path
    ' Param 2: Target Worksheet Object (wsNameList)
    ' Param 3: The exact name of the tab inside the Excel file ("NameList")
    Call LoadFileToWorksheet(vfilePath, wsNameList, "NameList")
End Sub
