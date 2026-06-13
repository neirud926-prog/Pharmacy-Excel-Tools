VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    wsScannedItemReport.Visible = xlSheetVeryHidden
'    ThisWorkbook.Save
'End Sub




Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim fso As Object
    Dim vbsPath As String
    Dim vbsContent As String
    Dim tempPath As String

    ' Release the shared DB connection so the Access lock is cleared cleanly
    Call CloseSharedConn

    tempPath = ThisWorkbook.FullName
    
    ' Check if this is a temp file
    If Left(ThisWorkbook.Name, 5) = "Temp_" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next
        
        ' Create VBScript to delete the file after closing
        vbsPath = "\\nltpha-nas01\PharmShare\COMMON\TempProgram\DeleteSelf_" & Format(Now, "yyyymmdd_hhmmss") & ".vbs"
        vbsContent = "WScript.Sleep 2000" & vbCrLf & _
                     "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
                     "If fso.FileExists(""" & tempPath & """) Then" & vbCrLf & _
                     "    fso.DeleteFile """ & tempPath & """, True" & vbCrLf & _
                     "End If" & vbCrLf & _
                     "If fso.FileExists(""" & vbsPath & """) Then" & vbCrLf & _
                     "    fso.DeleteFile """ & vbsPath & """, True" & vbCrLf & _
                     "End If"
        
        With fso.CreateTextFile(vbsPath, True)
            .Write vbsContent
            .Close
        End With
        
        ' Close workbook
        Application.DisplayAlerts = False
        ThisWorkbook.Close savechanges:=False ' Already saved if user chose Yes
        Application.DisplayAlerts = True
        
        ' Run VBScript to delete the file after a 2-second delay
        Shell "wscript.exe """ & vbsPath & """", vbHide
        Set fso = Nothing
    Else
        ' If not a temp file, just close normally
        Exit Sub
    End If
End Sub

Private Sub Workbook_Open()
    If Not Environ("username") = "wfw325" Then
        Call CreateTemp
    End If

    Call wsUnProtectAll
    Call SetQuery
    Call Update_Welcomemsg
    wsScannedItemReport.Visible = xlSheetVeryHidden
    Call wsProtectAll

    ' Show the menu straight away using the last-saved cached data, so users do
    ' not have to wait for the network Power Query refresh before they can work.
    wsMenu.Activate

    ' Refresh the Power Queries in the BACKGROUND (non-blocking). The heavy
    ' Record-table load (Update_wsRecord) is no longer done on open - the report
    ' forms reload it themselves when they are opened.
    Call RefreshAllInBackground
End Sub

' Kicks off a non-blocking refresh of every workbook query so the UI stays
' responsive while the data freshens in the background. Report code that needs
' synchronous results sets BackgroundQuery = False on its own connection.
Private Sub RefreshAllInBackground()
    On Error Resume Next
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeOLEDB Then
            conn.OLEDBConnection.BackgroundQuery = True
        ElseIf conn.Type = xlConnectionTypeODBC Then
            conn.ODBCConnection.BackgroundQuery = True
        End If
    Next conn
    ThisWorkbook.RefreshAll
    On Error GoTo 0
End Sub

Private Sub SetQuery()
    If Environ("computername") = "NLTPHS0002" Then
        ThisWorkbook.Queries("itemLocation").Formula = wsQuery.Range("B3").value
        'ThisWorkbook.Queries("GTIN").Formula = wsQuery.Range("C3").Value
    ElseIf Left(Environ("computername"), 6) = "NLTPMS" Then
        ThisWorkbook.Queries("itemLocation").Formula = wsQuery.Range("B4").value
        'ThisWorkbook.Queries("GTIN").Formula = wsQuery.Range("C4").Value
    Else
        ThisWorkbook.Queries("itemLocation").Formula = wsQuery.Range("B2").value
        'ThisWorkbook.Queries("GTIN").Formula = wsQuery.Range("C2").Value
    End If
    
End Sub

Private Sub Update_Welcomemsg()
    wsMenu.Shapes.Range(Array("WelcomeMsg")).TextFrame.Characters.text = "Welcome! " & GetUserFullName()
    If GetUserFullName() = "nltpmscclose" Then
        wsMenu.Shapes.Range(Array("WelcomeMsg")).TextFrame.Characters.text = "Welcome! " & StaffDisplayName(0)
    End If
    'Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text = "Welcome"
End Sub

Private Sub CreateTemp()
    Dim fso As Object
    Dim tempPath As String
    Dim tempFileName As String
    Dim userName As String
    Dim masterPath As String
    
    If Left(ThisWorkbook.Name, 4) = "Temp" Then Exit Sub
    
    ' Define the master file path (shared network drive)
    masterPath = ThisWorkbook.FullName
    
    ' Get the username (e.g., Windows username)
    userName = Environ("USERNAME")
    If userName = "nltpmscclose" Then
        Call AskCorpID
        userName = StaffDisplayName(1)
        
    End If
    
    ' Create a unique temporary file name with username and timestamp
    tempFileName = "Temp_Bar-Code Stock Replenishment Tool" & userName & ".xlsm"
    
    ' Define temporary path (local temp folder)
    tempPath = "\\nltpha-nas01\PharmShare\COMMON\TempProgram\" & tempFileName ' e.g., C:\Users\<User>\AppData\Local\Temp
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    ' Check if temp file already exists and delete it
    If fso.FileExists(tempPath) Then
        fso.DeleteFile tempPath, True
    End If
    ' Copy the master file to the temp location
    fso.CopyFile masterPath, tempPath, True
    If err.Number <> 0 Then
        MsgBox "Permission denied or failed to create temporary copy. Error: " & err.Description, vbCritical
        Set fso = Nothing
        ThisWorkbook.Close savechanges:=False
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Open the temp file

    ' Save the current workbook as the temp file and close the original
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs tempPath
    Workbooks.Open tempPath
    'ThisWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
    

    Set fso = Nothing
End Sub

Private Sub RemoveTemp()
    If Not Environ("username") = "wfw325" Then Workbooks.Open "\\nltpha-nas01\PharmShare\Programme Data\TempRemover.xlsm"
End Sub

Sub AskCorpID()
Dim userName, displayName As String
userName = InputBox("Please enter your CORP ID:", "User Information", "")
    displayName = GetNameByCorpID(CStr(userName))
    If userName <> "" And Len(displayName) > 0 Then
        MsgBox "Hello, " & userName & "!", vbInformation, "Welcome"
        StaffDisplayName(0) = displayName
        StaffDisplayName(1) = userName
        wsMem.Range("B1").value = displayName
    Else
        MsgBox "You did not enter a correct CORP ID.", vbExclamation, "Warning"
        ThisWorkbook.Close savechanges:=False
    End If
    
End Sub
