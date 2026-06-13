Attribute VB_Name = "DriveUser"

Function GetDriveUser()
    Const tmp As String = "$!$"                        ' << define any temporary placeholder
    
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    Dim oExec As Object

    Set oExec = oShell.exec("""\\nltpha-nas01\PharmShare\Programme Data\bat\checkdriveuser.bat""")
    Dim s As String
    Dim sLine As String
    While Not oExec.StdOut.AtEndOfStream
        sLine = Replace(oExec.StdOut.ReadLine, tmp, " ")
        If Left(sLine, 4) = "User" Then GetDriveUser = Left(Mid(sLine, 15, 8), Len(Mid(sLine, 15, 8)) - 1)
    Wend
    
End Function



Function GetNameByCorpID(corpID As String) As String
    On Error GoTo ErrHandler
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim connString As String
    Dim escapedCorpID As String
   
    escapedCorpID = Replace(corpID, "'", "''")
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
                 ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open connString
   

    sql = "SELECT Name FROM [NameList$] WHERE [CORP ID] = '" & escapedCorpID & "'"
    Set rs = conn.Execute(sql)
   
    If Not rs.EOF Then
        GetNameByCorpID = rs.Fields("Name").value
    Else
        GetNameByCorpID = ""
    End If
   
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function
   
ErrHandler:
    MsgBox "Error: " & err.Description & " (Code: " & err.Number & ")", vbCritical
    GetNameByCorpID = ""
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Function

Sub ktest()
    Debug.Print "Drive User: " & GetDriveUser()
    'Debug.Print "Name for Drive User: " & GetNameByCorpID(GetDriveUser())
End Sub

