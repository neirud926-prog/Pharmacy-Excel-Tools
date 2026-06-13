Attribute VB_Name = "mailService"
Sub eMail()

    Dim Source As Range
    Dim Dest As Workbook
    Dim wb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim OutApp As Object
    Dim OutMail As Object

    Set Source = Nothing
    On Error Resume Next
    Set Source = wsReport.Range("A1").CurrentRegion
    On Error GoTo 0

    If Source Is Nothing Then
        MsgBox "The source is not a range or the sheet is protected, please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set wb = ActiveWorkbook
    Set Dest = Workbooks.Add(xlWBATWorksheet)

    Source.Copy
    With Dest.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial Paste:=xlPasteValues
        .Cells(1).PasteSpecial Paste:=xlPasteFormats
        .Cells(1).Select
        Application.CutCopyMode = False
    End With

    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Stock Take Report -" & wb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

    If val(Application.Version) < 12 Then
        
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        
        FileExtStr = ".xlsx": FileFormatNum = 51
    End If

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Dest
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = "khw833@ha.org.hk"
            .CC = ""
            .BCC = ""
            .Subject = "Stock take Report " & VBA.Format(Date, "DD-MMM-YYYY")
             .Body = "Dear Macro, "
            .Attachments.Add Dest.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Send
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

