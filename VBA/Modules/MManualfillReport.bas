Attribute VB_Name = "MManualfillReport"

Function importTransaction() As Boolean
    Dim wbNew As Workbook
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    Application.ScreenUpdating = False
    
    MsgBox "Please select a Transaction File(Material_Transaction_YYYYMMDDXXXXXXXXX.xlsx)!"
    fileToOpen = Application.GetOpenFilename()

    
    If fileToOpen = False Then
        MsgBox "No file selected"
        Exit Function
    Else
        If fso.GetExtensionName(fileToOpen) = "xlsx" And Left(fso.GetFilename(fileToOpen), 21) = "Material_Transaction_" Then
            
            wsNL0ReplenTransaction.Range("A1:Z2000").value = ""
            Application.Workbooks.Open (fileToOpen)
            Set wbNew = ActiveWorkbook
            wbNew.Sheets(1).Range("A6").CurrentRegion.Copy wsNL0ReplenTransaction.Range("A1")
            
            wbNew.Close savechanges:=False
            importTransaction = True
        Else
            importTransaction = False
            MsgBox "msgbox Please Select a Correct file"
            Exit Function
        End If
    End If
    
    Dim iRow, ItemCount As Long
    Dim ItemCode As String
    
    Call NL0ReplenTransactionResize
   
End Function

Sub NL0ReplenTransactionResize()
     Dim ResizeRng As Range
     Dim RowCount As Long
     RowCount = wsNL0ReplenTransaction.Range("A1").End(xlDown).Row
    Set ResizeRng = wsNL0ReplenTransaction.Range("A1:W" & RowCount)
    
    wsNL0ReplenTransaction.ListObjects.Add(xlSrcRange, ResizeRng, , xlYes).Name = "ReplenTranscation"
End Sub
