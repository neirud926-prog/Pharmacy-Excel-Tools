Attribute VB_Name = "General"
Public EditGTIN As Boolean
Public LastGTIN As String
Sub showErrHandlingFrom()
    FrmErrorHandling.Show
End Sub
Sub showtestform()
    UserForm1.Show
End Sub
'Sub wsGTIN_Sort()
'    Call wsUnProtectAll
'    wsGTIN.Sort.SortFields.Clear
'    wsGTIN.Sort.SortFields.Add Key:=Range("A2"), _
'        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With wsGTIN.Sort
'        .SetRange wsGTIN.Range("A2", wsGTIN.Range("G2").End(xlDown))
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'    Call wsProtectAll
'End Sub
'Sub SortItemLoc()

'    wsItemLoc.ListObjects("itemLocation").Sort.SortFields.Clear
'    wsItemLoc.ListObjects("itemLocation").Sort.SortFields.Add Key:=Range("itemLocation[[#All],[Item Code]]"), SortOn:= _
'        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With wsItemLoc.ListObjects("itemLocation").Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'End Sub
Sub ShowUserMenu()
    frmUserMenu.Show
End Sub

Function GTINLookUpLoc(ByVal text As String, ByVal rng As Range, ByVal col As Integer)
    On Error GoTo Skip
    
    GTINLookUpLoc = Application.WorksheetFunction.VLookup(text, rng, col, 0)
    
Skip:
    If err.Number = 1004 Then
        MsgBox "This item has no locator!"
    End If
End Function

Sub fixt()
    Sheet7.Sort.SortFields.Clear
End Sub


Sub LanuchApp()
    frmApp.Show
End Sub
