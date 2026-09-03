VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntry 
   Caption         =   "Entry Form"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865.001
   OleObjectBlob   =   "frmEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
    If IsAllFilled And isFormatCorrect Then
        Dim iCount As Long
        iCount = Me.ListToAdd.ListCount
        Me.ListToAdd.AddItem
        Me.ListToAdd.List(iCount, 0) = Me.cbItemCode.Value
        Me.ListToAdd.List(iCount, 1) = Me.cbLoc.Value
        Me.ListToAdd.List(iCount, 2) = Me.cbLot.Value
        Me.ListToAdd.List(iCount, 3) = Me.cbUBDate.Value
        Me.ListToAdd.List(iCount, 4) = Me.cbQty.Value
        
        Me.cbItemCode.Value = ""
        Me.cbLoc.Value = ""
        Me.cbLot.Value = ""
        Me.cbUBDate.Value = ""
        Me.cbQty.Value = ""
    End If
    
    
End Sub




Private Sub btnSubmit_Click()
    Dim Target As Range
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=6
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=9
    For i = 0 To Me.ListToAdd.ListCount - 1
        Set Target = wsRecord.Range("A1").End(xlDown).Offset(1, 0)
        
        Target.Value = Date
        Target.Offset(0, 1).Value = Me.ListToAdd.List(i, 0)
        Target.Offset(0, 2).Value = Application.WorksheetFunction.VLookup(Me.ListToAdd.List(i, 0), wsLoc.Range("A:B"), 2, 0)
        Target.Offset(0, 3).Value = Me.ListToAdd.List(i, 1)
        Target.Offset(0, 4).Value = Me.ListToAdd.List(i, 2)
        Target.Offset(0, 5).Value = Me.ListToAdd.List(i, 3)
        Target.Offset(0, 6).Value = Me.ListToAdd.List(i, 4)
        Target.Offset(0, 7).Value = Environ("username")

    Next
    Unload Me
End Sub

Private Sub cbItemCode_Change()
    If Len(Me.cbItemCode) = 6 Then
    On Error Resume Next
        Me.lbItemDes.Caption = Application.WorksheetFunction.VLookup(Me.cbItemCode.Value, wsLoc.Range("A:C"), 2, 0)
        Me.cbLoc.List = ItemLocLookUp(Me.cbItemCode.Value)
    Else
        Me.lbItemDes.Caption = ""
    End If
    On Error GoTo 0
End Sub

Private Function IsAllFilled() As Boolean
    If Len(Me.cbItemCode) = 6 And Me.cbLoc.Value <> 0 And Me.cbLot.Value <> 0 And Me.cbQty.Value > 0 And Me.cbUBDate.Value <> 0 Then
        IsAllFilled = True
    Else
        IsAllFilled = False
    End If
End Function

Private Function isFormatCorrect() As Boolean
    If IsNumeric(Me.cbQty.Value) And IsDate(Me.cbUBDate.Value) Then
        isFormatCorrect = True
    Else
        isFormatCorrect = False
    End If
    
End Function



Private Sub cbUBDate_Change()
    If Len(Me.cbUBDate) = 8 And IsNumeric(Me.cbUBDate.Value) Then Me.cbUBDate.Value = Left(Me.cbUBDate.Value, 2) & "/" & MonthMatch(Mid(Me.cbUBDate.Value, 3, 2)) & "/" & Right(Me.cbUBDate.Value, 4)
End Sub

Private Function MonthMatch(ByVal num As Integer) As String
    arr = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    MonthMatch = arr(num - 1)
End Function



Private Sub lbItemDes_Click()

End Sub

Private Sub ListToAdd_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ListToAdd.RemoveItem (Me.ListToAdd.ListIndex)
End Sub



Private Sub UserForm_Initialize()
End Sub
