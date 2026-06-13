VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDrugGTIN 
   Caption         =   "Drug Info Maintenance"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   OleObjectBlob   =   "frmDrugGTIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDrugGTIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAddGTIN_Click()
    
    If Len(ItemDataMatch(Me.tbItemCode.value)) > 0 Then
        EditGTIN = False
        frmAdd.Show
    End If
    
End Sub


Private Sub cbBarCodemode_Change()
    Dim rng As Range
    Set rng = wsBarCodeSetting.Range("A:A").Find(what:=Me.tbItemCode.value)
    rng.Offset(0, 1).value = Me.cbBarCodemode.value
End Sub



Private Sub listGTIN_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Me.listGTIN.value = "" Then
        EditGTIN = True
        frmAdd.Show
    End If
End Sub


Public Sub ListGTIN_Refresh()
    Me.listGTIN.Clear
    
    Me.listGTIN.List = ListViewShowGTINDetails(Me.tbItemCode.value)
    
End Sub


Private Sub tbItemCode_Change()
    Me.tbItemCode.value = UCase(Me.tbItemCode.value)
    If Len(Me.tbItemCode) = 6 Then
        Me.listGTIN.Clear
        Me.lbDrugDes.Caption = ItemDataMatch(Me.tbItemCode, 2)
        If Me.lbDrugDes.Caption = "" Then Me.tbItemCode.value = ""
        
        If Not Me.lbDrugDes.Caption = "" Then
            Call ListGTIN_Refresh
            Me.btnAddGTIN.Enabled = True
        Else
            Me.btnAddGTIN.Enabled = False
        End If
        'Me.cbBarCodemode.Enabled = True
        'Me.cbBarCodemode.Value = Application.WorksheetFunction.VLookup(Me.tbItemCode.Value, wsBarCodeSetting.Range("A:B"), 2, 0)
    ElseIf Len(Me.tbItemCode) = 4 Then
        Me.listGTIN.Clear
        
    Else
        Me.listGTIN.Clear
        Me.lbDrugDes.Caption = ""
        Me.btnAddGTIN.Enabled = False
        EditGTIN = False
        Me.cbBarCodemode.value = ""
        Me.cbBarCodemode.Enabled = False
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.cbBarCodemode.AddItem "DM BarCode"
    Me.cbBarCodemode.AddItem "BarCode"
    
    

    EditGTIN = False
    Me.tbItemCode.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call Update_wsGTIN
End Sub
