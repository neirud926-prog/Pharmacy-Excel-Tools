VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmType 
   Caption         =   "Refill Type"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   OleObjectBlob   =   "frmType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnSave_Click()
    If Me.ob1 = True Then
        frmChecking.lbRefillType.Caption = Me.ob1.Caption
    ElseIf Me.ob2 = True Then
        frmChecking.lbRefillType.Caption = Me.ob2.Caption
    ElseIf Me.ob3 = True Then
        frmChecking.lbRefillType.Caption = Me.ob3.Caption
    Else
        MsgBox "Please Select!"
    End If
    Unload Me
    frmChecking.tbCodeScan.SetFocus
End Sub

Private Sub ob1_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.ob1 = True
End Sub


