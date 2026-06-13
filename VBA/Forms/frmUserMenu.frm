VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserMenu 
   Caption         =   "Menu"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2565
   OleObjectBlob   =   "frmUserMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btmShowCheckingFrm_Click()
    frmChecking.Show
    Unload Me
End Sub

Private Sub btnMaintenance_Click()
    frmDrugGTIN.Show
    Unload Me
End Sub

Private Sub btnOtherReport_Click()
    frmReplenManualFillReport.Show
    Unload Me
End Sub

Private Sub btnPrintReport_Click()
    frmPrintReport.Show
    Unload Me
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub BtnRecordTypeEdit_Click()
    frmRecordAmendment.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    If HasPermission = True Then
        Me.btnOtherReport.Enabled = True
    End If
End Sub
