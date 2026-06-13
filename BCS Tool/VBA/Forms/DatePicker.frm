VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "Date"
   ClientHeight    =   1185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2385
   OleObjectBlob   =   "DatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



 

Private Sub binSubmit_Click()

    frmApp.PrintBtnDate.Caption = VBA.Format(dpFrom.value(1), "dd-mmm-yyyy")
    Unload Me
    Call frmApp.DisplayTodayData(frmApp.PrintBtnDate.Caption)
End Sub

Private Sub cmbDateStart_Change()

End Sub

Private Sub UserForm_Initialize()
'    Call removeTudo(Me)
    Set dpFrom = New DateTimePicker
    
    With dpFrom
        .Add Me.cmbDate
        .Create Me, "DD/MM/YYYY"
        
        
    End With
End Sub


