VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFindTag 
   Caption         =   "Search"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "frmFindTag.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFindTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnFind_Click()

    If Len(Me.tbBarCode) > 0 And Len(Me.tbExp) > 0 And Len(Me.tbLot) > 0 Then

    Dim result As Variant
    result = FindCodeIndex(Me.tbBarCode.value, Me.tbLot.value, Me.tbExp.value)

    
        frmApp.AddLotStart.value = result(0)
        frmApp.AddLotEnd.value = result(1)
        
        frmApp.AddEXPStart.value = result(2)
        frmApp.AddEXPEnd.value = result(3)
        Unload Me
    Else
        MsgBox "Please enter all field before find"
    End If
End Sub
