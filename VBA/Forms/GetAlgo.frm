VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetAlgo 
   Caption         =   "GetAlgo"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6465
   OleObjectBlob   =   "GetAlgo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetAlgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton1_Enter()
    Me.tbBarCode.SetFocus
End Sub

Private Sub tbBarCode_AfterUpdate()
    AlgoNum = FindAlgo(Me.tbBarCode.value)
    Me.Algo.Caption = "Algo = " & AlgoNum
    Dim arr As Variant
    arr = ReadBarCode(Me.tbBarCode.value)
    
    Me.v1.Caption = arr(0)
    Me.v2.Caption = arr(1)
    Me.v3.Caption = arr(2)
    
    rowsAffected = PrepareUpdateSQL_ForUpdateAlgo("GTIN", arr(0), AlgoNum)
    Me.Rowupdated.Caption = rowsAffected & " rows updated."
    Me.tbBarCode.value = ""
End Sub

Private Sub tbBarCode_Change()
    Me.tbBarCode.SetFocus
End Sub


Private Sub TextBox1_Change()
    Me.tbBarCode.SetFocus
End Sub
