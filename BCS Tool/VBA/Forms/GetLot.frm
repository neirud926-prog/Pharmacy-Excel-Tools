VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetLot 
   Caption         =   "GetLot"
   ClientHeight    =   1440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5985
   OleObjectBlob   =   "GetLot.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Enter()
    Me.tbBarCode.SetFocus
End Sub

Private Sub tbBarCode_AfterUpdate()
        Dim GTINStr As String
    BarCodeLength = Len(Me.tbBarCode.value)
    
    If BarCodeLength < 16 And BarCodeLength > 0 Then
        'GTIN only
        'AddZeroCount = 14 - BarCodeLength
        
        'For i = 1 To AddZeroCount
        '    GTINStr = GTINStr & "0"
        'Next
        
        'Me.lbGTIN.Caption = GTINStr & Me.tbBarCode.Value
        
    ElseIf BarCodeLength >= 16 Then
    
        Dim ResultArr As Variant
        
        ResultArr = ReadBarCode(Me.tbBarCode.value)
        Dim loc, GTIN, Batch, Exp, item As String
        
        GTIN = UCase(ResultArr(0))
        Batch = UCase(ResultArr(1))
        Exp = UCase(ResultArr(2))
        
        iRow = wsGTIN.Columns(2).Find(what:=GTIN, LookAt:=xlPart).Row

        item = wsGTIN.Cells(iRow, 1).value
        
        Dim rng As Range
        Set rng = wsCurrentLot.Range("A:A").Find(what:=item)
        rng.Offset(0, 3).value = Batch
        rng.Offset(0, 4).value = Exp
        

    
    End If
    Me.tbBarCode.value = ""
    
End Sub

Private Sub tbBarCode_Change()
    Me.tbBarCode.SetFocus
End Sub


Private Sub TextBox1_Change()
    Me.tbBarCode.SetFocus
End Sub
