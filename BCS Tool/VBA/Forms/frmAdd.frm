VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdd 
   Caption         =   "Drug Info Maintenance"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10995
   OleObjectBlob   =   "frmAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAddGTIN_Click()
If Len(Me.lbGTIN.Caption) > 0 Then
    Dim rowsAffected As Long
    Dim CustomAlgoStr, AlgoStr As String
    
    If Me.cbCustomAlgo = True Then
        AlgoStr = "Custom"
        CustomAlgoStr = Me.GTINStart.value & "," & Me.GTINEnd.value & "," & Me.LotStart.value & "," & Me.LotEnd.value & "," & Me.EXPStart.value & "," & Me.EXPEnd.value
    End If
    
    If EditGTIN = False Then
        
            If Me.cbCustomAlgo = True Then
                rowsAffected = PrepareInsertSQL("GTIN", Me.lbItemCode.Caption, Me.lbGTIN.Caption, ApplyKey(Me.lbItemCode.Caption), Me.cbPackSize.value, AlgoStr, CustomAlgoStr)
            Else
                rowsAffected = PrepareInsertSQL("GTIN", Me.lbItemCode.Caption, Me.lbGTIN.Caption, ApplyKey(Me.lbItemCode.Caption), Me.cbPackSize.value)
            End If
            
        
    ElseIf EditGTIN = True Then
        If Me.cbCustomAlgo = True Then
            rowsAffected = PrepareUpdateSQL("GTIN", Me.lbGTIN.Caption, Me.lbItemCode.Caption, ApplyKey(Me.lbItemCode.Caption), Me.cbPackSize.value, AlgoStr, CustomAlgoStr)
        Else
            rowsAffected = PrepareUpdateSQL("GTIN", Me.lbGTIN.Caption, Me.lbItemCode.Caption, ApplyKey(Me.lbItemCode.Caption), Me.cbPackSize.value)
        End If
    End If
    
    MsgBox rowsAffected & " rows updated."
    
    frmDrugGTIN.tbItemCode.value = ""
    frmDrugGTIN.tbItemCode.SetFocus
    Unload Me
Else
    MsgBox "Incorrect data!"
End If
    
    
End Sub



Private Sub btnDelete_Click()
    Dim answer As Integer
    answer = MsgBox("Delete " & Me.lbItemCode.Caption & " " & Me.lbGTIN.Caption & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm?")
    If answer = vbYes Then
        rowsAffected = PrepareDeleteSQL("GTIN", Me.lbGTIN.Caption)
        MsgBox rowsAffected & " rows deleted."
        Unload Me
    End If
End Sub



Private Sub CheckBox1_Click()

End Sub

Private Sub btnHelper_Click()
    frmFindTag.Show
End Sub

Private Sub cbCustomAlgo_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.cbCustomAlgo = True Then frmpw.Show
    
End Sub



Private Sub cbCustomAlgo_Change()
    If Me.cbCustomAlgo = False Then
        Me.Width = 340
    End If
End Sub

Private Sub cbPackSize_Change()

End Sub

Private Sub CommandButton1_Click()

End Sub



Private Sub tbBarCode_AfterUpdate()
    Dim GTINStr As String
    
    If Me.cbCustomAlgo = False Then
        BarCodeLength = Len(Me.tbBarCode.value)
        
        If BarCodeLength < 16 And BarCodeLength > 0 Then
            'GTIN only
            AddZeroCount = 14 - BarCodeLength
            
            For i = 1 To AddZeroCount
                GTINStr = GTINStr & "0"
            Next
            
            Me.lbGTIN.Caption = GTINStr & Me.tbBarCode.value
        ElseIf BarCodeLength >= 16 Then
        
            Dim ResultArr As Variant
            ResultArr = ReadBarCode(Me.tbBarCode.value)
            If IsEmpty(ResultArr) Then
                Me.tbBarCode.value = ""
                Exit Sub
            End If
            Me.lbGTIN.Caption = UCase(ResultArr(0))
        
        Else
            Me.lbGTIN.Caption = ""
    
        End If
        
    ElseIf Me.cbCustomAlgo = True And Len(Me.GTINStart) > 0 And Len(Me.GTINEnd) > 0 And Len(Me.LotStart) > 0 And Len(Me.LotEnd) > 0 And Len(Me.EXPStart) > 0 And Len(Me.EXPEnd) > 0 Then
        
        Me.lbGTIN.Caption = Mid(Me.tbBarCode.value, Me.GTINStart.value, Me.GTINEnd.value - Me.GTINStart + 1)
        Me.lbBatch.Caption = Mid(Me.tbBarCode.value, Me.LotStart.value, Me.LotEnd.value - Me.LotStart + 1)
        Me.lbExpDate.Caption = Mid(Me.tbBarCode.value, Me.EXPStart.value, Me.EXPEnd.value - Me.EXPStart + 1)
        
    End If
End Sub





Private Sub tbBarCode_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If CapsLock = True Then
        MsgBox "CapsLock is On, Please turn off capslock before Scan"
        Me.tbBarCode.value = ""
    End If
End Sub

Private Sub tbBarCode_Change()
    Me.tbBarCode.value = KillSpecial(Me.tbBarCode.value)
End Sub

Private Sub UserForm_Activate()
    Me.lbItemCode.Caption = frmDrugGTIN.tbItemCode.value
    Me.lbDrugDes.Caption = frmDrugGTIN.lbDrugDes.Caption
    
    If EditGTIN = True Then
        Me.lbGTIN.Caption = frmDrugGTIN.listGTIN.List(, 1)
        
        Me.cbPackSize.value = frmDrugGTIN.listGTIN.List(, 3)
        If frmDrugGTIN.listGTIN.List(, 5) = "Custom" Then
            Me.cbCustomAlgo = True
        End If
        
        Me.btnDelete.Enabled = True
    End If
    Me.tbBarCode.SetFocus
End Sub


