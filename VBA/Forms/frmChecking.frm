VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChecking 
   Caption         =   "Replenishment Checking"
   ClientHeight    =   8385.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445.001
   OleObjectBlob   =   "frmChecking.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnAdd_Click()
    Dim RefNum, NewRefNum, StrNum As String
    Dim num As Integer
    RefNum = Me.lb_RefNum.Caption
    
    num = Right(RefNum, 3) + 1
    
    StrNum = CStr(num)
    While Len(StrNum) < 3
        StrNum = "0" & StrNum
    Wend
    
    NewRefNum = Left(RefNum, 11) & StrNum
    
    Me.lb_RefNum.Caption = NewRefNum
    
    If num > 1 Then
        Me.btnMinus.Visible = True
    
    End If
End Sub

Private Sub btnMinus_Click()
    Dim RefNum, NewRefNum, StrNum As String
    Dim num As Integer
    RefNum = Me.lb_RefNum.Caption
    
    num = Right(RefNum, 3) - 1
    
    StrNum = CStr(num)
    While Len(StrNum) < 3
        StrNum = "0" & StrNum
    Wend
    
    NewRefNum = Left(RefNum, 11) & StrNum
    
    Me.lb_RefNum.Caption = NewRefNum
    
    If num = 1 Then
        Me.btnMinus.Visible = False
    
    End If
End Sub

Private Sub btnChangeNum_Click()
    Dim sInput As String
    sInput = InputBox("Please enter the name of Reference No. : ")
    
    If Left(sInput, 11) = "NLTH-" & Format(Date, "yymmdd") Then
        lb_RefNum = sInput
    Else
        MsgBox "Change failed"
            
    End If
End Sub

Private Sub btnChangeType_Click()
    frmType.Show
End Sub



Private Sub btnPrintReport_Click()
    frmApp.MutiPage.value = 2
    Unload Me
End Sub


Private Sub btnSubmit_Enter()
    Me.tbCodeScan.SetFocus
End Sub



Private Sub frResult_Enter()
    Me.tbCodeScan.SetFocus
End Sub









Function old_GetRefNum()
    Get_RefNum = 1
    'MsgBox wsRecord.Cells(1048576, 1).End(xlUp).Value
    If DateDiff("d", wsRecord.Cells(1048576, 1).End(xlUp).value, Date) = 0 Then
        
        If wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 7).value = "Processing" Then
       
            If wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 5).value = "Out-Patient Daily Replenishment" And _
                Me.lbRefillType.Caption = "Replenishment from Main Store (TUE & THU)" Then
                
                Get_RefNum = Right(wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 6).value, 1) + 1
                
                    
            ElseIf wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 5).value = "Replenishment from Main Store (TUE & THU)" And _
                Me.lbRefillType.Caption = "Out-Patient Daily Replenishment" Then
                Get_RefNum = Right(wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 6).value, 1) + 1
            Else
                Get_RefNum = Right(wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 6).value, 1)
            End If
            'MsgBox "Processing"
            
            
        ElseIf wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 7).value = "Printed" Then
            'MsgBox "Printed"
            Get_RefNum = Right(wsRecord.Cells(1048576, 1).End(xlUp).Offset(0, 6).value, 1) + 1
        End If
    End If
End Function


Function Get_RefNum() As Long
    Dim sql As String
    Dim today As String
    Dim nextDate As String
    Dim result As Variant
    Dim lastRefNo As String
    Dim baseNum As Long
    
    ' Set today¡¦s date for filtering
    today = Format(Date, "mm/dd/yyyy") ' 10/09/2025
    nextDate = Format(DateAdd("d", 1, Date), "mm/dd/yyyy") ' 10/10/2025
    
    ' Query the latest Ref No for today from the database
    sql = "SELECT TOP 1 [Ref No], [Status], [Refill Type] FROM [Record] " & _
          "WHERE [Date and Time] >= #" & today & "# AND [Date and Time] < #" & nextDate & "# " & _
          "ORDER BY [Ref No] DESC;"
    result = DatabaseSQL(sql)
    
    If IsArray(result) And UBound(result, 1) > 0 Then
        lastRefNo = result(1, 1) ' First column is [Ref No]
        Dim lastStatus As String
        Dim lastRefillType As String
        lastStatus = result(1, 2) ' Second column is [Status]
        lastRefillType = result(1, 3) ' Third column is [Refill Type]
        
        ' Check if the last record is from today
            
            If lastStatus = "Processing" Then
                If (lastRefillType = "Out-Patient Daily Replenishment" And Me.lbRefillType.Caption = "Replenishment from Main Store (TUE & THU)") Or _
                   (lastRefillType = "Replenishment from Main Store (TUE & THU)" And Me.lbRefillType.Caption = "Out-Patient Daily Replenishment") Then
                    ' Extract the last digit and increment
                    baseNum = Val(Right(lastRefNo, 1)) + 1
                Else
                    baseNum = Val(Right(lastRefNo, 1))
                End If
            ElseIf lastStatus = "Printed" Then
                baseNum = Val(Right(lastRefNo, 1)) + 1
            Else
                baseNum = 1 ' Default if no matching status
            End If

    Else
        baseNum = 1 ' Default to 1 if no data or error
    End If
    
    Get_RefNum = baseNum
End Function



Private Sub Label7_Click()
        ref = InputBox("")
        Me.lb_RefNum.Caption = "ref"
End Sub

'lbBarCode
Private Sub tbCodeScan_AfterUpdate()
    Call wsUnProtectAll
    Dim iRow As Long
    If Me.tbCodeScan.value = "#ManualFill" Then
        Me.lbGTIN.Caption = "GTINNA-Manual"
    ElseIf Mid(Me.tbCodeScan.value, 12, 2) = "_B" And Right(Me.tbCodeScan.value, 2) = "aR" Then
        Me.lbQRCode.Caption = LookUpItemCode(Mid(Me.tbCodeScan.value, 13))
    ElseIf Len(Me.tbCodeScan.value) = 19 And Left(Me.tbCodeScan.value, 1) = "B" And Right(Me.tbCodeScan.value, 2) = "aR" Then
        Me.lbQRCode.Caption = LookUpItemCode(Me.tbCodeScan.value)
    ElseIf Len(Me.tbCodeScan.value) > 7 Then
        Me.lbBarCode.Caption = Me.tbCodeScan.value
        LastGTIN = Me.tbCodeScan.value
        

          Dim ResultArr As Variant
            Dim GTIN As String
            ResultArr = ReadBarCode(Me.tbCodeScan.value)
            If IsEmpty(ResultArr) Then
                Me.tbCodeScan.value = ""
                Exit Sub
            End If
            
            GTIN = ResultArr(0)
            
            Dim GtinResult As Variant
            GtinResult = ShowGTINDetailsByGTIN(GTIN)
        
        
        Me.lbGTIN.Caption = GtinResult(1, 1) & " - " & GtinResult(1, 2)
    End If
   
    ' Show Location
    If Len(Me.lbGTIN.Caption) > 10 And Right(Me.lbGTIN.Caption, 6) <> "Manual" Then
        Me.lbMatchResult.Caption = GTINLookUpLoc(Left(Me.lbGTIN.Caption, 6), wsItemLoc.Range("A:C"), 3)
        Me.lbMatchResult.Font.Size = 120
    Else
        Me.lbMatchResult.Caption = ""
    End If
   
    ' Checking
    Dim ResultMsg As String
    If Len(Me.lbQRCode.Caption) = 6 And Len(Me.lbGTIN.Caption) > 10 Then
        If Me.lbQRCode.Caption = Left(Me.lbGTIN.Caption, 6) Then
            ResultMsg = "Matched"
            Me.frResult.BackColor = &HC0FFC0
            Me.lbMatchResult.BackColor = &HC0FFC0
            Me.lbMatchResult.Font.Size = 120
            Me.lbMatchResult.Caption = Application.WorksheetFunction.VLookup(Left(Me.lbGTIN.Caption, 6), wsItemLoc.Range("A:C"), 3, 0)
        ElseIf Right(Me.lbGTIN.Caption, 6) = "Manual" Then
            ResultMsg = "Manual"
            Me.frResult.BackColor = &H80FFFF
            Me.lbMatchResult.BackColor = &H80FFFF
            Me.lbMatchResult.Font.Size = 108
            Me.lbMatchResult.Caption = Application.WorksheetFunction.VLookup(Me.lbQRCode.Caption, wsItemLoc.Range("A:C"), 3, 0)
        Else
            ' set size
            ResultMsg = "Warning: Not Match"
            Me.lbMatchResult.Caption = ResultMsg
            Me.frResult.BackColor = &HFF&
            Me.lbMatchResult.BackColor = &HFF&
            Me.lbMatchResult.Font.Size = 60
            Call NotMatchAlert(Left(Me.lbGTIN.Caption, 6))
        End If
        
        ' Get the last Lot and Exp for the Binshelf (previous to this insert)
        Dim binshelf As String
        binshelf = Me.lbQRCode.Caption
        Dim last As Variant
        last = GetLastLotAndExp(binshelf)
        
        ' Prepare parameters
        Dim param(0 To 13) As String
        param(0) = Format(Now, "dd-mmm-yyyy hh:mm")
        param(1) = Me.lbGTIN.Caption
        param(2) = Me.lbQRCode.Caption
        If ResultMsg = "Warning: Not Match" Then
            param(3) = Application.WorksheetFunction.VLookup(Left(Me.lbGTIN.Caption, 6), wsItemLoc.Range("A:C"), 3, 0) & "(" & Application.WorksheetFunction.VLookup(Me.lbQRCode.Caption, wsItemLoc.Range("A:C"), 3, 0) & ")"
        ElseIf ResultMsg = "Manual" Then
            param(3) = Application.WorksheetFunction.VLookup(Me.lbQRCode.Caption, wsItemLoc.Range("A:C"), 3, 0) & "(Manual)"
        Else
            param(3) = Application.WorksheetFunction.VLookup(Left(Me.lbGTIN.Caption, 6), wsItemLoc.Range("A:C"), 3, 0)
        End If
        param(4) = ResultMsg
        If GetUserFullName() = "nltpmscclose" Then
            param(8) = wsMem.Range("B1").value
        ElseIf GetUserFullName() = "User" Then
            param(8) = wsMem.Range("B1").value
        Else
            param(8) = GetUserFullName()
        End If
        param(5) = Me.lbRefillType.Caption
        param(6) = Me.lb_RefNum.Caption
        param(7) = "Processing"
        
        Dim DataResult As Variant
        DataResult = ReadBarCode(LastGTIN)
        Dim PackSz As String
        PackSz = ""
        On Error Resume Next
        PackSz = GetSolPackSize(DataResult(0))
        param(9) = DataResult(1)
        param(10) = DataResult(2) ' exp
        param(11) = PackSz ' Pack Size
        param(12) = last(0) ' LastLot (previous)
        param(13) = last(1) ' LastExp (previous)
        On Error GoTo 0
        
        ' Insert to DB
        Call InsertRecord(param(0), _
                          param(1), _
                          param(2), _
                          param(3), _
                          param(4), _
                          param(5), _
                          param(6), _
                          param(7), _
                          param(8), _
                          param(9), _
                          param(10), _
                          param(11), _
                          param(12), _
                          param(13))
                          
        'FindCodeIndex(Me.tbBarCode.value, Me.tbLot.value, Me.tbExp.value)
        On Error Resume Next
        If Len(last(0)) > 0 Then
            If DataResult(1) <> last(0) Then
                Me.lbLotChange.BackColor = &H80FFFF
                Me.lbLotChange.Caption = "Lot is changed, Form: Lot:" & last(0) & " Exp: " & last(1) & "To: Lot:" & DataResult(1) & " Exp: " & DataResult(2)
            End If
        End If
        On Error GoTo 0
        If ResultMsg = "Matched" Then
            If DataResult(1) <> "na----" And DataResult(2) <> "na----" And Len(DataResult(1)) > 0 And Len(DataResult(2)) > 0 Then
                AlgoStr = "Custom"
                result = FindCodeIndex(Me.lbBarCode.Caption, DataResult(1), DataResult(2))
                CustomAlgoStr = "3,16," & result(0) & "," & result(1) & "," & result(2) & "," & result(3)
                rowsAffected = PrepareUpdateSQL("GTIN", DataResult(0), Left(Me.lbGTIN.Caption, 6), ApplyKey(Left(Me.lbGTIN.Caption, 6)), "", AlgoStr, CustomAlgoStr)
        
                'MsgBox "i learn row:" & rowsAffected
            End If
        End If

        LastGTIN = ""
        Me.lbQRCode.Caption = ""
        Me.lbGTIN.Caption = ""
        Me.lbBarCode.Caption = ""
    Else
        Me.lbMatchResult.BackColor = &H80000014
        Me.frResult.BackColor = &H80000014
    End If
    Me.tbCodeScan.value = ""
    Me.tbCodeScan.SetFocus
    Call wsProtectAll
End Sub

Private Sub teststyle()

            ResultMsg = "Warning: Not Match"
            Me.lbMatchResult.Caption = ResultMsg
            Me.frResult.BackColor = &HFF&
            Me.lbMatchResult.BackColor = &HFF&
End Sub

Private Sub UpDateCurrentLot(ByVal ResultArr As Variant)
    
    
        
        Dim loc, GTIN, Batch, Exp, item, PackSize As String
        
        GTIN = UCase(ResultArr(0))
        Batch = UCase(ResultArr(1))
        Exp = UCase(ResultArr(2))
        
        iRow = wsGTIN.Columns(2).Find(what:=GTIN, LookAt:=xlPart).Row
        
        PackSize = wsGTIN.Cells(iRow, 9).value
        
        If Len(PackSize) > 0 Then PackSize = "(" & PackSize & ")"
        
        item = wsGTIN.Cells(iRow, 1).value & PackSize
        
        Dim rng As Range
        Set rng = wsCurrentLot.Range("A:A").Find(what:=item)
        
        rng.Offset(0, 1).value = Batch
        rng.Offset(0, 2).value = Exp

    
End Sub
Private Sub tbCodeScan_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If CapsLock = True Then
        MsgBox "CapsLock is On, Please turn off capslock before Scan"
        Me.tbCodeScan.value = ""
    End If
End Sub

Private Sub tbCodeScan_Change()
    Me.tbCodeScan.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    
    'frmType.Show
    
    If frmApp.Checkob1 Then
        Me.lbRefillType.Caption = frmApp.Checkob1.Caption
    ElseIf frmApp.Checkob2 Then
        Me.lbRefillType.Caption = frmApp.Checkob2.Caption
    ElseIf frmApp.Checkob3 Then
        Me.lbRefillType.Caption = frmApp.Checkob3.Caption
    End If
    
    Me.lb_RefNum.Caption = "NLTH-" & Format(Date, "YYMMDD") & "00" & Get_RefNum()
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Call Update_wsRecord
    ThisWorkbook.Save
End Sub


