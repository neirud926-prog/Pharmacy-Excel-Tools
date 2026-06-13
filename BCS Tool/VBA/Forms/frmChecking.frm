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

' Cached values carried between the GTIN scan and the QR scan that follows it,
' so the same barcode is not decoded - and the same GTIN row not queried -
' several times per check.
Private mDecoded As Variant   ' [GTIN, Lot, Exp] from the last scanned GTIN
Private mGtinRow As Variant   ' GTIN detail row from the last scanned GTIN
Private mBusy As Boolean      ' re-entrancy guard for tbCodeScan_AfterUpdate

' Location lookup that returns "" instead of raising when the item is missing.
Private Function LocLookup(ByVal key As String) As String
    Dim v As Variant
    v = Application.VLookup(key, wsItemLoc.Range("A:C"), 3, 0)
    If IsError(v) Then
        LocLookup = ""
    Else
        LocLookup = CStr(v)
    End If
End Function

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
    
    ' Set today��s date for filtering
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
    If mBusy Then Exit Sub
    mBusy = True
    Call wsUnProtectAll

    Dim code As String
    code = Me.tbCodeScan.value

    If code = "#ManualFill" Then
        Me.lbGTIN.Caption = "GTINNA-Manual"
    ElseIf Mid(code, 12, 2) = "_B" And Right(code, 2) = "aR" Then
        Me.lbQRCode.Caption = LookUpItemCode(Mid(code, 13))
    ElseIf Len(code) = 19 And Left(code, 1) = "B" And Right(code, 2) = "aR" Then
        Me.lbQRCode.Caption = LookUpItemCode(code)
    ElseIf Len(code) > 7 Then
        Me.lbBarCode.Caption = code
        LastGTIN = code

        ' Decode the barcode ONCE (no DB unless a custom algo is required) and
        ' cache it for the matching scan that follows.
        mDecoded = ReadBarCode(code)
        If IsEmpty(mDecoded) Then
            Call CleanupScan
            Exit Sub
        End If

        Dim GTIN As String
        GTIN = mDecoded(0)

        ' Fetch the GTIN detail row ONCE; reused below for the item name, the
        ' pack size and the algo learning instead of querying the DB 3-4 times.
        mGtinRow = ShowGTINDetailsByGTIN(GTIN)
        Me.lbGTIN.Caption = mGtinRow(1, 1) & " - " & mGtinRow(1, 2)
    End If

    Dim itemCode As String
    itemCode = Left(Me.lbGTIN.Caption, 6)

    ' Show Location
    If Len(Me.lbGTIN.Caption) > 10 And Right(Me.lbGTIN.Caption, 6) <> "Manual" Then
        Me.lbMatchResult.Caption = LocLookup(itemCode)
        Me.lbMatchResult.Font.Size = 120
    Else
        Me.lbMatchResult.Caption = ""
    End If

    ' Checking
    Dim ResultMsg As String
    If Len(Me.lbQRCode.Caption) = 6 And Len(Me.lbGTIN.Caption) > 10 Then
        If Me.lbQRCode.Caption = itemCode Then
            ResultMsg = "Matched"
            Me.frResult.BackColor = &HC0FFC0
            Me.lbMatchResult.BackColor = &HC0FFC0
            Me.lbMatchResult.Font.Size = 120
            Me.lbMatchResult.Caption = LocLookup(itemCode)
        ElseIf Right(Me.lbGTIN.Caption, 6) = "Manual" Then
            ResultMsg = "Manual"
            Me.frResult.BackColor = &H80FFFF
            Me.lbMatchResult.BackColor = &H80FFFF
            Me.lbMatchResult.Font.Size = 108
            Me.lbMatchResult.Caption = LocLookup(Me.lbQRCode.Caption)
        Else
            ' set size
            ResultMsg = "Warning: Not Match"
            Me.lbMatchResult.Caption = ResultMsg
            Me.frResult.BackColor = &HFF&
            Me.lbMatchResult.BackColor = &HFF&
            Me.lbMatchResult.Font.Size = 60
            Call NotMatchAlert(itemCode)
        End If

        ' Get the last Lot and Exp for the Binshelf (previous to this insert)
        Dim binshelf As String
        binshelf = Me.lbQRCode.Caption
        Dim last As Variant
        last = GetLastLotAndExp(binshelf)

        ' Reuse the decode from the GTIN scan rather than decoding again
        Dim DataResult As Variant
        If IsEmpty(mDecoded) Then
            DataResult = ReadBarCode(LastGTIN)
        Else
            DataResult = mDecoded
        End If

        ' Pack size from the cached GTIN row (was a separate DB lookup)
        Dim PackSz As String
        PackSz = ""
        On Error Resume Next
        If Not IsEmpty(mGtinRow) Then
            If mGtinRow(1, 4) <> "" Then PackSz = mGtinRow(1, 4) & "ML"
        End If
        On Error GoTo 0

        ' Resolve the current user ONCE (the ADSI lookup is slow)
        Dim fullName As String
        fullName = GetUserFullName()

        ' Prepare parameters
        Dim param(0 To 13) As String
        param(0) = Format(Now, "dd-mmm-yyyy hh:mm")
        param(1) = Me.lbGTIN.Caption
        param(2) = Me.lbQRCode.Caption
        If ResultMsg = "Warning: Not Match" Then
            param(3) = LocLookup(itemCode) & "(" & LocLookup(Me.lbQRCode.Caption) & ")"
        ElseIf ResultMsg = "Manual" Then
            param(3) = LocLookup(Me.lbQRCode.Caption) & "(Manual)"
        Else
            param(3) = LocLookup(itemCode)
        End If
        param(4) = ResultMsg
        If fullName = "nltpmscclose" Or fullName = "User" Then
            param(8) = wsMem.Range("B1").value
        Else
            param(8) = fullName
        End If
        param(5) = Me.lbRefillType.Caption
        param(6) = Me.lb_RefNum.Caption
        param(7) = "Processing"
        On Error Resume Next
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

        On Error Resume Next
        If Len(last(0)) > 0 Then
            If DataResult(1) <> last(0) Then
                Me.lbLotChange.BackColor = &H80FFFF
                Me.lbLotChange.Caption = "Lot is changed, Form: Lot:" & last(0) & " Exp: " & last(1) & "To: Lot:" & DataResult(1) & " Exp: " & DataResult(2)
            End If
        End If
        On Error GoTo 0

        ' Learn a custom decode position map for this GTIN on a good match, but
        ' only write to the DB when it actually differs from what is stored so
        ' we are not doing a redundant network UPDATE on every matched scan.
        If ResultMsg = "Matched" Then
            If DataResult(1) <> "na----" And DataResult(2) <> "na----" And Len(DataResult(1)) > 0 And Len(DataResult(2)) > 0 Then
                AlgoStr = "Custom"
                result = FindCodeIndex(Me.lbBarCode.Caption, DataResult(1), DataResult(2))
                CustomAlgoStr = "3,16," & result(0) & "," & result(1) & "," & result(2) & "," & result(3)

                Dim needUpdate As Boolean
                needUpdate = True
                On Error Resume Next
                If Not IsEmpty(mGtinRow) Then
                    If mGtinRow(1, 5) = AlgoStr And mGtinRow(1, 6) = CustomAlgoStr Then needUpdate = False
                End If
                On Error GoTo 0

                If needUpdate Then
                    rowsAffected = PrepareUpdateSQL("GTIN", DataResult(0), itemCode, ApplyKey(itemCode), "", AlgoStr, CustomAlgoStr)
                End If
            End If
        End If

        LastGTIN = ""
        mDecoded = Empty
        mGtinRow = Empty
        Me.lbQRCode.Caption = ""
        Me.lbGTIN.Caption = ""
        Me.lbBarCode.Caption = ""
    Else
        Me.lbMatchResult.BackColor = &H80000014
        Me.frResult.BackColor = &H80000014
    End If

    Call CleanupScan
End Sub

' Restores worksheet protection and resets the scan box. Called on every exit
' path of tbCodeScan_AfterUpdate.
Private Sub CleanupScan()
    Me.tbCodeScan.value = ""
    On Error Resume Next
    Me.tbCodeScan.SetFocus
    On Error GoTo 0
    Call wsProtectAll
    mBusy = False
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


