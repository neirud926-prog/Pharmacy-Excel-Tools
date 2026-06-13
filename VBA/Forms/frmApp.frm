VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmApp 
   Caption         =   "North Lantau Hospital Pharmacy"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "frmApp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------Btn Click Action------------------------------

Private Sub AddbtnHelper_Click()
    frmFindTag.Show
End Sub

Private Sub btnlb1_Click()
    Call PannelClicked(0)
End Sub
Private Sub btnlb2_Click()
    Call PannelClicked(1)
End Sub
Private Sub btnlb3_Click()
    Call PannelClicked(2)
End Sub
Private Sub btnlb4_Click()
    Call PannelClicked(3)
End Sub
Private Sub btnlb5_Click()
    Call PannelClicked(4)
End Sub
Private Sub btnlb6_Click()
    Call PannelClicked(5)
End Sub

Private Sub btnFrHome_Click()
    Call PannelClicked(0)
End Sub
Private Sub btnFrCheck_Click()
    Call PannelClicked(1)
End Sub
Private Sub btnFrPrint_Click()
    Call PannelClicked(2)
End Sub
Private Sub btnFrReport_Click()
    Call PannelClicked(3)
End Sub
Private Sub btnFrMaintenance_Click()
    Call PannelClicked(4)
End Sub
Private Sub btnFrFollowUp_Click()
    Call PannelClicked(5)
End Sub



Private Sub IconHome_Click()
    Call PannelClicked(0)
End Sub
Private Sub IconCheck_Click()
    Call PannelClicked(1)
End Sub


Private Sub IconPrint_Click()
    Call PannelClicked(2)
End Sub
Private Sub IconReport_Click()
    Call PannelClicked(3)
End Sub
Private Sub IconMaintenance_Click()
    Call PannelClicked(4)
End Sub
Private Sub IconFollowUp_Click()
    Call PannelClicked(5)
End Sub

Private Sub PannelClicked(ByVal btn As Integer)
    Me.MutiPage.value = btn
    Call PannelSizeSmall
    
    Dim clickedbtn As Integer
    clickedbtn = btn
    Select Case clickedbtn
    Case Is = 0
    Case Is = 1
    Case Is = 2
    Case Is = 3
    Case Is = 4
    End Select
    
End Sub


Private Sub ChecklbbtnStart_Click()
    Call CheckbtnStartisClicked
End Sub

Private Sub CheckFrbtnStart_Click()
    Call CheckbtnStartisClicked
End Sub

Private Sub CheckbtnStartisClicked()
    frmChecking.Show
End Sub

Private Sub PrintlbbtnPrint_Click()
    Call PrintbtnPrintReport
End Sub

Private Sub PrintFrbtnPrint_Click()
   Call PrintbtnPrintReport
End Sub

'---------UI Hover------------------------------------

Private Sub FrPannel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
End Sub

Private Sub FrCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
    Call PannelSizeSmall
End Sub

Private Sub FrHome_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
    Call PannelSizeSmall
End Sub

Private Sub FrMaintain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
    Call PannelSizeSmall
End Sub
Private Sub FrPrint_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
    Call PannelSizeSmall
End Sub

Private Sub FrReport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
    Call PannelSizeSmall
End Sub
Private Sub FrFollowUp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call BtnColorReset
    Call PannelSizeSmall
End Sub

Private Sub btnFrHome_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnFrHome.BackColor = ConstColor.Hoverblue
    Me.btnlb1.BackColor = ConstColor.Hoverblue
    Call PannelSizeBig
End Sub

Private Sub btnFrCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnFrCheck.BackColor = ConstColor.Hoverblue
    Me.btnlb2.BackColor = ConstColor.Hoverblue
    Call PannelSizeBig
End Sub

Private Sub btnFrPrint_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnFrPrint.BackColor = ConstColor.Hoverblue
    Me.btnlb3.BackColor = ConstColor.Hoverblue
    Call PannelSizeBig
End Sub

Private Sub btnFrReport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnFrReport.BackColor = ConstColor.Hoverblue
    Me.btnlb4.BackColor = ConstColor.Hoverblue
    Call PannelSizeBig
End Sub
Private Sub btnFrMaintain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnFrMaintain.BackColor = ConstColor.Hoverblue
    Me.btnlb5.BackColor = ConstColor.Hoverblue
    Call PannelSizeBig
End Sub
Private Sub btnFrFollowUp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnFrFollowUp.BackColor = ConstColor.Hoverblue
    Me.btnlb6.BackColor = ConstColor.Hoverblue
    Call PannelSizeBig
End Sub

Private Sub BtnColorReset()
    Me.btnFrHome.BackColor = RGB(221, 244, 255)
    Me.btnFrCheck.BackColor = RGB(221, 244, 255)
    Me.btnFrPrint.BackColor = RGB(221, 244, 255)
    Me.btnFrReport.BackColor = RGB(221, 244, 255)
    Me.btnFrMaintain.BackColor = RGB(221, 244, 255)
    Me.btnFrFollowUp.BackColor = RGB(221, 244, 255)
    Me.btnlb1.BackColor = RGB(221, 244, 255)
    Me.btnlb2.BackColor = RGB(221, 244, 255)
    Me.btnlb3.BackColor = RGB(221, 244, 255)
    Me.btnlb4.BackColor = RGB(221, 244, 255)
    Me.btnlb5.BackColor = RGB(221, 244, 255)
    Me.btnlb6.BackColor = RGB(221, 244, 255)
    
    '-other page
    Me.CheckFrbtnStart.BackColor = RGB(221, 244, 255)
    Me.PrintFrbtnPrint.BackColor = RGB(221, 244, 255)
    
End Sub

Private Sub PannelSizeSmall()
    Me.FrPannel.Width = 36
End Sub

Private Sub PannelSizeBig()
    Me.FrPannel.Width = 132
End Sub

Private Sub CheckFrbtnStart_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.CheckFrbtnStart.BackColor = ConstColor.Hoverblue
End Sub

Private Sub PrintFrbtnPrint_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.PrintFrbtnPrint.BackColor = ConstColor.Hoverblue
End Sub

'---------UI Hover------------------------------------ End


'-----------------------------Home Page-----------------------------
Private Sub HomePageInitialize()
    Me.HomelbWelcomeMsg.Caption = "Welcome! " & GetUserFullName()
    Call HomeListDisplayTodayData
    Call HomeLoadlbStatisticData
    
End Sub

Private Sub HomeLoadlbStatisticData()
    Dim uniqueRefNo As Long
    Dim totalLines As Long
    Dim nullCheckBy As Long
    Dim notNullCheckBy As Long
   
    uniqueRefNo = CountUniqueRefNoToday()
    totalLines = TotalLinesRecordToday()
    nullCheckBy = CountLinesErrHandlingCheckByNull()
    notNullCheckBy = CountLinesErrHandlingCheckByNotNull()
    AllNullCheckBy = AllCountLinesErrHandlingCheckByNull
   
    Me.HomelbNumReplen.Caption = uniqueRefNo
    Me.HomelbNumItem.Caption = totalLines
    Me.HomelbNumError.Caption = nullCheckBy
    Me.HomelbNumErrorDone.Caption = notNullCheckBy
    Me.HomelbNumAllError.Caption = AllNullCheckBy
    
End Sub

Public Sub HomeListDisplayTodayData(Optional ByVal newdate As String)
    Dim results As Variant
    Dim i As Long
    
    results = LoadTodayData(newdate)
    If IsArray(results) Then
        With Me.HomeListReportStatus
            Me.HomeListReportStatus.Clear
            '.ColumnCount = 5
            '.ColumnHeads = True
            '.ColumnWidths = "100;50;150;150;20"
            '.ColumnHeadings = "Ref No,Printed Status,Item Counted,Shelved By"
            '.RowSource = ""
            For i = 1 To UBound(results, 1)
                .AddItem results(i, 1)
                .List(i - 1, 1) = results(i, 2)
                .List(i - 1, 2) = results(i, 5)
            Next i
        End With
    Else
        MsgBox "No data found or an error occurred."
    End If
End Sub






'-------------------------Page Maintain---------------------------
Private Sub MainbtnAddGTIN_Click()
    
    If Len(ItemDataMatch(Me.MaintbItemCode.value)) > 0 Then
        EditGTIN = False
        
        Call MainLoadAddForm
    End If
    
End Sub

Private Sub MainLoadAddForm()
        MainFrAddFrm.Visible = True
        Me.AddlbItemCode.Caption = Me.MaintbItemCode.value
        Me.AddlbDrugDes.Caption = Me.MainlbDrugDes.Caption
    
        If EditGTIN = True Then
            Me.AddlbGTIN.Caption = Me.MainlistGTIN.List(, 1)
        
            Me.AddcbPackSize.value = Me.MainlistGTIN.List(, 3)
            If Me.MainlistGTIN.List(, 5) = "Custom" Then
                Me.AddcbCustomAlgo = True
        End If
        
        Me.AddbtnDelete.Enabled = True
        
        'Me.AddtbBarCode.SetFocus
        End If
        
End Sub
Private Sub MaincbBarCodemode_Change()
    Dim rng As Range
    Set rng = wsBarCodeSetting.Range("A:A").Find(what:=Me.tbItemCode.value)
    rng.Offset(0, 1).value = Me.MaincbBarCodemode.value
End Sub



Private Sub MainFrAddFrm_Click()

End Sub

Private Sub MainlistGTIN_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Me.MainlistGTIN.value = "" Then
        EditGTIN = True
        Call MainLoadAddForm
    End If
End Sub


Public Sub MainListGTIN_Refresh()
    Me.MainlistGTIN.Clear
    
    Me.MainlistGTIN.List = ListViewShowGTINDetails(Me.MaintbItemCode.value)
    
End Sub


Private Sub MaintbItemCode_Change()
    Me.MaintbItemCode.value = UCase(Me.MaintbItemCode.value)
    If Len(Me.MaintbItemCode) = 6 Then
        Me.MainlistGTIN.Clear
        Me.MainlbDrugDes.Caption = ItemDataMatch(Me.MaintbItemCode, 2)
        If Me.MainlbDrugDes.Caption = "" Then Me.MaintbItemCode.value = ""
        
        If Not Me.MainlbDrugDes.Caption = "" Then
            Call ListGTIN_Refresh
            Me.MainbtnAddGTIN.Enabled = True
        Else
            Me.MainbtnAddGTIN.Enabled = False
        End If
        'Me.cbBarCodemode.Enabled = True
        'Me.cbBarCodemode.Value = Application.WorksheetFunction.VLookup(Me.tbItemCode.Value, wsBarCodeSetting.Range("A:B"), 2, 0)
    ElseIf Len(Me.MaintbItemCode) = 4 Then
        Me.MainlistGTIN.Clear
        
    Else
        Me.MainlistGTIN.Clear
        Me.MainlbDrugDes.Caption = ""
        Me.MainbtnAddGTIN.Enabled = False
        EditGTIN = False
        Me.MaincbBarCodemode.value = ""
        Me.MaincbBarCodemode.Enabled = False
        Me.MainFrAddFrm.Visible = False
    End If
End Sub
Public Sub ListGTIN_Refresh()
    Me.MainlistGTIN.Clear
    
    Me.MainlistGTIN.List = ListViewShowGTINDetails(Me.MaintbItemCode.value)
End Sub

'--------------------------Page Main - > Add GTIN ---------------------------------
Private Sub AddbtnAddGTIN_Click()

If Len(Me.AddlbGTIN.Caption) > 0 Then

    Dim rowsAffected As Long
    Dim CustomAlgoStr, AlgoStr As String
    
    If Me.AddcbCustomAlgo = True Then
        AlgoStr = "Custom"
        CustomAlgoStr = Me.AddGTINStart.value & "," & Me.AddGTINEnd.value & "," & Me.AddLotStart.value & "," & Me.AddLotEnd.value & "," & Me.AddEXPStart.value & "," & Me.AddEXPEnd.value
    End If
    
    If EditGTIN = False Then
    
            If Me.AddcbCustomAlgo = True Then
                rowsAffected = PrepareInsertSQL("GTIN", Me.AddlbItemCode.Caption, Me.AddlbGTIN.Caption, ApplyKey(Me.AddlbItemCode.Caption), Me.AddcbPackSize.value, AlgoStr, CustomAlgoStr)
            Else
                rowsAffected = PrepareInsertSQL("GTIN", Me.AddlbItemCode.Caption, Me.AddlbGTIN.Caption, ApplyKey(Me.AddlbItemCode.Caption), Me.AddcbPackSize.value)
            End If
            
        
    ElseIf EditGTIN = True Then
        If Me.AddcbCustomAlgo = True Then
            rowsAffected = PrepareUpdateSQL("GTIN", Me.AddlbGTIN.Caption, Me.AddlbItemCode.Caption, ApplyKey(Me.AddlbItemCode.Caption), Me.AddcbPackSize.value, AlgoStr, CustomAlgoStr)
        Else
            rowsAffected = PrepareUpdateSQL("GTIN", Me.AddlbGTIN.Caption, Me.AddlbItemCode.Caption, ApplyKey(Me.AddlbItemCode.Caption), Me.AddcbPackSize.value)
        End If
    End If
    
    MsgBox rowsAffected & " rows updated."
    
    Me.MaintbItemCode.value = ""
    Me.MaintbItemCode.SetFocus
    
Else
    MsgBox "Incorrect data!"
End If
    
        
End Sub

Private Sub AddtbBarCode_AfterUpdate()
    Dim GTINStr As String
    
    If Me.AddcbCustomAlgo = False Then
        BarCodeLength = Len(Me.AddtbBarCode.value)
        
        If BarCodeLength < 16 And BarCodeLength > 0 Then
            'GTIN only
            AddZeroCount = 14 - BarCodeLength
            
            For i = 1 To AddZeroCount
                GTINStr = GTINStr & "0"
            Next
            
            Me.AddlbGTIN.Caption = GTINStr & Me.AddtbBarCode.value
        ElseIf BarCodeLength >= 16 Then
        
            Dim ResultArr As Variant
            ResultArr = ReadBarCode(Me.AddtbBarCode.value)
            If IsEmpty(ResultArr) Then
                Me.AddtbBarCode.value = ""
                Exit Sub
            End If
            Me.AddlbGTIN.Caption = UCase(ResultArr(0))
            Me.AddlbBatch.Caption = ResultArr(1)
            Me.AddlbExpDate.Caption = ResultArr(2)
        
        Else
            Me.AddlbGTIN.Caption = ""
            Me.AddlbBatch.Caption = ""
            Me.AddlbExpDate.Caption = ""
        End If
        
    ElseIf Me.AddcbCustomAlgo = True And Len(Me.AddGTINStart) > 0 And Len(Me.AddGTINEnd) > 0 And Len(Me.AddLotStart) > 0 And Len(Me.AddLotEnd) > 0 And Len(Me.AddEXPStart) > 0 And Len(Me.AddEXPEnd) > 0 Then
        
        Me.AddlbGTIN.Caption = Mid(Me.AddtbBarCode.value, Me.AddGTINStart.value, Me.AddGTINEnd.value - Me.AddGTINStart + 1)
        Me.AddlbBatch.Caption = Mid(Me.AddtbBarCode.value, Me.AddLotStart.value, Me.AddLotEnd.value - Me.AddLotStart + 1)
        Me.AddlbExpDate.Caption = Mid(Me.AddtbBarCode.value, Me.AddEXPStart.value, Me.AddEXPEnd.value - Me.AddEXPStart + 1)
        
    End If
End Sub

Private Sub AddcbCustomAlgo_Click()
    If MainAddFrCustomAlgo.Visible = False Then
        frmpw.Show
    Else
        Me.MainAddFrCustomAlgo.Visible = False
    End If
    
End Sub

Private Sub AddbtnDelete_Click()
    Dim answer As Integer
    answer = MsgBox("Delete " & Me.AddlbItemCode.Caption & " " & Me.AddlbGTIN.Caption & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm?")
    If answer = vbYes Then
        rowsAffected = PrepareDeleteSQL("GTIN", Me.AddlbGTIN.Caption)
        MsgBox rowsAffected & " rows deleted."
        Me.MainFrAddFrm.Visible = False
        Me.MaintbItemCode.value = ""
    End If
End Sub

'--------------------------Print page --------------------------------------------

Public Sub DisplayTodayData(Optional ByVal newdate As String)
    Dim results As Variant
    Dim i As Long
    
    results = LoadTodayData(newdate)
    If IsArray(results) Then
        With Me.PrintListRefNum
            Me.PrintListRefNum.Clear
            .ColumnCount = 5
            '.ColumnHeads = True
            .ColumnWidths = "90;60;145;150;20"
            '.ColumnHeadings = "Ref No,Printed Status,Item Counted,Shelved By"
            .RowSource = ""
            For i = 1 To UBound(results, 1)
                .AddItem results(i, 1)
                .List(i - 1, 1) = results(i, 2)
                .List(i - 1, 2) = results(i, 3)
                .List(i - 1, 3) = results(i, 4)
                .List(i - 1, 4) = results(i, 5)
            Next i
        End With
    Else
        MsgBox "No data found or an error occurred."
    End If
End Sub



Private Sub PrintBtnDate_Click()
    DatePicker.Show
End Sub

Private Sub PrintPageInitialize()
    Call DisplayTodayData
    Me.PrintBtnDate.Caption = Format(Now, "dd-mmm-yyyy")
End Sub

Private Sub PrintbtnPrintReport()
    If Me.PrintListRefNum.List(, 0) = "" Then Exit Sub
    Call PrintReport(Me.PrintListRefNum.List(, 0))
    Call PrintPageInitialize
End Sub
'--------------------------------------------Report Page------------

Private Sub ReportbtnGenManualReport_Click()
    frmReplenManualFillReport.Show
End Sub

Private Sub btnLastestItemLotReport_Click()
    Call SaveShelvingLastDetailsReport
End Sub

Private Sub btnItemFrequencyReport_Click()
    Call SaveShelvingReportToNewWorkbook
End Sub

'------------------------------------------Page Follow UP---------------------------------------------
Sub FollowUpPageInitialize()
    Call QuerySQLLoadError
    
    Dim arrNameList As Variant
    Dim lastRow As Long
    Dim rng As Range
    Me.FUcbFUaction.AddItem "Checked without problem"
    
    lastRow = wsNameList.Cells(wsNameList.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data found in column A of NameList worksheet.", vbExclamation
        Exit Sub
    End If
    
    Set rng = wsNameList.Range("A2:A" & lastRow)
    arrNameList = rng.value
    Me.FUcbNameList.Clear
    Dim i As Long
    For i = 1 To UBound(arrNameList, 1)
        If Not IsEmpty(arrNameList(i, 1)) Then
            Me.FUcbNameList.AddItem CStr(arrNameList(i, 1))
        End If
    Next i

End Sub

Sub QuerySQLLoadError()
    Dim results As Variant
    Dim i As Long
   
    ' Path is handled within DatabaseSQL, no need to specify here
    results = DatabaseSQL("SELECT * FROM [Error Handling Query FU]")
   
    ' Configure and populate the listbox
    With Me.FollowUpListErr
        .Clear ' Clear any existing items
        .ColumnCount = 10 ' Set to 10 columns as specified
        
        If IsArray(results) Then
            For i = 1 To UBound(results, 1)
                .AddItem ' Add a new row
                ' Fill the 10 columns for the current row
                For j = 0 To 9
                    .List(.ListCount - 1, j) = IIf(IsNull(results(i, j + 1)), "", results(i, j + 1))
                Next j
            Next i
        End If
    End With
End Sub

Private Sub FollowUpListErr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo err1
    With Me.FollowUpListErr
        If Len(.List(, 0)) > 0 Then
            Me.FUlbID.Caption = "ID: " & .List(, 0)
            Me.FUlbID.Tag = .List(, 0)
            Me.FUlbRefnum.Caption = "Ref. No.: " & .List(, 7)
            Me.FUlbDT.Caption = "Date and Time: " & .List(, 1)
            Me.FUlbError.Caption = "Error: " & .List(, 5)
            Me.FUlbGTIN.Caption = "Scanned GTIN(Drug): " & .List(, 2)
            Me.FUlbBinNum.Caption = "Scanned QR code(Bin-Shelf): " & .List(, 3)
            Me.FUlbLoc.Caption = "Correct Location: " & .List(, 4)
            Me.FUlbBy.Caption = "Scanned By: " & .List(, 8)
            
            Me.FUFrFollowUpForm.Visible = True
        Else
            Exit Sub
        End If
    End With
err1:
    If err.Number = 381 Then Exit Sub
    
    
End Sub

Private Sub FUbtnSubmitFollowUpForm_Click()
    On Error GoTo ErrorHandler
   
    Dim fuidValue As String
    Dim fuidLong As Long
    Dim reasonValue As String
    Dim checkByValue As String
    Dim sql As String
    Dim rowsAffected As Long
   
    ' Step 1: Get the FUID, cbReason, and cbCheckBy values from the form
    fuidValue = Trim(Me.FUlbID.Tag) ' Trim to remove any spaces
    reasonValue = Me.FUcbFUaction.value
    checkByValue = Me.FUcbNameList.value
   
    ' Step 2: Validate inputs
    If IsNull(fuidValue) Or fuidValue = "" Then
        MsgBox "Error: FUID is missing.", vbExclamation
        Exit Sub
    End If
   
    If Not IsNumeric(fuidValue) Then
        MsgBox "Error: FUID must be a numeric value.", vbExclamation
        Exit Sub
    End If
   
    fuidLong = CLng(fuidValue) ' Convert to Long Integer
   
    If IsNull(reasonValue) Or reasonValue = "" Then
        MsgBox "Please select a reason before submitting.", vbExclamation
        Exit Sub
    End If
   
    If IsNull(checkByValue) Or checkByValue = "" Then
        MsgBox "Please select a CheckBy value before submitting.", vbExclamation
        Exit Sub
    End If
   
    ' Step 3: Construct the SQL query with escaped values
    sql = "UPDATE ErrHandling SET FollowUpAction = '" & Replace(reasonValue, "'", "''") & "', CheckBy = '" & Replace(checkByValue, "'", "''") & "' WHERE FUID = " & fuidLong & ";"
   
    ' Step 4: Execute the update using DatabaseSQL
    rowsAffected = DatabaseSQL(sql)
   
    ' Step 5: Confirm success
    If rowsAffected > 0 Then
        MsgBox "Record updated successfully!", vbInformation
        Call HomePageInitialize
    Else
        MsgBox "No records were updated.", vbExclamation
    End If
    
    'Call RefreshQuery("Query - Error Handling Query")
    Call QuerySQLLoadError
   
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & err.Description, vbCritical
End Sub
'------------------------------------------Form Life Cycle--------------------------------------------
Private Sub UserForm_Initialize()
    Me.MutiPage.value = 0
    Call HomePageInitialize
    Call PrintPageInitialize
    Call FollowUpPageInitialize
    
    Me.MutiPage.Style = fmTabStyleNone
    
End Sub

Private Sub UserForm_Terminate()

End Sub
