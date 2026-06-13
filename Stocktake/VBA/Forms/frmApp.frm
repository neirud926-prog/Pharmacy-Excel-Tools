VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmApp 
   Caption         =   "North Lantau Hospital Pharmacy"
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21000
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

Private Sub btnChangeType_Click()
    Dim userInput As String
    
    With Me.btnChangeType
        If .Caption = "Daily" Then
            ' Switching to Annual ¡÷ ask for confirmation
            userInput = InputBox("To change to Annual mode, please type exactly:" & vbCrLf & vbCrLf & _
                                 "Annual" & vbCrLf & vbCrLf & _
                                 "(case-sensitive)", "Confirm Annual Mode")
            
            If userInput = "Annual" Then
                .Caption = "Annual"
                MsgBox "Mode changed to Annual.", vbInformation
            Else
                MsgBox "Change cancelled. Still in Daily mode.", vbInformation
            End If
            
        ElseIf .Caption = "Annual" Then
            ' Switching back to Daily ¡÷ no confirmation needed
            .Caption = "Daily"
            MsgBox "Mode changed back to Daily.", vbInformation
        End If
    End With
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
    Me.MutiPage.Value = btn
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
    'Me.CheckFrbtnStart.BackColor = RGB(221, 244, 255)
    'Me.PrintFrbtnPrint.BackColor = RGB(221, 244, 255)
    
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



Private Sub Label2_Click()

End Sub





'---------UI Hover------------------------------------ End

'------------------------------------------Form Life Cycle--------------------------------------------
Private Sub UserForm_Initialize()
    Me.MutiPage.Value = 1
    Me.MutiPage.Style = fmTabStyleNone
    With Me.ListTakenView
        .View = lvwReport              ' Column view
        .FullRowSelect = True          ' Select whole row
        .Gridlines = True              ' For readability
        .ColumnHeaders.Clear           ' Reset
        Me.ListTakenView.BackColor = vbWhite
        Me.ListTakenView.ForeColor = vbBlack
        
        ' Add 14 columns (unlimited possible)
        .ColumnHeaders.Add , , "Item Descriptions", 321
        .ColumnHeaders.Add , , "Item Code", 48
        .ColumnHeaders.Add , , "Col2", 49
        .ColumnHeaders.Add , , "Col3", 56
        .ColumnHeaders.Add , , "Col4", 90
        .ColumnHeaders.Add , , "Col5", 37
        .ColumnHeaders.Add , , "Col6", 48
        .ColumnHeaders.Add , , "Col7", 48
        .ColumnHeaders.Add , , "Col8", 64
        .ColumnHeaders.Add , , "Col9", 48
        .ColumnHeaders.Add , , "Col10", 48
        .ColumnHeaders.Add , , "Col11", 48
        .ColumnHeaders.Add , , "Col12", 48
        .ColumnHeaders.Add , , "Col13", 48
    End With
    Me.MutiPage.Value = 0
End Sub

Private Sub UserForm_Terminate()

End Sub

'-----------------------------dATA eNTRY
Private Sub PopulateListView()
    Dim itm As ListItem
    Dim missing As String
    Dim invalidNum As String
    Dim allValid As Boolean: allValid = True
    
    ' ------------------------------------------------
    ' 1. Collect all input fields once
    ' ------------------------------------------------
    Dim fields As Variant
    fields = Array( _
        Me.EntrylbDes.Caption, _
        Me.EntrytbItemCode.Value, _
        Me.EntrytbBinno.Caption, _
        Me.EntrytbBinQty.Value, _
        Me.EntrytbPPQty.Value, _
        Me.EntrytbEV54Qty.Value, _
        Me.EntrytbEV180Qty.Value, _
        Me.EntrytbTOSHOQty.Value, _
        Me.EntrytbMSQty.Value, _
        Me.EntrytbTUQty.Value, _
        Me.Entrytb277Qty.Value, _
        Me.EntrytbOPQty.Value, _
        Me.EntrytbIPQty.Value, _
        Me.EntrytbERPQty.Value _
    )
    
    ' ------------------------------------------------
    ' 2. Validate all fields
    ' ------------------------------------------------
    Dim i As Long
    For i = 0 To UBound(fields)
        Dim val As String
        val = Trim(fields(i) & "")
        
        If val = "" Then
            missing = missing & IIf(missing <> "", ", ", "") & "Field " & (i + 1)
            allValid = False
        ElseIf i >= 3 And i <= 13 Then   ' Qty fields (3 to 13)
            If Not IsNumeric(val) Then
                invalidNum = invalidNum & IIf(invalidNum <> "", ", ", "") & "Field " & (i + 1)
                allValid = False
            ElseIf CDbl(val) < 0 Then
                invalidNum = invalidNum & IIf(invalidNum <> "", ", ", "") & "Field " & (i + 1) & " (negative)"
                allValid = False
            End If
        End If
    Next i
    
    ' ------------------------------------------------
    ' 3. Show error if validation failed
    ' ------------------------------------------------
    If Not allValid Then
        Dim msg As String
        msg = "Cannot add item. Please correct the following:" & vbCrLf & vbCrLf
        
        If missing <> "" Then
            msg = msg & "Missing / empty:" & vbCrLf & missing & vbCrLf & vbCrLf
        End If
        
        If invalidNum <> "" Then
            msg = msg & "Invalid number:" & vbCrLf & invalidNum
        End If
        
        MsgBox msg, vbExclamation, "Input Validation Failed"
        
        ' Optional: focus first problematic field (you can improve this)
        Me.EntrytbItemCode.SetFocus
        Exit Sub
    End If
    
    ' ------------------------------------------------
    ' 4. All valid ? add to ListView
    ' ------------------------------------------------
    Application.ScreenUpdating = False
    
    Set itm = Me.ListTakenView.ListItems.Add(, , Me.EntrylbDes.Caption) ' Col0
    
    itm.SubItems(1) = UCase(Me.EntrytbItemCode.Value)      ' Col1
    itm.SubItems(2) = Me.EntrytbBinno.Caption       ' Col2
    itm.SubItems(3) = Me.EntrytbBinQty.Value        ' Col3
    itm.SubItems(4) = Me.EntrytbPPQty.Value         ' Col4
    itm.SubItems(5) = Me.EntrytbEV54Qty.Value       ' Col5
    itm.SubItems(6) = Me.EntrytbEV180Qty.Value      ' Col6
    itm.SubItems(7) = Me.EntrytbTOSHOQty.Value      ' Col7
    itm.SubItems(8) = Me.EntrytbMSQty.Value         ' Col8
    itm.SubItems(9) = Me.EntrytbTUQty.Value         ' Col9
    itm.SubItems(10) = Me.Entrytb277Qty.Value       ' Col10
    itm.SubItems(11) = Me.EntrytbOPQty.Value        ' Col11
    itm.SubItems(12) = Me.EntrytbIPQty.Value        ' Col12
    itm.SubItems(13) = Me.EntrytbERPQty.Value       ' Col13
    
    Application.ScreenUpdating = True
    
    ' Clear fields
    Call ClearEntryField
    Me.EntrytbItemCode.Value = ""
    
    ' Optional: give feedback
    ' MsgBox "Item added successfully.", vbInformation
End Sub
Private Sub EntrybtnSave_Click()
     Call PopulateListView
     Me.EntrytbItemCode.SetFocus
End Sub

Private Sub EntrytbItemCode_Change()
    If Len(Me.EntrytbItemCode.Value) = 6 Then
    On Error Resume Next
    Me.EntrylbDes.Caption = LookupRow(wsNLTItem.Range("A1").CurrentRegion, 1, Me.EntrytbItemCode.Value)(2)
        Call ApplyRestiction(Me.EntrytbItemCode.Value)
    On Error GoTo 0
    Else
        Call ClearEntryField
    End If
End Sub
Private Sub ClearEntryField()
    Me.EntrylbDes.Caption = ""
    
    Me.EntrytbBinno.Caption = ""    ' Col2
    Me.EntrytbBinQty.Value = ""   ' Col3
    Me.EntrytbPPQty.Value = ""   ' Col4
    Me.EntrytbEV54Qty.Value = ""   ' Col5
    Me.EntrytbEV180Qty.Value = ""   ' Col6
    Me.EntrytbTOSHOQty.Value = ""   ' Col7
    Me.EntrytbMSQty.Value = ""  ' Col8
    Me.EntrytbTUQty.Value = ""   ' Col9
    Me.Entrytb277Qty.Value = ""     ' Col10 (mixed types fine)
    Me.EntrytbOPQty.Value = ""        ' Col11
    Me.EntrytbIPQty.Value = ""         ' Col12
    Me.EntrytbERPQty.Value = ""
    '
End Sub

Private Sub ApplyRestiction(ByVal itemCode As String)
    Dim j As Long
    Dim result As Variant
    
    '--- 1. Lookup table: source worksheet + the TextBox you want to enable/disable
    Dim lookups As Variant
    lookups = Array( _
        Array(wsPPLoc, Me.EntrytbPPQty), _
        Array(wsEV54, Me.EntrytbEV54Qty), _
        Array(wsEV180, Me.EntrytbEV180Qty), _
        Array(wsTOSHO, Me.EntrytbTOSHOQty), _
        Array(wsMobileShelf, Me.EntrytbMSQty), _
        Array(wsTopupSMLoc, Me.EntrytbTUQty) _
    )
    
    '--- 2. Loop through every source sheet / textbox pair
    For j = LBound(lookups) To UBound(lookups)
        
        Dim srcWs As Worksheet
        Dim ctrl  As MSForms.Control
        
        Set srcWs = lookups(j)(0)            ' worksheet object
        Set ctrl = lookups(j)(1)             ' textbox control
        
        '--- Lookup the itemCode in column A of the source sheet
        result = LookupRow(srcWs.Range("A1").CurrentRegion, 1, itemCode)
        
        '--- Enable textbox only when the item exists in that sheet
        If IsEmpty(result) Then
            ctrl.Enabled = 0
            ctrl.BackColor = &H80000003
            ctrl.Value = 0
        Else
            ctrl.Enabled = 1
            ctrl.BackColor = &H80000005
        End If
        
        
    Next j
    
    Me.EntrytbBinno.Caption = LookupRow(wsBinShelfLocator.Range("A1").CurrentRegion, 1, itemCode)(2)

End Sub

Private Sub ListTakenView_DblClick()
    Dim itm As ListItem
    
    ' Use the already selected item (double-click auto-selects the row)
    Set itm = Me.ListTakenView.SelectedItem
    
    If itm Is Nothing Then Exit Sub   ' safety (should not happen on dblclick)
    
    If MsgBox("Remove """ & itm.Text & """ from the list?", _
              vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        
        Me.ListTakenView.ListItems.Remove itm.Index
        
        ' Optional: select next item or clear selection after delete
        If Me.ListTakenView.ListItems.Count > 0 Then
            Me.ListTakenView.ListItems(1).Selected = True
        End If
    End If
End Sub

Private Sub EntrybtnSubmit_Click()
     SaveListViewToAccess Me
     Me.ListTakenView.ListItems.Clear
     Me.EntrytbItemCode.SetFocus
     Call eMail
     
End Sub
'--------------------Print WS--------------------------------------
Private Sub btnPrint5items_Click()
    
    Call Print5itemsWorkSheet
    Unload Me
End Sub



'----------------------Report--------------------------------------
Private Sub ReportbtnDailyRecord_Click()
    Unload Me
    LoadDailyRecordsToReport
    Call ApplyReportFormatting  '
End Sub

Private Sub ReportbtnAnnualRecord_Click()
    Unload Me
    LoadAnnualRecordsToReport
    Call ApplyReportFormatting '
End Sub

'


