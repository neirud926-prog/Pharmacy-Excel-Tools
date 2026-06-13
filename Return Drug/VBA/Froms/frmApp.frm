VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmApp 
   Caption         =   "North Lantau Hospital Pharmacy"
   ClientHeight    =   8985.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   OleObjectBlob   =   "frmApp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' 1. MODULE VARIABLES & CONSTANTS
' ==============================================================================
Private scanStartTime As Double
Private isAutoFormatting As Boolean
Private targetBin As String
Private targetItemCode As String
Private isProcessing As Boolean
Private HoverEffectCollection As Collection
Public DisableHover As Boolean
Private isManualFill As Boolean

' UI Constants
Private Const PANEL_SMALL_WIDTH As Integer = 36
Private Const PANEL_BIG_WIDTH As Integer = 132

Private Sub PageScan_tbUniversal_AfterUpdate()
    PageScan_tbUniversal.SetFocus
End Sub

' ==============================================================================
' 2. FORM LIFECYCLE & CLASS SETUP
' ==============================================================================
Private Sub UserForm_Initialize()
    Me.MutiPage.Style = fmTabStyleNone
    Me.DataEntrylbRefNum.Caption = modDatabase.GetNextReferenceNumber()
    UpdateListView
    
    Me.MutiPage.Value = 0
    Me.MutiPage.Value = 1
    
    ' --- SETUP PAIRED HOVER EFFECTS ---
    Set HoverEffectCollection = New Collection
    
    ' Link the Frame, the Label, and the Image Icon together into one group!
    ' (Double check these names match exactly what is in your Properties window)
    Call RegisterHoverGroup(Me.btnFrHome, Me.btnlb1, Me.IconHome)
    Call RegisterHoverGroup(Me.btnFrCheck, Me.btnlb2, Me.IconCheck)
    Call RegisterHoverGroup(Me.btnFrPrint, Me.btnlb3, Me.IconPrint)
    Call RegisterHoverGroup(Me.btnFrReport, Me.btnlb4, Me.IconReport)
    Call RegisterHoverGroup(Me.btnFrMaintain, Me.btnlb5, Me.IconMaintenance) ' Note: Frame might be named btnFrMaintain or btnFrMaintenance
    Call RegisterHoverGroup(Me.btnFrFollowUp, Me.btnlb6, Me.IconFollowUp)
    ' ---------------------------------
    
    Call modDatabase.LoadRecentRefNumbers(Me.PageScan_cbRefnum)
    Call modDatabase.LoadRecentRefNumbers(Me.PR_cbShelvingWS)
    Call modDatabase.LoadRecentRefNumbers(Me.CC_cbRefnum)
End Sub

Private Sub UserForm_Terminate()
    ' Clean up memory when the form closes
    Set HoverEffectCollection = Nothing
End Sub

' --- Helper sub to easily build the Class Module collection ---
Private Sub RegisterHoverGroup(ByRef frmCtrl As MSForms.Frame, ByRef lblCtrl As MSForms.Label, ByRef iconCtrl As MSForms.Image)
    Dim btnHandler As clsHoverBtn
    Set btnHandler = New clsHoverBtn
    
    ' Send the 3 controls and a reference to THIS form (Me) to the class module
    btnHandler.Initialize frmCtrl, lblCtrl, iconCtrl, Me
    HoverEffectCollection.Add btnHandler
End Sub

' ==============================================================================
' 3. UI NAVIGATION & VISUALS
' ==============================================================================

' --- Panel Resizing Helpers ---
Public Sub PannelSizeSmall()
    Me.FrPannel.Width = PANEL_SMALL_WIDTH
End Sub

Public Sub PannelSizeBig()
    Me.FrPannel.Width = PANEL_BIG_WIDTH
End Sub

' --- Global Navigation Click Handler ---
Private Sub PannelClicked(ByVal btnIndex As Integer)
    ' 1. Lock the hover effect so it doesn't bounce back open
    DisableHover = True
    
    ' 2. Change the page
    Me.MutiPage.Value = btnIndex
    
    ' 3. Reset the colors and shrink the panel
    Call BtnColorReset
    Call PannelSizeSmall
    
    ' 4. Let Windows process the resize, then unlock the hover effect
    DoEvents
    DisableHover = False
End Sub

' --- Reset Button Colors (Must be PUBLIC so clsHoverBtn can trigger it) ---
Public Sub BtnColorReset()
    Dim defaultColor As Long
    defaultColor = RGB(221, 244, 255)
    
    ' Reset Frames
    Me.btnFrHome.BackColor = defaultColor
    Me.btnFrCheck.BackColor = defaultColor
    Me.btnFrPrint.BackColor = defaultColor
    Me.btnFrReport.BackColor = defaultColor
    Me.btnFrMaintain.BackColor = defaultColor
    Me.btnFrFollowUp.BackColor = defaultColor
    
    ' Reset Labels
    Me.btnlb1.BackColor = defaultColor
    Me.btnlb2.BackColor = defaultColor
    Me.btnlb3.BackColor = defaultColor
    Me.btnlb4.BackColor = defaultColor
    Me.btnlb5.BackColor = defaultColor
    Me.btnlb6.BackColor = defaultColor
End Sub

' --- Top Level Buttons ---
Private Sub AddbtnHelper_Click():       frmFindTag.Show: End Sub
Private Sub ChecklbbtnStart_Click():    frmChecking.Show: End Sub
Private Sub CheckFrbtnStart_Click():    frmChecking.Show: End Sub
Private Sub PrintlbbtnPrint_Click():    Call PrintbtnPrintReport: End Sub
Private Sub PrintFrbtnPrint_Click():    Call PrintbtnPrintReport: End Sub

' --- Navigation Clicks (Home = 0, Check = 1, Print = 2, Report = 3, Maintain = 4, FollowUp = 5) ---
Private Sub btnlb1_Click(): Call PannelClicked(0): End Sub
Private Sub btnlb2_Click(): Call PannelClicked(1): End Sub
Private Sub btnlb3_Click(): Call PannelClicked(2): End Sub
Private Sub btnlb4_Click(): Call PannelClicked(3): End Sub
Private Sub btnlb5_Click(): Call PannelClicked(4): End Sub
Private Sub btnlb6_Click(): Call PannelClicked(5): End Sub

Private Sub btnFrHome_Click():        Call PannelClicked(0): End Sub
Private Sub btnFrCheck_Click():       Call PannelClicked(1): End Sub
Private Sub btnFrPrint_Click():       Call PannelClicked(2): End Sub
Private Sub btnFrReport_Click():      Call PannelClicked(3): End Sub
Private Sub btnFrMaintain_Click():    Call PannelClicked(4): End Sub
Private Sub btnFrFollowUp_Click():    Call PannelClicked(5): End Sub

Private Sub IconHome_Click():         Call PannelClicked(0): End Sub
Private Sub IconCheck_Click():        Call PannelClicked(1): End Sub
Private Sub IconPrint_Click():        Call PannelClicked(2): End Sub
Private Sub IconReport_Click():       Call PannelClicked(3): End Sub
Private Sub IconMaintenance_Click():  Call PannelClicked(4): End Sub
Private Sub IconFollowUp_Click():     Call PannelClicked(5): End Sub

' --- Mouse Move Reset for Background Panels ---
' These reset the menu color and size if the mouse slips off the buttons onto the background.
Private Sub FrPannel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: End Sub
Private Sub FrCheck_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: Call PannelSizeSmall: End Sub
Private Sub FrHome_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: Call PannelSizeSmall: End Sub
Private Sub FrMaintain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: Call PannelSizeSmall: End Sub
Private Sub FrPrint_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: Call PannelSizeSmall: End Sub
Private Sub FrReport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: Call PannelSizeSmall: End Sub
Private Sub FrFollowUp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call BtnColorReset: Call PannelSizeSmall: End Sub

' ==============================================================================
' 4. BUSINESS LOGIC: DATA ENTRY
' ==============================================================================

Private Sub btnChangeType_Click()
    Dim userInput As String
    
    With Me.btnChangeType
        If .Caption = "Daily" Then
            userInput = InputBox("To change to Annual mode, please type exactly:" & vbCrLf & vbCrLf & _
                                 "Annual" & vbCrLf & vbCrLf & "(case-sensitive)", "Confirm Annual Mode")
            If userInput = "Annual" Then
                .Caption = "Annual"
                MsgBox "Mode changed to Annual.", vbInformation
            Else
                MsgBox "Change cancelled. Still in Daily mode.", vbInformation
            End If
        ElseIf .Caption = "Annual" Then
            .Caption = "Daily"
            MsgBox "Mode changed back to Daily.", vbInformation
        End If
    End With
End Sub

Private Sub DataEntrytbItemCode_Change()
    If Len(Me.DataEntrytbItemCode.Value) = 1 Then
        scanStartTime = Timer
        Me.DataEntrylbDes.Caption = ""
    End If
End Sub

Private Sub DataEntrytbItemCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim scanText As String, searchCode As String
    Dim lookupResult As Variant
    Dim timeElapsed As Double
    
    If KeyCode = 13 Or KeyCode = 9 Then
        scanText = Me.DataEntrytbItemCode.Value
        If Len(scanText) = 0 Then Exit Sub
        
        KeyCode = 0
        
        ' 1. Check Manual Entry speed limits
        If Me.DataEntryCbManualEntry.Value = False Then
            timeElapsed = Timer - scanStartTime
            If timeElapsed > 0.5 Then
                MsgBox "Manual typing is disabled. Please use the scanner or check 'Manual Entry'.", vbExclamation
                Me.DataEntrytbItemCode.Value = ""
                Exit Sub
            End If
        End If
        
        ' 2. Determine Search Code format & Convert to UPPERCASE
        If Me.DataEntryob_In.Value = True And Len(scanText) >= 10 Then
            searchCode = UCase(Mid(scanText, 5, 6)) ' <--- Wrapped in UCase()
        Else
            searchCode = UCase(scanText)            ' <--- Wrapped in UCase()
        End If
        
        ' 3. Process Lookup
        Me.DataEntrytbItemCode.Value = searchCode
        lookupResult = General.iLookup(searchCode, wsNLTItem.Range("A:B"), 2)
        
        If Not IsEmpty(lookupResult) And lookupResult <> "" Then
            Me.DataEntrylbDes.Caption = lookupResult
            Me.DataEntryQuantity.SetFocus
        Else
            Me.DataEntrylbDes.Caption = "Item Not Found"
        End If
    End If
End Sub

Private Sub DataEntrySubmitBtn_Click()
    Dim varReturnType As String, varEntryMethod As String
    Dim locResult As Variant
    Dim finalLocation As String
    
    ' 1. Basic Validation
    If Trim(Me.DataEntrytbItemCode.Value) = "" Or Not IsNumeric(Me.DataEntryQuantity.Value) Then
        MsgBox "Please check Item Code and Quantity.", vbExclamation
        Exit Sub
    End If

    ' 2. Determine Method and Type Safely
    varReturnType = IIf(Me.DataEntryob_In.Value = True, "IP", "OP")
    varEntryMethod = IIf(Me.DataEntryCbManualEntry.Value = True, "Manual", "Scan")

    ' 3. Safely Lookup Location Code (This prevents the silent crash!)
    locResult = General.iLookup(Me.DataEntrytbItemCode.Value, wsBinShelfLocator.Range("A:B"), 2)
    
    If IsEmpty(locResult) Or IsError(locResult) Then
        finalLocation = "UNKNOWN"
    Else
        finalLocation = CStr(locResult)
    End If

    ' 4. Save to Database
    modDatabase.SaveReturnRecord _
        returnType:=varReturnType, _
        itemCode:=Me.DataEntrytbItemCode.Value, _
        quantity:=CLng(Me.DataEntryQuantity.Value), _
        refNum:=Me.DataEntrylbRefNum.Caption, _
        locationCode:=finalLocation, _
        entryMethod:=varEntryMethod
        
    ' 5. Post-Submit Form Reset
    Me.DataEntrytbItemCode.Value = ""
    Me.DataEntryQuantity.Value = ""
    Me.DataEntrylbDes.Caption = ""
    Me.DataEntryCbManualEntry.Value = False ' Safety reset
    
    ' 6. Refresh the List
    UpdateListView
    Call modDatabase.LoadRecentRefNumbers(Me.PR_cbShelvingWS)
    ' Optional: Force the listview to visually repaint itself immediately
    Me.DataEntryListView.Refresh
    
    Me.DataEntrytbItemCode.SetFocus
End Sub

Private Sub DataEntryCbManualEntry_Click()
    If Me.DataEntryCbManualEntry.Value = True Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Manual entry is for items with damaged barcodes only." & vbCrLf & _
                          "Are you sure you want to enable manual typing?", _
                          vbYesNo + vbExclamation, "Confirm Manual Entry")
        
        If response = vbNo Then
            Me.DataEntryCbManualEntry.Value = False
        Else
            Me.DataEntrytbItemCode.Value = ""
            Me.DataEntrytbItemCode.SetFocus
        End If
    End If
End Sub

Private Sub DataEntryListView_DblClick()
    Dim selectedItem As MSComctlLib.ListItem
    Dim response As VbMsgBoxResult
    Dim recordIdToDelete As Long

    Set selectedItem = Me.DataEntryListView.selectedItem
    If selectedItem Is Nothing Then Exit Sub
    
    response = MsgBox("Are you sure you want to permanently delete the selected record?" & vbCrLf & _
                      "Item: " & selectedItem.SubItems(1), vbYesNo + vbQuestion, "Confirm Deletion")
                      
    If response = vbYes Then
        recordIdToDelete = CLng(selectedItem.Tag)
        modDatabase.DeleteReturnRecord recordIdToDelete
        UpdateListView
    End If
End Sub

Private Sub UpdateListView()
    Dim rs As Object
    Dim li As MSComctlLib.ListItem
    Dim refNum As String
    Dim desc As Variant
    
    refNum = Me.DataEntrylbRefNum.Caption
    If refNum = "" Then Exit Sub
    
    ' 1. Setup ListView Layout
    With Me.DataEntryListView
        .ListItems.Clear
        If .ColumnHeaders.Count = 0 Then
            .View = 3 ' lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .ColumnHeaders.Add , , "Time", 50
            .ColumnHeaders.Add , , "Item Code", 50
            .ColumnHeaders.Add , , "Description", 250
            .ColumnHeaders.Add , , "Qty", 40
            .ColumnHeaders.Add , , "Location", 50
            .ColumnHeaders.Add , , "Status", 60
        End If
    End With
    
    ' 2. Fetch and Populate Data
    Set rs = modDatabase.GetRecordsForRefNum(refNum)
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                ' Format handles Date/Time Nulls well, but we use "" & for the rest
                Set li = Me.DataEntryListView.ListItems.Add(, , Format(rs.fields("ReturnDateTime").Value, "hh:mm:ss"))
                
                ' Add "" & to prevent Invalid Use of Null errors!
                li.Tag = "" & rs.fields("ReturnID").Value
                li.SubItems(1) = "" & rs.fields("ItemCode").Value
                
                ' Lookup Description safely
                desc = General.iLookup(CStr(rs.fields("ItemCode").Value), wsNLTItem.Range("A:B"), 2)
                If IsEmpty(desc) Or IsError(desc) Then desc = "Unknown"
                
                li.SubItems(2) = desc
                li.SubItems(3) = "" & rs.fields("Quantity").Value
                li.SubItems(4) = "" & rs.fields("LocationCode").Value
                li.SubItems(5) = "" & rs.fields("RecordStatus").Value
                
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub

' ==============================================================================
' 5. BUSINESS LOGIC: SCAN PAGE ROUTING
' ==============================================================================

Private Sub PageScan_tbUniversal_Change()
    Dim currentText As String
    Dim waitStart As Double
    
    If isProcessing = True Then Exit Sub
    
    currentText = Me.PageScan_tbUniversal.Value
    If Len(currentText) = 0 Then Exit Sub
    
    If scanStartTime = 0 Then scanStartTime = Timer

    ' Type 1: Shelf Scan
    If InStr(1, currentText, "_") > 0 Then
        If UCase(Mid(currentText, 30, 2)) = "AR" Then
            ProcessUniversalScan currentText
        End If
        Exit Sub
    End If

    ' Type 2: Drug Label Scan OR Manual Fill QR Code
    If Len(currentText) = 6 Or Len(currentText) = 13 Or currentText = "#ManualFill" Then
        waitStart = Timer
        Do While Timer < waitStart + 0.15
            DoEvents
            If Len(Me.PageScan_tbUniversal.Value) <> Len(currentText) Then Exit Sub
        Loop
        
        ProcessUniversalScan Me.PageScan_tbUniversal.Value
    End If
End Sub

Private Sub ProcessUniversalScan(ByVal scanText As String)
    Dim timeElapsed As Double
    Dim limit As Double
    
    isProcessing = True
    timeElapsed = Timer - scanStartTime
    
    limit = IIf(InStr(1, scanText, "_") > 0, 2#, 0.8)
    
    ' Only block slow typing if it isn't our special Manual Fill code
    If timeElapsed > limit And scanText <> "#ManualFill" Then
        MsgBox "Manual entry detected. Please use the scanner.", vbCritical
        GoTo CleanUp
    End If
    
    If InStr(1, scanText, "_") > 0 Then
        HandleShelfScan scanText
    Else
        HandleDrugScan scanText
    End If

CleanUp:
    isProcessing = False
    scanStartTime = 0
    Me.PageScan_tbUniversal.Value = ""
    Me.PageScan_tbUniversal.SetFocus
End Sub

Private Sub HandleShelfScan(ByVal txt As String)
    Dim gtinCode As String, dbItemCode As String, shelfBin As Variant
    Dim matchResult As String
    Dim selectedRef As String
    Dim fkReturnID As Long
    Dim currentScanMethod As String
    
    selectedRef = Me.PageScan_cbRefnum.Value
    If selectedRef = "" Or selectedRef = "Please Select Ref. No." Then
        MsgBox "Please select a Reference Number from the dropdown before scanning the shelf!", vbExclamation
        Me.PageScan_tbUniversal.Value = ""
        Exit Sub
    End If
    
    If targetBin = "" Or targetBin = "BIN NOT FOUND" Then
        MsgBox "Please scan a drug label first!", vbExclamation
        Me.PageScan_tbUniversal.Value = ""
        Exit Sub
    End If
    
    ' Extract what shelf they just scanned
    gtinCode = Mid(txt, 13)
    dbItemCode = modDatabase.GetItemCodeFromGTIN(gtinCode)
    
    If dbItemCode = "" Then
        MsgBox "Hash (" & gtinCode & ") not found in Database.", vbExclamation
        Exit Sub
    End If
    
    shelfBin = General.iLookup(dbItemCode, wsBinShelfLocator.Range("A:B"), 2)
    
    ' --- MANUAL FILL LOGIC ---
    If isManualFill Then
        currentScanMethod = "Manual"
        ' In manual fill, the shelf they scan IS the item they want to shelve, so we look up dbItemCode
        fkReturnID = modDatabase.GetReturnID(selectedRef, dbItemCode)
        
        If fkReturnID > 0 Then
            matchResult = "Match"
            Me.PageScan_lbLocator.Caption = CStr(shelfBin) & " (Manual)"
            Me.PageScan_lbLocator.Font.Size = 72
            Me.PageScan_lbLocator.BackColor = yellow
            
            modDatabase.LogVerificationAttempt dbItemCode, gtinCode, matchResult, CStr(shelfBin), CStr(shelfBin), fkReturnID, currentScanMethod
            modDatabase.UpdateReturnBy fkReturnID
        Else
            matchResult = "Mismatch"
            Me.PageScan_lbLocator.Caption = "NOT IN LOT"
            Me.PageScan_lbLocator.Font.Size = 72
            Me.PageScan_lbLocator.BackColor = red
            
            modDatabase.LogVerificationAttempt dbItemCode, gtinCode, matchResult, "N/A", CStr(shelfBin), fkReturnID, currentScanMethod
        End If
        
        isManualFill = False
        targetBin = ""
        targetItemCode = ""
        Me.PageScan_tbUniversal.SetFocus
        Exit Sub
    End If
    
    ' --- NORMAL VERIFICATION LOGIC ---
    currentScanMethod = "Scan"
    Me.PageScan_lbLocator.Font.Size = 196
    
    ' ---> CRITICAL FIX: Look up the ReturnID using targetItemCode (the drug they are holding) <---
    fkReturnID = modDatabase.GetReturnID(selectedRef, targetItemCode)
    
    If CStr(shelfBin) = targetBin Then
        matchResult = "Match"
        Me.PageScan_lbLocator.BackColor = green
        
        ' Log against the targetItemCode!
        modDatabase.LogVerificationAttempt targetItemCode, gtinCode, matchResult, targetBin, CStr(shelfBin), fkReturnID, currentScanMethod
        
        ' Progress Tracking
        If fkReturnID > 0 Then
            Dim reqScans As Long, doneScans As Long
            modDatabase.CheckScanProgress fkReturnID, reqScans, doneScans
            
            If doneScans < reqScans Then
                Me.PageScan_lbLocator.Font.Size = 112
                Me.PageScan_lbLocator.Caption = targetBin & vbCrLf & "(" & doneScans & "/" & reqScans & ")"
            Else
                Me.PageScan_lbLocator.Caption = targetBin & vbCrLf & "(DONE)"
                modDatabase.UpdateReturnBy fkReturnID
            End If
        End If
        
        targetBin = ""
        targetItemCode = "" ' Reset
    Else
        matchResult = "Mismatch"
        Me.PageScan_lbLocator.Caption = targetBin
        Me.PageScan_lbLocator.BackColor = red
        
        ' ---> CRITICAL FIX: Log the error against targetItemCode <---
        modDatabase.LogVerificationAttempt targetItemCode, gtinCode, matchResult, targetBin, CStr(shelfBin), fkReturnID, currentScanMethod
    End If
    
    Me.PageScan_tbUniversal.SetFocus
End Sub

Private Sub HandleDrugScan(ByVal txt As String)
    Dim itemCode As String, locResult As Variant
    
    ' --- Intercept #ManualFill ---
    If txt = "#ManualFill" Then
        isManualFill = True
        Me.PageScan_lbLocator.Font.Size = 72
        targetBin = "MANUAL PENDING"
        targetItemCode = "" ' Reset since we don't have a drug yet
        Me.PageScan_lbLocator.Caption = "Manual Fill Pending"
        Me.PageScan_lbLocator.BackColor = Pannelblue
        Me.PageScan_tbUniversal.SetFocus
        Exit Sub
    End If
    
    itemCode = IIf(Len(txt) = 6, txt, Mid(txt, 5, 6))
    locResult = General.iLookup(itemCode, wsBinShelfLocator.Range("A:B"), 2)
    
    ' Reset font size
    Me.PageScan_lbLocator.Font.Size = 196
    
    If Not IsEmpty(locResult) And Not IsError(locResult) And locResult <> "" Then
        targetBin = CStr(locResult)
        targetItemCode = itemCode ' <--- NEW: Save the drug they are holding!
        Me.PageScan_lbLocator.Caption = targetBin
        Me.PageScan_lbLocator.BackColor = Mainblue
    Else
        Me.PageScan_lbLocator.Caption = "ERR"
        targetBin = ""
        targetItemCode = "" ' Reset on error
    End If
    Me.PageScan_tbUniversal.SetFocus
End Sub

'------------------------------------------
Private Sub PageScan_cbRefnum_Change()
   If Me.PageScan_cbRefnum.Value <> "Please Select Ref. No." Or Me.PageScan_cbRefnum.Value <> "" Then Me.PageScan_tbUniversal.Enabled = True
End Sub
Private Sub PR_cbShelvingWS_Change()
    If Me.PR_cbShelvingWS.Value <> "Please Select Ref. No." Or Me.PR_cbShelvingWS.Value <> "" Then
        Me.PR_btnPrintShelveWS.Enabled = True
        Me.PR_btnPrintShelveReport.Enabled = True
    End If
End Sub
' --- Triggers when the user picks a Reference Number ---
Private Sub CC_cbRefnum_Change()
    If Me.CC_cbRefnum.Value <> "" And Me.CC_cbRefnum.Value <> "Please Select Ref. No." Then
        UpdateCCListView
    Else
        Me.CC_ListView.ListItems.Clear
    End If
End Sub

' --- Populates the Status ListView ---
Private Sub UpdateCCListView()
    Dim rs As Object
    Dim li As MSComctlLib.ListItem
    Dim refNum As String
    
    ' Trackers for our progress logic
    Dim currentReturnID As Long
    Dim reqScans As Long
    Dim doneScans As Long
    Dim dbResult As String
    
    refNum = Me.CC_cbRefnum.Value
    If refNum = "" Or refNum = "Please Select Ref. No." Then Exit Sub
    
    ' 1. Set up the ListView Columns
    With Me.CC_ListView
        .ListItems.Clear
        If .ColumnHeaders.Count = 0 Then
            .View = 3 ' lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .ColumnHeaders.Add , , "Item Code", 60
            .ColumnHeaders.Add , , "Qty", 40
            .ColumnHeaders.Add , , "Drug Location", 60
            .ColumnHeaders.Add , , "DataEntry By", 60
            .ColumnHeaders.Add , , "Entry Method", 60
            .ColumnHeaders.Add , , "Shelf Location", 60
            .ColumnHeaders.Add , , "Result", 60
            .ColumnHeaders.Add , , "Progress", 60
            .ColumnHeaders.Add , , "Return By", 60
            .ColumnHeaders.Add , , "Scan Time", 120
             ' <--- NEW: Progress Column
        End If
    End With
    
    ' 2. Fetch the joined data
    Set rs = modDatabase.GetCheckStatusData(refNum)
    
    ' 3. Populate the list
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                ' Get the ID to check progress
                currentReturnID = rs.fields("ReturnID").Value
                modDatabase.CheckScanProgress currentReturnID, reqScans, doneScans
                
                Set li = Me.CC_ListView.ListItems.Add(, , "" & rs.fields("ItemCode").Value)
                
                li.SubItems(1) = "" & rs.fields("Quantity").Value
                li.SubItems(2) = "" & rs.fields("LocationCode").Value
                li.SubItems(3) = "" & rs.fields("DataEntryBy").Value
                li.SubItems(4) = "" & rs.fields("EntryMethod").Value
                li.SubItems(5) = "" & rs.fields("TargetBin").Value
                
                ' ---> NEW: SMART RESULT LOGIC <---
                dbResult = "" & rs.fields("VerificationResult").Value
                
                If doneScans >= reqScans Then
                    li.SubItems(6) = "Match"       ' 100% finished!
                ElseIf dbResult = "Mismatch" Then
                    li.SubItems(6) = "Mismatch"    ' They made a mistake on the last scan
                ElseIf doneScans > 0 Then
                    li.SubItems(6) = "Partial"     ' e.g., 1/3 scans done
                Else
                    li.SubItems(6) = "Pending"     ' 0 scans done
                End If
                
                li.SubItems(8) = "" & rs.fields("ScannedBy").Value
                
                If IsNull(rs.fields("ScanDateTime").Value) Then
                    li.SubItems(9) = ""
                Else
                    li.SubItems(9) = Format(rs.fields("ScanDateTime").Value, "hh:mm:ss")
                End If
                
                ' ---> NEW: DISPLAY THE PROGRESS <---
                li.SubItems(7) = doneScans & " / " & reqScans
                
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    ' --- Evaluate Button State for Counter Sign ---
    Dim allMatched As Boolean
    Dim i As Long
    allMatched = True
    
    If Me.CC_ListView.ListItems.Count > 0 Then
        For i = 1 To Me.CC_ListView.ListItems.Count
            ' Now it strictly checks our new smart Result column!
            If Me.CC_ListView.ListItems(i).SubItems(6) <> "Match" Then
                allMatched = False
                Exit For
            End If
        Next i
    Else
        allMatched = False
    End If
    
    Me.CC_btnCounterCheck.Enabled = allMatched
End Sub
'----------------------------------Page Print Report
Private Sub PR_btnPrintShelveWS_Click()
  GenerateShelvingList Me.PR_cbShelvingWS.Value
End Sub

Private Sub PR_btnPrintShelveReport_Click()
    GeneratePrintReport PR_cbShelvingWS.Value
End Sub

' --- Counter Check Button Click Event ---
Private Sub CC_btnCounterCheck_Click()
    Dim response As VbMsgBoxResult
    Dim corpId As String
    Dim refNum As String
    Dim lookupResult As Variant
    
    refNum = Me.CC_cbRefnum.Value
    
    ' 1. First Prompt: Confirm manual checks
    response = MsgBox("Confirm you have checked the manual filled item?", vbYesNo + vbQuestion, "Manual Check Confirmation")
    If response = vbNo Then Exit Sub
    
    ' 2. Second Prompt: Ask for CorpID Signature
    corpId = InputBox("Please enter your CorpID for Signature as Counter Sign Person:", "Counter Sign")
    If Trim(corpId) = "" Then Exit Sub ' Exit if blank
    
    ' 3. Validate CorpID exists in wsNameList Column D
    On Error Resume Next
    lookupResult = Application.Match(corpId, wsNameList.Range("C:C"), 0)
    On Error GoTo 0
    
    If IsError(lookupResult) Then
        MsgBox "Invalid CorpID. Could not find this ID in the system.", vbCritical, "Verification Failed"
        Exit Sub
    End If
    
    ' ---> NEW: SECURITY CHECK (Separation of Duties) <---
    If modDatabase.IsUserAlsoReturner(refNum, corpId) = True Then
        MsgBox "Security Alert: The Counter-Signer cannot be the same person who shelved the items!", vbCritical, "Separation of Duties Failed"
        Exit Sub
    End If
    
    ' 4. If valid, update the database!
    modDatabase.UpdateCounterCheck refNum, corpId
    
    MsgBox "Counter check completed successfully. Signed by: " & corpId, vbInformation, "Success"
    UpdateCCListView
End Sub

Private Sub PageScan_tbUniversal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' 9 = Tab key (sent by scanner), 13 = Enter key
    If KeyCode = 9 Or KeyCode = 13 Then
        
        KeyCode = 0 ' <-- CRITICAL: Tells Excel to ignore the Tab/Enter navigation action
        
        ' Safety Net: If the Change event didn't fire (e.g., non-standard barcode length),
        ' force it to process now that the scanner has finished sending data.
        If Trim(Me.PageScan_tbUniversal.Value) <> "" Then
            ProcessUniversalScan Me.PageScan_tbUniversal.Value
        End If
        
    End If
End Sub
