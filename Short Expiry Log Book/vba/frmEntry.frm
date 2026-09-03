VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntry 
   Caption         =   "Entry Form"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865.001
   OleObjectBlob   =   "frmEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' =====================================================================
'  frmEntry - Short Expiry entry form
'
'  Date integrity rules (see DateSafe module):
'    * The expiry typed into cbUBDate is parsed ONLY by TryParseExpiry.
'      Ambiguous input such as 11/1/2026 is refused.
'    * Once parsed, the box is rewritten as dd-Mmm-yyyy and a preview
'      (weekday + days to expiry) is shown next to it.
'    * The staging list keeps the date serial in a hidden column, so what
'      is saved is the number the user confirmed, not re-parsed text.
'    * Dates are written to the sheet with WriteDateCell (Value2 = Double),
'      never as text.
' =====================================================================

' ListToAdd columns
Private Const LC_CODE As Long = 0
Private Const LC_LOC As Long = 1
Private Const LC_LOT As Long = 2
Private Const LC_EXPIRY As Long = 3
Private Const LC_QTY As Long = 4
Private Const LC_SERIAL As Long = 5     ' hidden: CLng(expiry date)

Private mPreview As MSForms.Label
Private mUpdatingDateBox As Boolean

' ---------------------------------------------------------------------
'  Form setup
' ---------------------------------------------------------------------
Private Sub UserForm_Initialize()
    SetupStagingList
    SetupPreviewLabel
    Me.cbUBDate.ControlTipText = "ddmmyyyy  or  dd-Mmm-yyyy  (e.g. 01112026 / 01-Nov-2026)"
End Sub

Private Sub SetupStagingList()
    Dim widths As String, w As Single, i As Long
    With Me.ListToAdd
        If .ColumnCount < LC_SERIAL + 1 Then
            .ColumnCount = LC_SERIAL + 1
            widths = Trim$(.ColumnWidths)
            If Len(Replace(widths, ";", "")) > 0 Then
                widths = widths & ";0 pt"
            Else
                w = (.Width - 20) / LC_SERIAL
                widths = ""
                For i = 1 To LC_SERIAL
                    widths = widths & Format$(w, "0") & " pt;"
                Next i
                widths = widths & "0 pt"
            End If
            .ColumnWidths = widths
        End If
    End With
End Sub

' Uses a design-time label named lbDatePreview if one exists, otherwise
' creates one at run time to the right of the expiry box.
Private Sub SetupPreviewLabel()
    On Error Resume Next
    Set mPreview = Me.Controls("lbDatePreview")
    On Error GoTo 0
    If mPreview Is Nothing Then
        Set mPreview = Me.Controls.Add("Forms.Label.1", "lbDatePreview", True)
        With mPreview
            .Left = Me.cbUBDate.Left + Me.cbUBDate.Width + 6
            .Top = Me.cbUBDate.Top + 2
            .Width = Me.InsideWidth - .Left - 6
            .Height = Me.cbUBDate.Height
            .WordWrap = False
            .Font.Bold = True
        End With
    End If
    ShowPreview "", True
End Sub

Private Sub ShowPreview(ByVal txt As String, ByVal ok As Boolean)
    If mPreview Is Nothing Then Exit Sub
    mPreview.Caption = txt
    If ok Then
        mPreview.ForeColor = RGB(0, 110, 0)
    Else
        mPreview.ForeColor = RGB(190, 0, 0)
    End If
End Sub

' ---------------------------------------------------------------------
'  Expiry date box
' ---------------------------------------------------------------------
Private Sub cbUBDate_Change()
    If mUpdatingDateBox Then Exit Sub
    RefreshDatePreview False
End Sub

Private Sub cbUBDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RefreshDatePreview True
End Sub

' canonicalise:=True rewrites the box as dd-Mmm-yyyy once the text parses.
' While typing, only the 8-digit ddmmyyyy shortcut is rewritten (as before).
Private Sub RefreshDatePreview(ByVal canonicalise As Boolean)
    Dim s As String, d As Date, msg As String
    s = Trim$(Me.cbUBDate.Value)
    If Len(s) = 0 Then
        ShowPreview "", True
        Exit Sub
    End If
    If TryParseExpiry(s, d, msg) Then
        ShowPreview ExpiryPreview(d), (ExpiryPolicyWarning(d) = "")
        If canonicalise Or (Len(s) = 8 And IsNumeric(s)) Then
            If s <> ExpiryText(d) Then
                mUpdatingDateBox = True
                Me.cbUBDate.Value = ExpiryText(d)
                mUpdatingDateBox = False
            End If
        End If
    Else
        If canonicalise Then
            ShowPreview FirstLine(msg), False
        Else
            ShowPreview "...", False
        End If
    End If
End Sub

Private Function FirstLine(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbCrLf)
    If p > 0 Then FirstLine = Left$(s, p - 1) Else FirstLine = s
End Function

' ---------------------------------------------------------------------
'  Item code lookup
' ---------------------------------------------------------------------
Private Sub cbItemCode_Change()
    Dim v As Variant
    Me.lbItemDes.Caption = ""
    If Len(Me.cbItemCode.Value) = 6 Then
        v = Application.VLookup(Me.cbItemCode.Value, wsLoc.Range("A:C"), 2, 0)
        If Not IsError(v) Then
            Me.lbItemDes.Caption = v
            On Error Resume Next
            Me.cbLoc.List = ItemLocLookUp(Me.cbItemCode.Value)
            On Error GoTo 0
        End If
    End If
End Sub

' ---------------------------------------------------------------------
'  Add to staging list
' ---------------------------------------------------------------------
Private Sub btnAdd_Click()
    Dim d As Date, warn As String, iCount As Long

    If Not ValidateInputs(d) Then Exit Sub

    warn = ExpiryPolicyWarning(d)
    If Len(warn) > 0 Then
        If MsgBox(warn & vbCrLf & vbCrLf & "Add it anyway?", vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Check the expiry date") <> vbYes Then
            Me.cbUBDate.SetFocus
            Exit Sub
        End If
    End If

    If Len(Me.lbItemDes.Caption) = 0 Then
        If MsgBox("Item code '" & Me.cbItemCode.Value & "' is not in the ItemLocation list." & vbCrLf & _
                  "Add it anyway (Item Name will be blank)?", vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Unknown item code") <> vbYes Then
            Me.cbItemCode.SetFocus
            Exit Sub
        End If
    End If

    With Me.ListToAdd
        iCount = .ListCount
        .AddItem
        .List(iCount, LC_CODE) = Trim$(Me.cbItemCode.Value)
        .List(iCount, LC_LOC) = Trim$(Me.cbLoc.Value)
        .List(iCount, LC_LOT) = Trim$(Me.cbLot.Value)
        .List(iCount, LC_EXPIRY) = ExpiryText(d)
        .List(iCount, LC_QTY) = Trim$(Me.cbQty.Value)
        .List(iCount, LC_SERIAL) = CStr(CLng(d))
    End With

    Me.cbItemCode.Value = ""
    Me.cbLoc.Value = ""
    Me.cbLot.Value = ""
    Me.cbUBDate.Value = ""
    Me.cbQty.Value = ""
    ShowPreview "", True
    Me.cbItemCode.SetFocus
End Sub

Private Function ValidateInputs(ByRef d As Date) As Boolean
    Dim msg As String
    If Len(Trim$(Me.cbItemCode.Value)) <> 6 Then
        Fail "Item code must be 6 characters.", Me.cbItemCode
    ElseIf Len(Trim$(Me.cbLoc.Value)) = 0 Then
        Fail "Please select or type a location.", Me.cbLoc
    ElseIf Len(Trim$(Me.cbLot.Value)) = 0 Then
        Fail "Lot number is required.", Me.cbLot
    ElseIf Not TryParseExpiry(Me.cbUBDate.Value, d, msg) Then
        Fail msg, Me.cbUBDate
    ElseIf Not IsNumeric(Me.cbQty.Value) Then
        Fail "Quantity must be a number.", Me.cbQty
    ElseIf Val(Me.cbQty.Value) <= 0 Then
        Fail "Quantity must be greater than 0.", Me.cbQty
    Else
        ValidateInputs = True
    End If
End Function

Private Sub Fail(ByVal msg As String, ByVal ctl As MSForms.Control)
    MsgBox msg, vbExclamation, "Cannot add"
    On Error Resume Next
    ctl.SetFocus
    On Error GoTo 0
End Sub

Private Sub ListToAdd_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListToAdd.ListIndex >= 0 Then Me.ListToAdd.RemoveItem Me.ListToAdd.ListIndex
End Sub

' ---------------------------------------------------------------------
'  Save to the Record table  (all-or-nothing)
' ---------------------------------------------------------------------
Private Sub btnSubmit_Click()
    Dim lo As ListObject, rowRng As Range, addedRows As Collection
    Dim i As Long, d As Date, code As String, itemName As Variant
    Dim today As Date, errText As String

    If Me.ListToAdd.ListCount = 0 Then
        MsgBox "Nothing to save - add at least one item first.", vbInformation
        Exit Sub
    End If

    ' Every staged row must carry a valid serial before anything is written.
    For i = 0 To Me.ListToAdd.ListCount - 1
        If Not StagedDate(i, d) Then
            MsgBox "Row " & (i + 1) & " has no valid expiry date (" & Me.ListToAdd.List(i, LC_EXPIRY) & _
                   "). Remove it (double-click) and add it again.", vbExclamation
            Exit Sub
        End If
    Next i

    Set lo = wsRecord.ListObjects("Record")
    Set addedRows = New Collection
    today = Date

    On Error GoTo Failed
    Application.ScreenUpdating = False
    ClearRecordFilters lo

    For i = 0 To Me.ListToAdd.ListCount - 1
        StagedDate i, d
        code = Me.ListToAdd.List(i, LC_CODE)
        itemName = Application.VLookup(code, wsLoc.Range("A:B"), 2, 0)
        If IsError(itemName) Then itemName = ""

        Set rowRng = NextRecordRow(lo)
        addedRows.Add rowRng
        WriteDateCell rowRng.Cells(1, 1), today                    ' Date of Checking
        rowRng.Cells(1, 2).Value = code                             ' Item Code
        rowRng.Cells(1, 3).Value = itemName                         ' Item Name
        rowRng.Cells(1, 4).Value = Me.ListToAdd.List(i, LC_LOC)     ' Bin Shelf
        rowRng.Cells(1, 5).Value = Me.ListToAdd.List(i, LC_LOT)     ' Lot Number
        WriteDateCell rowRng.Cells(1, 6), d                         ' Use Before Date
        rowRng.Cells(1, 7).Value = Me.ListToAdd.List(i, LC_QTY)     ' Qty on hand
        rowRng.Cells(1, 8).Value = Environ$("username")             ' Entered by
    Next i

    Application.ScreenUpdating = True
    MsgBox Me.ListToAdd.ListCount & " record(s) saved.", vbInformation
    Unload Me
    Exit Sub

Failed:
    errText = Err.Description
    On Error Resume Next
    For i = addedRows.Count To 1 Step -1
        addedRows(i).ClearContents
        lo.ListRows(addedRows(i).Row - lo.HeaderRowRange.Row).Delete
    Next i
    Application.ScreenUpdating = True
    MsgBox "Nothing was saved. Error while writing to the Record table:" & vbCrLf & errText, vbCritical
End Sub

' Serial stored in the hidden column; falls back to re-parsing the
' canonical text (only happens if the list was populated by older code).
Private Function StagedDate(ByVal idx As Long, ByRef d As Date) As Boolean
    Dim s As String, msg As String
    s = Trim$(Me.ListToAdd.List(idx, LC_SERIAL) & "")
    If Len(s) > 0 And IsNumeric(s) Then
        d = CDate(CLng(s))
        StagedDate = True
    Else
        StagedDate = TryParseExpiry(Me.ListToAdd.List(idx, LC_EXPIRY) & "", d, msg)
    End If
End Function

' First blank row at the bottom of the table, or a newly added one.
Private Function NextRecordRow(ByVal lo As ListObject) As Range
    Dim lastRow As Range
    If lo.ListRows.Count > 0 Then
        Set lastRow = lo.ListRows(lo.ListRows.Count).Range
        If Application.WorksheetFunction.CountA(lastRow) = 0 Then
            Set NextRecordRow = lastRow
            Exit Function
        End If
    End If
    Set NextRecordRow = lo.ListRows.Add.Range
End Function
