Attribute VB_Name = "DateRepair"
Option Explicit
' =====================================================================
'  DateRepair - audit and repair of legacy date cells in the Record table
'
'  Run from Alt+F8:
'     1. AuditRecordDates   - builds a "Date Audit" sheet listing every
'                             cell that is text instead of a date, or a
'                             date that contradicts the checking date.
'                             Nothing in Record is changed.
'     2. Review the sheet.  Column "Apply" is pre-filled:
'           Y = safe, mechanical conversion (text -> real date, or an
'               impossible expiry whose day/month swap becomes plausible)
'           N = needs a human decision - change to Y if you agree with
'               the proposed value, or type your own date in "Proposed".
'     3. ApplyDateAudit     - backs up the Record sheet, then writes the
'                             "Y" rows as real date serials and notes the
'                             change in the Remarks column.
'
'  Legacy text such as "20/6/2021" is read as day/month/year (the site
'  convention).  Only the repair routines use that assumption; the entry
'  form never does.
' =====================================================================

Private Const AUDIT_SHEET As String = "Date Audit"
Private Const PLAUSIBLE_MONTHS_AHEAD As Long = 36   ' short-expiry stock: > 3 years is suspicious

' Audit sheet columns
Private Const AC_ROW As Long = 1
Private Const AC_CODE As Long = 2
Private Const AC_LOT As Long = 3
Private Const AC_COL As Long = 4
Private Const AC_CURRENT As Long = 5
Private Const AC_PROPOSED As Long = 6
Private Const AC_REASON As Long = 7
Private Const AC_APPLY As Long = 8
Private Const AC_RESULT As Long = 9

Private Const COL_CHECK As String = "Date of Checking"
Private Const COL_EXPIRY As String = "Use Before Date"
Private Const COL_REF As String = "Reference"
Private Const COL_REMARKS As String = "Remarks"

' ---------------------------------------------------------------------
'  1. AUDIT
' ---------------------------------------------------------------------
Public Sub AuditRecordDates()
    Dim lo As ListObject, body As Range, vals As Variant
    Dim n As Long, r As Long, sheetRow As Long
    Dim cA As Long, cB As Long, cE As Long, cF As Long, cL As Long
    Dim wsA As Worksheet, outRow As Long
    Dim code As String, lot As String
    Dim d As Date, msg As String, t As Date
    Dim haveCheck As Boolean, checkDate As Date, expDate As Date, swapped As Date
    Dim canSwap As Boolean, plausibleSwap As Boolean
    Dim lotIndex As Object, key As String, others As String
    Dim cntY As Long, cntN As Long

    Set lo = wsRecord.ListObjects("Record")
    ClearRecordFilters lo
    Set body = lo.DataBodyRange
    If body Is Nothing Then
        MsgBox "The Record table is empty.", vbInformation
        Exit Sub
    End If

    vals = body.Value2                      ' dates -> Double, text -> String
    n = UBound(vals, 1)
    cA = lo.ListColumns(COL_CHECK).Index
    cB = lo.ListColumns("Item Code").Index
    cE = lo.ListColumns("Lot Number").Index
    cF = lo.ListColumns(COL_EXPIRY).Index
    cL = lo.ListColumns(COL_REF).Index

    ' index of item+lot -> rows holding a real expiry date (for cross-checks)
    Set lotIndex = CreateObject("Scripting.Dictionary")
    For r = 1 To n
        If VarType(vals(r, cF)) = vbDouble Then
            key = LotKey(vals(r, cB), vals(r, cE))
            If Len(key) > 0 Then
                If lotIndex.Exists(key) Then
                    lotIndex(key) = lotIndex(key) & "," & r
                Else
                    lotIndex.Add key, CStr(r)
                End If
            End If
        End If
    Next r

    Set wsA = NewAuditSheet()
    outRow = 1

    For r = 1 To n
        sheetRow = body.Row + r - 1
        code = Trim$(vals(r, cB) & "")
        lot = Trim$(vals(r, cE) & "")

        ' ---- Date of Checking ---------------------------------------------
        haveCheck = False
        Select Case VarType(vals(r, cA))
            Case vbDouble
                haveCheck = True
                checkDate = CDate(vals(r, cA))
            Case vbString
                If TryParseExpiry(vals(r, cA), d, msg, True) Then
                    AddAudit wsA, outRow, sheetRow, code, lot, COL_CHECK, vals(r, cA), d, _
                             "Text converted to a real date (read as day/month/year).", "Y", cntY, cntN
                    haveCheck = True
                    checkDate = d
                Else
                    AddAudit wsA, outRow, sheetRow, code, lot, COL_CHECK, vals(r, cA), Empty, _
                             "Text is not a recognisable date - fix by hand.", "N", cntY, cntN
                End If
        End Select

        ' ---- Use Before Date ----------------------------------------------
        Select Case VarType(vals(r, cF))
            Case vbString
                If TryParseExpiry(vals(r, cF), d, msg, True) Then
                    AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, vals(r, cF), d, _
                             "Text converted to a real date (read as day/month/year).", "Y", cntY, cntN
                Else
                    AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, vals(r, cF), Empty, _
                             "Text is not a recognisable date - fix by hand.", "N", cntY, cntN
                End If

            Case vbDouble
                expDate = CDate(vals(r, cF))
                canSwap = (Day(expDate) <= 12 And Day(expDate) <> Month(expDate))
                If canSwap Then swapped = DateSerial(Year(expDate), Day(expDate), Month(expDate))
                plausibleSwap = False
                If canSwap And haveCheck Then
                    plausibleSwap = (swapped >= checkDate And swapped <= DateAdd("m", PLAUSIBLE_MONTHS_AHEAD, checkDate))
                End If

                If haveCheck And expDate < checkDate Then
                    If plausibleSwap Then
                        AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, ExpiryText(expDate), swapped, _
                                 "Expiry is BEFORE the checking date " & ExpiryText(checkDate) & _
                                 "; day and month look swapped (m/d/y PC).", "Y", cntY, cntN
                    Else
                        AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, ExpiryText(expDate), Empty, _
                                 "Expiry is BEFORE the checking date " & ExpiryText(checkDate) & " - review.", "N", cntY, cntN
                    End If

                ElseIf haveCheck And expDate > DateAdd("m", PLAUSIBLE_MONTHS_AHEAD, checkDate) Then
                    If plausibleSwap Then
                        AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, ExpiryText(expDate), swapped, _
                                 "Expiry is more than " & (PLAUSIBLE_MONTHS_AHEAD \ 12) & " years after checking " & _
                                 ExpiryText(checkDate) & "; the swapped date would be plausible.", "N", cntY, cntN
                    Else
                        AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, ExpiryText(expDate), Empty, _
                                 "Expiry is more than " & (PLAUSIBLE_MONTHS_AHEAD \ 12) & " years after checking " & _
                                 ExpiryText(checkDate) & " - review.", "N", cntY, cntN
                    End If

                ElseIf canSwap Then
                    ' same item + lot recorded elsewhere with the swapped date?
                    others = RowsWithDate(lotIndex, LotKey(vals(r, cB), vals(r, cE)), vals, cA, cF, CDbl(swapped), r, body.Row)
                    If Len(others) > 0 Then
                        AddAudit wsA, outRow, sheetRow, code, lot, COL_EXPIRY, ExpiryText(expDate), Empty, _
                                 "Same item + lot is recorded as " & ExpiryText(swapped) & " on sheet row(s) " & others & _
                                 " - one of them has day/month swapped. Review.", "N", cntY, cntN
                    End If
                End If
        End Select

        ' ---- Reference time stamp -----------------------------------------
        If VarType(vals(r, cL)) = vbString Then
            If TryParseStamp(vals(r, cL), t) Then
                AddAudit wsA, outRow, sheetRow, code, lot, COL_REF, vals(r, cL), t, _
                         "Text time stamp converted to a real date-time (read as day/month/year).", "Y", cntY, cntN
            Else
                AddAudit wsA, outRow, sheetRow, code, lot, COL_REF, vals(r, cL), Empty, _
                         "Text is not a recognisable date-time - fix by hand.", "N", cntY, cntN
            End If
        End If
    Next r

    wsA.Columns.AutoFit
    wsA.Activate
    MsgBox "Audit complete: " & (cntY + cntN) & " cell(s) listed on sheet '" & AUDIT_SHEET & "'." & vbCrLf & vbCrLf & _
           "   Apply = Y (safe, mechanical): " & cntY & vbCrLf & _
           "   Apply = N (needs review):     " & cntN & vbCrLf & vbCrLf & _
           "Review the sheet, change Apply / Proposed as needed, then run ApplyDateAudit.", vbInformation
End Sub

' ---------------------------------------------------------------------
'  2. APPLY
' ---------------------------------------------------------------------
Public Sub ApplyDateAudit()
    Dim wsA As Worksheet, lo As ListObject, last As Long, r As Long
    Dim planned As Long, applied As Long, skipped As Long
    Dim sheetRow As Long, colName As String, target As Range, proposed As Variant
    Dim remarksCol As Long, note As String, bakName As String

    On Error Resume Next
    Set wsA = ThisWorkbook.Worksheets(AUDIT_SHEET)
    On Error GoTo 0
    If wsA Is Nothing Then
        MsgBox "Run AuditRecordDates first.", vbExclamation
        Exit Sub
    End If

    last = wsA.Cells(wsA.Rows.Count, AC_ROW).End(xlUp).Row
    For r = 2 To last
        If UCase$(Trim$(wsA.Cells(r, AC_APPLY).Value & "")) = "Y" And VarType(wsA.Cells(r, AC_PROPOSED).Value2) = vbDouble Then planned = planned + 1
    Next r
    If planned = 0 Then
        MsgBox "No rows are marked Apply = Y with a proposed date.", vbInformation
        Exit Sub
    End If

    If MsgBox("Write " & planned & " corrected date(s) into the Record table?" & vbCrLf & vbCrLf & _
              "A backup copy of the Record sheet is made first, and every change is noted in the Remarks column.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Apply date audit") <> vbYes Then Exit Sub

    Set lo = wsRecord.ListObjects("Record")
    ClearRecordFilters lo
    remarksCol = lo.ListColumns(COL_REMARKS).Range.Column

    ' backup
    bakName = "Record_bak_" & Format$(Now, "yyyymmdd_hhnn")
    Application.DisplayAlerts = False
    wsRecord.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = bakName
    Application.DisplayAlerts = True

    Application.ScreenUpdating = False
    For r = 2 To last
        If UCase$(Trim$(wsA.Cells(r, AC_APPLY).Value & "")) = "Y" Then
            proposed = wsA.Cells(r, AC_PROPOSED).Value2
            If VarType(proposed) = vbDouble Then
                sheetRow = CLng(wsA.Cells(r, AC_ROW).Value)
                colName = wsA.Cells(r, AC_COL).Value & ""
                Set target = wsRecord.Cells(sheetRow, lo.ListColumns(colName).Range.Column)

                ' the row must still hold what the audit saw
                If CurrentText(target.Value2, (colName = COL_REF)) <> (wsA.Cells(r, AC_CURRENT).Value & "") Then
                    wsA.Cells(r, AC_RESULT).Value = "SKIPPED - cell now holds '" & CurrentText(target.Value2, (colName = COL_REF)) & "'"
                    skipped = skipped + 1
                Else
                    note = "[" & ExpiryText(Date) & "] " & colName & " repaired: was '" & wsA.Cells(r, AC_CURRENT).Value & "'"
                    If colName = COL_REF Then
                        WriteStampCell target, CDate(proposed)
                    Else
                        WriteDateCell target, CDate(proposed)
                    End If
                    With wsRecord.Cells(sheetRow, remarksCol)
                        If Len(.Value & "") > 0 Then
                            .Value = .Value & "; " & note
                        Else
                            .Value = note
                        End If
                    End With
                    wsA.Cells(r, AC_RESULT).Value = "Applied " & Format$(Now, "yyyy-mm-dd hh:nn")
                    applied = applied + 1
                End If
            End If
        End If
    Next r

    EnsureRecordDateFormats
    Application.ScreenUpdating = True

    MsgBox applied & " cell(s) repaired, " & skipped & " skipped." & vbCrLf & _
           "Backup sheet: " & bakName & vbCrLf & _
           "Delete the backup sheet once you are satisfied.", vbInformation
End Sub

' ---------------------------------------------------------------------
'  helpers
' ---------------------------------------------------------------------
Private Function NewAuditSheet() As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(AUDIT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = AUDIT_SHEET
    ws.Cells(1, AC_ROW).Value = "Sheet Row"
    ws.Cells(1, AC_CODE).Value = "Item Code"
    ws.Cells(1, AC_LOT).Value = "Lot Number"
    ws.Cells(1, AC_COL).Value = "Column"
    ws.Cells(1, AC_CURRENT).Value = "Current value"
    ws.Cells(1, AC_PROPOSED).Value = "Proposed"
    ws.Cells(1, AC_REASON).Value = "Reason"
    ws.Cells(1, AC_APPLY).Value = "Apply (Y/N)"
    ws.Cells(1, AC_RESULT).Value = "Result"
    ws.Rows(1).Font.Bold = True
    ws.Columns(AC_CURRENT).NumberFormat = "@"
    ws.Columns(AC_LOT).NumberFormat = "@"
    Set NewAuditSheet = ws
End Function

Private Sub AddAudit(ByVal ws As Worksheet, ByRef outRow As Long, ByVal sheetRow As Long, _
                     ByVal code As String, ByVal lot As String, ByVal colName As String, _
                     ByVal currentVal As Variant, ByVal proposed As Variant, ByVal reason As String, _
                     ByVal applyFlag As String, ByRef cntY As Long, ByRef cntN As Long)
    outRow = outRow + 1
    ws.Cells(outRow, AC_ROW).Value = sheetRow
    ws.Cells(outRow, AC_CODE).Value = code
    ws.Cells(outRow, AC_LOT).Value = lot
    ws.Cells(outRow, AC_COL).Value = colName
    ws.Cells(outRow, AC_CURRENT).Value = CurrentText(currentVal, (colName = COL_REF))
    If Not IsEmpty(proposed) Then
        If colName = COL_REF Then
            WriteStampCell ws.Cells(outRow, AC_PROPOSED), CDate(proposed)
        Else
            WriteDateCell ws.Cells(outRow, AC_PROPOSED), CDate(proposed)
        End If
    End If
    ws.Cells(outRow, AC_REASON).Value = reason
    ws.Cells(outRow, AC_APPLY).Value = applyFlag
    If applyFlag = "Y" Then cntY = cntY + 1 Else cntN = cntN + 1
End Sub

' One canonical text for a cell value, used both when auditing and when
' re-checking the cell before it is overwritten.
Private Function CurrentText(ByVal v As Variant, ByVal isStamp As Boolean) As String
    Select Case VarType(v)
        Case vbDouble
            If isStamp Then
                CurrentText = ExpiryText(CDate(v)) & " " & Format$(CDate(v), "hh:nn:ss")
            Else
                CurrentText = ExpiryText(CDate(v))
            End If
        Case vbEmpty
            CurrentText = ""
        Case Else
            CurrentText = CStr(v)
    End Select
End Function

Private Function LotKey(ByVal code As Variant, ByVal lot As Variant) As String
    If Len(Trim$(code & "")) = 0 Or Len(Trim$(lot & "")) = 0 Then Exit Function
    LotKey = UCase$(Trim$(code & "")) & "|" & UCase$(Trim$(lot & ""))
End Function

' Sheet rows (other than rowIdx) whose expiry serial equals target, for
' the given item+lot.  Rows whose own expiry is before their checking date
' are ignored - they are already listed as swapped and are not evidence.
Private Function RowsWithDate(ByVal lotIndex As Object, ByVal key As String, ByRef vals As Variant, _
                              ByVal cA As Long, ByVal cF As Long, ByVal target As Double, _
                              ByVal rowIdx As Long, ByVal firstSheetRow As Long) As String
    Dim parts() As String, i As Long, idx As Long, out As String, broken As Boolean
    If Len(key) = 0 Then Exit Function
    If Not lotIndex.Exists(key) Then Exit Function
    parts = Split(lotIndex(key), ",")
    For i = 0 To UBound(parts)
        idx = CLng(parts(i))
        If idx <> rowIdx Then
            If VarType(vals(idx, cF)) = vbDouble Then
                If vals(idx, cF) = target Then
                    broken = False
                    If VarType(vals(idx, cA)) = vbDouble Then broken = (vals(idx, cF) < vals(idx, cA))
                    If Not broken Then
                        If Len(out) > 0 Then out = out & ", "
                        out = out & (firstSheetRow + idx - 1)
                    End If
                End If
            End If
        End If
    Next i
    RowsWithDate = out
End Function

' "27/11/2024 10:17:12", "27/11/2024 10:17", "27/11/2024 3:05:00 PM", "27/11/2024"
Private Function TryParseStamp(ByVal txt As String, ByRef result As Date) As Boolean
    Dim s As String, p As Long, datePart As String, timePart As String
    Dim d As Date, msg As String, tp() As String, h As Long, m As Long, sec As Long
    Dim ampm As String

    s = Trim$(txt)
    p = InStr(s, " ")
    If p = 0 Then
        datePart = s
    Else
        datePart = Left$(s, p - 1)
        timePart = Trim$(Mid$(s, p + 1))
    End If
    If Not TryParseExpiry(datePart, d, msg, True) Then Exit Function

    If Len(timePart) > 0 Then
        If UCase$(Right$(timePart, 2)) = "AM" Or UCase$(Right$(timePart, 2)) = "PM" Then
            ampm = UCase$(Right$(timePart, 2))
            timePart = Trim$(Left$(timePart, Len(timePart) - 2))
        End If
        tp = Split(timePart, ":")
        If UBound(tp) < 1 Or UBound(tp) > 2 Then Exit Function
        If Not IsNumeric(tp(0)) Or Not IsNumeric(tp(1)) Then Exit Function
        h = CLng(tp(0)): m = CLng(tp(1))
        If UBound(tp) = 2 Then
            If Not IsNumeric(tp(2)) Then Exit Function
            sec = CLng(tp(2))
        End If
        If ampm = "PM" And h < 12 Then h = h + 12
        If ampm = "AM" And h = 12 Then h = 0
        If h < 0 Or h > 23 Or m < 0 Or m > 59 Or sec < 0 Or sec > 59 Then Exit Function
        result = d + TimeSerial(h, m, sec)
    Else
        result = d
    End If
    TryParseStamp = True
End Function
