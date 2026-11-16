Attribute VB_Name = "BinCycle"
'=== Module name: BinCycle ===========================================
' Daily stock-take rotation ("cycle") over every item in wsBinShelfLocator.
'
' Rule (agreed 16-Nov-2026, see Stocktake/BINSHELF_CYCLE_FIX.md):
'   - Replay every Daily record in [Record] in (RecDate, ID) order.
'   - A cycle is complete when every item in BinShelfLocator has at least
'     one Daily count. Counting then starts a new cycle.
'   - Next items = items NOT yet counted in the current cycle, in
'     BinShelfLocator (bin) order.
'   - Annual records are ignored here. Off-sequence / ad-hoc Daily entries
'     count the item as done but can never move or skip the rotation.
'
' Annual Cycle (agreed 16-Nov-2026) is SEPARATE from the Daily Cycle:
'   - One Annual Cycle = the Annual records dated in one Annual Year,
'     1-Jun to 31-May (start month = ANNUAL_START_MONTH).
'   - Complete when every BinShelfLocator item has an Annual count in that
'     Annual Year. Daily counts never count towards it.
'
' Replaces the old "bin after the latest saved record" logic, which let one
' off-sequence entry (e.g. 2A8A on 04-Sep-2026) jump the rotation and skip
' every 1*, P*, R*, S*, T* and MF*+ bin.
'====================================================================
Option Explicit

Private Const DB_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"

' First month of the Annual Year (6 = June: 1-Jun to 31-May next year).
Public Const ANNUAL_START_MONTH As Long = 6

' Result of replaying the Daily records against the current locator list.
Private Type CycleState
    ok As Boolean
    cyclesCompleted As Long
    lastCompletedOn As Variant    ' date the last full cycle finished (Empty if none)
    cycleStartedOn As Variant     ' date of first count in current cycle (Empty if none)
    countedInCycle As Long
    totalItems As Long
    locItems As Variant           ' 2-D array: (row, 1)=ItemCode, (row, 2)=BinShelf
    counted As Object             ' Dictionary: ItemCode -> True (current cycle)
End Type

'--------------------------------------------------------------------
' Reads BinShelfLocator (col A = ItemCode, col B = BinShelf, sorted by bin).
' locItems = 2-D array (row, 1..2); inLocator = Dictionary of ItemCodes.
'--------------------------------------------------------------------
Private Function LoadLocator(ByRef locItems As Variant, ByRef inLocator As Object) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, code As String

    On Error GoTo Fail
    Set ws = wsBinShelfLocator
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then GoTo Fail          ' need header + 2 rows for a 2-D array
    locItems = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 2)).Value

    Set inLocator = CreateObject("Scripting.Dictionary")
    inLocator.CompareMode = vbTextCompare
    For r = 1 To UBound(locItems, 1)
        code = Trim(locItems(r, 1) & "")
        If code <> "" Then If Not inLocator.Exists(code) Then inLocator(code) = True
    Next r
    LoadLocator = (inLocator.Count > 0)
    Exit Function
Fail:
    LoadLocator = False
End Function

'--------------------------------------------------------------------
' Replays all Daily records. Returns state.ok = False on any failure.
'--------------------------------------------------------------------
Private Function LoadCycleState() As CycleState
    Dim st As CycleState
    Dim inLocator As Object
    Dim conn As Object, rs As Object, data As Variant
    Dim i As Long, code As String

    On Error GoTo Fail

    '--- 1. Locator items (col A = ItemCode, col B = BinShelf, sorted by bin)
    If Not LoadLocator(st.locItems, inLocator) Then GoTo Fail
    st.totalItems = inLocator.Count

    Set st.counted = CreateObject("Scripting.Dictionary")
    st.counted.CompareMode = vbTextCompare

    '--- 2. Replay Daily records oldest -> newest
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    Set rs = conn.Execute("SELECT [ItemCode], [RecDate] FROM [Record] " & _
                          "WHERE Type = 'Daily' ORDER BY [RecDate] ASC, [ID] ASC")
    If Not rs.EOF Then data = rs.GetRows()   ' data(field, row)
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing

    If Not IsEmpty(data) Then
        For i = 0 To UBound(data, 2)
            code = Trim(data(0, i) & "")
            If inLocator.Exists(code) Then          ' items no longer in locator are ignored
                If st.counted.Count = 0 Then st.cycleStartedOn = data(1, i)
                If Not st.counted.Exists(code) Then st.counted(code) = True
                If st.counted.Count = st.totalItems Then
                    st.cyclesCompleted = st.cyclesCompleted + 1
                    st.lastCompletedOn = data(1, i)
                    st.counted.RemoveAll
                    st.cycleStartedOn = Empty
                End If
            End If
        Next i
    End If

    st.countedInCycle = st.counted.Count
    st.ok = True
    LoadCycleState = st
    Exit Function

Fail:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    st.ok = False
    LoadCycleState = st
End Function

'--------------------------------------------------------------------
' Next n items not yet counted in the current cycle, in bin order.
' Returns 2-D array (1..k, 1..2) = ItemCode, BinShelf (k <= n),
' or Empty on failure.
'--------------------------------------------------------------------
Public Function GetNextCycleItems(ByVal n As Long) As Variant
    Dim st As CycleState
    Dim r As Long, k As Long, code As String
    Dim picked() As Long, result() As Variant

    If n < 1 Then Exit Function
    st = LoadCycleState()
    If Not st.ok Then Exit Function

    ReDim picked(1 To n)
    For r = 1 To UBound(st.locItems, 1)
        code = Trim(st.locItems(r, 1) & "")
        If code <> "" Then
            If Not st.counted.Exists(code) Then
                k = k + 1
                picked(k) = r
                If k = n Then Exit For
            End If
        End If
    Next r
    If k = 0 Then Exit Function

    ReDim result(1 To k, 1 To 2)
    For r = 1 To k
        result(r, 1) = Trim(st.locItems(picked(r), 1) & "")
        result(r, 2) = Trim(st.locItems(picked(r), 2) & "")
    Next r
    GetNextCycleItems = result
End Function

'--------------------------------------------------------------------
' Number of full cycles completed so far; -1 on failure.
'--------------------------------------------------------------------
Public Function GetCyclesCompleted() As Long
    Dim st As CycleState
    st = LoadCycleState()
    If st.ok Then GetCyclesCompleted = st.cyclesCompleted Else GetCyclesCompleted = -1
End Function

'--------------------------------------------------------------------
' Human-readable progress line, e.g. for MsgBox.
'--------------------------------------------------------------------
Public Function GetCycleSummary() As String
    Dim st As CycleState, s As String
    st = LoadCycleState()
    If Not st.ok Then
        GetCycleSummary = "Cycle status unavailable (database or BinShelfLocator not reachable)."
        Exit Function
    End If
    s = "Daily cycle: " & st.countedInCycle & " of " & st.totalItems & " items counted"
    If Not IsEmpty(st.cycleStartedOn) Then s = s & " (started " & Format(st.cycleStartedOn, "dd-mmm-yyyy") & ")"
    s = s & "." & vbCrLf & "Cycles completed: " & st.cyclesCompleted
    If Not IsEmpty(st.lastCompletedOn) Then s = s & " (last finished " & Format(st.lastCompletedOn, "dd-mmm-yyyy") & ")"
    GetCycleSummary = s
End Function

'--------------------------------------------------------------------
' Call after a save. Notifies if that save finished a cycle.
'--------------------------------------------------------------------
Public Sub NotifyIfCycleCompleted(ByVal cyclesBeforeSave As Long)
    Dim st As CycleState
    If cyclesBeforeSave < 0 Then Exit Sub
    st = LoadCycleState()
    If Not st.ok Then Exit Sub
    If st.cyclesCompleted > cyclesBeforeSave Then
        MsgBox "Daily stock-take cycle complete - every bin has been counted." & vbCrLf & _
               "Finished on " & Format(st.lastCompletedOn, "dd-mmm-yyyy") & "." & vbCrLf & _
               "The next print starts a new cycle from the first bin.", vbInformation
    End If
End Sub

'====================================================================
' ANNUAL CYCLE (separate from the Daily Cycle)
'====================================================================

'--------------------------------------------------------------------
' 1-Jun (ANNUAL_START_MONTH) that starts the Annual Year containing d.
'--------------------------------------------------------------------
Public Function AnnualYearStart(ByVal d As Date) As Date
    If Month(d) >= ANNUAL_START_MONTH Then
        AnnualYearStart = DateSerial(Year(d), ANNUAL_START_MONTH, 1)
    Else
        AnnualYearStart = DateSerial(Year(d) - 1, ANNUAL_START_MONTH, 1)
    End If
End Function

Private Function AnnualYearLabel(ByVal yearStart As Date) As String
    AnnualYearLabel = Format(yearStart, "dd-mmm-yyyy") & " to " & _
                      Format(DateAdd("d", -1, DateAdd("yyyy", 1, yearStart)), "dd-mmm-yyyy")
End Function

'--------------------------------------------------------------------
' Annual counts of the Annual Year containing anyDateInYear.
' Dates are compared in VBA (not as SQL literals) so d/m vs m/d
' regional settings cannot shift the window.
'--------------------------------------------------------------------
Private Function LoadAnnualState(ByVal anyDateInYear As Date) As CycleState
    Dim st As CycleState
    Dim inLocator As Object
    Dim conn As Object, rs As Object, data As Variant
    Dim i As Long, code As String
    Dim yStart As Date, yEnd As Date, recDt As Variant

    On Error GoTo Fail

    If Not LoadLocator(st.locItems, inLocator) Then GoTo Fail
    st.totalItems = inLocator.Count
    Set st.counted = CreateObject("Scripting.Dictionary")
    st.counted.CompareMode = vbTextCompare

    yStart = AnnualYearStart(anyDateInYear)
    yEnd = DateAdd("yyyy", 1, yStart)            ' exclusive
    st.cycleStartedOn = yStart

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & ";"
    Set rs = conn.Execute("SELECT [ItemCode], [RecDate] FROM [Record] " & _
                          "WHERE Type = 'Annual' ORDER BY [RecDate] ASC, [ID] ASC")
    If Not rs.EOF Then data = rs.GetRows()
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing

    If Not IsEmpty(data) Then
        For i = 0 To UBound(data, 2)
            recDt = data(1, i)
            If IsDate(recDt) Then
                If CDate(recDt) >= yStart And CDate(recDt) < yEnd Then
                    code = Trim(data(0, i) & "")
                    If inLocator.Exists(code) Then
                        If Not st.counted.Exists(code) Then st.counted(code) = True
                    End If
                End If
            End If
        Next i
    End If

    st.countedInCycle = st.counted.Count
    If st.countedInCycle = st.totalItems Then st.cyclesCompleted = 1
    st.ok = True
    LoadAnnualState = st
    Exit Function

Fail:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    st.ok = False
    LoadAnnualState = st
End Function

Private Function DefaultDate(Optional ByVal anyDateInYear As Variant) As Date
    ' A missing Optional arrives here as an Error value, hence IsError.
    If IsMissing(anyDateInYear) Or IsError(anyDateInYear) Or IsEmpty(anyDateInYear) Then
        DefaultDate = Date
    Else
        DefaultDate = CDate(anyDateInYear)
    End If
End Function

'--------------------------------------------------------------------
' Items with no Annual count in the Annual Year, in bin order.
' 2-D array (1..k, 1..2) = ItemCode, BinShelf; Empty if none or on failure.
'--------------------------------------------------------------------
Public Function GetAnnualUncountedItems(Optional ByVal anyDateInYear As Variant) As Variant
    Dim st As CycleState, r As Long, k As Long, code As String
    Dim result() As Variant

    st = LoadAnnualState(DefaultDate(anyDateInYear))
    If Not st.ok Then Exit Function
    If st.totalItems - st.countedInCycle <= 0 Then Exit Function

    ReDim result(1 To st.totalItems - st.countedInCycle, 1 To 2)
    For r = 1 To UBound(st.locItems, 1)
        code = Trim(st.locItems(r, 1) & "")
        If code <> "" Then
            If Not st.counted.Exists(code) Then
                k = k + 1
                If k > UBound(result, 1) Then Exit For   ' duplicate codes in locator
                result(k, 1) = code
                result(k, 2) = Trim(st.locItems(r, 2) & "")
                st.counted(code) = True                  ' list each code once
            End If
        End If
    Next r
    GetAnnualUncountedItems = result
End Function

'--------------------------------------------------------------------
' Items still to count in the Annual Year; -1 on failure.
'--------------------------------------------------------------------
Public Function GetAnnualRemaining(Optional ByVal anyDateInYear As Variant) As Long
    Dim st As CycleState
    st = LoadAnnualState(DefaultDate(anyDateInYear))
    If st.ok Then GetAnnualRemaining = st.totalItems - st.countedInCycle Else GetAnnualRemaining = -1
End Function

Public Function GetAnnualSummary(Optional ByVal anyDateInYear As Variant) As String
    Dim st As CycleState
    st = LoadAnnualState(DefaultDate(anyDateInYear))
    If Not st.ok Then
        GetAnnualSummary = "Annual status unavailable (database or BinShelfLocator not reachable)."
        Exit Function
    End If
    GetAnnualSummary = "Annual cycle " & AnnualYearLabel(st.cycleStartedOn) & ": " & _
                       st.countedInCycle & " of " & st.totalItems & " items counted, " & _
                       (st.totalItems - st.countedInCycle) & " remaining" & _
                       IIf(st.cyclesCompleted = 1, " - COMPLETE.", ".")
End Function

'--------------------------------------------------------------------
' Call after an Annual save with GetAnnualRemaining() taken before it.
' Notifies only when the save brought the remaining count to zero.
'--------------------------------------------------------------------
Public Sub NotifyIfAnnualCompleted(ByVal remainingBeforeSave As Long)
    If remainingBeforeSave <= 0 Then Exit Sub
    If GetAnnualRemaining() = 0 Then
        MsgBox "Annual stock take complete - every bin has an Annual count for " & _
               AnnualYearLabel(AnnualYearStart(Date)) & ".", vbInformation
    End If
End Sub

'--------------------------------------------------------------------
' Writes the items not yet Annual-counted to wsReport and shows the
' summary. Pass any date in the wanted Annual Year (default: today),
' e.g. ReportAnnualUncounted DateSerial(2026, 3, 1) for 2025-26.
'--------------------------------------------------------------------
Public Sub ReportAnnualUncounted(Optional ByVal anyDateInYear As Variant)
    Dim items As Variant, d As Date
    d = DefaultDate(anyDateInYear)
    items = GetAnnualUncountedItems(d)

    wsReport.Cells.Clear
    wsReport.Range("A1:B1").Value = Array("ItemCode", "BinShelf")
    If Not IsEmpty(items) Then
        wsReport.Range("A2").Resize(UBound(items, 1), 2).Value = items
    End If
    wsReport.Columns.AutoFit
    MsgBox GetAnnualSummary(d), vbInformation
End Sub

Public Sub ShowAnnualStatus()
    ' Current and previous Annual Year
    MsgBox GetAnnualSummary(Date) & vbCrLf & vbCrLf & _
           GetAnnualSummary(DateAdd("yyyy", -1, Date)), vbInformation
End Sub
