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
'   - Annual records are ignored. Off-sequence / ad-hoc Daily entries count
'     the item as done but can never move or skip the rotation.
'
' Replaces the old "bin after the latest saved record" logic, which let one
' off-sequence entry (e.g. 2A8A on 04-Sep-2026) jump the rotation and skip
' every 1*, P*, R*, S*, T* and MF*+ bin.
'====================================================================
Option Explicit

Private Const DB_PATH As String = "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"

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
' Replays all Daily records. Returns state.ok = False on any failure.
'--------------------------------------------------------------------
Private Function LoadCycleState() As CycleState
    Dim st As CycleState
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim inLocator As Object
    Dim conn As Object, rs As Object, data As Variant
    Dim i As Long, code As String

    On Error GoTo Fail

    '--- 1. Locator items (col A = ItemCode, col B = BinShelf, sorted by bin)
    Set ws = wsBinShelfLocator
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo Fail
    st.locItems = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 2)).Value

    Set inLocator = CreateObject("Scripting.Dictionary")
    inLocator.CompareMode = vbTextCompare
    For r = 1 To UBound(st.locItems, 1)
        code = Trim(st.locItems(r, 1) & "")
        If code <> "" Then If Not inLocator.Exists(code) Then inLocator(code) = True
    Next r
    st.totalItems = inLocator.Count
    If st.totalItems = 0 Then GoTo Fail

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
