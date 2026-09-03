Attribute VB_Name = "General"
Option Explicit

Sub ShowEntryForm()
    frmEntry.Show
End Sub

Sub ShowReportForm()
    frmReport.Show
End Sub

' Removes any active filter on the Record table.  Falls back to clearing
' the two fields the old code used if ShowAllData is unavailable.
Public Sub ClearRecordFilters(Optional ByVal lo As ListObject)
    If lo Is Nothing Then Set lo = wsRecord.ListObjects("Record")
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    If Err.Number <> 0 Then
        Err.Clear
        lo.Range.AutoFilter Field:=6
        lo.Range.AutoFilter Field:=9
    End If
    On Error GoTo 0
End Sub

' Applies the unambiguous display formats to the three date columns of
' the Record table so every PC shows dd-mmm-yyyy whatever its regional
' settings.  Safe to run any time; it only changes formatting.
Public Sub EnsureRecordDateFormats()
    With wsRecord.ListObjects("Record")
        .ListColumns("Date of Checking").DataBodyRange.NumberFormat = EXPIRY_NUMBER_FORMAT
        .ListColumns("Use Before Date").DataBodyRange.NumberFormat = EXPIRY_NUMBER_FORMAT
        .ListColumns("Reference").DataBodyRange.NumberFormat = STAMP_NUMBER_FORMAT
    End With
End Sub

' Filter: Use Before Date within the given month (any status).
Sub UBDateRangeFilter(ByVal tgrYear As Integer, ByVal tgrMonth As Integer)
    Dim StartDate As Date, EndDate As Date
    Dim lo As ListObject
    Set lo = wsRecord.ListObjects("Record")
    ClearRecordFilters lo

    StartDate = DateSerial(tgrYear, tgrMonth, 1)
    EndDate = DateSerial(tgrYear, tgrMonth + 1, 1) - 1

    ' Criteria are numeric serials, so they work on every locale - but
    ' only for cells that hold real dates.  Run AuditRecordDates /
    ' ApplyDateAudit once to convert legacy text dates.
    lo.Range.AutoFilter Field:=6, Criteria1:=">=" & CDbl(StartDate), _
                        Operator:=xlAnd, Criteria2:="<=" & CDbl(EndDate)
End Sub

' Filter: Use Before Date on or before the end of the given month and
' status not yet "End".
Sub UBDateStatusFilter(ByVal tgrYear As Integer, ByVal tgrMonth As Integer)
    Dim EndDate As Date
    Dim lo As ListObject
    Set lo = wsRecord.ListObjects("Record")
    ClearRecordFilters lo

    EndDate = DateSerial(tgrYear, tgrMonth + 1, 1) - 1
    lo.Range.AutoFilter Field:=6, Criteria1:="<=" & CDbl(EndDate)
    lo.Range.AutoFilter Field:=9, Criteria1:="<>End"
End Sub

' Marks every visible (filtered) row as "End" and stamps the Reference
' column with the current date-time as a REAL date serial.  The old code
' stored Now in a String, which was written back as text and then
' re-interpreted by Excel using the PC's regional format.
Sub ChangeStatusToEnd()
    Dim rng As Range, cell As Range
    Dim stamp As Date
    stamp = Now

    On Error Resume Next
    Set rng = wsRecord.Range("I2", wsRecord.Range("I1048576").End(xlUp)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    For Each cell In rng
        cell.Value = "End"
        WriteStampCell cell.Offset(0, 3), stamp
    Next cell
End Sub

Sub SortByLoc()
    Dim lo As ListObject
    Set lo = wsRecord.ListObjects("Record")
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns("Bin Shelf").Range, SortOn:=xlSortOnValues, _
                        Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
