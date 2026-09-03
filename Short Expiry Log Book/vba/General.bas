Attribute VB_Name = "General"
Sub ShowEntryForm()
    frmEntry.Show
End Sub
Sub ShowReportForm()
    frmReport.Show
End Sub


Sub UBDateRangeFilter(ByVal tgrYear As Integer, ByVal tgrMonth As Integer)
    Dim StartDate, EndDate As Date
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=6
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=9
    
    StartDate = DateSerial(tgrYear, tgrMonth, 1)
    EndDate = DateSerial(tgrYear, tgrMonth + 1, 1) - 1
    
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=6, Criteria1:=">=" & CDbl(StartDate), Operator:=xlAnd, Criteria2:="<=" & CDbl(EndDate)
    
End Sub


Sub UBDateStatusFilter(ByVal tgrYear As Integer, ByVal tgrMonth As Integer)
    Dim StartDate, EndDate As Date
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=6
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=9
    
    EndDate = DateSerial(tgrYear, tgrMonth + 1, 1) - 1
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=6, Criteria1:="<=" & CDbl(EndDate)
    wsRecord.ListObjects("Record").Range.AutoFilter Field:=9, Criteria1:="<>End"
    
End Sub

Sub ChangeStatusToEnd()
    Dim rng As Range
    Dim Ref As String
    Ref = Now
    For Each rng In wsRecord.Range("I2", wsRecord.Range("I1048576").End(xlUp)).SpecialCells(xlCellTypeVisible)
        rng.Value = "End"
        rng.Offset(0, 3).Value = Ref
    Next
End Sub

Sub SortByLoc()

'

'
    wsRecord.ListObjects("Record").Sort.SortFields.Clear
    wsRecord.ListObjects("Record").Sort.SortFields.Add Key:=Range("Record[[#All],[Bin Shelf]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsRecord.ListObjects("Record").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


