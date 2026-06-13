Attribute VB_Name = "MReport"


Sub ReportSortByColor(ByVal rng As Range)
    wsReport.AutoFilter.Sort.SortFields.Clear
    wsReport.AutoFilter.Sort.SortFields.Add2 Key:=rng, SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:= _
        xlSortNormal
    wsReport.AutoFilter.Sort.SortFields.Add(rng, xlSortOnCellColor, xlDescending, , xlSortNormal).SortOnValue.Color = RGB(255, 199, 206)
    With wsReport.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub




Public Sub DataEntryForReport(ByVal RefNum As String)
   Call wsUnProtectAll
   
   Call GetRecordDetail(RefNum)
    
    Dim lastRow As Long
    lastRow = wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row
    
    Call ReportSortByColor(wsReport.Range("A8:A" & lastRow))
    wsReport.Range("F6").value = "Reference No:" & RefNum
    'Call DataEntryForFollowUp(RefNum)
    wsReport.AutoFilter.Sort.SortFields.Clear
    Call wsProtectAll
End Sub



Function IsAllMatch(ByVal RefNum As String) As Boolean
    
    Dim iRow, ItemCount As Long
    ItemCount = Application.WorksheetFunction.CountIf(wsRecord.Range("G:G"), RefNum)
    iRow = wsRecord.Range("G:G").Find(what:=RefNum).Row
    For i = 1 To ItemCount
        If Not wsRecord.Cells(iRow + i - 1, 5).value = "Matched" Then
            IsAllMatch = False
            Exit Function
        End If
    Next
    IsAllMatch = True
    
End Function

Sub PrintReport(ByVal RefNum As String)
    If RefNum = "" Then Exit Sub
        'If IsAllMatch(RefNum) = True Then
            Call DataEntryForReport(RefNum)
            If Len(wsReport.Range("A8").value) > 0 Then
                Dim PageCount As Integer
                PageCount = Application.WorksheetFunction.RoundUp(Application.WorksheetFunction.CountA(wsReport.Range("A8:A282")) / 28, 0)
                wsReport.Visible = xlSheetVisible
                
                Call wsUnProtectAll
                For i = 1 To PageCount
                
                    wsReport.Range("F5").value = i & " of " & PageCount
                
                    If Environ("computername") = "NLTPMS0010" Then
                        wsReport.PrintOut From:=i, To:=i, ActivePrinter:="HP LaserJet 400 M401 PCL 6 on nltpms0021"
                    ElseIf Environ("computername") = "NLTPHA-005" Then
                        wsReport.PrintOut From:=i, To:=i, ActivePrinter:="Kyocera CS 4054ci KX"
                    Else
                        wsReport.PrintOut From:=i, To:=i, ActivePrinter:="Kyocera TASKalfa 4054ci KX"
                    End If
                    
                Next
                wsReport.Visible = xlSheetVeryHidden
                
                 wsReport.Range("F5").value = "1 of 1"
                Call wsProtectAll
            Else
                MsgBox "Nothing to printout"
            End If
        'End If
End Sub



'-----------------------------------------------------------Other Report-------------------------------------
Sub GenManualFillReport(Optional ByVal tgrdate As String)

'Input Valadation
    wsRecordRetrived.Range("A2:M100").value = ""
    
    
    Dim iRow, ItemCount As Long
    Dim ItemCode As String
    
    Call GetRecordByTodayReplenType(tgrdate)
    ItemCount = wsRecordRetrived.Range("A1").CurrentRegion.Rows.Count
    
    Dim ResizeRng As Range
    Set ResizeRng = wsRecordRetrived.Range("A1:L" & ItemCount)
    wsRecordRetrived.ListObjects("ScannedItem").Resize ResizeRng
    
    
End Sub

Sub test3()
   Call GenManualFillReport("NLTH-251113002")
End Sub

'------------------------30Day Item fequency Report
Sub SaveShelvingReportToNewWorkbook()
    Dim results As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim filePath As String
   
    ' Get the data
    results = GetShelvingFrequencyLast30Days()
   
    ' Create a new workbook
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
   
    If IsArray(results) Then
        ' Add headers
        ws.Range("A1").value = "QR Code on Binshelf"
        ws.Range("B1").value = "Shelving Count"
        ws.Range("C1").value = "Gen. on: " & Now
        
        ' Populate data starting at A2
        For i = 1 To UBound(results, 1)
            ws.Cells(i + 1, 1).value = results(i, 1) ' [QR code on Binshelf]
            ws.Cells(i + 1, 2).value = results(i, 2) ' ShelvingCount
        Next i
        
        ' Auto-fit columns
        ws.Columns("A:B").AutoFit
        
        ' Prompt user to save the workbook
        filePath = Application.GetSaveAsFilename(InitialFileName:="Shelving_Frequency_Report_" & Format(Date, "yyyymmdd") & ".xlsx", _
                                                FileFilter:="Excel Files (*.xlsx), *.xlsx", _
                                                Title:="Save Shelving Frequency Report")
        If filePath <> "False" Then
            wb.SaveAs filePath
            MsgBox "Report saved successfully at " & filePath, vbInformation
        Else
            MsgBox "Save canceled. Workbook will remain open.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No data found or an error occurred. Workbook will be closed.", vbExclamation
        wb.Close savechanges:=False
    End If
   
    ' Clean up
    Set ws = Nothing
    Set wb = Nothing
End Sub

Sub SaveShelvingLastDetailsReport()
    Dim results As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim filePath As String
   
    ' Get the data without date filter
    results = GetShelvingLastDetails()
   
    ' Create a new workbook
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
   
    If IsArray(results) Then
        ' Add headers
        ws.Range("A1").value = "QR Code on Binshelf"
        ws.Range("B1").value = "Lot Of Last"
        ws.Range("C1").value = "Exp Of Last"
        ws.Range("D1").value = "Date and Time Of Last"
        
        ' Populate data starting at A2
        For i = 1 To UBound(results, 1)
            ws.Cells(i + 1, 1).value = results(i, 1) ' [QR code on Binshelf]
            ws.Cells(i + 1, 2).value = results(i, 2) ' LotOfLast
            ws.Cells(i + 1, 3).value = results(i, 3) ' ExpOfLast
            ws.Cells(i + 1, 4).value = results(i, 4) ' Date and TimeOfLast
        Next i
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
        
        ' Prompt user to save the workbook
        filePath = Application.GetSaveAsFilename(InitialFileName:="Shelving_Last_Details_Report_" & Format(Date, "yyyymmdd") & ".xlsx", _
                                                FileFilter:="Excel Files (*.xlsx), *.xlsx", _
                                                Title:="Save Shelving Last Details Report")
        If filePath <> "False" Then
            wb.SaveAs filePath
            'MsgBox "Report saved successfully at " & filePath, vbInformation
        Else
            'MsgBox "Save canceled. Workbook will remain open.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No data found or an error occurred. Workbook will be closed.", vbExclamation
        wb.Close savechanges:=False
    End If
   
    ' Clean up
    Set ws = Nothing
    Set wb = Nothing
End Sub

