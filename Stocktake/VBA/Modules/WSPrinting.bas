Attribute VB_Name = "WSPrinting"
Sub Print5itemsWorkSheet()
    Dim result As Variant
    Dim i As Long
    
    ' 1. Next 5 items NOT yet counted in the current Daily cycle, in bin order
    '    (BinCycle module). Items already counted in this cycle - including
    '    off-sequence entries - are never printed again until the cycle is done.
    result = BinCycle.GetNextCycleItems(5)
    
    If IsEmpty(result) Then
        MsgBox "No next Daily BinShelf found." & vbCrLf & BinCycle.GetCycleSummary(), vbExclamation
        Exit Sub
    End If
    
    ' 2. Write the ItemCodes to C3:C7 (row 3 to 7); clear unused rows
    '    ? This will trigger Worksheet_Change ? lookups happen automatically
    For i = 1 To 5
        If i <= UBound(result, 1) Then
            wsStocktakeSheet.Cells(2 + i, 3).Value = result(i, 1)   ' ItemCode
        Else
            wsStocktakeSheet.Cells(2 + i, 3).ClearContents
        End If
    Next i
    
    ' 3. Set date in G1
    wsStocktakeSheet.Range("G1").Value = Format(Date, "dd-mmm-yyyy")
    

    
    
    ' Optional: activate the sheet for user
    wsStocktakeSheet.Activate
    
    wsStocktakeSheet.PrintOut Collate:=True, IgnorePrintAreas:=False, ActivePrinter:="Kyocera TASKalfa 4054ci KX"
End Sub

Sub test333()
    MsgBox "Next bin: " & GetNextBinShelf() & vbCrLf & BinCycle.GetCycleSummary()
End Sub
