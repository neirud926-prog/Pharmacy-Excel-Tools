Attribute VB_Name = "WSPrinting"
Sub Print5itemsWorkSheet()
    Dim result As Variant
    Dim nextBin As String
    Dim i As Long
    
    ' 1. Get the next bin shelf (already filters Type = 'Daily')
    nextBin = GetNextBinShelf()
    
    If nextBin = "" Then
        MsgBox "No next Daily BinShelf found.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Get the current + next 4 rows (total 5)
    result = LookupRows(wsBinShelfLocator.Range("A1").CurrentRegion, _
        2, _
        nextBin, _
        5 _
    )
    
    If IsEmpty(result) Then
        MsgBox "Not enough consecutive BinShelves found starting from " & nextBin, vbExclamation
        Exit Sub
    End If
    
    ' 3. Write the 5 ItemCodes to C3:C7 (row 3 to 7)
    '    ? This will trigger Worksheet_Change ? lookups happen automatically
    For i = 1 To UBound(result, 1)           ' number of rows returned
        wsStocktakeSheet.Cells(2 + i, 3).Value = result(i, 1)   ' ItemCode = column 1 in result
    Next i
    
    ' 4. Set date in G1
    wsStocktakeSheet.Range("G1").Value = Format(Date, "dd-mmm-yyyy")
    

    
    
    ' Optional: activate the sheet for user
    wsStocktakeSheet.Activate
    
    wsStocktakeSheet.PrintOut Collate:=True, IgnorePrintAreas:=False, ActivePrinter:="Kyocera TASKalfa 4054ci KX"
End Sub

Sub test333()
    MsgBox GetNextBinShelf()
End Sub
