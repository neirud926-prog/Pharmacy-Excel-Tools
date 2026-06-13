Attribute VB_Name = "General"
Function iLookup(ByVal p1 As String, ByVal pRng As Range, ByVal pCol As Long)
On Error Resume Next
    iLookup = Application.WorksheetFunction.VLookup(p1, pRng, pCol, 0)
On Error GoTo 0
End Function



'====================================================================
' LookupRow – exact match in an unsorted column ? returns whole row
'====================================================================
'  rngTable   : contiguous range that contains the data (incl. headers)
'  lookupCol  : column **index inside rngTable** that holds the key
'  lookupVal  : value to find (exact match)
'  Return     : 1-D array (row values) or Empty if not found
'====================================================================
Public Function LookupRow(rngTable As Range, _
                          lookupCol As Long, _
                          lookupVal As Variant) As Variant

    Dim dict As Object                  ' Scripting.Dictionary
    Dim vData As Variant                ' 2-D array of the table
    Dim r As Long, lastRow As Long
    Dim key As Variant
    Dim result() As Variant
    
    '--- 1. Validate ------------------------------------------------
    If rngTable Is Nothing Then Exit Function
    If lookupCol < 1 Or lookupCol > rngTable.Columns.Count Then Exit Function
    
    '--- 2. Pull the whole block into a Variant array (fastest read)
    vData = rngTable.Value               ' 1-based, 2-D
    
    lastRow = UBound(vData, 1)
    If lastRow = 0 Then Exit Function   ' empty table
    
    '--- 3. Build a dictionary (key = column value, item = row number)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare     ' case-insensitive (change to vbBinaryCompare if needed)
    
    For r = 1 To lastRow                 ' row 1 = header, skip it
        key = vData(r, lookupCol)
        If Not IsError(key) Then
            If Not dict.Exists(key) Then dict(key) = r
        End If
    Next r
    
    '--- 4. Lookup --------------------------------------------------
    If dict.Exists(lookupVal) Then
        r = dict(lookupVal)              ' row index inside the array
        ReDim result(1 To rngTable.Columns.Count)
        Dim c As Long
        For c = 1 To rngTable.Columns.Count
            result(c) = vData(r, c)
        Next c
        LookupRow = result
    Else
        LookupRow = Empty                ' not found
    End If
End Function

Public Function LookupRows(rngTable As Range, _
                           lookupCol As Long, _
                           lookupVal As Variant, _
                           numRows As Long) As Variant
    Dim dict As Object                  ' Scripting.Dictionary
    Dim vData As Variant                ' 2-D array of the table
    Dim r As Long, lastRow As Long
    Dim key As Variant
    Dim result() As Variant
    Dim startRow As Long, endRow As Long
    
    '--- 1. Validate ------------------------------------------------
    If rngTable Is Nothing Then Exit Function
    If lookupCol < 1 Or lookupCol > rngTable.Columns.Count Then Exit Function
    If numRows < 1 Then Exit Function
    
    '--- 2. Pull the whole block into a Variant array (fastest read)
    vData = rngTable.Value               ' 1-based, 2-D
    
    lastRow = UBound(vData, 1)
    If lastRow < 2 Then Exit Function   ' empty table (assume header)
    
    '--- 3. Build a dictionary (key = column value, item = row number)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare     ' case-insensitive
    
    For r = 2 To lastRow                 ' start from data (skip header)
        key = vData(r, lookupCol)
        If Not IsError(key) Then
            If Not dict.Exists(key) Then dict(key) = r
        End If
    Next r
    
    '--- 4. Lookup & slice N rows -----------------------------------
    If dict.Exists(lookupVal) Then
        startRow = dict(lookupVal)       ' found data row
        endRow = Application.Min(startRow + numRows - 1, lastRow)
        
        ' Alloc 2-D result (endRow - startRow + 1 rows x cols)
        ReDim result(1 To (endRow - startRow + 1), 1 To UBound(vData, 2))
        
        Dim rr As Long, c As Long
        rr = 1
        For r = startRow To endRow
            For c = 1 To UBound(vData, 2)
                result(rr, c) = vData(r, c)
            Next c
            rr = rr + 1
        Next r
        
        LookupRows = result
    End If
End Function

Sub LaunchApp()
    frmApp.Show
End Sub

Sub ApplyReportFormatting()
    Dim lastRow As Long
    
    With wsReport
        .Visible = xlSheetVisible
        .Activate
        
        ' 1. ??????
        With .Range("A1").CurrentRegion
            .Borders.LineStyle = xlContinuous ' ????
            .Rows(1).Font.Bold = True        ' ???????
        End With
        
        ' 2. ??????
        .Range("A:A").NumberFormat = "@"               ' ItemCode ????
        .Range("F:F").NumberFormat = "yyyy-mm-dd"      ' ????
        
        ' 3. ?? E ??? 0 ?? (????????,????)
        lastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
        If lastRow > 1 Then ' ????????
            With .Range("E2:E" & lastRow)
                .FormatConditions.Delete ' ???????
                ' ????:?? 0 ??????????
                With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    .Interior.Color = RGB(255, 199, 206) ' ?????
                    .Font.Color = RGB(156, 0, 6)        ' ?????
                End With
            End With
        End If
    End With
End Sub
