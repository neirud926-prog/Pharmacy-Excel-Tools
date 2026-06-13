Attribute VB_Name = "ArrayFunction"
Public Function RemoveDuplicate(ByVal arr As Variant) As Variant
    Dim NewCol As Collection
    Dim newarr As Variant
    Set NewCol = New Collection
    On Error Resume Next
        For i = LBound(arr) To UBound(arr)
            NewCol.Add arr(i), arr(i)
        Next
    On Error GoTo 0
    
    ReDim newarr(NewCol.Count) As Variant
    For i = 1 To NewCol.Count
        newarr(i - 1) = NewCol.item(i)
    Next
    RemoveDuplicate = newarr
End Function

Public Function TwoDToOneDArr(ByVal arr As Variant) As Variant
    Dim newarr As Variant
    ReDim newarr(UBound(arr)) As Variant
    
    For i = LBound(arr) To UBound(arr)
        newarr(i) = arr(i, 1)
    Next
    
    TwoDToOneDArr = newarr
End Function

Public Function AddZeroForGTIN14(ByVal vgtin As String)
    If Len(vgtin) > 13 Then Exit Function
    For i = 1 To 14 - Len(vgtin)
        vgtin = "0" & vgtin
    Next
    AddZeroForGTIN14 = vgtin
End Function
