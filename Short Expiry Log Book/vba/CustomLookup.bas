Attribute VB_Name = "CustomLookup"
Function ItemLocLookUp(ByVal Code As String) As Variant
    Dim tgrRng As Range
    Dim arr As Variant
    Dim Coll As New Collection
    Dim LocType As String
    
    Set tgrRng = wsLoc.Range("A:A").Find(What:=Code)
    
    For i = 1 To 7
        iLoc = tgrRng.Offset(0, i + 1).Value
        If iLoc <> "" Then
            LocType = wsLoc.Cells(1, i + 2).Value
            Coll.Add LocType & ": " & iLoc
        End If
        
    Next
    
    ReDim arr(Coll.Count - 1) As Variant
    For i = 0 To Coll.Count - 1
        arr(i) = Coll.Item(i + 1)
    Next
    
    ItemLocLookUp = arr
End Function

