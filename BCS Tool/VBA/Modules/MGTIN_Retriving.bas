Attribute VB_Name = "MGTIN_Retriving"
Function ItemDataMatch(ByVal ItemCode As String, Optional ByVal col As Integer = 1)
    On Error GoTo err1
    Findings = Application.WorksheetFunction.VLookup(ItemCode, wsItemLoc.Range("A:C"), col, 0)
    ItemDataMatch = Findings
err1:
    If err.Number = 1004 Then
        MsgBox "Item not Found!"
    End If
End Function




Function Rereive_GTIN(ByVal ItemCode As String) As Variant

    Call wsGTIN_Sort
    
    Dim SchRng, Index1, IndexEnd, nextRng As Range
    Set SchRng = wsGTIN.Range("A:A")
    
    Set Index1 = SchRng.Find(what:=ItemCode)
    
    If Index1 Is Nothing Then GoTo myEnd
    
    Set nextRng = Index1
    Set nextRng = SchRng.FindNext(after:=nextRng)
    
    While Not nextRng.Row = Index1.Row
        
        Set IndexEnd = nextRng
        Set nextRng = SchRng.FindNext(after:=nextRng)
    Wend
    If IndexEnd = "" Then
        Set IndexEnd = Index1
    End If

    Rereive_GTIN = wsGTIN.Range("A" & Index1.Row, "F" & IndexEnd.Row)
    
myEnd:
End Function

Function GetItemList(ByVal Code As String) As Variant
    Dim Firstcell, Nextcell As Range
    Dim coll As New Collection
    Set Firstcell = wsItemLoc.Range("A1", wsItemLoc.Range("A1").End(xlDown)).Find(what:=Code, LookAt:=xlPart)
    coll.Add Firstcell.value
    Set Nextcell = Firstcell.FindNext
    
    While Not Nextcell = Firstcell
        coll.Add Nextcell.value
        Set Nextcell = Firstcell.FindNext
        'Debug.Print Nextcell.value
    Wend
    
End Function

Sub test22()
    GetItemList ("AMLO")
End Sub
