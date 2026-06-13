Attribute VB_Name = "MRecord"


Function GetSolPackSize(ByVal GTIN As String) As String
    Dim arr As Variant

    arr = ShowGTINDetailsByGTIN(GTIN)
    
    If arr(1, 4) <> "" Then
        'Debug.Print (arr(1, 4))
            GetSolPackSize = arr(1, 4) & "ML"
        Exit Function
    End If
    GetSolPackSize = ""
End Function

Function Get_TypeList() As Collection
    Dim TypeList As New Collection
    TypeList.Add "Out-Patient Daily Replenishment", "Out-Patient Daily Replenishment"
    TypeList.Add "Replenishment from Main Store (TUE & THU)", "Replenishment from Main Store (TUE & THU)"
    TypeList.Add "Other:", "Other:"
    Set Get_TypeList = TypeList
End Function





Sub test12345()
    MsgBox GetSolPackSize("04895103607258")
End Sub
