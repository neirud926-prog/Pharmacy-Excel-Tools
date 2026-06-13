Attribute VB_Name = "QRCodeFunction"
Sub DownloadAllQRImage()
    Dim arr As Variant
    Dim lastRow As Long
    lastRow = wsItemLoc.Range("A2").End(xlDown).Row
    arr = wsItemLoc.Range("A2:A" & lastRow)
    
    For i = LBound(arr) To UBound(arr)
        Call DownLoadQR(arr(i, 1))
    Next
End Sub


Sub insertQRImg(ByVal tgrRng As Range, ByVal ItemCode As String)

Dim photoNameAndPath As Variant
Dim photo As Picture

    If CheckQRExist(ItemCode, wsGTIN) = False Then
        photoNameAndPath = "O:\Programme Data\QRCode\" & ItemCode & ".png"
        
        Set photo = wsGTIN.Pictures.Insert(photoNameAndPath)
            With photo
                .Left = tgrRng.Left
                .Top = tgrRng.Top
                .Width = tgrRng.Width
                .Height = tgrRng.Height
                .Placement = 1
                .Name = ItemCode
            
            End With
    End If
    
End Sub

Function CheckQRExist(ByVal ItemCode As String, ByVal tgrWS As Worksheet) As Boolean
    Dim QrImg As Picture
    For Each QrImg In tgrWS.Pictures
        If QrImg.Name = ItemCode Then
            CJeckQRExist = True
            Exit Function
        End If
    Next
End Function
Sub CreateAllQRCodeIf()
    Dim lastRow As Long
    lastRow = wsGTIN.Range("A1").End(xlDown).Row
    For i = 2 To lastRow
        Call insertQRImg(wsGTIN.Cells(i, 8), wsGTIN.Cells(i, 1).value)
    Next
End Sub

Sub deleteall()
    On Error Resume Next
    Dim QrImg As Picture
    For Each QrImg In wsGTIN.Pictures
        QrImg.Delete
            
    Next
    On Error GoTo 0
End Sub

Sub insertQRForPrint(ByVal tgrRng As Range, ByVal ItemCode As String)

Dim photoNameAndPath As Variant
Dim photo As Picture

    If CheckQRExist(ItemCode, wsQRPrint) = False Then
        photoNameAndPath = "O:\Programme Data\QRCode\" & ItemCode & ".png"
        
        Set photo = wsQRPrint.Pictures.Insert(photoNameAndPath)
            With photo
                .Width = tgrRng.Width * 0.75
                .Height = tgrRng.Height * 0.75
                .Left = tgrRng.Left + tgrRng.Width * 0.25 / 2
                .Top = tgrRng.Top
                .Placement = 1
                .Name = ItemCode
                
            
            End With
    End If
    
End Sub

Sub PrintQRLabel()
    Call deleteall2
    Dim lastRow, iRow As Long
    lastRow = wsPrintBook.Range("A1").End(xlDown).Row
    
    
    
    For i = 1 To lastRow
        iRow = Application.WorksheetFunction.RoundDown(i / 13, 0)
        iCol = i - iRow * 13
        If iCol = 0 Then iCol = 1
        wsQRPrint.Cells(iRow + 1, iCol).value = wsPrintBook.Cells(i, 1).value
        Call insertQRForPrint(wsQRPrint.Cells(iRow + 1, iCol), wsQRPrint.Cells(iRow + 1, iCol).value)
    Next
    
End Sub

Sub deleteall2()
    On Error Resume Next
    Dim QrImg As Picture
    For Each QrImg In wsQRPrint.Pictures
        QrImg.Delete
            
    Next
    On Error GoTo 0
End Sub
