Attribute VB_Name = "MHashing"
Public Function BASE64SHA1(ByVal sTextToHash As String)

    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    Const cutoff As Integer = 16

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.GetBytes_4(sTextToHash)
    SharedSecretKey = asc.GetBytes_4(sTextToHash)
    enc.Key = SharedSecretKey

    bytes = enc.ComputeHash_2((TextToHash))
    BASE64SHA1 = EncodeBase64(bytes)
    BASE64SHA1 = Left(BASE64SHA1, cutoff)

    Set asc = Nothing
    Set enc = Nothing

End Function

Public Function ReplacePlus(ByVal Code As String)
    For i = 1 To Len(Code)
        If Mid(Code, i, 1) = "+" Then
            Code = Left(Code, i - 1) & "!" & Right(Code, Len(Code) - i)
        End If
    Next
    ReplacePlus = Code

End Function

Public Function ApplyKey(ByVal ItemCode As String)
    ApplyKey = ReplacePlus("B" & BASE64SHA1(ItemCode) & "aR")
End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing

End Function

Sub DownLoadQR(ByVal ItemCode As String, Optional ByVal Codetype As Integer = 0)
    '0 = item, '1= text
     Dim FileUrl, text As String
     Dim objXmlHttpReq As Object
     Dim objStream As Object
     text = ApplyKey(ItemCode)
    If Codetype = 0 Then
        FileUrl = "https://chart.googleapis.com/chart?chs=300x300&&cht=qr&chl=" & ApplyKey(ItemCode)
     ElseIf Codetype = 1 Then
        FileUrl = "https://chart.googleapis.com/chart?chs=300x300&&cht=qr&chl=" & ItemCode
     End If

     'example
     'FileUrl = "https://www.example.com/images/chart.jpg"

     Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
     objXmlHttpReq.Open "GET", FileUrl, False, "username", "password"
     objXmlHttpReq.send

     If objXmlHttpReq.status = 200 Then
          Set objStream = CreateObject("ADODB.Stream")
          objStream.Open
          objStream.Type = 1
          objStream.Write objXmlHttpReq.responseBody
          objStream.SaveToFile "O:\Programme Data\QRCode" & "\" & ItemCode & ".png", 2
          objStream.Close
     End If

End Sub


