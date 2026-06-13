Attribute VB_Name = "MReadingBarCode"
Public Function KillSpecial(ByVal Barcode As String) As String
    i = InStr(1, Barcode, "¡j")
    If i > 0 Then
        KillSpecial = Left(Barcode, i - 1) & Right(Barcode, Len(Barcode) - i)
    Else
        KillSpecial = Barcode
    End If
    
End Function
'
Public Function ReadBarCode(ByVal Barcode As String) As Variant
    Dim result As Variant
    result = TryCustomDecode(Barcode)
    If UBound(result) = 2 Then
        ReadBarCode = result
        'MsgBox "i using fastest method to read and done"
        Exit Function
    End If
    
    ReadBarCode = DecodeDistributor(Barcode)
    
End Function




Function AlgoTestP0(ByVal Barcode As String) As Variant
    Dim iTag90 As Integer
    Dim result(2) As Variant
    If Len(Barcode) <= 16 Then
        
        If Len(Barcode) < 15 Then
            While Not Len(Barcode) = 14
                Barcode = "0" & Barcode
            Wend
            result(0) = Barcode
        Else
            result(0) = Mid(Barcode, 3, 14)
        
        End If
        
        
        result(1) = "na----"
        result(2) = "na----"
    ElseIf Len(Barcode) > 16 And Len(Barcode) > 0 Then
        '0 = GTIN, 1 = Batch , 2 = EXP date
        
        result(0) = Mid(Barcode, 3, 14)
        
        iTag90 = InStr(1, Barcode, "90HK")
        'iTag90 = Len(BarCode) - 10
        
        If Mid(Barcode, 1, 2) = "01" And Mid(Barcode, 17, 2) = "10" Then
            'MsgBox "This is GSI128 with Reg. No."
            
        If iTag90 > 0 Then
            
            
            
            If Mid(Barcode, 17, 2) = "10" Then
                result(1) = Mid(Barcode, 19, iTag90 - 27)
            End If
            
            If Mid(Barcode, iTag90 - 8, 2) = "17" Then
                result(2) = Mid(Barcode, iTag90 - 6, 6)
            End If
        
        End If
        
        ElseIf Mid(Barcode, 1, 2) = "01" And Mid(Barcode, 17, 2) = "17" Then
            'MsgBox "This is DM"
            
            Dim batchLen As Integer
            batchLen = 26
            If iTag90 > 0 Then batchLen = batchLen + 10
            
            If Mid(Barcode, 25, 2) = "10" Then
                temp = Mid(Barcode, 27, Len(Barcode) - batchLen)
                If Len(temp) > 7 And InStr(1, temp, "21") Then
                    result(1) = Mid(temp, 1, InStr(1, temp, "21") - 1)
                Else
                    result(1) = temp
                End If
                
            End If
            
            If Mid(Barcode, 17, 2) = "17" Then
                result(2) = Mid(Barcode, 19, 6)
            End If
        
            
        End If
    End If
    
    AlgoTestP0 = result
    
End Function

Sub test123()
    Dim result As Variant
    For i = 2 To 37
        result = DecodeDistributor(wstest.Cells(i, 1).value)
        wstest.Cells(i, 2).value = result(0)
        wstest.Cells(i, 3).value = result(1)
        wstest.Cells(i, 4).value = result(2)

    Next
    
    
End Sub
Sub mytestdebug()
Dim arr As Variant
    
    arr = DecodeDistributor("0114987051142039")
    MsgBox arr(0)
End Sub


Function DecodeDistributor(ByVal Barcode As String)
    If Right(Barcode, 2) = "aR" Then Exit Function
    If Barcode = "" Then Exit Function
    
    
    'On Error Resume Next
    Barcode = RemoveBracket(Barcode)
    If Len(Barcode) <= 16 And Len(Barcode) > 0 Then
    
        thisresult = AlgoTestP0(Barcode)
    Else
        If Len(Barcode) > 16 Then
            
            If IsResultValid(AlgoTestP0(Barcode)) = True Then
                thisresult = AlgoTestP0(Barcode)
                DecodeDistributor = thisresult
                Exit Function
            End If
            If Mid(Barcode, Len(Barcode) - 7, 2) = "17" Then
                thisresult = AlgoTestP3(Barcode)
            ElseIf Mid(Barcode, Len(Barcode) - 15, 2) = "17" And Mid(Barcode, Len(Barcode) - 7, 2) = "11" Then
                thisresult = AlgoTestP4(Barcode)
            ElseIf Mid(Barcode, Len(Barcode) - 7, 2) = "15" Then
                thisresult = AlgoTestP6(Barcode)
            Else
                If IsResultValid(AlgoTestP1(Barcode)) = True Then
                    thisresult = AlgoTestP1(Barcode)
                ElseIf IsResultValid(AlgoTestP2(Barcode)) = True Then
                    thisresult = AlgoTestP2(Barcode)
                ElseIf IsResultValid(AlgoTestP5(Barcode)) = True Then
                    thisresult = AlgoTestP5(Barcode)
    
                End If
                
            End If
        End If
        
        
    End If
    
    If IsResultValid(thisresult) = True Then
            
            DecodeDistributor = thisresult
            Exit Function
    End If
    
'If IsEmpty(thisresult) = False Then
'        If UBound(thisresult) = 2 Then
'            If IsResultValid(thisresult) = True Then DecodeDistributor = thisresult
'        Else
'            DecodeDistributor = TryCustomDecode(Barcode)
'        End If
'    Else
'    'On Error GoTo 0
'
'        DecodeDistributor = TryCustomDecode(Barcode)
'End If
End Function

Function TryCustomDecode(ByVal Barcode As String) As Variant
    'test GTIN length
    If Len(Barcode) >= 16 Then
        Dim GTIN As String
        GTIN = Mid(Barcode, 3, 14)
        Dim result As Variant
        result = ShowGTINDetailsByGTIN(GTIN)
        If IsEmpty(result) Then
            Exit Function
        End If
        
        If UBound(result) > 0 Then
        
            'For j = LBound(result, 2) To UBound(result, 2)
                'Debug.Print result(1, j),
            'Next j
           If result(1, 5) = "Custom" Then
                Dim arr(2), Index As Variant
                Index = Split(result(1, 6), ",")
                arr(0) = Mid(Barcode, Index(0), Index(1) - Index(0) + 1)
                arr(1) = Mid(Barcode, Index(2), Index(3) - Index(2) + 1)
                arr(2) = Mid(Barcode, Index(4), Index(5) - Index(4) + 1)
                TryCustomDecode = arr
            Else
                TryCustomDecode = Array(Barcode)
           End If
        
        Else
            TryCustomDecode = Array(Mid(Barcode, 3, 14))
        End If
    Else
        TryCustomDecode = Array(Barcode)
    End If
    
End Function


Function FindAlgo(ByVal Barcode As String)
    If Right(Barcode, 2) = "aR" Then Exit Function
    If Barcode = "" Then Exit Function
    
    
    'On Error Resume Next
    
    If Len(Barcode) <= 16 And Len(Barcode) > 0 Then
    
        thisresult = AlgoTestP0(Barcode)
        FindAlgo = 0
        Exit Function
    Else
        If Len(Barcode) > 16 Then
            
            If IsResultValid(AlgoTestP0(Barcode)) = True Then
                thisresult = AlgoTestP0(Barcode)
                FindAlgo = "0"
                Exit Function
            End If
            If Mid(Barcode, Len(Barcode) - 7, 2) = "17" Then
                thisresult = AlgoTestP3(Barcode)
                If IsResultValid(thisresult) = True Then FindAlgo = "3"
                Exit Function
            ElseIf Mid(Barcode, Len(Barcode) - 15, 2) = "17" And Mid(Barcode, Len(Barcode) - 7, 2) = "11" Then
                thisresult = AlgoTestP4(Barcode)
                If IsResultValid(thisresult) = True Then FindAlgo = "4"
                Exit Function
            ElseIf Mid(Barcode, Len(Barcode) - 7, 2) = "15" Then
                thisresult = AlgoTestP6(Barcode)
                If IsResultValid(thisresult) = True Then FindAlgo = "6"
                Exit Function
            Else
                If IsResultValid(AlgoTestP1(Barcode)) = True Then
                    thisresult = AlgoTestP1(Barcode)
                    FindAlgo = "1"
                    Exit Function
                ElseIf IsResultValid(AlgoTestP2(Barcode)) = True Then
                    thisresult = AlgoTestP2(Barcode)
                    FindAlgo = "2"
                    Exit Function
                ElseIf IsResultValid(AlgoTestP5(Barcode)) = True Then
                    thisresult = AlgoTestP5(Barcode)
                    FindAlgo = "5"
                    Exit Function
    
                End If
                
            End If
        End If
    End If
    
    

    'On Error GoTo err
'err:
    'If err.Number > 0 Then MsgBox err.Number & " Undefined Algo"
        

End Function



Function Pattern_Found(ByVal Barcode As String, ByVal tagA As String, ByVal tagB As String, ByVal diff As String) As Variant
    
    Dim arr_tag(1) As Variant
    
    Dim coll_tagA As New Collection
    Dim coll_tagB As New Collection
    
    Firsttag_A = 1
    While Not Firsttag_A = 0
        Firsttag_A = InStr(Firsttag_A + 1, Barcode, tagA)
        If Firsttag_A > 0 Then coll_tagA.Add Firsttag_A
    Wend
    Firsttag_B = 1
    While Not Firsttag_B = 0
        Firsttag_B = InStr(Firsttag_B + 1, Barcode, tagB)
        If Firsttag_B > coll_tagA(1) Then coll_tagB.Add Firsttag_B
    Wend
    
    For i = 1 To coll_tagA.Count
        For r = 1 To coll_tagB.Count
            If coll_tagB(r) - coll_tagA(i) = diff Then
            
                arr_tag(0) = coll_tagA(i)
                arr_tag(1) = coll_tagB(r)
            End If
        Next
    Next
    
    Pattern_Found = arr_tag
    
End Function

Function AlgoTestP1(ByVal Barcode As String) As Variant

    Dim result(2) As Variant
    '"17" Count
    'find first 17
    Dim Firsttag_17, Firsttag_10, last_tag17 As Integer
    
    Dim tag_17, tag_10 As Integer
    
    Dim arr_tag As Variant
    On Error Resume Next
    arr_tag = Pattern_Found(Barcode, "17", "10", 8)
    tag_10 = arr_tag(1)
    tag_17 = arr_tag(0)
    
    On Error GoTo 0
    
        If tag_10 - tag_17 = 8 Then
            result(0) = Mid(Barcode, 3, 14)
            result(2) = Mid(Barcode, tag_17 + 2, 6)
            'check for SN
            If InStr(Mid(Barcode, tag_10 + 2, 14), "21") > 0 Then
            'has serial num
                Dim tag21 As Integer
                tag21 = InStr(tag_10 + 2, Barcode, "21")
                result(1) = Mid(Barcode, tag_10 + 2, tag21 - tag_10 - 2)
        
            Else
                result(1) = Mid(Barcode, tag_10 + 2, Len(Barcode) - tag_10 + 1)
                        'no SN and batch at right
            End If
            AlgoTestP1 = result
            Exit Function
        
        End If


    result(0) = False
    AlgoTestP1 = result
End Function
Function AlgoTestP2(ByVal Barcode As String) As Variant
     Dim result(2) As Variant
     
    Dim tag_17, tag_11 As Integer
    
    Dim arr_tag As Variant
    On Error Resume Next
    arr_tag = Pattern_Found(Barcode, "17", "11", 8)
    tag_11 = arr_tag(1)
    tag_17 = arr_tag(0)
    On Error GoTo 0
        
        If tag_11 - tag_17 = 8 Then
            result(0) = Mid(Barcode, 3, 14)
            result(2) = Mid(Barcode, Firsttag_17 + 2, 6)
            
            Dim tag10 As Integer
            tag10 = tag_11
            
            For r = 1 To 3
                tag10 = InStr(tag10 + 1, Barcode, "10")
                If tag10 - tag_11 = 8 Then
                    result(1) = Mid(Barcode, tag10 + 2, Len(Barcode) - tag10 + 1)
                    Exit For
                End If
            Next
            
            AlgoTestP2 = result
            Exit Function
            
        End If

    result(0) = False
    AlgoTestP2 = result
End Function

Function AlgoTestP3(ByVal Barcode As String) As Variant
    '17 at last
    
    '11-> 17
    If Mid(Barcode, Len(Barcode) - 7, 2) = "17" Then
        
        Dim tag_17, tag_11 As Integer
        
        Dim result(2) As Variant
        '"11" Count
        'find first 11
        Dim arr_tag As Variant
           
        On Error Resume Next
        arr_tag = Pattern_Found(Barcode, "11", "17", 8)
        tag_17 = arr_tag(1)
        tag_11 = arr_tag(0)
        On Error GoTo 0
            
            If tag_17 - tag_11 = 8 Then
                result(0) = Mid(Barcode, 3, 14)
                result(2) = Mid(Barcode, Len(Barcode) - 5, 6)
                'check for SN
                
                result(1) = Mid(Barcode, 19, Len(Barcode) - 34)
               
                            
                AlgoTestP3 = result
                Exit Function
                
            End If

    
        If Mid(Barcode, 17, 2) = "10" And Len(Barcode) - 26 < 11 Then
            result(0) = Mid(Barcode, 3, 14)
            result(2) = Mid(Barcode, Len(Barcode) - 5, 6)
            result(1) = Mid(Barcode, 19, Len(Barcode) - 26)

            AlgoTestP3 = result
            Exit Function
        End If
        
    End If
    'result(0) = False
    AlgoTestP3 = result

End Function

Function AlgoTestP4(ByVal Barcode As String) As Variant
    Dim result(2) As Variant
    '17 at last

    If Mid(Barcode, Len(Barcode) - 15, 2) = "17" And Mid(Barcode, Len(Barcode) - 7, 2) = "11" Then
        result(0) = Mid(Barcode, 3, 14)
        result(1) = Mid(Barcode, 19, Len(Barcode) - 34)
        result(2) = Mid(Barcode, Len(Barcode) - 13, 6)
       
        AlgoTestP4 = result
        
        Exit Function
    End If
    result(0) = False
    AlgoTestP4 = result

End Function

Function AlgoTestP5(ByVal Barcode As String) As Variant
     Dim result(2) As Variant
     
     '17->21 (8)
     Dim tag_17, tag_21 As Integer
     
    Dim arr_tag As Variant
    On Error Resume Next
    arr_tag = Pattern_Found(Barcode, "17", "21", 8)
    tag_21 = arr_tag(1)
    tag_17 = arr_tag(0)
    On Error GoTo 0
    
        
        If tag_21 - tag_17 = 8 Then
            result(0) = Mid(Barcode, 3, 14)
            
            result(1) = Mid(Barcode, 19, tag_17 - 19)
            result(2) = Mid(Barcode, tag_17 + 2, 6)

            AlgoTestP5 = result
            Exit Function
            
        End If
    
    result(0) = False
    AlgoTestP5 = result
End Function

Function AlgoTestP6(ByVal Barcode As String) As Variant
     Dim result(2) As Variant
     
     '10->15 (8)
     Dim tag_10, tag_15 As Integer
     
    Dim arr_tag As Variant
    On Error Resume Next
    arr_tag = Pattern_Found(Barcode, "10", "15", 8)
    tag_15 = arr_tag(1)
    tag_10 = arr_tag(0)
    On Error GoTo 0
    
        
        If tag_15 - tag_10 = 8 Then
            result(0) = Mid(Barcode, 3, 14)
            
            result(1) = Mid(Barcode, 19, 6)
            result(2) = Mid(Barcode, tag_15 + 2, 6)

            AlgoTestP6 = result
            Exit Function
            
        End If
    
    result(0) = False
    AlgoTestP6 = result
End Function

Sub test()
    
    
    Dim Barcode As String
    Dim GTINStr As String
    
    Barcode = "4895175601208"
    '01 00771313216104 21 673173 17 251031 10 TP0763
    BarCodeLength = Len(Barcode)
    
    If BarCodeLength < 16 And BarCodeLength > 0 Then
        'GTIN only
        'AddZeroCount = 14 - BarCodeLength
        
        'For i = 1 To AddZeroCount
        '    GTINStr = GTINStr & "0"
        'Next
        
        'Me.lbGTIN.Caption = GTINStr & Me.tbBarCode.Value
        
    ElseIf BarCodeLength >= 16 Then
    
    Dim arr As Variant
    arr = ReadBarCode(Barcode)
    
    
    
    MsgBox arr(0)
    MsgBox arr(1)
    MsgBox arr(2)
    
    GTIN = arr(0)
    
    
    iRow = wsGTIN.Columns(2).Find(what:=GTIN, LookAt:=xlPart).Row

    
    End If
End Sub

Function IsResultValid(ByVal result As Variant) As Boolean
    Dim arr As Variant
    If result(0) = False Then
        IsResultValid = False
        Exit Function
    End If
    arr = result
    If Len(arr(0)) = 14 And Len(arr(2)) = 6 And Len(arr(1)) > 0 Then IsResultValid = True

End Function

Public Function RemoveBracket(ByVal Code As String) As String
    Dim result As String
    result = Replace(Code, "(", "")
    result = Replace(result, ")", "")
    RemoveBracket = result
End Function

Public Function FindCodeIndex(ByVal vBarCode As String, ByVal vLot As String, ByVal vExp As String) As Variant

    Dim vIndex(3) As Variant
    If Len(vBarCode) > 0 And Len(vExp) > 0 And Len(vLot) > 0 Then
        vIndex(0) = InStr(vBarCode, vLot)
        vIndex(1) = vIndex(0) + Len(vLot) - 1
        
        vIndex(2) = InStr(vBarCode, vExp)
        vIndex(3) = vIndex(2) + Len(vExp) - 1
    Else
        FindCodeIndex = False
        Exit Function
    End If
    
    FindCodeIndex = vIndex
End Function
Sub demybug()
    MsgBox ReadBarCode("0105021730030926215494160972830110240500234117261231")(0)
    'MsgBox ReadBarCode("05021730030926")(2)

    
End Sub
