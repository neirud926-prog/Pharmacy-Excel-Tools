Attribute VB_Name = "BugFix"
Sub FixGTINto14()
    
    totalRow = wsGTIN.Range("A2", wsGTIN.Range("A2").End(xlDown)).Count
    
    For i = 1 To totalRow
        While Not Len(wsGTIN.Cells(i + 1, 2).value) = 14
            wsGTIN.Cells(i + 1, 2).value = "0" & wsGTIN.Cells(i + 1, 2).value
        Wend
    Next
End Sub


Sub test()
    
    
    ResultArr = ReadBarCode("0108002660039583218422995650201066202PC0317270531")
End Sub

'01 08002660039583 21 842064118026 10 66202PC0317270531
