Attribute VB_Name = "Alert"
Sub AlertSpeak(ByVal words As String)
    Application.Speech.Speak words & "is not matched, Please check again!"
End Sub
Function itemCodeStop(ByVal ItemCode As String) As String
   
    itemCodeStop = Mid(ItemCode, 1, 1) & "," & Mid(ItemCode, 2, 1) & "," & Mid(ItemCode, 3, 1) & "," & Mid(ItemCode, 4, 1) & "," & _
                    Mid(ItemCode, 5, 1) & "," & Mid(ItemCode, 6, 1) & ","
        

End Function
Sub NotMatchAlert(ByVal ItemCode As String)
    Call AlertSpeak(itemCodeStop(ItemCode))
End Sub
