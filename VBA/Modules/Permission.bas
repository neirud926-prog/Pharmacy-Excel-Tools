Attribute VB_Name = "Permission"

Function GetUserFullName()
    Set WSHnet = CreateObject("WScript.Network")
    userName = WSHnet.userName
    UserDomain = WSHnet.UserDomain
    Set objUser = GetObject("WinNT://" & UserDomain & "/" & userName & ",user")
    GetUserFullName = objUser.FullName
End Function

Function HasPermission() As Boolean

    
    If InStr(GetUserFullName, "P(Pharm)") > 0 Then
        HasPermission = True
    ElseIf InStr(GetUserFullName, "RP(Pharm)") > 0 Then
        HasPermission = True
        
    ElseIf InStr(GetUserFullName, "SD(Pharm)") > 0 Then
        HasPermission = True
    ElseIf GetUserFullName = "Daniel WONG, NLTH D(Pharm)" Then
        HasPermission = True
    Else
        HasPermission = False
    End If
    
End Function

Sub wsProtectAll()
    'wsGTIN.Protect Password:="NLTH"
    'wsNameList.Protect Password:="NLTH"
    wsReport.Protect Password:="NLTH"
    'wsRecord.Protect Password:="NLTH"
    wsCurrentLot.Protect Password:="NLTH"
    
End Sub

Sub wsUnProtectAll()
    wsGTIN.Unprotect Password:="NLTH"
    wsNameList.Unprotect Password:="NLTH"
    wsReport.Unprotect Password:="NLTH"
    'wsRecord.Unprotect Password:="NLTH"
    wsCurrentLot.Unprotect Password:="NLTH"
End Sub
