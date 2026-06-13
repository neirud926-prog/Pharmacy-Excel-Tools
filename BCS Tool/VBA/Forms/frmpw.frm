VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmpw 
   Caption         =   "Verify"
   ClientHeight    =   1305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3195
   OleObjectBlob   =   "frmpw.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmpw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnSubmit_Click()
    If Me.tb.value = "admin" Then
        frmApp.AddcbCustomAlgo = True
        If frmApp.AddcbCustomAlgo = True Then
            frmApp.MainAddFrCustomAlgo.Visible = True
        End If
    Else
        frmApp.AddcbCustomAlgo = False
        If frmApp.AddcbCustomAlgo = False Then
           frmApp.MainAddFrCustomAlgo.Visible = False
        End If
    End If
    Unload Me
End Sub




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Me.tb.value = "admin" Then
        frmAdd.cbCustomAlgo = True
        If frmAdd.cbCustomAlgo = True Then
            frmAdd.Width = 540
        End If
    Else
        frmAdd.cbCustomAlgo = False
        If frmAdd.cbCustomAlgo = False Then
            frmAdd.Width = 340
        End If
    End If
    Unload Me
End Sub
