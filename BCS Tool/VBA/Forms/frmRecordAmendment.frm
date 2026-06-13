VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecordAmendment 
   Caption         =   "Record Amendment"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5295
   OleObjectBlob   =   "frmRecordAmendment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRecordAmendment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnSearch_Click()

    Me.cbOldType.value = GetRecordTypeByRefNo(Me.cbRefNum.value)
    
    If Me.cbOldType.value = "" Then Exit Sub
    Dim thisTypeList As New Collection
    Set thisTypeList = Get_TypeList()
    
    thisTypeList.Remove (Me.cbOldType.value)
    
    For i = 1 To thisTypeList.Count
        Me.cbNewType.AddItem thisTypeList.item(i)
    Next
End Sub








Private Sub btnUpdate_Click()
    If Len(Me.cbNewType.value) > 0 Then
        Dim rowsAffected As Long
        rowsAffected = PrepareUpdateSQL_ForUpdateType("Record", Me.cbRefNum.value, Me.cbNewType.value)
        MsgBox rowsAffected & " rows updated."
        
        Unload Me
    End If
End Sub

Private Sub UserForm_Click()

End Sub
