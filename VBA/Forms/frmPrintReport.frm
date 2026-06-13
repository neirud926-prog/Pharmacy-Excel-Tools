VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintReport 
   Caption         =   "PrintReport"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   OleObjectBlob   =   "frmPrintReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnPrint_Click()
    If Me.ListRefNum.List(, 0) = "" Then Exit Sub
    Call PrintReport(Me.ListRefNum.List(, 0))
    'Call ChangeStatusTpPrinted(Me.ListRefNum.List(, 0))
    ThisWorkbook.Save
    Unload Me
End Sub



Private Sub btnReprint_Click()
    
    Dim sInput As String
    sInput = InputBox("Please enter the name of Reference No. : ")
    
    If Left(sInput, 4) = "NLTH" Then
    
        Call PrintReport(sInput)
        'Call ChangeStatusTpPrinted(sInput)
        Call Update_wsRecord
        ThisWorkbook.Save
        Unload Me
    End If
End Sub



Private Sub ListRefNum_Click()

End Sub

Private Sub UserForm_Initialize()
    Call Update_wsRecord
    On Error Resume Next
    Me.ListRefNum.List = GetProcessingRefno
    On Error GoTo 0
End Sub

Private Sub ListViewRefresh()
    On Error Resume Next
    Dim arr, newarr As Variant
    Dim coll As New Collection
    Dim iRow, EndRow As Integer
    iRow = wsRecord.Range("H:H").Find(what:="Processing").Row
    EndRow = wsRecord.Range("G" & iRow).End(xlDown).Row
    
    
    If EndRow = 0 And Not iRow = 0 Then
        EndRow = iRow
        arr = wsRecord.Range("G" & iRow, "G" & EndRow)
        Me.ListRefNum.AddItem arr
        
    Else
        arr = wsRecord.Range("G" & iRow, "H" & EndRow)
        
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = "Processing" Then
                coll.Add arr(i, 1), arr(i, 1)
            End If
        Next
        
        ReDim newarr(coll.Count - 1)
        For i = 1 To coll.Count
            newarr(i - 1) = coll.item(i)
        Next
        Me.ListRefNum.List = newarr
    End If
    
    On Error GoTo 0
End Sub
