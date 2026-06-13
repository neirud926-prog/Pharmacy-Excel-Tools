VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReplenManualFillReport 
   Caption         =   "Manual fill Report"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "frmReplenManualFillReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReplenManualFillReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnGenReprot_Click()
    Dim fileSelected As Boolean
    fileSelected = importTransaction
    Call GenManualFillReport(Me.tbRefNum.value)
    wsScannedItemReport.ListObjects("ManualFIllReport2").Refresh
    wsScannedItemReport.Visible = xlSheetVisible
    Unload Me
    Unload frmApp
End Sub

Private Sub cbListRefresh()
    On Error Resume Next
    Dim arr, newarr As Variant
    Dim coll As New Collection
    Dim iRow, EndRow As Integer
    iRow = wsRecord.Range("F:F").Find(what:="Replenishment from Main Store (TUE & THU)").Row
    EndRow = wsRecord.Range("G" & iRow).End(xlDown).Row
    
    
    If EndRow = 0 And Not iRow = 0 Then
        EndRow = iRow
        arr = wsRecord.Range("G" & iRow, "G" & EndRow)
        Me.tbRefNum.AddItem arr
        
    Else
        arr = wsRecord.Range("F" & iRow, "G" & EndRow)
        
        For i = LBound(arr) To UBound(arr)
            If arr(i, 1) = "Replenishment from Main Store (TUE & THU)" Then
                coll.Add arr(i, 2), arr(i, 2)
            End If
        Next
        
        ReDim newarr(coll.Count - 1)
        For i = 1 To coll.Count
            newarr(i - 1) = coll.item(coll.Count - i + 1)
        Next
        Me.tbRefNum.List = newarr
    End If
    
    On Error GoTo 0
End Sub

Private Function LoadRefNumFromDB() As Variant
    Dim sql As String
    Dim result As Variant
    
    ' Query Ref No where Refill Type is 'Replenishment from Main Store (TUE & THU)'
    sql = "SELECT DISTINCT [Ref No] FROM [Record] WHERE [Refill Type] = 'Replenishment from Main Store (TUE & THU)' ORDER BY [Ref No] DESC;"
    result = DatabaseSQL(sql)
    
    If IsArray(result) Then
        If UBound(result, 1) > 0 Then
            LoadRefNumFromDB = Application.Transpose(Application.Index(result, Evaluate("ROW(1:" & (UBound(result, 1) - LBound(result, 1) + 1) & ")"), 1))
        Else
            LoadRefNumFromDB = Array() ' Return empty array if no data
        End If
    Else
        LoadRefNumFromDB = Array() ' Return empty array on error
    End If
End Function




Private Sub Label1_Click()

End Sub




Private Sub tbRefNum_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.tbRefNum.value = Format(Date, "yyyy-mm-dd")
End Sub
