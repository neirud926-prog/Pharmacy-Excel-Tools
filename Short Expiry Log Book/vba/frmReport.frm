VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "Print Report"
   ClientHeight    =   2325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4020
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub btnPrint_Click()
    Call UBDateStatusFilter(Me.cbYear.Value, Me.cbMonth.Value)
    Call SortByLoc
    Dim prtRng As Range
    Set prtRng = wsRecord.Range("B1", wsRecord.Range("G1").End(xlDown))
    prtRng.PrintOut Copies:=1, Collate:=True, ActivePrinter:="Kyocera COLOR TASKalfa 4052ci (Pharm)"
    Call ChangeStatusToEnd
End Sub

Private Sub UserForm_Initialize()
    Me.cbYear.List = Array(Year(Date), Year(Date) + 1)
    Me.cbMonth.List = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
End Sub
