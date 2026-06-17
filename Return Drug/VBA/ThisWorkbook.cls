VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    wsReport.Visible = xlSheetHidden

    ' Load only what the app needs immediately (bin-shelf locator is used by
    ' scanning) so the workbook becomes usable right away. The remaining
    ' network-file loads are deferred to run just after open completes, instead
    ' of making the user wait for all seven files up front.
    LoadBinshelfLocator
    Application.OnTime Now, "LoadSecondaryLocators"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Release any cached DB connections so Access locks clear cleanly.
    On Error Resume Next
    modDb.CloseAllConns
    On Error GoTo 0
End Sub
