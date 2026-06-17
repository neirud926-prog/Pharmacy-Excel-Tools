Attribute VB_Name = "modSurplusScreening"
Option Explicit

' ============================================================================
' modSurplusScreening - Surplus / Shortage lookup for the Return Drug app
' ----------------------------------------------------------------------------
' A scanned item code is checked against the Stocktake database's stock-take
' variance (Quantity = physical count - ERP system count) and classified as
' SURPLUS or SHORTAGE.
'
' Data source : Stocktake.accdb, table [Record]
'   Quantity = TotalSum - ERP  (positive = more physical than ERP = surplus)
' Record used : the latest record for the item (TOP 1 by RecDate, any Type)
' Rule        : variance >= cutoff (20)  -> SURPLUS
'               variance <  cutoff (20)  -> SHORTAGE
'
' This module returns values only; the form decides how to colour the result.
' ============================================================================

Private Const STOCKTAKE_DB As String = _
    "\\nltpha-nas01\PharmShare\Programme Data\Database\Stocktake.accdb"

' Single cut-off: at exactly 20 the item counts as SURPLUS.
Public Const SURPLUS_CUTOFF As Double = 20

' Result codes returned by ClassifySurplus / used by the form.
Public Const RES_SURPLUS  As String = "SURPLUS"
Public Const RES_SHORTAGE As String = "SHORTAGE"

' --- Extract the 6-character item code from a scanned string -----------------
' OP format: "ABCD01"          (length 6)  -> use as-is
' IP format: "XXXXABCD01XXX"   (length>=10)-> characters 5..10 (Mid(s,5,6))
Public Function ParseScanToItemCode(ByVal scanText As String) As String
    scanText = Trim(scanText)
    If Len(scanText) = 6 Then
        ParseScanToItemCode = UCase$(scanText)
    ElseIf Len(scanText) >= 10 Then
        ParseScanToItemCode = UCase$(Mid$(scanText, 5, 6))
    Else
        ' Unexpected length - hand back what we got (caller validates)
        ParseScanToItemCode = UCase$(scanText)
    End If
End Function

' --- Latest stock-take variance for an item code ----------------------------
' Returns the variance (Quantity). found = False when the item has no
' stock-take record at all (caller should show the "no data" state).
Public Function GetItemVariance(ByVal itemCode As String, ByRef found As Boolean) As Double
    On Error GoTo Fail

    found = False
    GetItemVariance = 0

    Dim conn As Object, cmd As Object, rs As Object
    Set conn = modDb.GetConn(STOCKTAKE_DB)      ' reused, persistent connection

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = "SELECT TOP 1 Quantity FROM [Record] " & _
                      "WHERE ItemCode = ? ORDER BY RecDate DESC, ID DESC"
    cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 50, itemCode) ' adVarChar

    Set rs = cmd.Execute
    If Not rs.EOF Then
        If Not IsNull(rs.fields(0).Value) Then
            GetItemVariance = CDbl(rs.fields(0).Value)
            found = True
        End If
    End If

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Exit Function

Fail:
    ' Drop the connection so a fresh one is opened next time, then report "no data"
    modDb.DropConn STOCKTAKE_DB
    found = False
    GetItemVariance = 0
End Function

' --- Classify a variance into SURPLUS / SHORTAGE ----------------------------
Public Function ClassifySurplus(ByVal variance As Double) As String
    If variance >= SURPLUS_CUTOFF Then
        ClassifySurplus = RES_SURPLUS
    Else
        ClassifySurplus = RES_SHORTAGE
    End If
End Function
