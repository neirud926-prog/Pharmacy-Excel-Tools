Attribute VB_Name = "modDb"
Option Explicit

' ============================================================================
' modDb - cached, reusable ADO connections (performance)
' ----------------------------------------------------------------------------
' Opening an ADODB connection to an Access database on the network share is the
' slowest part of every query. This module keeps ONE open connection per
' database path and hands it back on demand, so repeated scans/queries reuse
' the same connection instead of paying the network handshake each time.
' Connections are dropped and reopened automatically if they go bad.
'
' Usage:
'   Set conn = modDb.GetConn("\\server\share\db.accdb")
'   ' ...run queries, but DO NOT close conn...
' Call modDb.CloseAllConns on workbook close.
' ============================================================================

Private mConns As Object   ' Scripting.Dictionary: dbPath -> ADODB.Connection

Public Function GetConn(ByVal dbPath As String) As Object
    On Error GoTo Reopen

    If mConns Is Nothing Then Set mConns = CreateObject("Scripting.Dictionary")

    If mConns.Exists(dbPath) Then
        Dim cached As Object
        Set cached = mConns(dbPath)
        If Not cached Is Nothing Then
            If cached.State = 1 Then        ' adStateOpen
                Set GetConn = cached
                Exit Function
            End If
        End If
    End If

Reopen:
    On Error GoTo 0
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Persist Security Info=False;"

    If mConns Is Nothing Then Set mConns = CreateObject("Scripting.Dictionary")
    Set mConns(dbPath) = conn               ' add or replace
    Set GetConn = conn
End Function

' Drop a single connection (e.g. after an error on that database) so the next
' GetConn opens a fresh one.
Public Sub DropConn(ByVal dbPath As String)
    On Error Resume Next
    If mConns Is Nothing Then Exit Sub
    If mConns.Exists(dbPath) Then
        Dim c As Object
        Set c = mConns(dbPath)
        If Not c Is Nothing Then If c.State = 1 Then c.Close
        mConns.Remove dbPath
    End If
    On Error GoTo 0
End Sub

' Close every cached connection. Call this from Workbook_BeforeClose.
Public Sub CloseAllConns()
    On Error Resume Next
    If mConns Is Nothing Then Exit Sub
    Dim k As Variant, c As Object
    For Each k In mConns.Keys
        Set c = mConns(k)
        If Not c Is Nothing Then If c.State = 1 Then c.Close
    Next k
    mConns.RemoveAll
    Set mConns = Nothing
    On Error GoTo 0
End Sub
