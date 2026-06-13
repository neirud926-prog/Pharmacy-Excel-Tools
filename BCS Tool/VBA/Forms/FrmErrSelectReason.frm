VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmErrSelectReason 
   Caption         =   "Select Reason"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7485
   OleObjectBlob   =   "FrmErrSelectReason.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmErrSelectReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSubmit_Click()
    On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim fuidValue As String
    Dim fuidLong As Long
    Dim reasonValue As String
    Dim checkByValue As String
    Dim dbPath As String
    
    ' Step 1: Define the path to the Access database
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
    
    ' Step 2: Get the FUID, cbReason, and cbCheckBy values from the form
    fuidValue = Trim(Me.lbDetails.Tag) ' Trim to remove any spaces
    reasonValue = Me.cbReason.value
    checkByValue = Me.cbNameList.value
    
    ' Step 3: Validate inputs
    If IsNull(fuidValue) Or fuidValue = "" Then
        MsgBox "Error: FUID is missing.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(fuidValue) Then
        MsgBox "Error: FUID must be a numeric value.", vbExclamation
        Exit Sub
    End If
    
    fuidLong = CLng(fuidValue) ' Convert to Long Integer
    
    If IsNull(reasonValue) Or reasonValue = "" Then
        MsgBox "Please select a reason before submitting.", vbExclamation
        Exit Sub
    End If
    
    If IsNull(checkByValue) Or checkByValue = "" Then
        MsgBox "Please select a CheckBy value before submitting.", vbExclamation
        Exit Sub
    End If
    
    ' Step 4: Set up the ADODB connection
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    
    ' Step 5: Set up the Command object
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = "UPDATE ErrHandling SET FollowUpAction = ?, CheckBy = ? WHERE FUID = ?"
        .CommandType = adCmdText
        
        ' Step 6: Define parameters
        .Parameters.Append .CreateParameter("FollowUpAction", adVarChar, adParamInput, 255, reasonValue)
        .Parameters.Append .CreateParameter("CheckBy", adVarChar, adParamInput, 255, checkByValue)
        .Parameters.Append .CreateParameter("FUID", adInteger, adParamInput, , fuidLong)
        
        ' Step 7: Execute the command
        .Execute
    End With
    
    ' Step 8: Confirm success
    MsgBox "Record updated successfully!", vbInformation
    
    ' Step 9: Close the form
    Unload Me
    
    ' Step 10: Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    
    ' Step 11: Refresh query
    Call RefreshQuery("Query - Error Handling Query")
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set cmd = Nothing
    Set conn = Nothing
    Unload Me

End Sub
Private Sub UserForm_Initialize()
    
    Dim arrNameList As Variant
    Dim lastRow As Long
    Dim rng As Range
    
    Me.cbReason.AddItem "Checked without problem"
    
    lastRow = wsNameList.Cells(wsNameList.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "No data found in column A of NameList worksheet.", vbExclamation
        Exit Sub
    End If
    
    Set rng = wsNameList.Range("A2:A" & lastRow)
    arrNameList = rng.value
    Me.cbNameList.Clear
    Dim i As Long
    For i = 1 To UBound(arrNameList, 1)
        If Not IsEmpty(arrNameList(i, 1)) Then
            Me.cbNameList.AddItem CStr(arrNameList(i, 1))
        End If
    Next i

End Sub

