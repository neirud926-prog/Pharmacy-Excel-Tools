VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmErrorHandling 
   Caption         =   "Follow Up"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14640
   OleObjectBlob   =   "FrmErrorHandling.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmErrorHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ListErr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    ' Get the index of the selected item
    Dim selectedIndex As Integer
    selectedIndex = Me.ListErr.ListIndex
    
    ' Check if an item is selected
    If selectedIndex >= 0 Then
        ' Array to store the 10 column values
        Dim values(0 To 9) As String
        Dim i As Integer
        
        ' Retrieve values from all columns of the selected row
        For i = 0 To 9
            values(i) = Me.ListErr.List(selectedIndex, i)
        Next i
        
        ' Format the data into a display string
        Dim displayText As String
        displayText = ""
        Dim columnNames As Variant
        ' Define column names based on your data structure
        columnNames = Array("ID", "Date and Time", "Code on Drug", "Code on Binshelf", "Location", "Result", "Type", "Ref No.", "Fill By")
        
        ' Build the display text with column names and values
        For i = 0 To 8
            displayText = displayText & columnNames(i) & ": " & values(i) & vbCrLf
        Next i
        

        FrmErrSelectReason.lbDetails.Caption = displayText
        FrmErrSelectReason.lbDetails.Tag = Me.ListErr.List(0, 0)
        FrmErrSelectReason.Show
        
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim dbPath As String
    Dim sql As String
    Dim i As Integer
    
    ' Path to your Access database
    dbPath = "\\nltpha-nas01\PharmShare\Programme Data\Database\Replenishment.accdb"
    
    ' The query to execute
    sql = "SELECT * FROM [Error Handling Query FU]"
    
    ' Create and open the connection
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    
    ' Open the recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Configure the listbox
        With Me.ListErr
            .Clear              ' Clear any existing items
            .ColumnCount = 10   ' Set to 10 columns as specified
        
        ' Populate the listbox with data
        Do While Not rs.EOF
   
            .AddItem  ' Add a new row
            ' Fill the 10 columns for the current row
            For i = 0 To 9
                .List(.ListCount - 1, i) = IIf(IsNull(rs.Fields(i).value), "", rs.Fields(i).value)
            Next i
        
        rs.MoveNext  ' Move to the next record
        Loop
        End With
    
    
    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    wsFollowUp.ListObjects("Error_Handling_Query").QueryTable.BackgroundQuery = False
    wsFollowUp.ListObjects("Error_Handling_Query").QueryTable.Refresh
End Sub
