VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Set a reference to the Microsoft ActiveX Data Objects 2.X Library

Private mcnConnect As ADODB.Connection

Public Sub Class_Terminate()

On Error Resume Next
' Close and destroy the connection object.
mcnConnect.Close
Set mcnConnect = Nothing

End Sub

Public Sub Initialize(ByRef sConnection As String)
' Create the connection to the database.

Set mcnConnect = New ADODB.Connection
mcnConnect.ConnectionString = sConnection
mcnConnect.Open

End Sub

'Retrieving Data

Public Sub GetCustomerData(ByRef rngDestination As Excel.Range)

Dim rsData As ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT * FROM [Top Ten Orders by Sales Amount]"
Set rsData = New ADODB.Recordset

rsData.Open sSQL, mcnConnect, adOpenForwardOnly, adLockReadOnly

If Not rsData.EOF Then
    rngDestination.CopyFromRecordset rsData 'copy from recordset does not retrieve headers
Else
    MsgBox "No data located.", vbCritical, "Error!"
End If

Dim lColumn As Long, lRow As Long

rsData.MoveFirst

If Not rsData.EOF Then 'looping recordset by rows
' Add the column headers.
    For lColumn = 0 To rsData.Fields.Count - 1
        With Sheet2.Range("A1")
        .Offset(0, lColumn).Value = _
        rsData.Fields(lColumn).Name
        End With
    Next lColumn
' Add the data.
    lRow = 1
    Do While Not rsData.EOF
        For lColumn = 0 To rsData.Fields.Count - 1
            With Sheet2.Range("A1")
            .Offset(lRow, lColumn).Value = _
            rsData.Fields(lColumn).Value
            End With
        Next lColumn
    lRow = lRow + 1
    rsData.MoveNext
    Loop
Else
MsgBox "No data located.", vbCritical, "Error!"
End If


rsData.Close

End Sub
