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

Private Sub Class_Terminate()
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

Public Sub GetCustomerData(ByRef rngDestination As Excel.Range)
Dim rsData As ADODB.Recordset
Dim sSQL As String
sSQL = "SELECT * FROM Customers"
Set rsData = New ADODB.Recordset
rsData.Open sSQL, mcnConnect, adOpenForwardOnly, adLockReadOnly
If Not rsData.EOF Then
rngDestination.CopyFromRecordset rsData
End If
rsData.Close
End Sub
