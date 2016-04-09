Attribute VB_Name = "BusinessLogic"
Option Explicit
Public gclsDataAccess As CDataAccess

Public Sub Auto_Open()

Dim sConnection As String
sConnection = "Provider=SQLOLEDB;" & _
"Data Source=MyServer;" & _
"Initial Catalog=Northwind;" & _
"Integrated Security=SSPI"

Set gclsDataAccess = New CDataAccess
gclsDataAccess.Initialize sConnection

End Sub
Public Sub TestDataAccess()

gclsDataAccess.GetCustomerData Sheet1.Range(“A1”)

End Sub

