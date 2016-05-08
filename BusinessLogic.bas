Attribute VB_Name = "BusinessLogic"
Option Explicit

Public gclsDataAccess As CDataAccess

Public Sub Auto_Open()

Dim sConnection As String
'sConnection = "Provider=SQLOLEDB;" & _
'"Data Source=MyServer;" & _
'"Initial Catalog=Northwind;" & _
'"Integrated Security=SSPI"

sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=C:\Users\Elciak\Documents\sample access db\Northwind.accdb" '& _
'"User ID=UserName;" & _
'"Password=Password;"

Set gclsDataAccess = New CDataAccess
gclsDataAccess.Initialize sConnection

End Sub
Public Sub TestDataAccess()

gclsDataAccess.GetCustomerData Sheet1.Range("A1")
gclsDataAccess.Class_Terminate

End Sub

Sub test()
Call Auto_Open
Call TestDataAccess

End Sub
