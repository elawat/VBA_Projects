Attribute VB_Name = "ConnectAccess"
Option Explicit

'stored query in Access db
'UPDATE Books SET Books.notes = [NewNote] WHERE Books.title_id = [ID];

Public Sub UpdateAccess()

Dim cmAccess As ADODB.Command
Dim objParams As ADODB.Parameters
Dim lAffected As Long
Dim sPath As String
Dim sConnect As String

'Get the database path
sPath = "C:\Users\Elciak\Documents\sample access db"
If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"

'Create the connection string.
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & sPath & "Books2010.accdb;"

'Create the Command object.
Set cmAccess = New ADODB.Command
cmAccess.ActiveConnection = sConnect
cmAccess.CommandText = "UpdateBookNote" 'name of saved query in Access
cmAccess.CommandType = adCmdStoredProc 'We tell ADO that we are running a stored query by setting the
                                        'CommandType property of the Command object to adCmdStoredProc.

'Create and append the parameters of stored query.
Set objParams = cmAccess.Parameters 'For each parameter in the Access parameter query
                                    'we must create a Parameter object and add it to the Parameters collection.
                                    'These parameters must be created and added to the collection in exactly
                                    'the same order as they appear in the SQL of the Access parameter query.
objParams.Append cmAccess.CreateParameter("NewNote", _
adVarChar, adParamInput, 20)
objParams.Append cmAccess.CreateParameter("ID", _
adVarChar, adParamInput, 6)
Set objParams = Nothing

'Load the parameters and execute the query.
cmAccess.Parameters("NewNote").Value = "test"
cmAccess.Parameters("ID").Value = "TC3218"
cmAccess.Execute lAffected, , adExecuteNoRecords

'Verify the correct number of records updated.
If lAffected <> 1 Then

MsgBox "Error updating record.", vbCritical, "Error!"
End If

'deleting
'sSQL = "DELETE FROM Shippers " & _
'"WHERE CompanyName = 'Excellent Shipping';"
''Create and execute the Command object.
'Set cmAccess = New ADODB.Command
'cmAccess.ActiveConnection = sConnect
'cmAccess.CommandText = sSQL
'cmAccess.CommandType = adCmdText
'cmAccess.Execute lAffected, , adExecuteNoRecords

Set cmAccess = Nothing

End Sub
