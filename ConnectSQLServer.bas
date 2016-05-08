Attribute VB_Name = "ConnectSQLServer"
' Description:  Exporting data from excel to existing table in SQL Server
'               using ADOBD

'Option Explicit

Type ExportSourceType
    sServerName As String
    sTableName As String
    sarrColNames() As String
End Type

Sub test()

Dim tdbExport As ExportSourceType
    tdbExport.sServerName = "DESKTOP-4491DNA\SQLEXPRESS"
    tdbExport.sTableName = "Authors"
    ReDim tdbExport.sarrColNames(1 To 10)
    tdbExport.sarrColNames(1) = "ID"
    tdbExport.sarrColNames(2) = "lname"
    tdbExport.sarrColNames(3) = "fname"
    tdbExport.sarrColNames(4) = "phone"
    tdbExport.sarrColNames(5) = "city"
    tdbExport.sarrColNames(6) = "county"
    tdbExport.sarrColNames(7) = "postcode"
    tdbExport.sarrColNames(8) = "sex"
    tdbExport.sarrColNames(9) = "salary"
    tdbExport.sarrColNames(10) = "topic"

Dim aData() As Variant
aData = ThisWorkbook.Sheets(1).Range("A1").CurrentRegion.Value

Call ExportToSQLServer(aData, True, tdbExport)

End Sub


Sub ExportToSQLServer(varrExportData() As Variant, bHasHeaders As Boolean, tSource As ExportSourceType)
'Comments: Export data passed as array into SQL Server
'Libraries: Microsoft ActiveX Data Objects 2.8 Library
'Arguments: varrExportData - array to export to db; bHasHeaders - true if array contain headers;
'           tSource - user defined type; server name, table name and columns
'Returns: n.a.
'Date Developer Action
' ———————————————————————————————
'14/03/2016 EW

Dim conn As New ADODB.Connection
Dim sSQL As String
Dim sConnString As String
Dim iSkipHeaders As Integer
Dim icounter As Integer 'loop through array with data
Dim iCounter2 As Integer 'loop through columns in table to bulid sql statement
Dim iCounter3 As Integer ''loop through values in array to insert to bulid sql statement

iSkipHeaders = 0

If bHasHeaders = True Then
iSkipHeaders = 1
End If
       
'Open a connection to SQL Server
sConnString = "Provider=SQLOLEDB;Data Source=" & tSource.sServerName & ";Initial Catalog=TestDB;Integrated Security=SSPI;"
conn.Open sConnString
                       
'Loop through array
For icounter = LBound(varrExportData, 1) + iSkipHeaders To UBound(varrExportData, 1)
    'build sql statement
     sSQL = "INSERT INTO [" & tSource.sTableName & "] ("
     For iCounter2 = LBound(tSource.sarrColNames) To UBound(tSource.sarrColNames)
        If iCounter2 <> UBound(tSource.sarrColNames) Then
        sSQL = sSQL & tSource.sarrColNames(iCounter2) & ", "
        Else
        sSQL = sSQL & tSource.sarrColNames(iCounter2) & ") "
        End If
     Next iCounter2
     sSQL = sSQL & "Values ("
     For iCounter3 = LBound(varrExportData, 2) To UBound(varrExportData, 2)
        If iCounter3 <> UBound(varrExportData, 2) Then
        sSQL = sSQL & "'" & varrExportData(icounter, iCounter3) & "',"
        Else
        sSQL = sSQL & "'" & varrExportData(icounter, iCounter3) & "');"
        End If
     Next iCounter3
     
'Generate and execute sql statement to import values in array to SQL Server table
conn.Execute sSQL

Next icounter
        
MsgBox "Imported."
                
conn.Close
Set conn = Nothing
             
End Sub


Sub ExportToSQLServerRange()
'from range, looping through rows
'references: Microsoft ActiveX Data Objects 2.8 Library
'Date Developer Action
' ———————————————————————————————
'14/03/2016 EW adjusted http://tomaslind.net/2013/12/26/export-data-excel-to-sql-server/

Dim conn As New ADODB.Connection
Dim iRowNo As Integer
Dim sReference As String
Dim sTitle As String
  
With Sheets("References")
            
    'Open a connection to SQL Server
    'conn.Open "Provider=SQLOLEDB;Data Source=server_name;Initial Catalog=Sepon;Integrated Security=SSPI;"
                
    'Skip the header row
    iRowNo = 2
                
    'Loop until empty cell in CustomerId
    Do Until .Cells(iRowNo, 1) = ""
        sReference = .Cells(iRowNo, 1)
        sTitle = .Cells(iRowNo, 2)
                        
        'Generate and execute sql statement to import the excel rows to SQL Server table
        conn.Execute "insert into [References] (Reference, Title) values ('" & sReference & "', '" & sTitle & "');"
         
        iRowNo = iRowNo + 1
    Loop
                
    MsgBox "Imported."
                
    conn.Close
Set conn = Nothing
             
End With
 
End Sub


Public mcnSQLServer As ADODB.Connection

Public Sub ConnectToSQLServer()
Const sSOURCE As String = "ConnectToSQLServer"
Dim lAttempt As Long
Dim sConnect As String

On Error GoTo ErrorHandler
'Create the connection string.
sConnect = "Provider=SQLOLEDB;" & _
"Data Source=DESKTOP-4491DNA\SQLEXPRESS;" & _
"Initial Catalog=TestDB;" & _
"Integrated Security=SSPI"

'Attempt to open the connection.
Application.StatusBar = "Attempting to connect..."
Set mcnSQLServer = New ADODB.Connection
mcnSQLServer.ConnectionString = sConnect
mcnSQLServer.Open
'Close connection to enable connection pooling.
mcnSQLServer.Close
Stop

ErrorExit:
Application.StatusBar = False
Exit Sub

ErrorHandler:
'We will try to make the connection three times before bailing out.
If lAttempt < 3 And mcnSQLServer.Errors.Count > 0 Then
    If mcnSQLServer.Errors(0).NativeError = 17 Then
        Application.StatusBar = "Retrying connection..."
        lAttempt = lAttempt + 1
        Resume
    Else
        Resume ErrorExit
    End If
Else
    Resume ErrorExit
End If
'If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
'    Stop
'    Resume
'    Else
'    Resume ErrorExit
'End If

End Sub


Public Sub RunStoredProcAsMethod()

Dim cnSQLServer As ADODB.Connection
Dim rsData As ADODB.Recordset
Dim sConnect As String

'Clear the destination worksheet.
Sheet1.UsedRange.Clear

'Create the connection string.
sConnect = "Provider=SQLOLEDB;" & _
"Data Source=DESKTOP-4491DNA\SQLEXPRESS;" & _
"Initial Catalog=TestDB;" & _
"Integrated Security=SSPI"

'Create the Connection and Recordset objects.
Set cnSQLServer = New ADODB.Connection
Set rsData = New ADODB.Recordset

'Open the connection and execute the stored procedure.
cnSQLServer.Open sConnect
cnSQLServer.spGetAuthorInfo "Berkeley", rsData 'stored procedure in SQL Server, takimg one arg, returing recordset

'Make sure we got records back
If Not rsData.EOF Then
    'Dump the contents of the recordset onto the worksheet.
    ThisWorkbook.Sheets(1).Range("A1").CopyFromRecordset rsData
Else
    MsgBox "Error: No records returned.", vbCritical
End If

'Clean up our ADO objects.
rsData.Close
If CBool(cnSQLServer.State And adStateOpen) Then cnSQLServer.Close
Set cnSQLServer = Nothing
Set rsData = Nothing
End Sub


'insert records using command object and stored procedure

Public Sub Insert()

Dim cmAccess As ADODB.Command
Dim objParams As ADODB.Parameters
Dim lAffected As Long
Dim sPath As String
Dim sConnect As String
Dim sSQL As String

Dim aData() As Variant

aData = ThisWorkbook.Sheets(1).Range("A1").CurrentRegion.Value


'Create the connection string.
sConnect = "Provider=SQLOLEDB;" & _
"Data Source=DESKTOP-4491DNA\SQLEXPRESS;" & _
"Initial Catalog=TestDB;" & _
"Integrated Security=SSPI"

'Create the Command object.
Set cmSQLServer = New ADODB.Command
cmSQLServer.ActiveConnection = sConnect
cmSQLServer.CommandText = "spAddEmployee"
cmSQLServer.CommandType = adCmdStoredProc

'Create and append the parameters of stored query.
Set objParams = cmSQLServer.Parameters
 With objParams
    .Append cmSQLServer.CreateParameter("ID", adVarChar, adParamInput, 50)
    .Append cmSQLServer.CreateParameter("fname", adVarChar, adParamInput, 50)
    .Append cmSQLServer.CreateParameter("lname", adVarChar, adParamInput, 50)
    .Append cmSQLServer.CreateParameter("job_id", adSmallInt, adParamInput)
    .Append cmSQLServer.CreateParameter("job_lvl", adSmallInt, adParamInput)
    .Append cmSQLServer.CreateParameter("pub_id", adVarChar, adParamInput, 10)
    .Append cmSQLServer.CreateParameter("hire_date", adDate, adParamInput, 10)
  End With
Set objParams = Nothing

'Load the parameters and execute the query.
For icounter = LBound(aData, 1) + 1 To UBound(aData)
cmSQLServer.Parameters("ID").Value = aData(icounter, 1)
cmSQLServer.Parameters("fname").Value = aData(icounter, 2)
cmSQLServer.Parameters("lname").Value = aData(icounter, 3)
cmSQLServer.Parameters("job_id").Value = aData(icounter, 4)
cmSQLServer.Parameters("job_lvl").Value = aData(icounter, 5)
cmSQLServer.Parameters("pub_id").Value = aData(icounter, 6)
cmSQLServer.Parameters("hire_date").Value = aData(icounter, 7)
cmSQLServer.Execute lAffected, , adExecuteNoRecords

'Verify the correct number of records updated.
If lAffected <> 1 Then

MsgBox "Error updating record.", vbCritical, "Error!"
End If
Next icounter

Set cmSQLServer = Nothing

End Sub



Public Sub ExtractMultipleRecordsets()

Dim rsData As ADODB.Recordset

'Clear the destination worksheet.
Sheet1.UsedRange.Clear

'Procedure from Listing 19-16
sConnect = "Provider=SQLOLEDB;" & _
"Data Source=DESKTOP-4491DNA\SQLEXPRESS;" & _
"Initial Catalog=TestDB;" & _
"Integrated Security=SSPI"

'Attempt to open the connection.
Set mcnSQLServer = New ADODB.Connection
mcnSQLServer.ConnectionString = sConnect
mcnSQLServer.Open


'Create and open the Recordset object.
Set rsData = New ADODB.Recordset
rsData.Open "spGetLookupTables", mcnSQLServer, _
adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

If Not rsData.EOF Then
'The first recordset contains the Suppliers list.
Sheet1.Range("A1").CopyFromRecordset rsData
Set rsData = rsData.NextRecordset

'The second recordset contains the Customers list.
Sheet1.Range("D1").CopyFromRecordset rsData
Set rsData = rsData.NextRecordset

'The third recordset contains the Shippers list.
Sheet1.Range("G1").CopyFromRecordset rsData
Set rsData = rsData.NextRecordset
'There is no need to clean up the Recordset object at this
'point. It will be closed and set to Nothing automatically
'by ADO after the last call to the NextRecordset method.

Else
MsgBox "No data located.", vbCritical, "Error!"
End If

'Close the pooled connection
mcnSQLServer.Close
End Sub
