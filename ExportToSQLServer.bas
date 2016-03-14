Attribute VB_Name = "ExportToSQLServer"
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
    tdbExport.sTableName = "ReferencesTest"
    ReDim tdbExport.sarrColNames(1 To 2)
    tdbExport.sarrColNames(1) = "Reference"
    tdbExport.sarrColNames(2) = "Title"

Dim aData() As Variant
aData = ThisWorkbook.Sheets(1).Range("A1:B10").Value

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
Dim iCounter As Integer 'loop through array with data
Dim iCounter2 As Integer 'loop through columns in table to bulid sql statement
Dim iCounter3 As Integer ''loop through values in array to insert to bulid sql statement

iSkipHeaders = 0

If bHasHeaders = True Then
iSkipHeaders = 1
End If
       
'Open a connection to SQL Server
sConnString = "Provider=SQLOLEDB;Data Source=" & tSource.sServerName & ";Initial Catalog=Sepon;Integrated Security=SSPI;"
conn.Open sConnString
                       
'Loop through array
For iCounter = LBound(varrExportData, 1) + iSkipHeaders To UBound(varrExportData, 1)
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
        sSQL = sSQL & "'" & varrExportData(iCounter, iCounter3) & "',"
        Else
        sSQL = sSQL & "'" & varrExportData(iCounter, iCounter3) & "');"
        End If
     Next iCounter3
     
'Generate and execute sql statement to import values in array to SQL Server table
conn.Execute sSQL

Next iCounter
        
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

