Attribute VB_Name = "ExportToSQLServer"
' Description:  Exporting data from excel to existing table in SQL Server
'               using ADOBD

Option Explicit

Sub ExportToSQLServer()

'references: Microsoft ActiveX Data Objects 2.8 Library

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
