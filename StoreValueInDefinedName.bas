Attribute VB_Name = "StoreValueInDefinedName"
'Description:  Storing values in defined names instead of cells

 
Sub SaveValueInDefinedName(sName As String, vReferVal As Variant, Optional ByVal ws As Worksheet)
'Comments: Saves defined name with its value
'Libraries: standard
'Arguments: name and value of defined name to be saved,
'           optionally worksheets objectif name should be worksheet level
'Returns: n.a.
'Date Developer Action
' ———————————————————————————————
'09/04/2016 EW created
         
If ws Is Nothing Then
    ThisWorkbook.Names.Add sName, vReferVal 'workbook level name
Else
    ws.Names.Add sName, vReferVal 'sheet level name
End If

End Sub

Function GetNameRefersTo(sDefName As String) As String
'Comments: Retrievs value of defined name
'Libraries: standard
'Arguments: name of defined name
'Returns: value of defined name as string
'Date Developer Action
' ———————————————————————————————
'09/04/2016 EW adjusted from http://www.cpearson.com/excel/DefinedNames.aspx

Dim sDefNameString As String
Dim bHasRef As Boolean
Dim rngDefName As Range
Dim nDefName As Name
    
Set nDefName = ThisWorkbook.Names(sDefName)
    
On Error Resume Next
    
Set rngDefName = nDefName.RefersToRange 'error if name does not refer to range
If Err.Number = 0 Then ' no error
        bHasRef = True
Else
        bHasRef = False
End If
    
If bHasRef = True Then
        sDefNameString = rngDefName.Text
Else
        sDefNameString = nDefName.RefersTo
    If StrComp(Mid(S, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
            ' text constant
            sDefNameString = Mid(sDefNameString, 3, Len(S) - 3)
        Else
            ' numeric contant
            sDefNameString = Mid(sDefNameString, 2)
    End If
End If

GetNameRefersTo = sDefNameString

End Function
