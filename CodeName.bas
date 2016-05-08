Attribute VB_Name = "CodeName"
'Going through Pearson blog codename section
'http://www.cpearson.com/excel/RenameProblems.aspx

Option Explicit

'While a Worksheet has a property to get the CodeName property,
'there is no such proprety that will give you a Worksheet object from a CodeName property.
'However, some simple VBA code can provide this functionality.

'debug.Print Sheet2.Name
'Sheet2 is a codename

'This procedure takes a string variable containing a CodeName
'and returns a Worksheet object associated with the code name.
Function GetWorksheetFromCodeName(CodeName As String) As Worksheet

Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
    If StrComp(WS.CodeName, CodeName, vbTextCompare) = 0 Then '0 - equal
        Set GetWorksheetFromCodeName = WS
        Exit Function
    End If
Next WS

End Function

'getting worksheet name from defined range on this sheet

Function GetWorksheetFromName(NameText As String) As Worksheet
    With ThisWorkbook
        Set GetWorksheetFromName = .Names(NameText).RefersToRange.Worksheet
    End With
End Function

'Changing The CodeName Of A Worksheet
'ThisWorkbook.VBProject.VBComponents("Sheet1").Name = "SummarySheet"
'or manual in VBA Editor

Sub test()

Dim WS As Worksheet
Set WS = GetWorksheetFromCodeName("SummarySheet")
Debug.Print WS.Name

Set WS = GetWorksheetFromName("Summary")
Debug.Print WS.Name

End Sub
