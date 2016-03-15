Attribute VB_Name = "UsefulFunctions"
Option Explicit

Function GetPath() As String
'Comments: Opens file dialog and return selected by user path as string
'Libraries: standard
'Arguments: none
'Returns: selected by user path as string (if cancel selected - zero lenght string returned)

Dim sPath As String
Dim oFileDialog As Object

'set and open dialog box to choose file
Set oFileDialog = Application.FileDialog(msoFileDialogFilePicker)
oFileDialog.Title = "Select Index Trades report"
oFileDialog.AllowMultiSelect = False

If oFileDialog.Show = -1 Then 'if user clicks ok
    sPath = oFileDialog.SelectedItems(1)
End If

GetPath = sPath

'check if cancel selected in calling sub
'If Len(Dir(sPath)) = 0 Then
'    MsgBox "You did not select any file."
'    Exit Sub
'End If

End Function

Function ExtractData(sPath As String, sRange As String) As Variant()
'Comments: Returns data as arrays from selected file, from current region for given range
'Libraries: standard
'Arguments: path as string and range
'Returns: array of variant

Dim wkReport As Workbook
Dim aTrades() As Variant

Set wkReport = Workbooks.Open(sPath, , , , , , , , , , , , , True) 'open with local setting to avoid switching to American date format
'wkReport.Sheets(1).Columns(2).NumberFormat = "0"
'wkReport.Sheets(1).Columns(1).NumberFormat = "0"
aTrades = wkReport.Sheets(1).Range(sRange).CurrentRegion.Value

ExtractData = aTrades

wkReport.Close (False)

End Function

Function CheckIfEmpty(aData() As Variant, iIndex As Integer) As Boolean
'Comments: Checks if there is empty element in array
'Libraries: standard
'Arguments: array to check as variant and integer to specify dimension to browse
'Returns: true if there is emptu element

Dim icounter As Integer

CheckIfEmpty = False

For icounter = LBound(aData, 1) To UBound(aData, 1)

    If aData(icounter, iIndex) = "" Then
        CheckIfEmpty = True
        Exit For
    End If
Next icounter
        
End Function
