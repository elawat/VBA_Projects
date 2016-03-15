Attribute VB_Name = "UsefulSubs"
Option Explicit

Sub ListFiles(ByVal sSourceFolderPath As String, ByVal ws As Worksheet, ByVal rngForName As Range, rngForPath As Range)
'Comments: List names and paths of all the files in the folder
'          and copies them to specified range
'Libraries: standard
'Arguments: path of the folder to list files from as string; worksheet and ranges where to copy names and paths


Dim oFSO As Object
Dim oSourceFolder As Object
Dim oFileItem As Object
Dim iPathRow As Integer
Dim iPathCol As Integer
Dim iNameRow As Integer
Dim iNameCol As Integer

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oSourceFolder = oFSO.GetFolder(sSourceFolderPath)

iPathRow = rngForPath.row
iPathCol = rngForPath.Column
iNameRow = rngForName.row
iNameCol = rngForName.Column


For Each oFileItem In oSourceFolder.Files
  'display file properties
  ws.Cells(iNameRow, iNameCol).Value = oFileItem.Name
  ws.Cells(iPathRow, iPathCol).Value = oFileItem.Path

  iNameRow = iNameRow + 1
  iPathRow = iPathRow + 1

Next oFileItem

'ws.Columns(iNameCol).AutoFit
'ws.Columns(iPathCol).AutoFit


Set oFileItem = Nothing
Set oSourceFolder = Nothing
Set oFSO = Nothing

End Sub

Sub Identify_Export(ByVal ws As Worksheet, aData() As Variant, aFind() As String, rngStart As Range, sOutput As String)
'Comments: Check if some value exists in array
'Libraries: standard

Dim icounter As Integer
Dim iCounter2 As Integer

For icounter = LBound(aData, 1) To UBound(aData, 1)
    For iCounter2 = LBound(aFind) To UBound(aFind)
        If InStr(CStr(aData(icounter, 2)), aFind(iCounter2)) Then
            ws.Cells(rngStart.row - 1 + icounter, rngStart.Column).Value = sOutput
        End If
    Next iCounter2
Next icounter

End Sub
