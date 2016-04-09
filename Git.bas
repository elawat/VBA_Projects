Attribute VB_Name = "Git"
' Description:  Exporting VBA modules and classes

Option Explicit

'Call ExportSourceFiles("C:\Users\Elciak\Documents\VBA_Projects")

Public Sub ExportSourceFiles(destPath As String)
'Comments: Exports VBA modules and classes into specified repository
'Libraries: Microsoft Visual Basic for Applications Extensibility 5.3
'Arguments: path where to export as string
'Returns: n.a.
'Date Developer Action
' ———————————————————————————————
'13/03/2016 EW adjusted from https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/

Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export (destPath & "\" & component.Name & ToFileExtension(component.Type))
        End If
    Next
    MsgBox "Exported"
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
'Comments: Coverts VBA component type into string
'Libraries: Microsoft Visual Basic for Applications Extensibility 5.3
'Arguments: path where to export as string
'Returns: VBA component type as string
'Date Developer Action
' ———————————————————————————————
'13/03/2016 EW adjusted from https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/

Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function

Public Sub RemoveAllModules()
'remove all code but Git module
Dim project As VBProject
Set project = Application.VBE.ActiveVBProject
 
Dim comp As VBComponent
For Each comp In project.VBComponents
If Not comp.Name = "Git" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
project.VBComponents.Remove comp
End If
Next
End Sub

Public Sub ImportSourceFiles(sourcePath As String)
'import all modules but Git module
Dim file As String
file = Dir(sourcePath)
While file <> vbNullString And Right(file, 7) <> "Git.bas"
Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & "\" & file
file = Dir
Wend
End Sub

