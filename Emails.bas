Attribute VB_Name = "Emails"
'Using Outlook objetc library

Option Explicit
Sub SendChart()
'Comments: Paste excel chart into email
'Libraries:
'Arguments:
'Returns: n.a.
'Date Developer Action
' ———————————————————————————————
'14/04/2016 EW

Dim olapp As Object
Dim olmail As Object
Dim sPath As String

Set olapp = CreateObject("Outlook.application")
Set olmail = olapp.createitem(0)
sPath = "C:\Users\Elciak\Desktop\Chart1.png"

ThisWorkbook.Sheets(1).ChartObjects("Chart 1").Chart.Export sPath

With olmail
.display
.to = "test@gmail.com"
.Subject = "Chart"
.HTMLBody = .HTMLBody & "<img src=" & sPath & ">"

End With

End Sub
