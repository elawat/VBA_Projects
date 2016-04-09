Attribute VB_Name = "UDF_Library"
'Function Library Add-ins
'containing user definied functions (UDF)

Option Explicit

Sub RegisterFunction(sFunctionName As String, sDescription As String, iCategory As Integer)
'Comments: Registers function with Excel Function Wizard
'Libraries: standard
'Arguments: funtion name, its description and category to which should be assigned
'           categories:
'               1 Financial
'               2 Date & Time
'               3 Math & Trig
'               4 Statistical
'               5 Lookup & Reference
'               6 Database
'               7 Text
'               8 Logical
'               9 Information
'               10 Commands
'               11 Customizing
'               12 Macro Control
'               13 DDE/External
'               14 User Defined

sDescription = "Calculates age in years"

Application.MacroOptions Macro:=sFunctionName, Description:=sDescription, Category:=iCategory

End Sub


Public Function Age(ByRef dDoB As Date) As Integer
Attribute Age.VB_Description = "Calculates age in years"
Attribute Age.VB_ProcData.VB_Invoke_Func = " \n2"
'Comments: Calculates age in years
'Libraries: standard
'Arguments: date of birth
'Returns: age in years

If dDoB = 0 Then
      Age = ""

Else
      
    If DateSerial(Year(Date), Month(dDoB), Day(dDoB)) > Date Then
        Age = Year(Date) - Year(dDoB) - 1
    Else
        Age = Year(Date) - Year(dDoB)
    End If

End If

End Function


