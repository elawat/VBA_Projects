Attribute VB_Name = "DatesFunctions"
' Description:  Function related to dates

Option Explicit

Function LastWorkingDay(dtDate As Date)
'Comments: Does not take into account bank holidays.
'Arguments: date from the month for which last working day (not weekend) should be found
'Returns: last week day of month
'Date Developer Action
' ———————————————————————————————
'25/03/2016 EW created

LastWorkingDay = LastDayOfMonth(dtDate)
Dim iCounter As Integer

For iCounter = 1 To 7
If Format(LastWorkingDay, "ddd") = "Sat" Or Format(LastWorkingDay, "ddd") = "Sun" Then
    LastWorkingDay = LastWorkingDay - 1
Else
    LastWorkingDay = LastWorkingDay
    Exit For
End If
Next iCounter

End Function

Function ChosenLastDayOfMonth(sDDD As String, dDate As Date)
'Comments: To find last selected day of week of month
'Arguments: 1) selected day of week in format "Mon", "Tue" etc.
'           2) date from the month for which last selected day should be found
'Returns: last selected day of month
'Date Developer Action
' ———————————————————————————————
'15/03/2016 EW created

Dim dLastDay As Date
Dim iCounter As Integer

dLastDay = LastDayOfMonth(dDate)

For iCounter = 1 To 7
    If Format(dLastDay, "ddd") = sDDD Then
        ChosenLastDayOfMonth = dLastDay
        Exit Function
    Else
        dLastDay = DateAdd("d", -1, dLastDay)
    End If
Next iCounter

End Function

Function LastDayOfMonth(dDate As Date) As Date

LastDayOfMonth = DateSerial(Year(dDate), Month(dDate) + 1, 0)

End Function

Function FirstDayOfMonth(dDate As Date) As Date

FirstDayOfMonth = DateSerial(Year(dDate), Month(dDate), 1)

End Function

