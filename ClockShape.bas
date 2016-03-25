Attribute VB_Name = "ClockShape"
' Description:  Changes HH shape in GMT to Clock Shape (50HH) for British Summer Time (BST)

Option Explicit

Function ChangeToClockShape(aShape() As Variant) As Variant()
'Comments: Cahnges GMT HH shape to Clock Shape for BSM
'Arguments: HH GMT shape as an array of variant
'Returns: Clock Shape
'Date Developer Action
' ———————————————————————————————
'25/03/2016 EW created

Dim aFinalShape() As Variant
ReDim aFinalShape(1 To 51, 1 To 1)
Dim iCounter As Long
Dim icounter2 As Long
Dim dtMar As Date, dtOct As Date


For iCounter = LBound(aShape, 1) To UBound(aShape, 1) 'loop through each record/row in GMT shape(day)
    ReDim Preserve aFinalShape(1 To 51, 1 To iCounter)
    dtMar = ChosenLastDayOfMonth("Sun", DateSerial(Year(aShape(iCounter, 1)), 3, 1))
    dtOct = ChosenLastDayOfMonth("Sun", DateSerial(Year(aShape(iCounter, 1)), 10, 1))
   
    For icounter2 = LBound(aShape, 2) To UBound(aShape, 2) 'do for each column/cell
        
            Select Case aShape(iCounter, 1) 'different logic for different day
            Case Is < dtMar, Is > dtOct 'no changes, HH49 and 50 - 0
                If icounter2 < 48 Then
                    aFinalShape(icounter2, iCounter) = aShape(iCounter, icounter2)
                Else
                    aFinalShape(icounter2, iCounter) = aShape(iCounter, icounter2)
                    aFinalShape(icounter2 + 2, iCounter) = 0
                End If
            Case Is = dtMar 'shift from HH 3
                If icounter2 < 4 Then
                    aFinalShape(icounter2, iCounter) = aShape(iCounter, icounter2)
                ElseIf icounter2 < 6 Then
                    aFinalShape(icounter2, iCounter) = 0
                    aFinalShape(icounter2 + 2, iCounter) = aShape(iCounter, icounter2)
                ElseIf icounter2 < 48 Then
                    aFinalShape(icounter2 + 2, iCounter) = aShape(iCounter, icounter2)
                Else
                    ReDim Preserve aFinalShape(1 To 51, 1 To iCounter + 1)
                    aFinalShape(icounter2 - 46, iCounter + 1) = aShape(iCounter, icounter2)
                    aFinalShape(icounter2 + 2, iCounter) = 0
                End If
            Case Is = dtOct 'all shifted by 2 HH, last two HH as 49 and 50
                If icounter2 = 1 Then
                    aFinalShape(icounter2, iCounter) = aShape(iCounter, icounter2)
                Else
                    aFinalShape(icounter2 + 2, iCounter) = aShape(iCounter, icounter2)
                End If
            Case Is > dtMar, Is < dtOct 'all shifted by 2 HH
                If icounter2 = 1 Then
                    aFinalShape(icounter2, iCounter) = aShape(iCounter, icounter2)
                ElseIf icounter2 < 48 Then
                    aFinalShape(icounter2 + 2, iCounter) = aShape(iCounter, icounter2)
                Else
                    ReDim Preserve aFinalShape(1 To 51, 1 To iCounter + 1)
                    aFinalShape(icounter2 - 46, iCounter + 1) = aShape(iCounter, icounter2)
                    aFinalShape(icounter2 + 2, iCounter) = 0
                End If
            End Select
                
    Next icounter2
 Next iCounter


ChangeToClockShape = aFinalShape

End Function

Sub test()

Dim aGMT() As Variant
Dim aClockShape() As Variant

ThisWorkbook.Sheets(1).Columns(1).NumberFormat = "0"

aGMT = ThisWorkbook.Sheets(1).Range("A1").CurrentRegion.Value

aClockShape() = ChangeToClockShape(aGMT)

ThisWorkbook.Sheets(2).Range("A1").Resize(UBound(aClockShape, 2), UBound(aClockShape, 1)).Value = Application.Transpose(aClockShape)

ThisWorkbook.Sheets(1).Columns(1).NumberFormat = "dd/mm/yyyy"
ThisWorkbook.Sheets(2).Columns(1).NumberFormat = "dd/mm/yyyy"

MsgBox "Converted to clock shape"

End Sub

