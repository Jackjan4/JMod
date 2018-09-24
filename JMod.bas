Attribute VB_Name = "JMod"
' Module by Jan-Philipp Roslan
' Contains utility function for better code maintenance and to reduce duplicated code
' Version: 0.5
' Date: 24.09.18
' Time: 09:55


Public Const DoubleMin As Double = -1.79769313486231E+308
Public Const DoubleMax As Double = 1.79769313486231E+308

' Using a CommandButton for toggling an area
Sub FoldingButton(ByVal area As String, ByVal row As Boolean, ByRef btn As CommandButton, ByVal openednstr As String, ByVal closedstr As String)

    Dim Zelle As String
    Dim hidden As Boolean

    ' VarInit
    Application.ScreenUpdating = False
    Zelle = ActiveCell.Address

    If btn.BackColor = RGB(153, 255, 153) Then
        With btn
            .BackColor = RGB(255, 192, 0)
            .Caption = closedstr
            .ForeColor = RGB(192, 0, 0)
            .Font.Bold = True
        End With
        ' Area will be hidden
        hidden = True
    Else
        With btn
            .BackColor = RGB(153, 255, 153)
            .Caption = openednstr
            .ForeColor = RGB(0, 0, 0)
            .Font.Bold = False
        End With

        ' Area will be shown
        hidden = False
    End If

    range(Zelle).Select

    ' Apply area change
    If (row = True) Then
        range(area).EntireRow.hidden = hidden
    Else
        range(area).EntireColumn.hidden = hidden
    End If

    ' Update screen
    Application.ScreenUpdating = True

End Sub

' Check if cell is empty
Function CellIsEmpty(ByVal range As range)

    If range.Value = "" Or IsEmpty(range) Then
        CellIsEmpty = True
    Else
        CellIsEmpty = False
    End If

End Function


' Value of cell
Function Val(ByVal cell As Variant)

    Select Case TypeName(cell)

    Case "string", "String":
        Val = range(cell).Value

    Case "Range", "range":
        Val = cell.Value
    End Select
End Function


' Check if named range exists
Function RangeExists(ByVal name As String) As Boolean
    Dim Test As range
    On Error Resume Next
    Set Test = ActiveSheet.range(name)
    RangeExists = Err.Number = 0
End Function


' Average
Function Avg(ByVal range As range)

    Avg = Application.WorksheetFunction.Average(range)

End Function


' Pi
Function Pi()

    Pi = 4 * Atn(1)

End Function


'Right pad
Function RightPad(ByVal source As String, ByVal inserter As String, ByVal wishedLength As Integer)

    Dim result As String
    result = source

    While Len(result) < wishedLength
        result = result & inserter

    Wend

    RightPad = result

End Function



' If assignment
Function IfA(ByVal expression As Boolean, trueVal As Variant, falseVal As Variant)

If expression = True Then
    IfA = trueVal
Else
    IfA = falseVal
End If
End Function




