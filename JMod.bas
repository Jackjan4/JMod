Attribute VB_Name = "JMod"
' Module by Jan-Philipp Roslan
' Contains utility function for better code maintenance and to reduce duplicated code
' Version: 0.4
' Date: 16.8.16
' Time: 15:52




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

    If range.value = "" Or IsEmpty(range) Then
        CellIsEmpty = True
    Else
        CellIsEmpty = False
    End If

End Function


' Value of cell
Function Val(ByVal cell As Variant)

    Select Case TypeName(cell)

        Case "string", "String":
            Val = range(cell).value

        Case "Range", "range":
            Val = cell.value
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





