Sub PaintedCellRect()

Const sideCount As Integer = 10
Dim startRowIndex As Integer, startColIndex As Integer
Dim currentCellColor As Long

startRowIndex = ActiveCell.Row
startColIndex = ActiveCell.Column

For currentRowIndex = startRowIndex To startRowIndex + sideCount - 1
    For currentColIndex = startColIndex To startColIndex + sideCount - 1
        If (currentRowIndex + currentColIndex) Mod 2 = 0 Then
            Cells(currentRowIndex, currentColIndex).Interior.ColorIndex = 6
        Else
            Cells(currentRowIndex, currentColIndex).Interior.ColorIndex = 4
        End If
    Next
Next

End Sub