Sub FourColorPaintedRect()

Const sideCount As Integer = 10
Dim startRowIndex As Integer, startColIndex As Integer
Dim yellowColor As Long, greenColor As Long, blueColor As Long, redColor As Long
Dim currentBigCellColor As Long

startRowIndex = ActiveCell.Row
startColIndex = ActiveCell.Column

yellowColor = RGB(255, 255, 0)
greenColor = RGB(0, 255, 0)
blueColor = RGB(0, 0, 255)
redColor = RGB(255, 0, 0)

currentCellColorForOddRowIndex = yellowColor
currentCellColorForEvenRowIndex = blueColor

For currentRowIndex = startRowIndex To startRowIndex + sideCount - 1
    For currentColIndex = startColIndex To startColIndex + sideCount - 1
        If currentRowIndex Mod 2 = 1 Then
            Cells(currentRowIndex, currentColIndex).Interior.Color = currentCellColorForOddRowIndex
            
            If currentCellColorForOddRowIndex = yellowColor Then
                currentCellColorForOddRowIndex = greenColor
            Else
                currentCellColorForOddRowIndex = yellowColor
            End If
        Else
            Cells(currentRowIndex, currentColIndex).Interior.Color = currentCellColorForEvenRowIndex
            
            If currentCellColorForEvenRowIndex = blueColor Then
                currentCellColorForEvenRowIndex = redColor
            Else
                currentCellColorForEvenRowIndex = blueColor
            End If
        End If
    Next
Next

End Sub