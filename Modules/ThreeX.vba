Sub ThreeX()

Const sideCountForBigCell As Integer = 3
Const sideCountForSquareRect As Integer = 9

Dim startRowIndex As Integer, startColIndex As Integer
Dim topLeftRowIndex As Integer, topLeftColIndex As Integer
Dim bottomRightRowIndex As Integer, bottomRightColIndex As Integer
Dim yellowColor As Long, blueColor As Long
Dim currentBigCellColor As Long

startRowIndex = ActiveCell.Row
startColIndex = ActiveCell.Column

yellowColor = RGB(255, 255, 0)
blueColor = RGB(0, 0, 255)
currentBigCellColor = yellowColor

For currentRowIndex = startRowIndex To startRowIndex + sideCountForSquareRect - 1 Step sideCountForBigCell
    For currentColIndex = startColIndex To startColIndex + sideCountForSquareRect - 1 Step sideCountForBigCell
        topLeftRowIndex = currentRowIndex
        topLeftColIndex = currentColIndex
        bottomRightRowIndex = currentRowIndex + sideCountForBigCell - 1
        bottomRightColIndex = currentColIndex + sideCountForBigCell - 1
        
        Range(Cells(topLeftRowIndex, topLeftColIndex), Cells(bottomRightRowIndex, bottomRightColIndex)).Interior.Color = currentBigCellColor
        
        If currentBigCellColor = yellowColor Then
            currentBigCellColor = blueColor
        Else
            currentBigCellColor = yellowColor
        End If
    Next
Next
        
End Sub