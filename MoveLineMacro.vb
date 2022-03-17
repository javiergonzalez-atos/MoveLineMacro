Sub MoveLineToRight()
    DrawTodayLine 1
End Sub

Sub MoveLineToLeft()
    DrawTodayLine -1
End Sub

Sub DrawTodayLine(lineDirection As Integer)

    Dim startCell As Range
    Dim lastCell As Range
    Dim actualLinePosition As Range
    Dim newLinePosition As Range
    Dim borderColorIndex As Integer
    
    Set startCell = Range("$A$3")
    Set lastCell = Cells(startCell.End(xlDown).Row - 2, startCell.End(xlToRight).Column)

    borderColorIndex = 44 'Orange
    
    Set actualLinePosition = GetColorLinePosition(startCell, lastCell, borderColorIndex)
    
    EraseLineBorders actualLinePosition
    
    Set newLinePosition = actualLinePosition.Offset(0, lineDirection)
    
    DrawColorBorders newLinePosition, borderColorIndex

End Sub

Sub DrawColorBorders(lineRange As Range, colorIndex As Integer)
 
    With lineRange.Borders
        .LineStyle = xlContinuous
        .colorIndex = colorIndex
        .TintAndShade = 0
        .Weight = xlThick
    End With

End Sub

Sub EraseLineBorders(lineRange As Range)

    With lineRange.Borders
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub

Function GetColorLinePosition(currentCell As Range, lastCell As Range, borderColorIndex As Integer) As Range
    
    For i = currentCell.Column To lastCell.Column
        
        If currentCell.Borders.colorIndex = borderColorIndex Then
            Set GetColorLinePosition = Range(currentCell, Cells(lastCell.Row, currentCell.Column))
            Exit For
        End If
        Set currentCell = currentCell.Offset(0, 1)
    Next
    

End Function