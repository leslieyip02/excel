Sub confusionMatrix()
    Worksheets("Training Predictions").Activate
    Dim nRow, nCol, predictionColumn, actualColumn As Integer
    nRow = ActiveSheet.UsedRange.Rows.Count
    nCol = ActiveSheet.UsedRange.Columns.Count

    predictionColumn = nCol
    actualColumn = nCol - 1

    Dim actualRange As Range
    Set actualRange = Range(Cells(2, actualColumn - 1), Cells(nRow, actualColumn - 1))
    Dim uniques As Variant
    uniques = ActiveSheet.Evaluate("Unique(" & actualRange.Address & ")")

    Dim row, col As Integer
    row = 2
    col = nCol + 2
    
    Dim i As Integer
    i = 0
    For Each Unique In uniques
        i = i + 1
        Cells(row + i, col).Value = Unique
        Cells(row, col + i).Value = Unique
    Next Unique

    Dim rng As Range
    Set rng = Range(Cells(row, col), Cells(row + i, col + i))
    rng.Borders.LineStyle = xlContinuous
    For j = 1 To i
        For k = 1 To i
            Cells(row + j, col + k).Value = 0
        Next k
    Next j

    Dim predicted, actual As Integer
    For j = 2 To nRow
        predicted = Cells(j, predictionColumn).Value
        actual = Cells(j, actualColumn).Value
    Next j
End Sub
