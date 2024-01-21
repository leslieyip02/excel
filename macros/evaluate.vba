Sub ConfusionMatrix(sheetName As String)
    Worksheets(sheetName).Activate
    Dim nRow, nCol, predictionColumn, actualColumn As Integer
    nRow = Worksheets(sheetName).UsedRange.Rows.Count
    nCol = Worksheets(sheetName).UsedRange.Columns.Count

    predictionColumn = nCol
    actualColumn = nCol - 1

    Dim actualRange As Range
    Set actualRange = Range(Cells(2, actualColumn), Cells(nRow, actualColumn))
    Dim uniques As Variant
    uniques = ActiveSheet.Evaluate("Unique(" & actualRange.Address & ")")

    Dim row, col As Integer
    row = 2
    col = nCol + 2
    
    Dim i As Integer
    i = 0
    For Each Unique In uniques
        Cells(row + i + 1, col).Value = "predict_" & i
        Cells(row, col + i + 1).Value = "actual_" & i
        i = i + 1
    Next Unique

    Dim rng As Range
    Set rng = Range(Cells(row, col), Cells(row + i, col + i))
    rng.Borders.LineStyle = xlContinuous
    For j = 1 To i
        For k = 1 To i
            Cells(row + j, col + k).Value = 0
        Next k
    Next j

    Dim predicted, actual, a As Integer
    For j = 2 To nRow
        predicted = Cells(j, predictionColumn).Value
        actual = Cells(j, actualColumn).Value
        Cells(row + predicted + 1, col + actual + 1).Value = Cells(row + predicted + 1, col + actual + 1).Value + 1
    Next j

    Dim correct, tp, fp, tn, fn As Integer
    Dim accuracy, precision, recall, f1 As Double
    correct = 0
    For a = 1 To i
        correct = correct + Cells(row + a, col + a).Value
    Next a
    Dim currentRow As Integer
    currentRow = row + i + 2
    Cells(currentRow, col).Value = "accuracy"
    Cells(currentRow, col + 1).Value = correct / (nRow - 1)
    currentRow = currentRow + 1

    For a = 1 To i
        tp = Cells(row + a, col + a).Value
        Cells(row + a, col + a).Interior.Color = vbGreen
        fp = 0
        fn = 0
        For b = 1 To i
            If a <> b Then
                fp = fp + Cells(row + a, col + b).Value
                fn = fp + Cells(row + b, col + a).Value
            End If
        Next b

        If (tp + fp = 0) Then
            precision = 0
        Else
            precision = tp / (tp + fp)
        End If
        Cells(currentRow, col).Value = "precision_" & (a - 1)
        Cells(currentRow, col + 1).Value = precision
        currentRow = currentRow + 1

        If (tp + fn = 0) Then
            recall = 0
        Else
            recall = tp / (tp + fn)
        End If
        Cells(currentRow, col).Value = "recall_" & (a - 1)
        Cells(currentRow, col + 1).Value = recall
        currentRow = currentRow + 1

        If (precision + recall) = 0 Then
            f1 = 0
        Else
            f1 = 2 * precision * recall / (precision + recall)
        End If
        Cells(currentRow, col).Value = "f1_" & (a - 1)
        Cells(currentRow, col + 1).Value = f1
        currentRow = currentRow + 1
    Next a
End Sub


