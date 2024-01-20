Function meanSquaredError(predictions As Variant) As Variant
    ' Takes a numSamples by 1 2d matrix
    ' Returns a numSamples by numLabels one hot encoded matrix
    Worksheets("Training Data").Activate
    Dim nRow, nCol, predictionColumn, actualColumn As Integer
    nRow = ActiveSheet.UsedRange.Rows.Count
    nCol = ActiveSheet.UsedRange.Columns.Count

    Dim actualRange As Range
    Set actualRange = Range(Cells(2, nCol), Cells(nRow, nCol))
    Dim uniques As Variant
    uniques = ActiveSheet.Evaluate("Unique(" & actualRange.Address & ")")
    Dim nUnique As Integer
    nUnique = 0
    For Each Unique in uniques
        nUnique = nUnique + 1
    Next Unique

    Dim oneHotEncoded(nRow - 1, nUnique) as Integer
    Dim predicted as Integer
    For i = 1 to nRow
        predicted = predictions(i, 0)
        For j = 0 to nUnique - 1
            If j = predicted Then
                oneHotEncoded(i, j) = 1
            Else
                oneHotEncoded(i, j) = 0
            End If
        Next j
    Next i

    meanSquaredError = oneHotEncoded
End Function
