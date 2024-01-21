Public nLayers As Integer
Public Activation As Variant
Public Alpha As Double
Public Epoch As Integer

Sub Main()
    nLayers = 2
    Activation = Array(1, 2)
    Alpha = 0.002
    Epochs = 160

    For i = 1 To Epochs
        Iterate(i)
    Next i

    Call Predict
End Sub

Sub Iterate(epoch as Integer)
    ' Load input data
    Dim nSamples, nFeatures As Integer
    nSamples = Worksheets("Training Data").UsedRange.Rows.Count - 1
    nFeatures = Worksheets("Training Data").UsedRange.Columns.Count - 1
    ReDim X(nFeatures, nSamples) As Double
    ' Transpose
    ' -1 sample since last index is the actual label
    For i = 1 To nFeatures - 1
        For j = 1 To nSamples
            X(i - 1, j - 1) = Worksheets("Training Data").Cells(j + 1, i).Value
        Next j
    Next i

    ' Forward Prop
    ' Didn't manage to wrap in a function
    Dim A1 As Variant
    A1 = ForwardStep(X, 1)

    Dim A2 As Variant
    A2 = ForwardStep(A1, 2)

    ' Backprop

    Dim labelsRange As Range
    Set labelsRange = Range(Cells(2, nFeatures + 1), Cells(nSamples + 1, nFeatures + 1))
    Dim Uniques As Variant
    Uniques = Worksheets("Training Data").Evaluate("Unique(" & labelsRange.Address & ")")
    Dim nLabels As Integer
    nLabels = 0
    For Each Unique In Uniques
        nLabels = nLabels + 1
    Next Unique

    ' One-hot encode labels
    Dim label As Integer
    ReDim Y(nSamples, nLabels) As Double
    For i = 1 To nSamples
        label = Worksheets("Training Data").Cells(i + 1, nFeatures + 1).Value
        For j = 0 To nLabels - 1
            If j = label Then
                Y(i - 1, j) = 1
            Else
                Y(i - 1, j) = 0
            End If
        Next j
    Next i
    
    Dim YTranspose As Variant
    YTranspose = TransposeMatrix(Y)

    ' Get loss
    Dim loss as Double
    Dim MSE As Double
    MSE = 0
    ReDim DZ2(nLabels, nSamples) As Double
    For i = 0 To nLabels - 1
        For j = 0 To nSamples - 1
            loss = A2(i, j) - YTranspose(i, j)
            DZ2(i, j) = loss
            MSE = MSE + loss * loss
        Next j
    Next i
    MSE = MSE / nSamples
    If epoch mod 10 = 0 Then
        Worksheets("Main").Cells(7 + epoch / 10, 2).Value = "loss_" & epoch
        Worksheets("Main").Cells(7 + epoch / 10, 3).Value = MSE
    End If

    Dim DW2 As Variant
    DW2 = DotProduct(DZ2, TransposeMatrix(A1))
    Dim nRows, nCols As Integer
    nRows = UBound(DW2, 1)
    nCols = UBound(DW2, 2)
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            DW2(i, j) = DW2(i, j) / nSamples
        Next j
    Next i

    ' TODO: DB2

    Dim W2 As Variant
    W2 = LoadMatrix("Layer_2")

    Dim DZ1 As Variant
    DZ1 = DotProduct(TransposeMatrix(W2), DZ2)
    nRows = UBound(DZ1, 1)
    nCols = UBound(DZ1, 2)
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            If DZ1(i, j) < 0 Then
                DZ1(i, j) = 0
            Else
                DZ1(i, j) = 1
            End If
        Next j
    Next i

    Dim DW1 As Variant
    DW1 = DotProduct(DZ1, TransposeMatrix(X))
    nRows = UBound(DW1, 1)
    nCols = UBound(DW1, 2)
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            DW1(i, j) = DW1(i, j) / nSamples
        Next j
    Next i

    '  TODO: DB1

    ' Gradient descent
    ' Update W1
    nRows = Worksheets("Layer_1").UsedRange.Rows.Count
    nCols = Worksheets("Layer_1").UsedRange.Columns.Count
    For i = 1 To nRows
        For j = 1 To nCols
            Worksheets("Layer_1").Cells(i, j).Value = Worksheets("Layer_1").Cells(i, j).Value - Alpha * DW1(j - 1, i - 1)
        Next j
    Next i

    ' Update W2
    nRows = Worksheets("Layer_2").UsedRange.Rows.Count
    nCols = Worksheets("Layer_2").UsedRange.Columns.Count
    For i = 1 To nRows
        For j = 1 To nCols
            Worksheets("Layer_2").Cells(i, j).Value = Worksheets("Layer_2").Cells(i, j).Value - Alpha * DW2(j - 1, i - 1)
        Next j
    Next i

End Sub

Function LoadMatrix(sheetName As String) As Double()
    Dim nRow, nCol As Integer
    nRows = Worksheets(sheetName).UsedRange.Rows.Count
    nCols = Worksheets(sheetName).UsedRange.Columns.Count

    ReDim matrix(nRows, nCols) As Double
    For i = 1 To nRows
        For j = 1 To nCols
            matrix(i - 1, j - 1) = Worksheets(sheetName).Cells(i, j).Value
        Next j
    Next i

    LoadMatrix = matrix
End Function

Function CopyMatrix(matrix As Variant) As Variant
    Dim nRows, nCols As Integer
    nRows = UBound(matrix, 1)
    nCols = UBound(matrix, 2)

    ReDim Copied(nRows, nCols) As Double
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            Copied(i, j) = matrix(i, j)
        Next j
    Next i

    CopyMatrix = Copied
End Function

Function TransposeMatrix(matrix As Variant) As Variant
    Dim nRows, nCols As Integer
    nRows = UBound(matrix, 1)
    nCols = UBound(matrix, 2)

    ReDim Transposed(nCols, nRows) As Double
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            Transposed(j, i) = matrix(i, j)
        Next j
    Next i

    TransposeMatrix = Transposed
End Function

Function ForwardStep(inputMatrix As Variant, sheetIndex As Integer) As Double()
    Dim sheetName As String
    sheetName = "Layer_" + CStr(sheetIndex)

    Dim nRows, nCols, nColsInput As Integer
    nRows = Worksheets(sheetName).UsedRange.Rows.Count
    nCols = Worksheets(sheetName).UsedRange.Columns.Count

    ReDim weightMatrix(nRows, nCols) As Double
    weightMatrix = LoadMatrix(sheetName)

    nColsInput = UBound(inputMatrix, 2)

    Dim Z As Variant
    Z = DotProduct(TransposeMatrix(weightMatrix), inputMatrix)

    ReDim a(nCols, nColsInput) As Double
    If Activation(sheetIndex - 1) = 1 Then
        ' ReLU
        For i = 0 To nCols - 1
            For j = 0 To nColsInput - 1
                If Z(i, j) < 0 Then
                    a(i, j) = 0
                Else
                    a(i, j) = Z(i, j)
                End If
            Next j
        Next i
    ElseIf Activation(sheetIndex - 1) = 2 Then
        ' Softmax
        For i = 0 To nCols - 1
            Dim expSum As Double
            expSum = 0

            For j = 0 To nColsInput - 1
                expSum = expSum + Exp(Z(i, j))
            Next j

            For j = 0 To nColsInput - 1
                a(i, j) = Exp(Z(i, j)) / expSum
            Next j
        Next i
    End If

    ' Save for Z sheet
    sheetName = "Z_" + CStr(sheetIndex)
    For i = 1 To nCols
        For j = 1 To nColsInput
            Worksheets(sheetName).Cells(i, j).Value = Z(i - 1, j - 1)
        Next j
    Next i

    ' Save for A sheet
    sheetName = "A_" + CStr(sheetIndex)
    For i = 1 To nCols
        For j = 1 To nColsInput
            Worksheets(sheetName).Cells(i, j).Value = a(i - 1, j - 1)
        Next j
    Next i

    ForwardStep = a
End Function

Function DotProduct(matrix1 As Variant, matrix2 As Variant) As Double()
    Dim nMatrix1Rows, nMatrix1Cols, nMatrix2Rows, nMatrix2Cols As Integer
    nMatrix1Rows = UBound(matrix1, 1)
    nMatrix1Cols = UBound(matrix1, 2)
    nMatrix2Rows = UBound(matrix2, 1)
    nMatrix2Cols = UBound(matrix2, 2)

    ReDim matrix3(nMatrix1Rows, nMatrix2Cols) As Double
    For i = 0 To nMatrix1Rows - 1
        For j = 0 To nMatrix2Cols - 1
            For k = 0 To nMatrix2Rows - 1
                matrix3(i, j) = matrix3(i, j) + matrix1(i, k) * matrix2(k, j)
            Next k
        Next j
    Next i

    DotProduct = matrix3
End Function

Sub Predict()
    ' Load input data
    Dim Splits as Variant
    Splits = Array("Training", "Test")

    For Each splitPrefix in Splits
        Dim nSamples, nFeatures As Integer
        nSamples = Worksheets(splitprefix & " Data").UsedRange.Rows.Count - 1
        nFeatures = Worksheets(splitprefix & " Data").UsedRange.Columns.Count - 1
        ReDim X(nFeatures, nSamples) As Double
        ' Transpose
        ' -1 sample since last index is the actual label
        For i = 1 To nFeatures - 1
            For j = 1 To nSamples
                X(i - 1, j - 1) = Worksheets(splitprefix & " Data").Cells(j + 1, i).Value
            Next j
        Next i

        ' Forward Prop
        ' Didn't manage to wrap in a function
        Dim A1 As Variant
        A1 = ForwardStep(X, 1)

        Dim A2 As Variant
        A2 = ForwardStep(A1, 2)

        ' Get maximum confidence value per sample
        Dim nLabels, predictedIndex As Integer
        nLabels = UBound(A2, 1)

        ReDim Predictions(nSamples) As Integer
        For i = 0 to nSamples - 1
            predictedIndex = 0
            For j = 1 to nLabels
                If A2(j, i) > A2(predictedIndex, i) Then
                    predictedIndex = j
                End If
            Next j
            Predictions(i) = predictedIndex
        Next i

        Dim splitPredictions As String
        splitPredictions = splitprefix & " Predictions"
        nCol = Worksheets(splitPredictions).UsedRange.Columns.Count
        Worksheets(splitPredictions).Cells(1, nCol + 1).Value = "predictions"
        For i = 1 To nSamples
            Worksheets(splitPredictions).Cells(i + 1, nCol + 1).Value = Predictions(i - 1)
        Next i

        Call ConfusionMatrix(splitPredictions)

    Next splitPrefix
End Sub