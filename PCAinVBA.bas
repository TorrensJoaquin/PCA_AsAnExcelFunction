Function PerformPCA(AR As Range, Optional LowDimension = 2, Optional Iterations = 1)
    Application.Volatile
    A = reindexRange(AR)
    ''Create a random matrix
    Dim mOriginal As Long, mNew As Long, n As Long, p As Long, i As Long, j As Long, k As Long, iter As Long
    Dim C As Variant
    Dim Q As Variant
    Dim E As Variant
    Dim LowerDimensionalApproximation As Long
    ''Decide the size of the lower dimensional approximation
    ''The minimum size of the variance matrix will be 10x10 and
    ''cannot be smaller than the requested size plus two.
    mOriginal = UBound(A, 2)
    If mOriginal > 10 Then
        If mOriginal - LowDimension > 2 Then
            LowerDimensionalApproximation = 10
        Else
            LowerDimensionalApproximation = LowDimension + 2
        End If
        If mOriginal > LowerDimensionalApproximation + 1 Then
            Q = CreateARandomMatrix(mOriginal, LowerDimensionalApproximation)
            A = fmMult(A, Q)
            Erase Q
        End If
    End If
    ''Data Normalization
    mNew = UBound(A, 2)
    p = UBound(A, 1)
    Dim Average As Double
    Dim StandardDeviation As Double
    For j = 0 To mNew
        Average = 0
        For i = 0 To p
            Average = Average + A(i, j)
        Next i
        Average = Average / (p + 1)
        StandardDeviation = 0
        For i = 0 To p
            A(i, j) = A(i, j) - Average
            StandardDeviation = StandardDeviation + A(i, j) ^ 2
        Next i
        StandardDeviation = Sqr(StandardDeviation / p)
        For i = 0 To p
            A(i, j) = A(i, j) / StandardDeviation
        Next i
    Next j
    ''Covariance Matrix
    ReDim C(0 To mNew, 0 To mNew)
    For i = 0 To mNew
        For j = 0 To mNew
            For k = 0 To p
                C(i, j) = C(i, j) + A(k, i) * A(k, j)
            Next k
            C(i, j) = C(i, j) / p
        Next j
    Next i
    ''PCA
    Dim mRows As Long, nCols As Long
    R = createIdentity(mNew, mNew)
    For iter = 0 To Iterations
        C = fmMult(R, C) ' A=R*Q
        For j = 0 To mNew
            R(j, j) = 0
            For i = 0 To mNew
                R(j, j) = R(j, j) + C(i, j) ^ 2
            Next
            R(j, j) = Sqr(R(j, j))
            For i = 0 To mNew
                C(i, j) = C(i, j) / R(j, j)
            Next
            For k = (j + 1) To mNew
                R(j, k) = 0
                For i = 0 To mNew
                    R(j, k) = R(j, k) + C(i, j) * C(i, k)
                Next
                For i = 0 To mNew
                    C(i, k) = C(i, k) - C(i, j) * R(j, k)
                Next
            Next
        Next
    Next iter
    Erase R
    E = createIdentity(mNew, mNew)
    E = fmMult(E, C) ' E: Eigenvectors
    ''Perform de Dimensionality Reduction
    ''A = reindexRange(AR)
    ReDim C(0 To p, 0 To LowDimension - 1)
    For i = 0 To p
        For j = 0 To mNew
            For k = 0 To LowDimension - 1
                C(i, k) = C(i, k) + A(i, j) * E(j, k)
            Next k
        Next j
    Next i
    PerformPCA = C
End Function
Private Function fmMult(A As Variant, B As Variant) As Variant
    'Assumes that A,B are 1-based variant arrays

    Dim m As Long, n As Long, p As Long, i As Long, j As Long, k As Long
    Dim C As Variant

    m = UBound(A, 1)
    p = UBound(A, 2)
    If UBound(B, 1) <> p Then
        Debug.Print "Not Defined!"
        Exit Function
    End If
    n = UBound(B, 2)

    ReDim C(0 To m, 0 To n)
    For i = 0 To m
        For j = 0 To n
            For k = 0 To p
                C(i, j) = C(i, j) + A(i, k) * B(k, j)
            Next k
        Next j
    Next i
    fmMult = C
End Function
Private Function createIdentity(mRows, nCols) As Variant
    Dim A As Variant
    ReDim A(mRows, nCols)
    Dim i As Long
    Dim j As Long
    For i = 0 To mRows
        For j = 0 To nCols
            If i = j Then
                A(i, j) = 1
            Else
                A(i, j) = 0
            End If
        Next j
    Next i
    createIdentity = A
End Function
Private Function reindexRange(A As Range) As Variant
    B = A.Value2
    Dim C As Variant
    ReDim C(UBound(B, 1) - 1, UBound(B, 2) - 1)
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(B, 1) - 1
        For j = 0 To UBound(B, 2) - 1
            C(i, j) = B(i + 1, j + 1)
        Next
    Next
    reindexRange = C
End Function
Function CreateARandomMatrix(DimensionA As Long, DimensionB As Long)
    Dim FinalArray As Variant
    Dim i As Long
    Dim j As Long
    ReDim FinalArray(0 To DimensionA, 0 To DimensionB)
    For i = 0 To DimensionA
        For j = 0 To DimensionB
            FinalArray(i, j) = WorksheetFunction.NormInv(Rnd(), 0, 1)
        Next
    Next
    CreateARandomMatrix = FinalArray
End Function
