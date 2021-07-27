Function PerformPCA(AR As Range, LowDimension)
    ''Data Normalization
    A = reindexRange(AR)
    Dim m As Long, n As Long, p As Long, i As Long, j As Long, k As Long, iter As Long
    Dim C As Variant
    m = UBound(A, 2)
    p = UBound(A, 1)
    Dim Average As Double
    Dim StandardDeviation As Double
    For j = 0 To m
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
    ReDim C(0 To m, 0 To m)
    For i = 0 To m
        For j = 0 To m
            For k = 0 To p
                C(i, j) = C(i, j) + A(k, i) * A(k, j)
            Next k
        Next j
    Next i
    For i = 0 To m
        For j = 0 To m
            C(i, j) = C(i, j) / p
        Next j
    Next i
    ''PCA
    Dim mRows As Long, nCols As Long
    R = createIdentity(m + 1, m + 1)
    E = createIdentity(m + 1, m + 1)
    For iter = 0 To 1000
        C = fmMult(R, C) ' A=R*Q
        For j = 0 To m
            R(j, j) = 0
            For i = 0 To m
                R(j, j) = R(j, j) + C(i, j) ^ 2
            Next
            R(j, j) = Sqr(R(j, j))
            For i = 0 To m
                C(i, j) = C(i, j) / R(j, j)
            Next
            For k = (j + 1) To m
                R(j, k) = 0
                For i = 0 To m
                    R(j, k) = R(j, k) + C(i, j) * C(i, k)
                Next
                For i = 0 To m
                    C(i, k) = C(i, k) - C(i, j) * R(j, k)
                Next
            Next
        Next
        E = fmMult(E, C) ' E: Eigenvectors
    Next iter
    ''Perform de Dimensionality Reduction
    ReDim C(0 To p, 0 To LowDimension - 1)
    For i = 0 To p
        For j = 0 To m
            For k = 0 To LowDimension - 1
                C(i, k) = C(i, k) + A(i, j) * E(j, k)
            Next k
        Next j
    Next i
    PerformPCA = C
End Function
Private Function fmMult(A As Variant, B As Variant) As Variant
    Dim m As Long, n As Long, p As Long, i As Long, j As Long, k As Long
    Dim C As Variant

    If TypeName(A) = "Range" Then A = A.Value
    If TypeName(B) = "Range" Then B = B.Value

    m = UBound(A, 1)
    p = UBound(A, 2)
    If UBound(B, 1) <> p Then
        MatrixProduct = "Not Defined!"
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
    ReDim A(mRows - 1, nCols - 1)
    Dim i As Long
    Dim j As Long
    For i = 0 To mRows - 1
        For j = 0 To nCols - 1
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
