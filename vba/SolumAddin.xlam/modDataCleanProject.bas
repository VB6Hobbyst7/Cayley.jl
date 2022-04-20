Attribute VB_Name = "modDataCleanProject"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sConditionalZScores
' Author     : Philip Swannell
' Date       : 22-Jul-2019
' Purpose    : Wraps sConditionalMultiVariate to analyse many rows of data in "ObservedReturns" for each row
'              we calculate the conditional ZScore for each element, conditional on the other elements in the row. i.e.
'              i.e. by how many "conditional standard deviations" does the observation differ from the conditional mean?
' Parameters :
'  Mu             : The vector mean of the returns distribution. Pass as 1-d array or 1-column or 1-row 2-d array (or Range object)
'  Sigma          : The covariance matrix of the returns
'  ObservedReturns: Matrix, each row consitutes one day's observed returns, so the calculation of time series -> returns must be done
'                   before this function is called
' -----------------------------------------------------------------------------------------------------------------------
Function sConditionalZScores(ByVal mu, sigma, ObservedReturns, Optional ReturnFullDetails As Boolean)
Attribute sConditionalZScores.VB_Description = "Returns a vector of conditional z-scores for a vector observation of a multivariate normal Y~N(Mu,Sigma)."
Attribute sConditionalZScores.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim a() As Double
          Dim ConditionalExpectations() As Double
          Dim ConditionalSDs()
          Dim ConditionalZScores() As Double
          Dim Headers
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim Mu1 As Double
          Dim Mu2() As Double
          Dim MuBar
          Dim NC As Long
          Dim NR As Long
          Dim ReturnArray
          Dim Sigma12Sigma22Inv
          Dim SigmaBar
          Dim ZScores() As Double

1         On Error GoTo ErrHandler
          'Coerce Mu, Sigma, ObservedReturns to be 2-d arrays (as opposed to Range objects)
2         Force2DArrayRMulti mu, sigma
3         Force2DArrayR ObservedReturns, NR, NC

          'Input Validation
4         CheckCovMatrix sigma, "Sigma", True
5         CheckNumericMatrix ObservedReturns, "ObservedReturns"
6         If sNRows(mu) = 1 Then mu = sArrayTranspose(mu) 'For convenience flip to column vector
7         If sNRows(mu) <> NC Then Throw "Mu must be a one-row or one column array with the same number of elements as there are columns in ObservedReturns"
8         CheckNumericMatrix mu, "Mu"

9         ReDim ZScores(1 To NR, 1 To NC)
10        ReDim ConditionalZScores(1 To NR, 1 To NC)
11        ReDim a(1 To NC - 1, 1 To 1)
12        ReDim Mu2(1 To NC - 1, 1 To 1)
13        ReDim ConditionalExpectations(1 To NR, 1 To NC)
14        ReDim ConditionalSDs(1 To 1, 1 To NC)

15        For j = 1 To NC

16            For i = 1 To NR
                  'Partition Mu into Mu1 (1 element) and Mu2 (all other elements). Also set A, the value of the "conditioned upon" sub-vector.
17                For k = 1 To NC
18                    If k < j Then
19                        Mu2(k, 1) = mu(k, 1)
20                        a(k, 1) = ObservedReturns(i, k)
21                    ElseIf k = j Then
22                        Mu1 = mu(k, 1)
23                    ElseIf k > j Then
24                        Mu2(k - 1, 1) = mu(k, 1)
25                        a(k - 1, 1) = ObservedReturns(i, k)
26                    End If
27                Next k
28                If i = 1 Then
                      'Call sConditionalMultiVariate to get SigmaBar and Sigma12Sigma22Inv for the current j
                      'We ignore the returned-by-reference MuBar since it's only valid for i = 1. Instead use SigmaBar and Sigma12Sigma22Inv which are independent of i
29                    sConditionalMultiVariate mu, sigma, j, a, MuBar, SigmaBar, Sigma12Sigma22Inv
30                End If
                  
                  'Calculate MuBar - this is the formula given at _
                   https://stats.stackexchange.com/questions/30588/deriving-the-conditional-distributions-of-a-multivariate-normal-distribution
31                MuBar = Mu1
32                For k = 1 To NC - 1
33                    MuBar = MuBar + Sigma12Sigma22Inv(k) * (a(k, 1) - Mu2(k, 1))
34                Next k

35                ConditionalExpectations(i, j) = MuBar
36                If i = 1 Then
37                    ConditionalSDs(1, j) = Sqr(SigmaBar(1, 1))
38                End If
39                ConditionalZScores(i, j) = (ObservedReturns(i, j) - ConditionalExpectations(i, j)) / ConditionalSDs(1, j)
40                ZScores(i, j) = (ObservedReturns(i, j) - mu(j, 1)) / Sqr(sigma(j, j))
41            Next i
42        Next j

          'Construct a "full details" array, for use on "CurveViewer" sheet of demo workbook DataCleaning.xlsm
43        If ReturnFullDetails Then
44            ReturnArray = sArrayTranspose(sArrayStack( _
                  ConditionalExpectations, _
                  ConditionalSDs, _
                  ConditionalZScores, _
                  sArrayTranspose(mu), _
                  sArrayPower(sArrayTranspose(sDiagonal(sigma)), 0.5), _
                  ZScores))

45            Headers = sReshape("", 1, sNCols(ReturnArray))
46            For i = 1 To NR
47                Headers(1, i) = "Cond Mean" & IIf(NR = 1, "", " " & CStr(i))
48                Headers(1, NR + 1 + i) = "Cond ZScore" & IIf(NR = 1, "", " " & CStr(i))
49                Headers(1, 2 * NR + 3 + i) = "ZScore" & IIf(NR = 1, "", " " & CStr(i))
50            Next i
51            Headers(1, NR + 1) = "Cond SD"
52            Headers(1, 2 * NR + 2) = "Mean"
53            Headers(1, 2 * NR + 3) = "SD"
54            ReturnArray = sArrayStack(Headers, ReturnArray)
55            sConditionalZScores = ReturnArray
56        Else
57            sConditionalZScores = ConditionalZScores
58        End If

59        Exit Function
ErrHandler:
60        sConditionalZScores = "#sConditionalZScores (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'https://stats.stackexchange.com/questions/30588/deriving-the-conditional-distributions-of-a-multivariate-normal-distribution
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sConditionalMultiVariate
' Author     : Philip Swannell
' Date       : 19-Jul-2019
' Purpose    : We have a multivariate normal distribution Y with (vector) mean Mu and covariance matrix Sigma.
'              Consider partitioning Mu and Y into
'              Mu = [Mu1,Mu2]
'               Y = [Y1,Y2]
'with a similar partition of Sigma into
'              Sigma =  [Sigma11, Sigma12]
'                       [Sigma21, Sigma22]
'The function returns the mean and variance of the conditional distribution of Y2 given that Y1 = A
'See
'
' Arguments
' Mu        : The vector mean Mu of Y. Enter as a 1-column array.
' Sigma     : The covariance matrix Sigma of Y.
' Y1Indices : Column vector of indices of the elements of Y that are in Y1. If Y1Indices is {2;3} then
'             Y1 is elements 2 and 3 of Y and Y2 is the remaining elements.
' A         : A column vector of values. The function returns statistics of Y1 conditional on Y2=A.
' retMuBar  : For use from VBA only. By reference argument set to the mean of the conditional
'             distribution.
' retSigmaBar: For use from VBA only. By reference argument set to the covariance of the conditional
'             distribution.
' retSigma12Sigma22Inv: For use from VBA only. By reference argument set to the matrix product of Sigma12 and the
'             inverse of Sigma22.
'
' Notes     : See
'             https://stats.stackexchange.com/questions/30588/deriving-the-conditional-distributions-of-a-multivariate-normal-distribution
' -----------------------------------------------------------------------------------------------------------------------
Function sConditionalMultiVariate(ByVal mu As Variant, ByVal sigma As Variant, ByVal Y1Indices, ByVal a, Optional ByRef retMuBar, Optional ByRef retSigmaBar, Optional ByRef retSigma12Sigma22Inv As Variant)
Attribute sConditionalMultiVariate.VB_Description = "Given a multi-variate normal Y ~ N(Mu,Sigma),  returns mean (MuBar) and variance (SigmaBar) of a sub-vector (Y1) of Y conditional on Y2 (the complement of Y1) being equal to A. The first column of the return is MuBar and remaining columns are SigmaBar."
Attribute sConditionalMultiVariate.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim Mu1
          Dim Mu2
          Dim MuBar
          Dim N As Long
          Dim N1 As Long
          Dim N2 As Long
          Dim Sigma11 As Variant
          Dim Sigma12 As Variant
          Dim Sigma21 As Variant
          Dim Sigma22 As Variant
          Dim Sigma22Inv
          Dim SigmaBar
          Dim Y2Indices

1         On Error GoTo ErrHandler

          'Convert inputs into 2-dimensional arrays
2         Force2DArrayRMulti mu, sigma, Y1Indices, a
3         N = sNRows(sigma)
          
          'Validation
4         CheckCovMatrix sigma, "Sigma"
5         If sNRows(mu) <> N Or sNCols(mu) <> 1 Then Throw "Mu must be 1-column array with the same number of rows as Sigma"
6         CheckNumericColumnVector mu, "Mu"
7         CheckIndices Y1Indices, N, "Y1Indices"

8         Y2Indices = ComplementOf(Y1Indices, N)
9         N1 = sNRows(Y1Indices)
10        N2 = sNRows(Y2Indices)

11        If sNRows(a) <> N2 Then Throw "The number of rows in A should be the number of rows in Mu less the number of rows in Y1Indicator i.e. " + CStr(N) + " - " + CStr(N1) + " = " + CStr(N2) + ", but it has " + CStr(sNRows(a)) + " row(s)"
12        If sNCols(a) <> 1 Then Throw "A should be a one-column array of numbers, but it has more than one column"
          
13        CheckNumericColumnVector a, "A"

          'Partition the covariance matrix into four. sIndex provides vectorized indexing.
14        Sigma11 = sIndex(sigma, Y1Indices, sArrayTranspose(Y1Indices))
15        Sigma12 = sIndex(sigma, Y1Indices, sArrayTranspose(Y2Indices))
16        Sigma21 = sIndex(sigma, Y2Indices, sArrayTranspose(Y1Indices))
17        Sigma22 = sIndex(sigma, Y2Indices, sArrayTranspose(Y2Indices))

          'Partition Mu into Mu1 and Mu2
18        Mu1 = sIndex(mu, Y1Indices)
19        Mu2 = sIndex(mu, Y2Indices)
20        Sigma22Inv = MInverse(Sigma22)

21        retSigma12Sigma22Inv = MMult(Sigma12, Sigma22Inv)
          'These are the two formulas given at https://stats.stackexchange.com/questions/30588/deriving-the-conditional-distributions-of-a-multivariate-normal-distribution
22        MuBar = sArrayAdd(Mu1, MMult(retSigma12Sigma22Inv, sArraySubtract(a, Mu2)))
23        SigmaBar = sArraySubtract(Sigma11, MMult(Sigma12, MMult(Sigma22Inv, Sigma21)))

          'Set by reference arguments
24        retMuBar = MuBar
25        retSigmaBar = SigmaBar

26        sConditionalMultiVariate = sArrayRange(MuBar, SigmaBar)

27        Exit Function
ErrHandler:
28        Throw "#sConditionalMultiVariate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CheckNumericColumnVector
' Author     : Philip Swannell
' Date       : 22-Jul-2019
' Purpose    : Validator for a column-matrix of numbers
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckNumericColumnVector(x As Variant, VariableName As String)
          Dim i As Long
1         On Error GoTo ErrHandler
2         If NumDimensions(x) <> 2 Then Throw VariableName + " must have 2 dimensions"
3         If sNCols(x) <> 1 Then Throw VariableName + " must have only one column"

4         For i = LBound(x, 1) To UBound(x, 1)
5             If Not IsNumber(x(i, 1)) Then Throw VariableName + " must be a 1-column array of numbers but it has a non-number at position " + CStr(i)
6         Next

7         Exit Function
ErrHandler:
8         Throw "#CheckNumericColumnVector (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CheckNumericMatrix
' Author     : Philip Swannell
' Date       : 22-Jul-2019
' Purpose    : Validator for argument being 2-dimensional array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckNumericMatrix(Data As Variant, VariableName As String)
          Dim i As Long
          Dim j As Long

1         On Error GoTo ErrHandler
2         For i = LBound(Data, 1) To UBound(Data, 1)
3             For j = LBound(Data, 2) To UBound(Data, 2)
4                 If Not IsNumber(Data(i, j)) Then Throw "Non-number found in " + VariableName + " at row " + CStr(i) + ", column " + CStr(j)
5             Next j
6         Next i

7         Exit Function
ErrHandler:
8         Throw "#CheckNumericMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CheckCovMatrix
' Author     : Philip Swannell
' Date       : 19-Jul-2019
' Purpose    : Simple validation on the Covariance matrix, checks matrix is square, all elements numeric and is symmetric.
'              Could check matrix is invertible, but currently don't do that.
' Parameters :
'  Sigma:
'  N        :  Number of rows in Sigma
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckCovMatrix(sigma As Variant, VariableName As String, Optional CheckInvertible As Boolean = False)
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim j As Long
          Dim N As Long
2         N = sNRows(sigma)
3         If sNCols(sigma) <> N Then Throw ("Sigma must be a square matrix")
4         For i = 1 To N
5             For j = 1 To i
6                 If Not (IsNumber(sigma(i, j))) Then
7                     Throw "Non number found at (" & CStr(i) & "," & CStr(j) & ") in matrix " & VariableName
8                 ElseIf Not (IsNumber(sigma(j, i))) Then
9                     Throw "Non number found at (" & CStr(j) & "," & CStr(i) & ") in matrix " & VariableName
10                ElseIf sigma(i, j) <> sigma(j, i) Then
11                    Throw VariableName & " must be symmetric but element (" & CStr(i) & "," & CStr(j) & ") is not equal to element (" & CStr(j) & "," & CStr(i) & ")"
12                End If
13            Next j
14        Next i

15        If CheckInvertible Then
              Dim Res
16            On Error Resume Next
17            Res = MInverse(sigma)
18            On Error GoTo ErrHandler
19            If IsEmpty(Res) Then
20                Throw VariableName & " is singular!"
21            End If
22        End If
23        Exit Function
ErrHandler:
24        Throw "#CheckCovMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CheckIndices
' Author     : Philip Swannell
' Date       : 19-Jul-2019
' Purpose    : Check that passed in variable Indices is a one-column, 2-dimensional array, whose elements are integers in
'              the range 1 to N and is strictly monotonic. Throws an error if not.
' Parameters :
'  Indices     :
'  N           :
'  VariableName: Name of the variable for constructing error description
' -----------------------------------------------------------------------------------------------------------------------
Private Function CheckIndices(Indices As Variant, N As Long, VariableName As String)
          
          Dim i As Long

1         On Error GoTo ErrHandler

2         If sNCols(Indices) <> 1 Then Throw VariableName & " must be a one-column array giving the indexes (in the range 1 to " + CStr(N) + ") of the elements of Y1"
3         If sNRows(Indices) >= N Then Throw VariableName & " must be a one-column array with fewer elements than there are rows in Sigma"
          'Because method ComplementOf assumes 1-based indexing
4         If LBound(Indices, 1) <> 1 Then Throw "Assertion failed: expected variable " & VariableName & " to be indexed from 1, but it is indexed from " + CStr(LBound(Indices, 1))

5         For i = 1 To sNRows(Indices)
6             If Not (IsNumber(Indices(i, 1))) Then
7                 Throw VariableName & " must be a column array of integers in the range 1 to " + CStr(N) + ") arranged in increasing order, but element " + CStr(i) + " is not a number"
8             ElseIf Indices(i, 1) <> CLng(Indices(i, 1)) Then
9                 Throw VariableName & " must be a column array of integers in the range 1 to " + CStr(N) + ") arranged in increasing order, but element " + CStr(i) + " is not an integer"
10            ElseIf Indices(i, 1) < 1 Then
11                Throw VariableName & " must be a column array of integers in the range 1 to " + CStr(N) + ") arranged in increasing order, but element " + CStr(i) + " is less than 1"
12            ElseIf Indices(i, 1) > N Then
13                Throw VariableName & " must be a column array of integers in the range 1 to " + CStr(N) + ") arranged in increasing order, but element " + CStr(i) + " is greater than " + CStr(N)
14            ElseIf i > 1 Then
15                If Indices(i, 1) <= Indices(i - 1, 1) Then
16                    Throw VariableName & " must be a column array of integers in the range 1 to " + CStr(N) + ") arranged in increasing order, but element " + CStr(i) + " is not greater than element " + CStr(i - 1)
17                End If
18            End If
19        Next i

20        Exit Function
ErrHandler:
21        Throw "#CheckIndices (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MMult
' Author     : Philip Swannell
' Date       : 19-Jul-2019
' Purpose    : Wrap Excel function MMULT
' -----------------------------------------------------------------------------------------------------------------------
Private Function MMult(Array1, Array2)
1         On Error GoTo ErrHandler
2         MMult = Application.WorksheetFunction.MMult(Array1, Array2)
3         Exit Function
ErrHandler:
4         Throw "#MMult (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MInverse
' Author     : Philip Swannell
' Date       : 19-Jul-2019
' Purpose    : Wrap Excel function MINVERSE
' -----------------------------------------------------------------------------------------------------------------------
Private Function MInverse(TheArray)
1         On Error GoTo ErrHandler
2         MInverse = Application.WorksheetFunction.MInverse(TheArray)
3         Exit Function
ErrHandler:
4         Throw "#MInverse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ComplementOf
' Author     : Philip Swannell
' Date       : 19-Jul-2019
' Purpose    : A specialised (for speed) set difference function.
'              If Y1Indices is a monotonic increasing (i.e. Y1Indices(k,1) < Y1Indices(k+1,1) ) subset of the integers
'              1 to N then the return is the monotonic complement.
' -----------------------------------------------------------------------------------------------------------------------
Private Function ComplementOf(Y1Indices As Variant, N As Long)

          Dim i As Long
          Dim NumIn1  As Long
          Dim NumIn2 As Long
          Dim ReadAt As Long
          Dim Result() As Long
          Dim WriteAt As Long
          
1         On Error GoTo ErrHandler
2         ReDim Result(1 To N - sNRows(Y1Indices), 1 To 1)
          
3         ReadAt = 1
4         WriteAt = 0
5         NumIn1 = sNRows(Y1Indices)
6         NumIn2 = N - NumIn1
          
7         For i = 1 To N
8             If ReadAt > NumIn1 Then
9                 WriteAt = WriteAt + 1
10                Result(WriteAt, 1) = i
11            ElseIf Y1Indices(ReadAt, 1) = i Then
12                ReadAt = ReadAt + 1
13            Else
14                WriteAt = WriteAt + 1
15                Result(WriteAt, 1) = i
16            End If
17        Next i

18        ComplementOf = Result

19        Exit Function
ErrHandler:
20        Throw "#ComplementOf (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sArrayDifferenceReport
' Author     : Philip Swannell
' Date       : 25-Jul-2019
' Purpose    : For two arrays (top row, left col treated as headers) reshapes the contents (no headers) into columns
'              and returns a "report" on the element-wise differences. Columns are:
'              Row - Integer row index into the Arrays
'              Col - Integer column index into the Arrays
'              LeftLabel - the relevant left label from the arrays i.e. Array1(Row,1), equivalently Array2(Row,1)
'              TopLabel - the relevant top label from the arrays i.e. Array1(1,Col), equivalently Array2(1,Col)
'              Value1 - Array1(Row,Col)
'              Value2 - Array2(Row,Col)
'              Abs(Value1) - Abs(Array1(Row,Col))
'              Abs(Value2) - Abs(Array2(Row,Col))
' Parameters :
'  Array1: An array of data, top row and left column taken to be "Headers", remaining elements numeric.
'  Array2: A second similar array, must be of same dimensions with identical headers.
'  SortCol: The return is sorted on this column, in descending order
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayDifferenceReport(Array1, Array2, Optional SortCol As Long, Optional Name1 As String = "Value1", Optional Name2 As String = "Value2")
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim LeftHeaders1
          Dim LeftHeaders2
          Dim NC As Long
          Dim NC1 As Long
          Dim NC2 As Long
          Dim NR As Long
          Dim NR1 As Long
          Dim NR2 As Long
          Dim ReturnArray() As Variant
          Dim TopHeaders1
          Dim TopHeaders2
          Dim Values1
          Dim Values2

1         On Error GoTo ErrHandler
2         Force2DArrayR Array1, NR1, NC1
3         Force2DArrayR Array2, NR2, NC2

4         If NR1 <> NR2 Or NC1 <> NC2 Then Throw "Array1 and Array2 must have the same dimensions"
5         TopHeaders1 = sSubArray(Array1, 1, 2, 1)
6         TopHeaders2 = sSubArray(Array2, 1, 2, 1)
7         LeftHeaders1 = sSubArray(Array1, 2, 1, , 1)
8         LeftHeaders2 = sSubArray(Array2, 2, 1, , 1)
9         Values1 = sSubArray(Array1, 2, 2)
10        Values2 = sSubArray(Array2, 2, 2)
11        NR = NR1 - 1
12        NC = NC1 - 1

13        If Not sArraysIdentical(TopHeaders1, TopHeaders2) Then Throw "The header row at the top of Array1 and Array2 must be identical, but they are not"
14        If Not sArraysIdentical(LeftHeaders1, LeftHeaders2) Then Throw "The header column at the left of Array1 and Array2 must be identical, but they are not"

15        ReDim ReturnArray(1 To NR * NC + 1, 1 To 8)

16        ReturnArray(1, 1) = "Row"
17        ReturnArray(1, 2) = "Col"
18        ReturnArray(1, 3) = "LeftLabel"
19        ReturnArray(1, 4) = "TopLabel"
20        ReturnArray(1, 5) = Name1
21        ReturnArray(1, 6) = Name2
22        ReturnArray(1, 7) = "Abs(" & Name1 & ")"
23        ReturnArray(1, 8) = "Abs(" & Name2 & ")"

24        k = 2
25        For i = 1 To NR
26            For j = 1 To NC
27                ReturnArray(k, 1) = i + 1 'Add one so that we get an index into Array1 rather than Values1
28                ReturnArray(k, 2) = j + 1 'Ditto
29                ReturnArray(k, 3) = LeftHeaders1(i, 1)
30                ReturnArray(k, 4) = TopHeaders1(1, j)
31                ReturnArray(k, 5) = Values1(i, j)
32                ReturnArray(k, 6) = Values2(i, j)
33                ReturnArray(k, 7) = IIf(IsNumber(Values1(i, j)), Abs(Values1(i, j)), "#Non-number!")
34                ReturnArray(k, 8) = IIf(IsNumber(Values2(i, j)), Abs(Values2(i, j)), "#Non-number!")
35                k = k + 1
36            Next j
37        Next i

38        If SortCol >= 1 And SortCol <= 8 Then
39            ReturnArray = sSortedArray(ReturnArray, SortCol, , , False, , , , , 1)
40        End If

41        sArrayDifferenceReport = ReturnArray

42        Exit Function
ErrHandler:
43        sArrayDifferenceReport = "#sArrayDifferenceReport (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
