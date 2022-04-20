Attribute VB_Name = "modMathsB"
Option Explicit
Private Enum EnmExtrapType
    ExtrapFlat = 0
    ExtrapNone = 1
    ExtrapLinear = 2
    ExtrapPolynomial = 3
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCorrelationTriangle
' Author    : Philip Swannell
' Date      : 11-Nov-2016
' Purpose   : Converts a currency correlation matrix from one numeraire to another. See "The Shape of
'             Things in a Currency Trio"
'             http://www.frbsf.org/economic-research/files/wpjl99-04a.pdf
' Arguments
' Numeraire : A three-letter currency code that does not appear in CcyList
' CcyList   : A column array of three-letter currency codes.
' NewNumeraire: A three-letter currency code that does appear in CcyList.
' CorrMatrix: A correlation matrix where element i,j is the correlation between CcyList(i)/Numeraire and
'             CcyList(j)/Numeraire
' Vols      : A column array of volatilities where element i is the volatility of CcyList(i)/Numeraire
'
' Notes     : The return is a correlation matrix where element(i,j) is the correlation between
'             CcyList(i)/NewNumeraire and CcyList(j)/NewNumeraire where Numeraire has
'             replaced NewNumeraire in CcyList.
' -----------------------------------------------------------------------------------------------------------------------
Function sCorrelationTriangle(Numeraire As String, CCyList As Variant, NewNumeraire As String, CorrMatrix As Variant, Vols As Variant)
Attribute sCorrelationTriangle.VB_Description = "Converts a currency correlation matrix from one numeraire to another. See ""The Shape of Things in a Currency Trio"" http://www.frbsf.org/economic-research/files/wpjl99-04a.pdf"
Attribute sCorrelationTriangle.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim Found As Boolean
          Dim i As Long
          Dim j As Long
          Dim N As Long
          Dim NewCCyList As Variant
          Dim NewCorrMatrix As Variant
          Dim tmp As Double
          Dim VolA As Double
          Dim VolB As Double
          Dim VolC As Double
          Dim VolMatrix

1         On Error GoTo ErrHandler

2         Force2DArrayRMulti CCyList, CorrMatrix, Vols
3         If sNCols(CCyList) <> 1 Then Throw "CCyList must have one column"
4         If sNCols(Vols) <> 1 Then Throw "Vols must have one column"
5         N = sNRows(CCyList)
6         If sNRows(Vols) <> N Then Throw "Vols must have the same number of rows as CCylist"

7         If N <> sNRows(CorrMatrix) Then
8             Throw "CorrMatrix must have the same number of rows and columns as CCyList has rows"
9         ElseIf N <> sNCols(CorrMatrix) Then
10            Throw "CorrMatrix must have the same number of rows and columns as CCyList has rows"
11        End If

12        For i = 1 To N
13            If Not IsNumber(Vols(i, 1)) Then Throw "Vols must be numbers"
14            If Vols(i, 1) < 0 Then Throw "Vols must be positive"
15            If VarType(CCyList(i, 1)) <> vbString Then Throw "CcyList must be strings"
16            If (CCyList(i, 1)) = Numeraire Then Throw "Numeraire cannot appear in CCyList"
17            If CCyList(i, 1) = NewNumeraire Then Found = True
18            If Not IsNumber(CorrMatrix(i, i)) Then Throw "On-diagonal elements of CorrMatrix must be 1"
19            If CorrMatrix(i, i) <> 1 Then Throw "On-diagonal elements of CorrMatrix must be 1"
20        Next i
21        If Not Found Then Throw "NewNumeraire must appear in CcyList"
22        For i = 1 To N
23            For j = 1 To i - 1
24                If Not IsNumber(CorrMatrix(i, j)) Or Not IsNumber(CorrMatrix(j, 1)) Then Throw "All elements of CorrMatrix must be numbers"
25                If CorrMatrix(i, j) <> CorrMatrix(j, i) Then Throw "CorrMatrix must be symmetric, but element " + CStr(i) + "," + CStr(j) + " is not equal to element " + CStr(j) + "," + CStr(i)
26                If CorrMatrix(i, j) < -1 Or CorrMatrix(i, j) > 1 Then Throw "All elements of CorrMatrix must be between -1 and 1"
27            Next j
28        Next i

29        NewCCyList = sArrayIf(sArrayEquals(CCyList, NewNumeraire), Numeraire, CCyList)

30        VolMatrix = sReshape(0, N, N)        'VolMatrix is indexed by ccylist. Element i,j is the vol of CCyList(i) vs CCyList(j)

31        For i = 1 To N
32            For j = 1 To i - 1
33                tmp = Vols(i, 1) ^ 2 + Vols(j, 1) ^ 2 - 2 * Vols(i, 1) * Vols(j, 1) * CorrMatrix(i, j)
34                If tmp < 0 Then Throw "Cannot imply vol for " + CCyList(i) + CCyList(j)
35                VolMatrix(i, j) = Sqr(tmp)
36                VolMatrix(j, i) = VolMatrix(i, j)
37            Next j
38        Next i

          Dim IoN        'IndexOfNumeraire
39        IoN = sMatch(NewNumeraire, CCyList)

40        NewCorrMatrix = sIdentityMatrix(N)
41        For i = 1 To N
42            For j = 1 To i - 1
43                If i <> IoN And j <> IoN Then
44                    VolA = VolMatrix(i, IoN)
45                    VolB = VolMatrix(j, IoN)
46                    VolC = VolMatrix(i, j)
47                ElseIf i = IoN Then
48                    VolA = Vols(i, 1)
49                    VolB = VolMatrix(j, IoN)
50                    VolC = Vols(j, 1)
51                ElseIf j = IoN Then
52                    VolA = VolMatrix(i, IoN)
53                    VolB = Vols(j, 1)
54                    VolC = Vols(i, 1)
55                End If
56                NewCorrMatrix(i, j) = (VolA ^ 2 + VolB ^ 2 - VolC ^ 2) / (2 * VolA * VolB)
57                NewCorrMatrix(j, i) = NewCorrMatrix(i, j)
58            Next j
59        Next i

60        sCorrelationTriangle = NewCorrMatrix

61        Exit Function
ErrHandler:
62        sCorrelationTriangle = "#sCorrelationTriangle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'TODO DOCUMENT THIS!
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sCramerVonMises
' Author     : Philip Swannell
' Date       : 23-Mar-2019
' Purpose    : Takes a sample, adjusts to have zero mean and unit variance and then calculates the Cramer-von Mises distance to N(0,1)
'              https://en.wikipedia.org/wiki/Cram%C3%A9r%E2%80%93von_Mises_criterion
' -----------------------------------------------------------------------------------------------------------------------
Function sCramerVonMises(ByVal Sample As Variant)
Attribute sCramerVonMises.VB_Description = "Tests Sample for normality. After adjustment to have have zero mean and unit variance, the Cramer-von Mises statistic is calculated versus the standard normal distribution. See https://en.wikipedia.org/wiki/Cram%C3%A9r%E2%80%93von_Mises_criterion"
Attribute sCramerVonMises.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim Av As Double
          Dim ChooseVector
          Dim i As Long
          Dim N As Long
          Dim NC As Long
          Dim NR As Long
          Dim SD As Double
          Dim x As Double
          Dim y As Double

1         On Error GoTo ErrHandler
2         Force2DArrayR Sample

3         NR = sNRows(Sample)
4         NC = sNCols(Sample)
5         N = NC * NR
6         If NC <> 1 Then
7             Sample = sReshape(Sample, N, 1)
8         End If
9         ChooseVector = sArrayIsNumber(Sample)
10        If Not sAll(ChooseVector) Then
11            Sample = sMChoose(Sample, sArrayIsNumber(Sample))
12            N = sNRows(Sample)
13        End If
14        Sample = sSortedArray(Sample)
15        Av = Application.WorksheetFunction.Average(Sample)
16        SD = Application.WorksheetFunction.StDev_P(Sample)

          Dim Result As Double
17        Result = 1 / (12 * N)
18        For i = 1 To N
19            x = ((2 * i) - 1) / 2 / N
20            y = (Sample(i, 1) - Av) / SD
21            Result = Result + (x - func_normsdistdd(y)) ^ 2
22        Next i

23        sCramerVonMises = Result

24        Exit Function
ErrHandler:
25        sCramerVonMises = "#sCramerVonMises (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sEigen
' Author    : Philip Swannell
' Date      : 05-Jun-2017
' Purpose   : Computes the eigenvalues and eigenvectors for a real symmetric positive definite matrix
'             using the "JK Method".  The first column of the return matrix contains the
'             eigenvalues and the remaining columns contain the eigenvectors.
' Arguments
' M         : The input matrix, must be symmetric positive definite.
'
' Notes     : See: KAISER,H.F. (1972) "THE JK METHOD: A PROCEDURE FOR FINDING THE EIGENVALUES OF A REAL
'             SYMMETRIC MATRIX", The Computer Journal, VOL.15, 271-273.
'             Code from http://www.freevbcode.com/ShowCode.asp?ID=9209,
'             with changes to error handling to work from Excel and more checking of data passed in
'             for more informative errors.
'             Original code was for VB6 with function name EIGEN_JK.
'             Original code could give incorrect sign of eigen values, (always returned positive eigen values)
'             got correction from
' https://www.experts-exchange.com/questions/24824063/VBA-Function-for-Eigen-decomposition-returns-only-positive-eigenvalues.html
' -----------------------------------------------------------------------------------------------------------------------
Function sEigen(ByVal Matrix As Variant) As Variant
Attribute sEigen.VB_Description = "Computes the eigenvalues and eigenvectors for a real symmetric positive definite matrix using the ""JK Method"".  The first column of the return matrix contains the eigenvalues and the remaining columns contain the eigenvectors."
Attribute sEigen.VB_ProcData.VB_Invoke_Func = " \n23"

          '***************************************************************************
          '**  Function computes the eigenvalues and eigenvectors for a real        **
          '**  symmetric positive definite matrix using the "JK Method".  The       **
          '**  first column of the return matrix contains the eigenvalues and       **
          '**  the rest of the p+1 columns contain the eigenvectors.                **
          '**  See:                                                                 **
          '**  KAISER,H.F. (1972) "THE JK METHOD: A PROCEDURE FOR FINDING THE       **
          '**  EIGENVALUES OF A REAL SYMMETRIC MATRIX", The Computer Journal,       **
          '**  VOL.15, 271-273.                                                     **
          '***************************************************************************

          Dim a() As Variant
          Dim Cos_ As Double
          Dim Cos2 As Double
          Dim Cot2 As Double
          Dim den As Double
          Dim Ematrix() As Double
          Dim hold As Double
          Dim i As Long
          Dim iter As Long
          Dim j As Long
          Dim k As Long
          Dim NUM As Double
          Dim p As Long
          Dim Sin_ As Double
          Dim Sin2 As Double
          Dim Tan2 As Double
          Dim Test As Double
          Dim tmp As Double
          Const eps As Double = 1E-16
          Const MaxIter = 20    'was 15 in the code copied from web pages above but not hard to find cases when 15 not enough but 20 works

1         On Error GoTo ErrHandler

2         Force2DArrayR Matrix

3         a = Matrix

4         If LBound(a, 1) <> 1 Or LBound(a, 2) <> 1 Then Throw "Matrix must have lower bounds of 1"
5         If UBound(a, 1) <> UBound(a, 2) Then Throw "Matrix must be a square matrix with the same number of rows as columns"
6         p = UBound(a, 1)

7         For i = 1 To p
8             For j = 1 To i
9                 If Not IsNumber(a(i, j)) Then Throw "All elements of Matrix must be numbers but element " + CStr(i) + "," + CStr(j) + " is not"
10                If Not IsNumber(a(j, i)) Then Throw "All elements of Matrix must be numbers but element " + CStr(j) + "," + CStr(i) + " is not"
11                If a(j, i) <> a(i, j) Then Throw "Matrix must be a symmetric positive definite, but it's not symmetric since element " + CStr(i) + "," + CStr(j) + " is not a equal to element " + CStr(j) + "," + CStr(i)
12            Next j
13        Next i

14        ReDim Ematrix(1 To p, 1 To p + 1)

15        For iter = 1 To MaxIter

              'Orthogonalize pairs of columns in upper off diag
16            For j = 1 To p - 1
17                For k = j + 1 To p

18                    den = 0#
19                    NUM = 0#
                      'Perform single plane rotation
20                    For i = 1 To p
21                        NUM = NUM + 2 * a(i, j) * a(i, k)   ': numerator eq. 11
22                        den = den + (a(i, j) + a(i, k)) * _
                              (a(i, j) - a(i, k))             ': denominator eq. 11
23                    Next i

                      'Skip rotation if aij is zero and correct ordering
24                    If Abs(NUM) < eps And den >= 0 Then Exit For

                      'Perform Rotation
25                    If Abs(NUM) <= Abs(den) Then
26                        Tan2 = Abs(NUM) / Abs(den)          ': eq. 11
27                        Cos2 = 1 / Sqr(1 + Tan2 * Tan2)     ': eq. 12
28                        Sin2 = Tan2 * Cos2                  ': eq. 13
29                    Else
30                        Cot2 = Abs(den) / Abs(NUM)          ': eq. 16
31                        Sin2 = 1 / Sqr(1 + Cot2 * Cot2)     ': eq. 17
32                        Cos2 = Cot2 * Sin2                  ': eq. 18
33                    End If

34                    Cos_ = Sqr((1 + Cos2) / 2)              ': eq. 14/19
35                    Sin_ = Sin2 / (2 * Cos_)                ': eq. 15/20

36                    If den < 0 Then
37                        tmp = Cos_
38                        Cos_ = Sin_                         ': table 21
39                        Sin_ = tmp
40                    End If

41                    Sin_ = Sgn(NUM) * Sin_                  ': sign table 21

                      'Rotate
42                    For i = 1 To p
43                        tmp = a(i, j)
44                        a(i, j) = tmp * Cos_ + a(i, k) * Sin_
45                        a(i, k) = -tmp * Sin_ + a(i, k) * Cos_
46                    Next i

47                Next k
48            Next j

              'Test for convergence
49            Test = Application.WorksheetFunction.SumSq(a)
50            If Abs(Test - hold) < eps And iter > 5 Then Exit For
51            hold = Test
52        Next iter

53        If iter = MaxIter + 1 Then
              'Failure to converge often caused by matrix not +ve semi definite and sCholesky effectively tests for that
              'Implicit assumption that the JK method fails when Matrix not +ve definite
54            ThrowIfError sCholesky(Matrix)
55            Throw "JK Iteration has not converged."
56        End If

          'Compute eigenvalues/eigenvectors
57        For j = 1 To p
              'Compute eigenvalues
58            For k = 1 To p
59                Ematrix(j, 1) = Ematrix(j, 1) + a(k, j) ^ 2
60            Next k
61            Ematrix(j, 1) = Sqr(Ematrix(j, 1))    'PGS - this will always give positive Eigen values, but corrected below...

              'Normalize eigenvectors
62            For i = 1 To p
63                If Ematrix(j, 1) <= 0 Then
64                    Ematrix(i, j + 1) = 0
65                Else
66                    Ematrix(i, j + 1) = a(i, j) / Ematrix(j, 1)
67                End If
68            Next i

              Dim M As Long
69            If Ematrix(j, 1) > 0 Then
                  ' Find biggest component
70                M = 1
71                For i = 2 To p
72                    If Abs(a(i, j)) > Abs(a(M, j)) Then
73                        M = i
74                    End If
75                Next i
                  ' Calculate m'th component of Matrix and jth eigenvector
76                tmp = 0
77                For i = 1 To p
78                    tmp = tmp + Matrix(M, i) * a(i, j)
79                Next i
                  ' recalculate eigenvalue as quotient output/input
80                Ematrix(j, 1) = tmp / a(M, j)
81            End If

82        Next j

83        sEigen = Ematrix

84        Exit Function
ErrHandler:
85        sEigen = "#sEigen (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestsEigen
' Author    : Philip Swannell
' Date      : 07-Jun-2017
' Purpose   : Test harness. Not too hard to find cases where sEigen does not converge
'             approx 1 time in 5,000 for the randomly generated symmetric matrix below, but
'             did not find cases where EigeVectors and EigenValues did not "work"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestsEigen()

          Dim Check
          Dim EigenValues As Variant
          Dim EigenVectors As Variant
          Dim i As Long
          Dim Matrix
          Dim p As Long
          Dim Res
          Dim Res1
          Dim Res2

1         On Error GoTo ErrHandler
2         For i = 1 To 1000
3             p = 10

4             Matrix = sRandomVariable(p, p)
5             Matrix = sArrayAdd(Matrix, sArrayTranspose(Matrix))
6             Res = sEigen(Matrix)
7             If sIsErrorString(Res) Then Stop
8             EigenValues = sSubArray(Res, 1, 1, , 1)
9             EigenVectors = sSubArray(Res, 1, 2)

10            Res1 = Application.WorksheetFunction.MMult(Matrix, EigenVectors)
11            Res2 = sArrayMultiply(sArrayTranspose(EigenValues), EigenVectors)

12            Check = sMaxOfArray(sArrayAbs(sArraySubtract(Res1, Res2)))
13            If Check > 0.0000000001 Then Stop
              'Debug.Print Check

14        Next

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#TestsEigen (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sInterp
' Author    : Philip Swannell
' Date      : 27-Apr-2015
' Purpose   : Interpolation function. xArrayAscending and yArray define a function y = f(x) and the
'             return from the function is the interpolated values of f(x) at the points of
'             xValues.
'             See also the associated allocation function sAllocExposures.
' Arguments
' xArrayAscending: An array of numbers, may be either a single column or a single row.  Must be in ascending
'             order - smaller numbers at the top (left) , larger numbers at the bottom
'             (right).
' yArray    : A single column or single row array of values - must have the same number of elements as
'             xArray. Values must be numeric unless InterpType is FlatFromLeft or
'             FlatToRight in which case non-numbers are permitted.
' xValues   : An array of numbers , for which the function will return interpolated values of y. May be
'             any size of array.
' InterpType: Determines interpolation scheme used. Allowed: Linear, FlatFromLeft, FlatToRight,
'             BlendedQuadratic, FMM, Natural, Periodic, Hyman, monoH.FC
' Signature : 2 characters to set left and right extrapolation beyond the bounds of xArrayAscending.
'             F = Flat, X = Linear eXtrapolation, P = Polynomial, N = None. If omitted
'             defaults from InterpType: Linear > NN, FlatFromLeft > NF, FlatToRight > FN
'
' Notes     : More information on InterpType:
'             BlendedQuadratic:
'             If A<B<x<C<D for adjacent "knots" (grid-points) A,B,C,D and f is the
'             quadratic fit at A,B,C and g the quadratic fit at B,C,D  sInterp(x) = w1 *
'             f(x) + w2 * g(x), w1 = (C-x)/(C-B), w2 = 1-w1. BlendedQuadratic has useful
'             "locality" properties since the interpolated value at a point x is determined
'             only by four knots, two to its left and two two its right.
'             FMM, Natural, Periodic, Hyman, monoH.FC
'             For these InterpTypes the function sInterp is a wrap to the R function
'             spline. For details see
'             https://stat.ethz.ch/R-manual/R-devel/library/stats/html/smooth.spline.html
'             https://stackoverflow.com/questions/22509016/splinefun-with-method-fmm
' -----------------------------------------------------------------------------------------------------------------------
Function sInterp(ByVal xArrayAscending As Variant, ByVal yArray As Variant, ByVal xValues As Variant, _
        Optional ByVal InterpType As String = "Linear", Optional ByRef Signature As String = vbNullString)
Attribute sInterp.VB_Description = "Interpolation function. xArrayAscending and yArray define a function y = f(x) and the return from the function is the interpolated values of f(x) at the points of xValues.\nSee also the associated allocation function sAllocExposures."
Attribute sInterp.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim c As Variant
          Dim i As Long
          Dim interpTypeCode As Long
          Dim j As Long
          Dim LeftExtrapType As EnmExtrapType
          Dim M As Long
          Dim N As Long
          Dim Result() As Variant
          Dim RightExtrapType As EnmExtrapType
          Dim SearchRes As Long
          Dim sizeOfxArray As Long
          Dim x1 As Double
          Dim X2 As Double
          Dim y1 As Double
          Dim y2 As Double
          Const badSignature = "Unrecognised Signature. Must be a 2 letter string for left and right extrapolation styles. F = Flat, X = Linear, P Polynomial, N = No extrapolation."
          Const badSignature2 = "Invalid signature. When InterpType is FlatFromLeft or FlatToRight then Signature must be one of FF, NN, FN or NF"

1         On Error GoTo ErrHandler
2         Force2DArrayR xArrayAscending: Force2DArrayR yArray: Force2DArrayR xValues
3         If sNCols(xArrayAscending) > 1 And sNRows(xArrayAscending) = 1 Then xArrayAscending = sArrayTranspose(xArrayAscending)
4         If sNCols(yArray) > 1 And sNRows(yArray) = 1 Then yArray = sArrayTranspose(yArray)

5         If sNCols(xArrayAscending) <> 1 Then Throw "xArrayAscending must have either 1 column or 1 row"
6         If sNCols(yArray) < 1 Then Throw "yArray must have either 1 column or 1 row "
7         sizeOfxArray = sNRows(xArrayAscending)
8         If sizeOfxArray <> sNRows(yArray) Then Throw "xArrayAscending and yArray must have the same number of elements"

9         Select Case LCase$(Replace(InterpType, " ", vbNullString))
              Case "linear"
10                interpTypeCode = 1
11                If Signature = vbNullString Then Signature = "NN"
12            Case "flatfromleft"
13                interpTypeCode = 2
14                If Signature = vbNullString Then Signature = "NF"
15                If InStr(UCase$(Signature), "X") > 0 Then Throw badSignature2
16            Case "flattoright"
17                interpTypeCode = 3
18                If Signature = vbNullString Then Signature = "FN"
19                If InStr(UCase$(Signature), "X") > 0 Then Throw badSignature2
20            Case "fmmspline", "naturalspline", "periodicspline", "hymanspline", "monoh.fcspline", "fmm", "natural", "periodic", "hyman", "monoh.fc"
21                interpTypeCode = 4
22                If Signature = vbNullString Then Signature = "PP"
23            Case "blendedquadratic"
24                interpTypeCode = 5
25                If Signature = vbNullString Then Signature = "PP"
26            Case Else
27                Throw "InterpType must be one of Linear, FlatFromLeft or FlatToRight, BlendedQuadratic, FMM, Natural, Periodic, Hyman, monoH.FC"
28        End Select

29        If Len(Signature) <> 2 Then Throw badSignature
30        Select Case UCase$(Left$(Signature, 1))
              Case "P"
31                If interpTypeCode < 4 Then Throw "Polynomial extrapolation not supported for InterpType " + InterpType
32                LeftExtrapType = ExtrapPolynomial
33            Case "F"
34                LeftExtrapType = ExtrapFlat
35            Case "X"
36                If interpTypeCode >= 4 Then Throw "Linear extrapolation not supported for InterpType " + InterpType
37                LeftExtrapType = ExtrapLinear
38            Case "N"
39                LeftExtrapType = ExtrapNone
40                If sColumnMin(xValues)(1, 1) < xArrayAscending(1, 1) Then
41                    If interpTypeCode < 4 Then
42                        Throw "#Extrapolation not allowed. First character of Signature must be F (Flat extrapolation) or X (linear eXtrapolation)!"
43                    Else
44                        Throw "#Extrapolation not allowed. First character of Signature must be F (Flat extrapolation) or P (Polynomial extrapolation)!"
45                    End If
46                End If
47            Case Else
48                Throw badSignature
49        End Select
50        Select Case UCase$(Right$(Signature, 1))
              Case "P"
51                If interpTypeCode < 4 Then Throw "Polynomial extrapolation not supported for InterpType " + InterpType
52                RightExtrapType = ExtrapPolynomial
53            Case "F"
54                RightExtrapType = ExtrapFlat
55            Case "X"
56                If interpTypeCode >= 4 Then Throw "Linear extrapolation not supported for InterpType " + InterpType
57                RightExtrapType = ExtrapLinear
58            Case "N"
59                RightExtrapType = ExtrapNone
60                If sColumnMax(xValues)(1, 1) > xArrayAscending(sizeOfxArray, 1) Then
61                    If interpTypeCode < 4 Then
62                        Throw "#Extrapolation not allowed. Second character of Signature must be F (Flat extrapolation) or X (linear eXtrapolation)!"
63                    Else
64                        Throw "#Extrapolation not allowed. Second character of Signature must be F (Flat extrapolation) or P (Polynomial extrapolation)!"
65                    End If
66                End If
67            Case Else
68                Throw badSignature
69        End Select

70        i = 0
71        For Each c In xArrayAscending
72            i = i + 1
73            If Not IsNumberOrDate(c) Then Throw "Non number found in xArrayAscending"
74            If i > 1 Then If c <= xArrayAscending(i - 1, 1) Then Throw "xArrayAscending must be in ascending order, but element " + CStr(i) + " is not greater than element " + CStr(i - 1)
75        Next
76        If Not (interpTypeCode = 2 Or interpTypeCode = 3) Then
77            For Each c In yArray
78                If Not IsNumberOrDate(c) Then Throw "Non number found in yArray"
79            Next
80        End If

          'End of input checking, now call the sub routines for non linear interpolation
81        If interpTypeCode = 4 Then
82            InterpType = Replace(LCase$(InterpType), "spline", vbNullString)
83            If InterpType = "monoh.fc" Then InterpType = "monoH.FC"
84            sInterp = ThrowIfError(InterpSpline(xArrayAscending, yArray, xValues, InterpType, LeftExtrapType, RightExtrapType))
85            Exit Function
86        ElseIf interpTypeCode = 5 Then
87            sInterp = ThrowIfError(InterpBlendedQuadratic(xArrayAscending, yArray, xValues, LeftExtrapType, RightExtrapType))
88            Exit Function
89        End If

90        N = sNRows(xValues): M = sNCols(xValues)
91        ReDim Result(1 To N, 1 To M)
92        If interpTypeCode = 1 Then        'Linear
93            For i = 1 To N
94                For j = 1 To M
95                    If Not IsNumberOrDate(xValues(i, j)) Then
96                        Result(i, j) = "#Non number in xValues!"
97                    Else
98                        SearchRes = ThrowIfError(sBinaryChopSearch(xValues(i, j), xArrayAscending, sizeOfxArray))
99                        If SearchRes = 0 Then
100                           If LeftExtrapType = ExtrapFlat Then
101                               Result(i, j) = yArray(1, 1)
102                           ElseIf LeftExtrapType = ExtrapLinear Then
103                               x1 = xArrayAscending(1, 1): X2 = xArrayAscending(2, 1)
104                               y1 = yArray(1, 1): y2 = yArray(2, 1)
                                  ' Result(i, j) = y1 * (x2 - xValues(i, j)) / (x2 - x1) + y2 * (xValues(i, j) - x1) / (x2 - x1)    'TODO - safe error handling at this line
105                               Result(i, j) = SafeInterp(xValues(i, j), x1, X2, y1, y2)
106                           ElseIf LeftExtrapType = ExtrapNone Then
                                  'should never hit this case as input checking should have trapped this condition...
107                               Result(i, j) = "#Extrapolation not allowed. First character of Signature must be F (Flat extrapolation) or X (linear eXtrapolation)!"
108                           End If
109                       ElseIf xValues(i, j) > xArrayAscending(sizeOfxArray, 1) Then
110                           If RightExtrapType = ExtrapFlat Then
111                               Result(i, j) = yArray(sizeOfxArray, 1)
112                           ElseIf RightExtrapType = ExtrapLinear Then
113                               x1 = xArrayAscending(sizeOfxArray - 1, 1): X2 = xArrayAscending(sizeOfxArray, 1)
114                               y1 = yArray(sizeOfxArray - 1, 1): y2 = yArray(sizeOfxArray, 1)
                                  ' Result(i, j) = y1 * (x2 - xValues(i, j)) / (x2 - x1) + y2 * (xValues(i, j) - x1) / (x2 - x1)    'TODO - safe error handling at this line
115                               Result(i, j) = SafeInterp(xValues(i, j), x1, X2, y1, y2)
116                           ElseIf RightExtrapType = ExtrapNone Then
117                               Result(i, j) = "#Extrapolation not allowed. Second character of Signature must be F (Flat extrapolation) or X (linear eXtrapolation)!"
118                           End If
119                       Else
120                           x1 = xArrayAscending(SearchRes, 1): X2 = xArrayAscending(SearchRes + 1, 1)
121                           y1 = yArray(SearchRes, 1): y2 = yArray(SearchRes + 1, 1)
                              'Result(i, j) = y1 * (x2 - xValues(i, j)) / (x2 - x1) + y2 * (xValues(i, j) - x1) / (x2 - x1)    'TODO - safe error handling at this line
122                           Result(i, j) = SafeInterp(xValues(i, j), x1, X2, y1, y2)
123                       End If
124                   End If
125               Next j
126           Next i
127       ElseIf interpTypeCode = 2 Then        'FlatFromLeft
128           For i = 1 To N
129               For j = 1 To M
130                   If Not IsNumberOrDate(xValues(i, j)) Then
131                       Result(i, j) = "#Non number in xValues!"
132                   Else

133                       SearchRes = ThrowIfError(sBinaryChopSearch(xValues(i, j), xArrayAscending, sizeOfxArray))
134                       If SearchRes = 0 Then        'xValues(i, j)<xArrayAscending(1,1)
135                           If LeftExtrapType = ExtrapFlat Then
136                               Result(i, j) = yArray(1, 1)
137                           ElseIf LeftExtrapType = ExtrapNone Then
                                  'should never hit this case as input checking should have trapped this condition...
138                               Result(i, j) = "#Extrapolation not allowed. First character of Signature must be F to allow."
139                           End If
140                       ElseIf xValues(i, j) >= xArrayAscending(sizeOfxArray, 1) Then
141                           If RightExtrapType = ExtrapFlat Or xValues(i, j) = xArrayAscending(sizeOfxArray, 1) Then
142                               Result(i, j) = yArray(sizeOfxArray, 1)
143                           ElseIf RightExtrapType = ExtrapNone Then
144                               Result(i, j) = "#Extrapolation not allowed. Second character of Signature must be F to allow."
145                           End If
146                       Else
147                           Result(i, j) = yArray(SearchRes, 1)
148                       End If
149                   End If
150               Next j
151           Next i
152       ElseIf interpTypeCode = 3 Then        'FlatToRight
153           For i = 1 To N
154               For j = 1 To M
155                   If Not IsNumberOrDate(xValues(i, j)) Then
156                       Result(i, j) = "#Non number in xValues!"
157                   Else
158                       SearchRes = ThrowIfError(sBinaryChopSearch(xValues(i, j), xArrayAscending, sizeOfxArray))
159                       If SearchRes = 0 Then        'xValues(i, j)<xArrayAscending(1,1)
160                           If (LeftExtrapType = ExtrapFlat) Then
161                               Result(i, j) = yArray(1, 1)
162                           ElseIf LeftExtrapType = ExtrapNone Then
163                               Result(i, j) = "#Extrapolation not allowed. First character of Signature must be F to allow."
164                           End If
165                       ElseIf xValues(i, j) >= xArrayAscending(sizeOfxArray, 1) Then
166                           If RightExtrapType = ExtrapFlat Or xValues(i, j) = xArrayAscending(sizeOfxArray, 1) Then
167                               Result(i, j) = yArray(sizeOfxArray, 1)
168                           ElseIf RightExtrapType = ExtrapNone Then
169                               Result(i, j) = "#Extrapolation not allowed. Second character of Signature must be F to allow."
170                           End If
171                       Else
172                           If xValues(i, j) = xArrayAscending(SearchRes, 1) Then
173                               Result(i, j) = yArray(SearchRes, 1)
174                           Else
175                               Result(i, j) = yArray(SearchRes + 1, 1)
176                           End If
177                       End If
178                   End If
179               Next j
180           Next i
181       End If

182       sInterp = Result
183       Exit Function
ErrHandler:
184       sInterp = "#sInterp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: InterpSpline
' Purpose:  Wrap to base R's function spline, via my R function InterpSpline, with handling of flat extrapolation in this VBA layer
' Author: Philip Swannell
' Date: 23-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Private Function InterpSpline(ByVal xArrayAscending As Variant, ByVal yArray As Variant, ByVal xValues As Variant, _
        ByVal SplineType As String, LeftExtrapType As EnmExtrapType, RightExtrapType As EnmExtrapType)
          Static HaveChecked As Boolean
          Static N As Long
          Static Result As Variant
1         On Error GoTo ErrHandler
2         If Not HaveChecked Then CheckR "InterpSpline", gPackagesSAI, gRSourcePath + "SolumAddin.R": HaveChecked = True
3         Result = ThrowIfError(Application.Run("BERT.Call", "InterpSpline", xArrayAscending, yArray, xValues, SplineType))
4         If LeftExtrapType = ExtrapFlat Then
5             Result = sArrayIf(sArrayLessThan(xValues, xArrayAscending(1, 1)), yArray(1, 1), Result)
6         End If
7         If RightExtrapType = ExtrapFlat Then
8             N = sNRows(xArrayAscending)
9             Result = sArrayIf(sArrayGreaterThan(xValues, xArrayAscending(N, 1)), yArray(N, 1), Result)
10        End If
11        InterpSpline = Result

12        Exit Function
ErrHandler:
13        InterpSpline = "#InterpSpline (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: InterpBlendedQuadratic
' Purpose:  Interpolation by fitting quadratics to successive triplets. Where we have A<B<x<C<D for knots A,B,C,D
' then Interp(x) = W1 * f(x) + W2 * g(x), f is the quadratic fitting at A,B,C and g that fitting at C,D,E. W1 = (C-x)/(C-D), W2 = 1-W1
' Parameter xArrayAscending (Variant):
' Parameter yArray (Variant):
' Parameter xValues (Variant):
' Parameter LeftExtrapType (EnmExtrapType): Only need to support ExtrapFlat and ExtrapPolynomial
' Parameter RightExtrapType (EnmExtrapType):
' Author: Philip Swannell
' Date: 22-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Private Function InterpBlendedQuadratic(ByVal xArrayAscending As Variant, ByVal yArray As Variant, ByVal xValues As Variant, LeftExtrapType As EnmExtrapType, RightExtrapType As EnmExtrapType)
1         On Error GoTo ErrHandler
2         Force2DArrayRMulti xArrayAscending, yArray, xValues
          Dim i As Long
          Dim M As Long
          Dim N As Long
          Dim QuadraticCoeffs
          Dim Res
          Dim xVec(1 To 3, 1 To 1) As Double
          Dim yVec(1 To 3, 1 To 1) As Double
3         N = sNRows(xArrayAscending)
4         M = sNRows(xValues)
5         QuadraticCoeffs = sReshape(0, N - 2, 3)
6         For i = 1 To N - 2
7             xVec(1, 1) = xArrayAscending(i, 1)
8             xVec(2, 1) = xArrayAscending(i + 1, 1)
9             xVec(3, 1) = xArrayAscending(i + 2, 1)
10            yVec(1, 1) = yArray(i, 1)
11            yVec(2, 1) = yArray(i + 1, 1)
12            yVec(3, 1) = yArray(i + 2, 1)
13            Res = PolyFit(xVec, yVec)
14            QuadraticCoeffs(i, 1) = Res(1, 1)
15            QuadraticCoeffs(i, 2) = Res(2, 1)
16            QuadraticCoeffs(i, 3) = Res(3, 1)
17        Next i
18        InterpBlendedQuadratic = QuadraticCoeffs
          Dim Result
          Dim RowNum As Long
          Dim Weight1 As Double
          Dim Weight2 As Double
19        Result = sReshape(0, M, 1)

20        For i = 1 To M
21            If LeftExtrapType = ExtrapFlat And xValues(i, 1) < xArrayAscending(1, 1) Then
22                Result(i, 1) = yArray(1, 1)
23            ElseIf RightExtrapType = ExtrapFlat And xValues(i, 1) > xArrayAscending(N, 1) Then
24                Result(i, 1) = yArray(N, 1)
25            Else
26                RowNum = sBinaryChopSearch(xValues(i, 1), xArrayAscending, N)
27                If RowNum <= 1 Then
28                    Result(i, 1) = QuadraticCoeffs(1, 1) * xValues(i, 1) ^ 2 + _
                          QuadraticCoeffs(1, 2) * xValues(i, 1) + _
                          QuadraticCoeffs(1, 3)
29                ElseIf RowNum >= N - 1 Then
30                    Result(i, 1) = QuadraticCoeffs(N - 2, 1) * xValues(i, 1) ^ 2 + _
                          QuadraticCoeffs(N - 2, 2) * xValues(i, 1) + _
                          QuadraticCoeffs(N - 2, 3)
31                Else
32                    Weight1 = (xArrayAscending(RowNum + 1, 1) - xValues(i, 1)) / (xArrayAscending(RowNum + 1, 1) - xArrayAscending(RowNum, 1))
33                    Weight2 = 1 - Weight1

34                    Result(i, 1) = Weight1 * (QuadraticCoeffs(RowNum - 1, 1) * xValues(i, 1) ^ 2 + _
                          QuadraticCoeffs(RowNum - 1, 2) * xValues(i, 1) + _
                          QuadraticCoeffs(RowNum - 1, 3)) + _
                          Weight2 * (QuadraticCoeffs(RowNum, 1) * xValues(i, 1) ^ 2 + _
                          QuadraticCoeffs(RowNum, 2) * xValues(i, 1) + _
                          QuadraticCoeffs(RowNum, 3))
35                End If
36            End If
37        Next i

38        InterpBlendedQuadratic = Result

39        Exit Function
ErrHandler:
40        InterpBlendedQuadratic = "#InterpBlendedQuadratic (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: PolyFit
' Purpose: Returns the coefficients of an n-1 order polynomial that exactly fit the n points defined by xVec and yVec
' Parameter xVec ():
' Parameter yVec ():
' Author: Philip Swannell
' Date: 22-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Private Function PolyFit(xVec, yVec)
          Dim a() As Double
          Dim i As Long
          Dim j As Long
          Dim N As Long
1         On Error GoTo ErrHandler
          '2         Force2DArrayRMulti xVec, yVec 'Remove when tested
2         N = sNRows(xVec)
3         ReDim a(1 To N, 1 To N)
4         For i = 1 To N
5             For j = 1 To N
6                 a(i, j) = xVec(i, 1) ^ (N - j)
7             Next j
8         Next i

9         PolyFit = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(a), yVec)

10        Exit Function
ErrHandler:
11        Throw "#PolyFit (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBinaryChopSearch
' Author    : Philip Swannell
' Purpose   : The function returns the position in LookupColumn of the largest element that
'             is smaller than LookupValue. Returns zero if LookUpValue is smaller than the first element of LookupColumn
' Parameter LookupValue (): a number
' Parameter LookupColumn (): single column array of numbers in ascending order (i.e. small numbers at top, big numbers at bottom)
' Parameter N (Long): the number of rows in LookupColumn
' Date      : 27-Apr-2015
' -----------------------------------------------------------------------------------------------------------------------
Private Function sBinaryChopSearch(LookupValue, LookupColumn, N As Long)
          Dim Bottom As Long
          Dim Middle As Long
          Dim Top As Long

1         On Error GoTo ErrHandler
          '  Force2DArrayR LookupColumn ' not necessary except when testing this fn from a sheet

2         Top = 1: Bottom = N
3         Do While Bottom - Top > 1
4             Middle = (Top + Bottom) / 2
5             If LookupColumn(Middle, 1) < LookupValue Then
6                 Top = Middle
7             ElseIf LookupColumn(Middle, 1) = LookupValue Then
8                 Top = Middle: Bottom = Middle
9             Else
10                Bottom = Middle
11            End If
12        Loop
13        If LookupValue < LookupColumn(Top, 1) Then
14            sBinaryChopSearch = Top - 1
15        ElseIf LookupValue = LookupColumn(Top, 1) Then
16            sBinaryChopSearch = Top
17        ElseIf LookupValue <= LookupColumn(Bottom, 1) Then
18            sBinaryChopSearch = Top
19        ElseIf LookupValue > LookupColumn(Bottom, 1) Then
20            sBinaryChopSearch = Bottom
21        End If

22        Exit Function
ErrHandler:
23        sBinaryChopSearch = "#sBinaryChopSearch (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
