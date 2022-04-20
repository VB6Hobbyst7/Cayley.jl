Attribute VB_Name = "modMathsA"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : bsCore
' Author    : Philip Swannell
' Date      : 30-Apr-2015
' Purpose   : Black-Scholes formula for undiscounted value of European calls, puts, digitals
'             and forwards
' 12/10/16    Moved to this module from modLowerLevelFunctions (Private Module) )so that we
'             can call from the Cayley workbook
' -----------------------------------------------------------------------------------------------------------------------
Public Function bsCore(OptStyle As EnmOptStyle, Forward, Strike, Volatility, Time)
1         On Error GoTo ErrHandler
          Dim dMinus As Double
          Dim dPlus As Double
          Dim VolRootT

2         If Not IsNumber(Forward) Then Throw "Forward must be a number"
3         If Not IsNumber(Strike) Then Throw "Strike must be a number"
4         If Not IsNumber(Volatility) Then Throw "Volatility must be a number"
5         If Not IsNumber(Time) Then Throw "Time must be a number"
6         If Time < 0 Then Throw "Time must be positive or zero"

7         Select Case OptStyle
              Case OptStyleBuy
8                 bsCore = Forward - Strike
9                 Exit Function
10            Case OptStyleSell
11                bsCore = Strike - Forward
12                Exit Function
13        End Select

14        If Time = 0 Then
15            Select Case OptStyle
                  Case OptStyleCall
16                    bsCore = IIf(Forward > Strike, Forward - Strike, 0)
17                Case OptStylePut
18                    bsCore = IIf(Forward > Strike, 0, Strike - Forward)
19                Case optStyleUpDigital
20                    bsCore = IIf(Forward > Strike, 1, 0)        '> or >=
21                Case optStyleDownDigital
22                    bsCore = IIf(Forward > Strike, 0, 1)        '> or >=
23                Case Else
24                    Throw "Unhandled OptionStyle"
25            End Select
26        Else
27            VolRootT = Volatility * Sqr(Time)
28            dPlus = Log(Forward / Strike) / VolRootT + 0.5 * VolRootT
29            dMinus = dPlus - VolRootT
30            Select Case OptStyle
                  Case OptStyleCall
31                    bsCore = func_normsdist(dPlus) * Forward - func_normsdist(dMinus) * Strike
32                Case OptStylePut
33                    bsCore = func_normsdist(-dMinus) * Strike - func_normsdist(-dPlus) * Forward
34                Case optStyleUpDigital
35                    bsCore = func_normsdist(dMinus)
36                Case optStyleDownDigital
37                    bsCore = func_normsdist(-dMinus)
38                Case Else
39                    Throw "Unhandled OptionStyle"
40            End Select
41        End If

42        Exit Function
ErrHandler:
43        bsCore = "#bsCore (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sAllocExposures
' Author    : Philip Swannell
' Date      : 09-Nov-2017
' Purpose   : This function is the allocator to sInterp and allocates risks to interpolated values back
'             to risks to the "grid-points" of the interpolation.
' Arguments
' xValues   : A column vector of numbers to which interpolation was previously performed (as the xValues
'             argument to sInterp).
' Exposures : A column vector of the exposures to the interpolated values.
' xArrayAscending: A column vector of the x-coordinates of the "grid-points" to which interpolation was
'             performed to yield "yValues" to which Exposures are the risks. Must be in
'             ascending order (lowest value at the top).
' InterpType: A string determining the interpolation scheme used. Allowed values: Linear, FlatFromLeft,
'             FlatToRight. If omitted defaults to Linear.
' Signature : A 2 character string to set left (1st character) and right (2nd) extrapolation beyond the
'             bounds of xArrayAscending.
'             F = Flat, X = Linear eXtrapolation, N = None. If omitted defaults from
'             InterpType: Linear > NN, FlatFromLeft > NF, FlatToRight > FN
' -----------------------------------------------------------------------------------------------------------------------
Function sAllocExposures(xValues, Exposures, xArrayAscending, Optional InterpType As String = "Linear", Optional Signature As String)
Attribute sAllocExposures.VB_Description = "This function is the allocator to sInterp and allocates risks to interpolated values back to risks to the ""grid-points"" of the interpolation."
Attribute sAllocExposures.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim i As Long
          Dim InterpRes
          Dim Result
          Const Error1 = "xValues and Exposures must be 1-column arrays of numbers with the the same number of elements in each of them"
          Const Error2 = "xArrayAscending must be 1-column array of numbers in ascending order"
          Dim NR As Long
          Dim NRxAA As Long

          'Check arguments
1         On Error GoTo ErrHandler
2         Force2DArrayRMulti xValues, Exposures, xArrayAscending
3         NR = sNRows(xValues)
4         If sNCols(xValues) <> 1 Or sNCols(Exposures) <> 1 Or sNRows(Exposures) <> NR Then Throw Error1
5         For i = 1 To NR
6             If Not (IsNumberOrDate(xValues(i, 1))) Then
7                 Throw "Found non-number at row " + CStr(i) + " of xValues"
8             ElseIf Not (IsNumberOrDate(Exposures(i, 1))) Then
9                 Throw "Found non-number at row " + CStr(i) + " of Exposures"
10            End If
11        Next i

12        If sNCols(xArrayAscending) <> 1 Then Throw Error2
13        NRxAA = sNRows(xArrayAscending)
14        For i = 1 To NRxAA
15            If Not (IsNumberOrDate(xArrayAscending(i, 1))) Then Throw Error2
16            If i > 1 Then If xArrayAscending(i, 1) <= xArrayAscending(i - 1, 1) Then Throw Error2
17        Next i

18        Result = sReshape(0, NRxAA, 1)
19        InterpRes = ThrowIfError(sInterp(xArrayAscending, sIntegers(NRxAA), xValues, InterpType, Signature))
          'Note that signature is passed by Reference to sInterp and may be modified (sensibly defaulted) by that function

          Dim IndxHi As Long
          Dim IndxLo As Long
          Dim LeftExtrap As Boolean
          Dim RightExtrap As Boolean
          Dim WeightHi As Double
          Dim WeightLo As Double
20        LeftExtrap = UCase$(Left$(Signature, 1)) = "X"
21        RightExtrap = UCase$(Right$(Signature, 1)) = "X"

22        For i = 1 To NR
23            ThrowIfError InterpRes(i, 1)    'will generate sensible error message if extrapolation is not allowed and _
                                               yet elements of xValues are outside the range of xArrayAscending
24            IndxLo = CLng(InterpRes(i, 1) - 0.5)
25            IndxHi = IndxLo + 1
26            If IndxLo = 0 And LeftExtrap Then
27                IndxLo = 1
28                IndxHi = 2
29            ElseIf IndxLo = NRxAA And RightExtrap Then
30                IndxLo = NRxAA - 1
31                IndxHi = NRxAA
32            End If
33            WeightLo = IndxHi - InterpRes(i, 1)
34            WeightHi = InterpRes(i, 1) - IndxLo
35            If WeightLo <> 0 Then Result(IndxLo, 1) = Result(IndxLo, 1) + WeightLo * Exposures(i, 1)
36            If WeightHi <> 0 Then Result(IndxHi, 1) = Result(IndxHi, 1) + WeightHi * Exposures(i, 1)
37        Next i
38        sAllocExposures = Result
39        Exit Function
ErrHandler:
40        sAllocExposures = "#sAllocExposures (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCholesky
' Author    : Philip Swannell
' Date      : 2-May-2015
' Purpose   : Returns the Cholesky decomposition of SymmetricMatrix. i.e. a matrix A such that A times A
'             transpose is equal to the input SymmetricMatrix. Useful within Monte Carlo
'             simulation to obtain random deviates with a desired correlation matrix.
' Arguments
' SymmetricMatrix: The matrix, which must be symmetric and positive definite.
'  VBA code adapted from http://vbadeveloper.net/numericalmethodsvbacholeskydecomposition.pdf
' -----------------------------------------------------------------------------------------------------------------------
Function sCholesky(SymmetricMatrix As Variant)
Attribute sCholesky.VB_Description = "Returns the Cholesky decomposition of SymmetricMatrix. i.e. a matrix A such that A times A transpose is equal to the input SymmetricMatrix. Useful within Monte Carlo simulation to obtain random deviates with a desired correlation matrix."
Attribute sCholesky.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim c As Long
          Dim E As Variant
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim N As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR SymmetricMatrix

          Dim Element As Double
          Dim L_Lower() As Double
3         N = sNRows(SymmetricMatrix)
4         c = sNCols(SymmetricMatrix)
5         If N <> c Then Throw "SymmetricMatrix must be square"

6         For Each E In SymmetricMatrix
7             If Not IsNumber(E) Then Throw "SymmetricMatrix must be numeric"
8         Next

9         For i = 1 To N
10            For j = 1 To i - 1
11                If SymmetricMatrix(i, j) <> SymmetricMatrix(j, i) Then Throw "SymmetricMatrix must be symmetric"
12            Next j
13        Next i

14        ReDim L_Lower(1 To N, 1 To N)
15        For i = 1 To N
16            For j = 1 To N
17                Element = SymmetricMatrix(i, j)
18                For k = 1 To i - 1
19                    Element = Element - L_Lower(i, k) * L_Lower(j, k)
20                Next k
21                If i = j Then
22                    If Element < 0 Then Throw "SymmetricMatrix must be positive definite"
23                    L_Lower(i, i) = Sqr(Element)
24                ElseIf i < j Then
25                    If L_Lower(i, i) = 0 Then Throw "SymmetricMatrix must be positive definite"
26                    L_Lower(j, i) = Element / L_Lower(i, i)
27                End If
28            Next j
29        Next i
30        sCholesky = L_Lower
31        Exit Function
ErrHandler:
32        sCholesky = "#sCholesky (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCorrelateTimeSeries
' Author    : Philip Swannell
' Date      : 02-May-2017
' Purpose   : Returns the correlation matrix for time series data. Each column is assumed to be either
'             Brownian or Geometric Brownian. For each pair of columns only rows where both
'             columns contain numbers are considered when calculating the correlation for
'             that pair
' Arguments
' TimeSeries: A set of multiple data series. Should be a matrix with one column for each series.
'
' ColumnTypeIndicators: TRUE to indicate all time series have normal increments (Brownian), FALSE to indicate that
'             all have log-normal increments (Geometric Brownian). If some series are
'             normal and others log-normal enter a row or column of logical values.
' UseMAD    : If TRUE, then a MAD estimate of correlation is returned. MAD (Median Absolute Deviation)
'             is a robust estimate immune to a proportion of "bad" data. If FALSE, then the
'             sample correlations are returned. See sCorrelationEstimate for more details.
' -----------------------------------------------------------------------------------------------------------------------
Function sCorrelateTimeSeries(TimeSeries As Variant, Optional ByVal ColumnTypeIndicators As Variant, Optional UseMAD As Boolean = False, Optional ReturnCovariance As Boolean)
Attribute sCorrelateTimeSeries.VB_Description = "Correlation matrix for time series. Each column may be either Brownian (normal increments), Geometric Brownian (log-normal increments) or ""raw"" normal data. Correlation between pairs of columns is calculated using only rows where both columns are numeric."
Attribute sCorrelateTimeSeries.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim ChooseVector As Variant
          Dim ChooseVectors As Variant
          Dim Columns() As Variant
          Dim ColumnsFirstDiff() As Variant
          Dim CTIs() As Long
          Dim i As Long
          Dim j As Long
          Dim Mads() As Double
          Dim Medians() As Double
          Dim N As Long
          Dim RhoMad() As Double
          Dim ScaledColumns() As Variant
          Dim StandardDeviations() As Double
          Dim tmp As Variant
          Const NIErr As String = "Invalid ColumnTypeIndicators. Must be a single value or an array (1 column or 1 row) of values with the same number of elements as TimeSeries has columns. Allowed values: TRUE (successive differences assumed normal) FALSE (successive ratios assumed normal) ""Raw"" (data itself assumed normal)"

          Dim Indicators() As Boolean

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti TimeSeries, ColumnTypeIndicators
3         N = sNCols(TimeSeries)

4         ReDim CTIs(1 To N) ' 1 = Normal increments, 2 = log normal increments, 3 = Raw
5         ReDim Medians(1 To N)
6         ReDim StandardDeviations(1 To N)
7         ReDim Mads(1 To N)
8         ReDim Columns(1 To N)
9         ReDim ColumnsFirstDiff(1 To N)
10        ReDim ScaledColumns(1 To N)
11        ReDim RhoMad(1 To N, 1 To N)
12        ReDim ChooseVectors(1 To N)
13        ReDim Indicators(1 To N)        ' Indicator(i) is True if non-numbers appear in column(i)

14        If sNRows(ColumnTypeIndicators) = 1 And sNCols(ColumnTypeIndicators) = 1 Then
15            For i = 1 To N
16                CTIs(i) = ValidateCTI(ColumnTypeIndicators(1, 1), NIErr)
17            Next i
18        ElseIf sNRows(ColumnTypeIndicators) = N And sNCols(ColumnTypeIndicators) = 1 Then
19            For i = 1 To N
20                CTIs(i) = ValidateCTI(ColumnTypeIndicators(i, 1), NIErr)
21            Next i
22        ElseIf sNRows(ColumnTypeIndicators) = 1 And sNCols(ColumnTypeIndicators) = N Then
23            For i = 1 To N
24                CTIs(i) = ValidateCTI(ColumnTypeIndicators(1, i), NIErr)
25            Next i
26        Else
27            Throw NIErr
28        End If

29        For i = 1 To N
30            Columns(i) = sSubArray(TimeSeries, 1, i, , 1)
31            ChooseVectors(i) = sArrayIsNumber(Columns(i))
32            Indicators(i) = IsNumber(sMatch(False, ChooseVectors(i)))
33            If Not Indicators(i) Then
34                Select Case CTIs(i)
                      Case 1
35                        ColumnsFirstDiff(i) = sFirstDifference(Columns(i))
36                    Case 2
37                        ColumnsFirstDiff(i) = sFirstRatio(Columns(i))
38                    Case 3
39                        ColumnsFirstDiff(i) = Columns(i)
40                End Select
41                If UseMAD Then
42                    Medians(i) = Application.WorksheetFunction.Median(ColumnsFirstDiff(i))
43                    Mads(i) = Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(ColumnsFirstDiff(i), Medians(i))))
44                    If ReturnCovariance Then
45                        StandardDeviations(i) = 1.4826022185056 * Mads(i)  ' 1/NORM.S.INV(0.75)
46                    End If
47                    ScaledColumns(i) = sArrayDivide(sArraySubtract(ColumnsFirstDiff(i), Medians(i)), Mads(i))
48                Else
49                    If ReturnCovariance Then
50                        StandardDeviations(i) = Application.WorksheetFunction.StDev_S(ColumnsFirstDiff(i))
51                    End If
52                End If
53            End If
54        Next

55        For i = 1 To N
56            For j = 1 To i - 1
57                If Indicators(i) Or Indicators(j) Then        'non numbers exist in one or both columns
58                    ChooseVector = sArrayAnd(ChooseVectors(i), ChooseVectors(j))
59                    If sArrayCount(ChooseVector) < 3 Then
                          'Throw "Cannot estimate correlation between column " + CStr(i) + " and column " + CStr(j) + " because there are not at least two rows for which both columns contain numbers"
60                        tmp = 0
61                    Else
                          Dim TempData1
                          Dim TempData2
62                        Select Case CTIs(i)
                              Case 1
63                                TempData1 = sFirstDifference(sMChoose(Columns(i), ChooseVector))
64                            Case 2
65                                TempData1 = sFirstRatio(sMChoose(Columns(i), ChooseVector))
66                            Case 3
67                                TempData1 = sMChoose(Columns(i), ChooseVector)
68                        End Select

69                        Select Case CTIs(j)
                              Case 1
70                                TempData2 = sFirstDifference(sMChoose(Columns(j), ChooseVector))
71                            Case 2
72                                TempData2 = sFirstRatio(sMChoose(Columns(j), ChooseVector))
73                            Case 3
74                                TempData2 = sMChoose(Columns(j), ChooseVector)
75                        End Select
76                        If UseMAD Then
77                            tmp = MADCorrelation(TempData1, TempData2)
78                            If VarType(tmp) = vbString Then        'Arrgh MADCorrelation fails e.g. Median Absolute Deviation of one of the two data series is not in the _
                                                                      slightest normal, very high number of repeats. Best we can do is take sample correlation
79                                tmp = Application.WorksheetFunction.Correl(TempData1, TempData2)
80                            End If
81                        Else
82                            tmp = Application.WorksheetFunction.Correl(TempData1, TempData2)
83                        End If

84                    End If
85                Else
86                    If UseMAD Then
87                        tmp = -100
88                        On Error Resume Next
89                        tmp = 0.5 * (Application.WorksheetFunction.Median(sArrayAbs(sArrayAdd(ScaledColumns(i), ScaledColumns(j)))) - _
                              Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(ScaledColumns(i), ScaledColumns(j)))))
90                        On Error GoTo ErrHandler
91                        If tmp <> -100 Then
92                            tmp = RhoFromRhoMAD(CDbl(tmp))
93                        Else
94                            tmp = Application.WorksheetFunction.Correl(ColumnsFirstDiff(i), ColumnsFirstDiff(j))
95                        End If
96                    Else
97                        tmp = Application.WorksheetFunction.Correl(ColumnsFirstDiff(i), ColumnsFirstDiff(j))
98                    End If
99                End If
100               RhoMad(j, i) = tmp
101               RhoMad(i, j) = tmp
102           Next j
103           RhoMad(i, i) = 1
104       Next i

          'It's inefficient to calculate the covariance by first calculating the correlation and then scaling, should revisit the code...
105       If ReturnCovariance Then
106           For i = 1 To N
107               For j = 1 To i
108                   RhoMad(i, j) = RhoMad(i, j) * StandardDeviations(i) * StandardDeviations(j)
109                   RhoMad(j, i) = RhoMad(i, j)
110               Next j
111           Next i
112       End If

113       sCorrelateTimeSeries = RhoMad

114       Exit Function
ErrHandler:
115       sCorrelateTimeSeries = "#sCorrelateTimeSeries (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ValidateCTI(x As Variant, ErrorString As String)
1         If VarType(x) = vbBoolean Then
2             ValidateCTI = IIf(x, 1, 2)
3         ElseIf VarType(x) = vbString Then
4             If LCase(x) = "raw" Then
5                 ValidateCTI = 3
6             Else
7                 Throw ErrorString
8             End If
9         Else
10            Throw ErrorString
11        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCorrelation
' Author    : Philip Swannell
' Date      : 04-Dec-2019
' Purpose   : Returns the correlation matrix of a set of data series. The output will be a square
'             correlation matrix whose size is equal to the number of columns in the input.
' Arguments
' DataSeries: A set of multiple data series. Should be a matrix with one column for each series.
'
' AllowNonNumbers: If TRUE correlation between column pairs takes into account only rows where both columns
'             contain numbers; and the return may not be positive semi definite. If FALSE
'             (the default) the function returns an error if there are non-numbers in
'             DataSeries.
' -----------------------------------------------------------------------------------------------------------------------
Function sCorrelation(DataSeries As Variant, Optional AllowNonNumbers As Boolean)
Attribute sCorrelation.VB_Description = "Returns the correlation matrix of a set of data series. The output will be a square correlation matrix whose size is equal to the number of columns in the input."
Attribute sCorrelation.VB_ProcData.VB_Invoke_Func = " \n23"
1         On Error GoTo ErrHandler
2         sCorrelation = sCorrelationOrCovariance(DataSeries, AllowNonNumbers, False)
3         Exit Function
ErrHandler:
4         sCorrelation = "#sCorrelation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCovariance
' Author    : Philip Swannell
' Date      : 04-Dec-2019
' Purpose   : Returns the covariance matrix of a set of data series. The output will be a square
'             covariance matrix whose size is equal to the number of columns in the input.
' Arguments
' DataSeries: A set of multiple data series. Should be a matrix with one column for each series.
'
' AllowNonNumbers: If TRUE covariance between column pairs takes into account only rows where both columns
'             contain numbers; and the return may not be positive semi definite. If FALSE
'             (the default) the function returns an error if there are non-numbers in
'             DataSeries.
' -----------------------------------------------------------------------------------------------------------------------
Function sCovariance(DataSeries As Variant, Optional AllowNonNumbers As Boolean)
Attribute sCovariance.VB_Description = "Returns the covariance matrix of a set of data series. The output will be a square covariance matrix whose size is equal to the number of columns in the input."
Attribute sCovariance.VB_ProcData.VB_Invoke_Func = " \n23"
1         On Error GoTo ErrHandler
2         sCovariance = sCorrelationOrCovariance(DataSeries, AllowNonNumbers, True)
3         Exit Function
ErrHandler:
4         sCovariance = "#sCovariance (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sCorrelationOrCovariance
' Author     : Philip Swannell
' Date       : 04-Dec-2019
' Purpose    : Implements both sCorrelation and sCovariance
' -----------------------------------------------------------------------------------------------------------------------
Private Function sCorrelationOrCovariance(DataSeries As Variant, AllowNonNumbers As Boolean, DoCov As Boolean)
          Dim Columns() As Variant
          Dim i As Long
          Dim j As Long
          Dim M() As Double
          Dim NC As Long
          Dim NR As Long
          Dim tmp() As Boolean
          Dim TmpCount As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR DataSeries

3         NC = sNCols(DataSeries)
4         NR = sNRows(DataSeries)

5         If NR < 2 Then Throw "DataSeries must have at least two rows"

6         If Not AllowNonNumbers Then
7             For i = 1 To NR
8                 For j = 1 To NC
9                     If Not IsNumberOrDate(DataSeries(i, j)) Then
10                        Throw "Non-number found in DataSeries at row " + CStr(i) + ", column " + CStr(j)
11                    End If
12                Next j
13            Next i
14        Else
              Dim ChooseVectors() As Variant
              Dim Counts() As Variant
15            ReDim ChooseVectors(1 To NC)
16            ReDim Counts(1 To NC)
17            For j = 1 To NC
18                ReDim tmp(1 To NR, 1 To 1)
19                TmpCount = 0
20                For i = 1 To NR
21                    If IsNumberOrDate(DataSeries(i, j)) Then
22                        tmp(i, 1) = True
23                        TmpCount = TmpCount + 1
24                    End If
25                Next i
26                If TmpCount < 2 Then Throw "There must be at least 2 numbers in each column, but column " + CStr(j) + " does not have two numbers"
27                ChooseVectors(j) = tmp
28                Counts(j) = TmpCount
29            Next j
30        End If

31        ReDim M(1 To NC, 1 To NC)
32        ReDim Columns(1 To NC)

33        For i = 1 To NC
34            Columns(i) = sSubArray(DataSeries, 1, i, , 1)
35        Next i

          Dim ChooseVector

36        For i = 1 To NC
37            For j = 1 To IIf(DoCov, i, i - 1)
38                If AllowNonNumbers Then
39                    If Counts(i) = NR And Counts(j) = NR Then
40                        M(i, j) = CorrelOrCov(Columns(i), Columns(j), DoCov)
41                    Else
42                        ChooseVector = sArrayAnd(ChooseVectors(i), ChooseVectors(j))
43                        If sArrayCount(ChooseVector) < 2 Then
44                            Throw "Cannot calculate " + IIf(DoCov, "covariance", "correlation") + " between column " + CStr(i) + " and column " + CStr(j) + " because there are not at least two rows for which both columns contain numbers"
45                        End If
46                        M(i, j) = CorrelOrCov(sMChoose(Columns(i), ChooseVector), sMChoose(Columns(j), ChooseVector), DoCov)
47                    End If
48                Else
49                    M(i, j) = CorrelOrCov(Columns(i), Columns(j), DoCov)
50                End If
51                M(j, i) = M(i, j)
52            Next j
53            If Not DoCov Then
54                M(i, i) = 1
55            End If
56        Next i
57        sCorrelationOrCovariance = M

58        Exit Function
ErrHandler:
59        Throw "#sCorrelationOrCovariance (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function CorrelOrCov(Array1, Array2, DoCov As Boolean)
1         On Error GoTo ErrHandler
2         If DoCov Then
3             CorrelOrCov = Application.WorksheetFunction.Covariance_P(Array1, Array2)
4         Else
5             CorrelOrCov = Application.WorksheetFunction.Correl(Array1, Array2)
6         End If
7         Exit Function
ErrHandler:
8         Throw "#CorrelOrCov (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCorrelationEstimate
' Author    : Philip Swannell
' Date      : 19-Jun-2016
' Purpose   : Returns a robust estimate of the correlation matrix of a set of data series.
' Arguments
' DataSeries: A set of multiple data series. Should be a matrix with one column for each series.
'
' AllowNonNumbers: If TRUE correlation estimation between each column pair takes into account only rows where
'             both columns contain numbers. If FALSE (the default) the function returns an
'             error if there are non-numbers in DataSeries.
'
' Notes     : The estimate is robust to the data series containing some small proportion of "bad" data
'             i.e. data from a different and unknown distribution. The function first
'             calculates the Median Absolute Deviation correlation coefficients (RhoMAD)
'             and then solves for Rho using the fact that for a bivariate normal with
'             correlation Rho, RhoMAD = SquareRoot((1+Rho)/2) - SquareRoot((1-Rho)/2).
'
'             If the attempt to calculate RhoMAD between two-columns of DataSeries fails
'             (e.g. in the presence of large numbers of repeated items) then the
'             corresponding element in the return from the function will simply be the
'             sample correlation between those columns.
'
'             Note that the return from this function is not guaranteed to be positive semi
'             definite, though in practice with real-world data sets it is usually so.
'
'             See page 519 of
'             Gideon, Rudy A. (2007) "The Correlation Coefficients," Journal of Modern
'             Applied Statistical Methods: Vol. 6: Iss. 2, Article 16.
'             Available at: http://digitalcommons.wayne.edu/jmasm/vol6/iss2/16
' -----------------------------------------------------------------------------------------------------------------------
Function sCorrelationEstimate(DataSeries As Variant, Optional AllowNonNumbers As Boolean)
Attribute sCorrelationEstimate.VB_Description = "Returns a robust estimate of the correlation matrix of a set of data series."
Attribute sCorrelationEstimate.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim ChooseVector As Variant
          Dim ChooseVectors As Variant
          Dim Columns() As Variant
          Dim i As Long
          Dim j As Long
          Dim Mads() As Double
          Dim Medians() As Double
          Dim N As Long
          Dim RhoMad() As Double
          Dim ScaledColumns() As Variant
          Dim tmp As Variant

          Dim Indicators() As Boolean

1         On Error GoTo ErrHandler
2         Force2DArrayR DataSeries
3         N = sNCols(DataSeries)

4         ReDim Medians(1 To N)
5         ReDim Mads(1 To N)
6         ReDim Columns(1 To N)
7         ReDim ScaledColumns(1 To N)
8         ReDim RhoMad(1 To N, 1 To N)
9         ReDim ChooseVectors(1 To N)
10        ReDim Indicators(1 To N)        ' Indicator(i) is True if non-numbers appear in column(i)

11        For i = 1 To N
12            Columns(i) = sSubArray(DataSeries, 1, i, , 1)
13            If AllowNonNumbers Then
14                ChooseVectors(i) = sArrayIsNumber(Columns(i))
15                Indicators(i) = IsNumber(sMatch(False, ChooseVectors(i)))
16            End If
17            If Not Indicators(i) Then
18                Medians(i) = Application.WorksheetFunction.Median(Columns(i))
19                Mads(i) = Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(Columns(i), Medians(i))))
20                ScaledColumns(i) = sArrayDivide(sArraySubtract(Columns(i), Medians(i)), Mads(i))
21            End If
22        Next

23        For i = 1 To N
24            For j = 1 To i - 1
25                If Indicators(i) Or Indicators(j) Then        'non numbers exist in one or both columns
26                    ChooseVector = sArrayAnd(ChooseVectors(i), ChooseVectors(j))
27                    If sArrayCount(ChooseVector) < 2 Then
                          'Throw "Cannot estimate correlation between column " + CStr(i) + " and column " + CStr(j) + " because there are not at least two rows for which both columns contain numbers"
28                        tmp = 0
29                    Else
30                        tmp = MADCorrelation(sMChoose(Columns(i), ChooseVector), sMChoose(Columns(j), ChooseVector))
31                        If VarType(tmp) = vbString Then        'Arrgh MADCorrelation fails e.g. Median Absolute Deviation of one of the two data series is not in the _
                                                                  slightest normal, very high number of repeats. Best we can do is take sample correlation
32                            tmp = Application.WorksheetFunction.Correl(sMChoose(Columns(i), ChooseVector), sMChoose(Columns(j), ChooseVector))
33                        End If
34                    End If
35                Else
36                    tmp = -100
37                    On Error Resume Next
38                    tmp = 0.5 * (Application.WorksheetFunction.Median(sArrayAbs(sArrayAdd(ScaledColumns(i), ScaledColumns(j)))) - _
                          Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(ScaledColumns(i), ScaledColumns(j)))))
39                    On Error GoTo ErrHandler
40                    If tmp <> -100 Then
41                        tmp = RhoFromRhoMAD(CDbl(tmp))
42                    Else
43                        tmp = Application.WorksheetFunction.Correl(Columns(i), Columns(j))
44                    End If
45                End If

46                RhoMad(j, i) = tmp
47                RhoMad(i, j) = tmp
48            Next j
49            RhoMad(i, i) = 1
50        Next i

51        sCorrelationEstimate = RhoMad

52        Exit Function
ErrHandler:
53        sCorrelationEstimate = "#sCorrelationEstimate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MADCorrelation
' Author    : Philip Swannell
' Date      : 21-Jun-2016
' Purpose   : Robust estimator of correlation between two data series that may have "bad" data
'             called by sCorrelationEstimate when non-numbers appear since in that case
'             sCorrelation cannot use its cached variables for the ScaledColumns
' -----------------------------------------------------------------------------------------------------------------------
Private Function MADCorrelation(Series1, Series2)
1         On Error GoTo ErrHandler
          Dim MAD1 As Double
          Dim MAD2 As Double
          Dim Median1 As Double
          Dim Median2 As Double
          Dim RhoMad As Double
          Dim ScaledSeries1 As Variant
          Dim ScaledSeries2 As Variant

2         Force2DArrayRMulti Series1, Series2

3         If sNCols(Series1) <> 1 Or sNCols(Series2) <> 1 Or sNRows(Series1) <> sNRows(Series2) Or sNRows(Series1) < 2 Then
4             Throw "Series1 and Series2 must both have one column and the same number of rows and more than one row"
5         End If

6         Median1 = Application.WorksheetFunction.Median(Series1)
7         Median2 = Application.WorksheetFunction.Median(Series2)
8         MAD1 = Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(Series1, Median1)))
9         MAD2 = Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(Series2, Median2)))

10        ScaledSeries1 = sArrayDivide(sArraySubtract(Series1, Median1), MAD1)
11        ScaledSeries2 = sArrayDivide(sArraySubtract(Series2, Median2), MAD2)

12        RhoMad = 0.5 * Application.WorksheetFunction.Median(sArrayAbs(sArrayAdd(ScaledSeries1, ScaledSeries2))) _
              - 0.5 * Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(ScaledSeries1, ScaledSeries2)))

13        MADCorrelation = RhoFromRhoMAD(RhoMad)

14        Exit Function
ErrHandler:
15        MADCorrelation = "#MADCorrelation (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RhoFromRhoMAD
' Author    : Philip Swannell
' Date      : 19-Jun-2016
' Purpose   : Where y = SquareRoot((1+x)/2) - SquareRoot((1-x)/2)
'             This method solves for x in terms of y via Newton-Raphson
' -----------------------------------------------------------------------------------------------------------------------
Private Function RhoFromRhoMAD(RhoMad As Double)
          Dim fpxn
          Dim fxn As Double
          Dim xn As Double
          Dim xnp1 As Double
          Const Epsilon As Double = 0.000000000000001
          Dim i As Long

1         On Error GoTo ErrHandler

2         If RhoMad <= -1 + Epsilon Then
3             RhoFromRhoMAD = -1
4             Exit Function
5         ElseIf RhoMad >= 1 - Epsilon Then
6             RhoFromRhoMAD = 1
7             Exit Function
8         End If

9         xn = RhoMad
TryAgain:
10        fxn = (Sqr((1 + xn) / 2) - Sqr((1 - xn) / 2)) - RhoMad
11        fpxn = 0.5 * (1 / Sqr(2 + 2 * xn) + 1 / Sqr(2 - 2 * xn))
12        xnp1 = xn - fxn / fpxn
13        If Abs(xnp1 - xn) < 2 * Epsilon Then
14            RhoFromRhoMAD = xnp1
15            Exit Function
16        End If
17        xn = SafeMax(SafeMin(xnp1, 1 - Epsilon), Epsilon - 1)
18        i = i + 1
19        If i > 20 Then Throw "Failed to converge" 'I don't think that can ever happen. PGS 18-July-19
20        GoTo TryAgain

21        Exit Function
ErrHandler:
22        RhoFromRhoMAD = "#RhoFromRhoMAD (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
