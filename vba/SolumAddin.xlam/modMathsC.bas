Attribute VB_Name = "modMathsC"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sKendallTau
' Author    : Philip Swannell
' Date      : 12-Dec-2017
' Purpose   : Calculates Kendall rank correlation coefficient between two data series, with optional
'             conversion to Pearson correlation (assuming bi-variate normal).
'
'             Element i,j of the return is the Kendall Tau between column i of Data1 and
'             column j of Data2.
' Arguments
' Data1     : An array of numbers, each column being regarded as a separate data series. May contain
'             non-numbers, which are ignored for calculation of Kendall Tau.
' Data2     : An optional second array of numbers. If omitted, then Data2 is treated as being the same
'             as Data1.
' ConvertToPearson: If FALSE, then the Kendall Tau is returned, otherwise, If TRUE, then the equivalent
'             Pearson correlation is returned.
'
'             Tau = sin(p t / 2)
'
' Notes     : References:
'             https://en.m.wikipedia.org/wiki/Kendall_rank_correlation_coefficient
'
'             https://stats.stackexchange.com/questions/133460/proof-of-the-relation-between-kendalls-tau-and-pearsons-rho-for-the-gaussian-c
'
'             See also:
'             sCorrelationEstimate
' -----------------------------------------------------------------------------------------------------------------------
Function sKendallTau(Data1, Optional Data2, Optional ConvertToPearson As Boolean)
Attribute sKendallTau.VB_Description = "Calculates Kendall rank correlation coefficient between two data series, with optional conversion to Pearson correlation (assuming bi-variate normal).\n\nElement i,j of the return is the Kendall Tau between column i of Data1 and column j of Data2."
Attribute sKendallTau.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim i As Long
          Dim j As Long
          Dim NA
          Dim NC As Long
          Dim NR As Long
          Dim Result
          Static HaveCalledAlready As Boolean

1         On Error GoTo ErrHandler

2         If Not HaveCalledAlready Then
3             CheckR "sKendallTauR", gPackagesSAI, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()"
4             HaveCalledAlready = True
5         End If

          'Work around current bug in BERT2 see
          'https://github.com/sdllc/Basic-Excel-R-Toolkit/issues/154
6         If IsMissing(Data2) Or IsEmpty(Data2) Then
7             Data2 = CVErr(xlErrNA)
8         End If

9         Result = ThrowIfError(Application.Run("BERT.Call", "KendallTau", Data1, Data2, ConvertToPearson))
10        Force2DArray Result
11        NR = sNRows(Result): NC = sNCols(Result)

12        NA = CVErr(xlErrNA)
13        For i = 1 To NR
14            For j = 1 To NC
15                If Not (IsNumeric(CStr(Result(i, j)))) Then        'handle what happens when an entire column of the input data is non-numeric, R's NA becomes a "1.#QNAN" here in VBA...
16                    Result(i, j) = NA
17                End If
18            Next
19        Next

20        sKendallTau = Result
21        Exit Function
ErrHandler:
22        sKendallTau = "#sKendallTau (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sKendallTauOLD - retained only for testing sKendallTau
' Author    : Philip
' Date      : 27-Jun-2017
' Purpose   : Calculates Kendall Tau for a pair of data series. Naive algorithm, order N-squared
' Can convert to Pearson (standard) correlation under the assumption of bivariate normal
' See https://stats.stackexchange.com/questions/133460/proof-of-the-relation-between-kendalls-tau-and-pearsons-rho-for-the-gaussian-c
' https://en.m.wikipedia.org/wiki/Kendall_rank_correlation_coefficient
' Knight's faster algotrithm is described at http://adereth.github.io/blog/2013/10/30/efficiently-computing-kendalls-tau/
' -----------------------------------------------------------------------------------------------------------------------
Function sKendallTauOLD(ByVal Series1, ByVal Series2, ConvertToPearson As Boolean, AllowNonNumbers As Boolean)
1         On Error GoTo ErrHandler
2         Force2DArrayRMulti Series1, Series2
          Dim Denom As Double
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim N As Long
          Dim NR As Long
          Dim Numer As Long
          Dim Tau As Double
          Dim Ties1 As Long
          Dim Ties2 As Long
          Const PI = 3.14159265358979
          Dim DoThis As Boolean

3         NR = sNRows(Series1)

4         If sNCols(Series1) <> 1 Then
5             Throw "Series1 must have one column"
6         ElseIf sNCols(Series1) <> 1 Then
7             Throw "Series2 must have one column"
8         ElseIf sNRows(Series2) <> NR Then
9             Throw "Series1 and Series2 must have the same number of rows"
10        End If

11        If AllowNonNumbers Then
12            For i = 1 To NR
13                If IsNumber(Series1(i, 1)) Then
14                    If IsNumber(Series2(i, 1)) Then
15                        N = N + 1
16                    End If
17                End If
18            Next i
19        Else
20            N = NR
21            For i = 1 To NR
22                If Not IsNumber(Series1(i, 1)) Then
23                    Throw "Non-number found at row " + CStr(i) + " of Series1 consider setting AllowNonNumbers to TRUE"
24                ElseIf Not IsNumber(Series2(i, 1)) Then
25                    Throw "Non-number found at row " + CStr(i) + " of Series2 consider setting AllowNonNumbers to TRUE"
26                End If
27            Next i
28        End If

29        DoThis = True
30        For i = 2 To NR
31            For j = 1 To i - 1
32                If AllowNonNumbers Then
33                    DoThis = IsNumber(Series1(i, 1)) And IsNumber(Series2(i, 1)) And IsNumber(Series1(j, 1)) And IsNumber(Series2(j, 1))
34                End If
35                If DoThis Then
36                    k = Sgn(Series1(i, 1) - Series1(j, 1)) * Sgn(Series2(i, 1) - Series2(j, 1))
37                    Numer = Numer + k
38                    If k = 0 Then
39                        If Series1(i, 1) = Series1(j, 1) Then
40                            Ties1 = Ties1 + 1
41                        End If
42                        If Series2(i, 1) = Series2(j, 1) Then
43                            Ties2 = Ties2 + 1
44                        End If
45                    End If
46                End If

47            Next j
48        Next i

49        Denom = Sqr(((N * (N - 1) / 2) - Ties1) * ((N * (N - 1) / 2) - Ties2))

50        Tau = Numer / Denom
51        If ConvertToPearson Then
52            sKendallTauOLD = Sin(Tau * PI / 2)
53        Else
54            sKendallTauOLD = Tau
55        End If
56        Exit Function
ErrHandler:
57        sKendallTauOLD = "#sKendallTauOLD (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: SmoothTransition
' Purpose:    Defines a smooth function (infinitely differentiable) that has value y1 at x<=x1, y2 at x>= x2 and smooth in between
' See https://en.wikipedia.org/wiki/Non-analytic_smooth_function
' Author: Philip Swannell
' Date: 23-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Function SmoothTransition(ByVal x As Double, Optional x1 As Double = 0, Optional X2 As Double = 1, Optional y1 As Double = 0, Optional y2 As Double = 1)
          Dim fx As Double
          Dim Res As Double
1         On Error GoTo ErrHandler
2         x = (x - x1) / (X2 - x1)
3         fx = F(x)
4         Res = F(x) / (fx + F(1 - x))
5         SmoothTransition = y1 + (y2 - y1) * Res
6         Exit Function
ErrHandler:
7         Throw "#SmoothTransition (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function F(x As Double)
1         On Error GoTo ErrHandler
2         If x <= 0 Then
3             F = 0
4         Else
5             F = Exp(-1 / x)
6         End If
7         Exit Function
ErrHandler:
8         Throw "#f (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRandomSetSeed
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Sets the seed of one of the random number generators used by sRandomVariable.
' Arguments
' GeneratorName: Allowed values:
'             Wichmann-Hill
'             Mersenne-Twister
'             Note that the VBA-Rnd random number generator cannot be seeded.
' Seed      : Any Long (i.e. a whole number  between -2,147,483,648 and 2,147,483,647). Can also be
'             omitted, in which case a value based on the current system time will be used
'             instead.
' -----------------------------------------------------------------------------------------------------------------------
Function sRandomSetSeed(ByVal GeneratorName As String, Optional ByVal Seed As Variant)
Attribute sRandomSetSeed.VB_Description = "Sets the seed of one of the random number generators used by sRandomVariable."
Attribute sRandomSetSeed.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim CorrectedGenName As String
          Const SeedError = "Seed must be a Long (whole number between -2,147,483,648 and 2,147,483,647) or omitted to use a seed based on the current system time."
          Const MaxLong = 2147483647
          Const Minlong = -2147483648#

1         On Error GoTo ErrHandler
2         GeneratorName = Replace(LCase$(GeneratorName), " ", vbNullString)

3         If IsMissing(Seed) Or IsEmpty(Seed) Then Seed = CLng(Timer * 60)
4         If Not IsNumber(Seed) Then Throw SeedError
5         If Seed > MaxLong Or Seed < Minlong Then Throw SeedError
6         If Seed <> CLng(Seed) Then Throw SeedError

7         On Error Resume Next
8         Seed = CLng(Seed)
9         On Error GoTo ErrHandler

10        If VarType(Seed) <> vbLong Then Throw SeedError

11        Select Case GeneratorName
              Case "wichmann-hill", "wichmannhill"
12                CorrectedGenName = "Wichmann-Hill"
13                RndM Seed
14            Case "mersennetwister", "mersenne-twister"
15                CorrectedGenName = "Mersenne-Twister"
16                init_genrand Seed
17            Case "vbarnd", "vba-rnd"
18                CorrectedGenName = "VBA-Rnd"
                  'https://msdn.microsoft.com/en-us/library/office/gg264511(v=office.15).aspx _
                   "To repeat sequences of random numbers, call Rnd with a negative argument immediately _
                   before using Randomize with a numeric argument. Using Randomize with the same value for _
                   number does not repeat the previous sequence" BUT my use of Rnd with negative argument _
                   does not seem to have desired effect. :-(
19                Rnd -1
20                Randomize Seed
21            Case Else
22                Throw "Unrecognised GeneratorName - allowed names are: Wichmann-Hill, Mersenne-Twister or VBA-Rnd"
23        End Select
24        sRandomSetSeed = CorrectedGenName + " RNG seeded with '" + Format$(Seed, "###,##0") + "' at " + Format$(Now, "hh:mm:ss")

25        Exit Function
ErrHandler:
26        sRandomSetSeed = "#sRandomSetSeed (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRandomVariable
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Simulates random variables from a given distribution. Returns an array of IID random
'             variables from either:
'             Uniform: Uniform on the interval [0,1]
'             Normal: Normal with zero mean and unit variance.
'             Integer: Integer uniform on the set {1,2,...,Arg_1}
' Arguments
' NumRows   : The number of rows in the return. In the case of Sobol or Shifted Sobol sequences this
'             number should be 2^n - 1 for some n to avoid bias (non-zero mean) in the
'             return.
' NumCols   : The number of columns in the return. If omitted defaults to 1.
' DistributionName: Allowed values:
'             Uniform: Uniform distribution on the interval [0,1]
'             Normal: Normal distribution with zero mean and unit variance.
'             Integer: Integer uniform distribution on the set {1,2,...,Arg_1}
' GeneratorName: Allowed values: Wichmann-Hill, Mersenne-Twister, VBA-Rnd, Sobol, Shifted-Sobol
' Arg_1     : Gives the upper bound of the distribution when DistributionName = Integer
'
' Notes     : The user can choose the Random Number Generator (RNG) employed:
'             Wichmann-Hill
'             A compromise RNG, that may be good enough for most purposes. It produces
'             numbers at a rate of 6.3 million per second.
'             Mersenne-Twister
'             A very good RNG but slow when implemented in VBA, about 830,000 numbers per
'             second.
'             VBA-Rnd
'             This is the RNG built in to VBA's Rnd function. It's fast but with a short
'             period of only 16,777,216. Produces numbers at a rate of 15.8 million per
'             second - so takes less than two seconds to "get back to where it started."
'             Sobol or Shifted-Sobol
'             Return is the NumCols-dimensional Sobol sequence. Note that NumRows should be
'             2^n -1 for some n (e.g. 2^10 - 1 = 1023) in order that the sequences are
'             unbiased (have mean 0). In the case of Shifted-Sobol the return is the sum of
'             a standard Sobol sequence and a random add-on.
'
'             Speed test data as of Nov 2018 on Intel i7-6700 CPU @ 3.40GHz.
' -----------------------------------------------------------------------------------------------------------------------
Function sRandomVariable(NumRows As Long, Optional NumCols As Long = 1, _
        Optional ByVal DistributionName As String = "Uniform", _
        Optional ByVal GeneratorName As String = "Wichmann-Hill", _
        Optional ByVal Arg_1 As Double, _
        Optional ByVal Seed As Variant)
Attribute sRandomVariable.VB_Description = "Simulates random variables from a given distribution. Returns an array of IID random variables from either:\nUniform: Uniform on the interval [0,1]\nNormal: Normal with zero mean and unit variance.\nInteger: Integer uniform on the set {1,2,...,Arg_1}"
Attribute sRandomVariable.VB_ProcData.VB_Invoke_Func = " \n23"

1         Application.Volatile
          Dim FlipToInteger As Boolean
          Dim FlipToNormal As Boolean
          Dim i As Long
          Dim j As Long
          Dim Pow2 As Long
          Dim Result() As Double
          Dim SobolShift As Boolean
          Dim UseMersenne As Boolean
          Dim UseRnd As Boolean
          Dim UseSobol As Boolean
          Dim UseWH As Boolean

2         On Error GoTo ErrHandler

3         DistributionName = LCase$(DistributionName)
4         GeneratorName = Replace(LCase$(GeneratorName), " ", vbNullString)

5         Select Case GeneratorName
              Case "sobol"
6                 UseSobol = True
7                 SobolShift = False
8             Case "shifted-sobol", "shiftedsobol"
9                 UseSobol = True
10                SobolShift = True
11            Case "wichmann-hill", "wichmannhill"
12                UseWH = True        '   Period 6,953,607,871,644
13            Case "mersennetwister", "mersenne-twister"        'period 2^19937 - 1 and equidistribution in 623 consecutive dimensions
14                UseMersenne = True
15            Case "vbarnd", "vba-rnd"
16                UseRnd = True        ' Not recommended! Has period of 2^24 = 16,777,216
17            Case Else
18                Throw "Unrecognised GeneratorName - allowed names are: Wichmann-Hill, Mersenne-Twister, VBA-Rnd, Sobol and Shifted-Sobol"
19        End Select

20        If Not IsMissing(Seed) Then
21            ThrowIfError sRandomSetSeed(GeneratorName, Seed)
22        End If

23        Select Case LCase$(DistributionName)
              Case "uniform"
24                FlipToNormal = False
25            Case "normal"
26                FlipToNormal = True
27            Case "integer"
28                FlipToInteger = True
29                Arg_1 = CLng(Arg_1)
30                If Arg_1 <= 0 Then Throw "Arg_1 must be a positive integer"
31            Case Else
32                Throw "Unrecognised DistributionName - allowed names are: Uniform, Normal, Integer"
33        End Select

34        If NumRows < 1 Then Throw "NumRows must be positive"
35        If NumCols < 1 Then Throw "NumCols must be positive"
36        ReDim Result(1 To NumRows, 1 To NumCols)

37        If UseSobol Then
38            If NumCols > 500 Then Throw "Sobol sequences cannot exceed 500 dimensions"
39            Pow2 = Log(NumRows + 1) / Log(2)
40            If 2 ^ Pow2 - 1 <> NumRows Then Throw "Sobol sequence must be generated for 2^n -1 sample for some n"

              Dim Sobol As clsSobol
41            Set Sobol = New clsSobol
42            Sobol.SetData NumCols, NumRows, SobolShift
43            Result = Sobol.GetSobolSequence()
44            Set Sobol = Nothing
45            If FlipToNormal Then
46                For i = 1 To NumRows
47                    For j = 1 To NumCols
48                        Result(i, j) = func_normsinv(Result(i, j))
49                    Next j
50                Next i
51            End If
52        ElseIf UseWH And FlipToNormal Then
53            For i = 1 To NumRows
54                For j = 1 To NumCols
55                    Result(i, j) = func_normsinv(RndM())
56                Next j
57            Next i
58        ElseIf UseWH And Not FlipToNormal Then
59            For i = 1 To NumRows
60                For j = 1 To NumCols
61                    Result(i, j) = RndM()
62                Next j
63            Next i
64        ElseIf UseRnd And FlipToNormal Then
65            For i = 1 To NumRows
66                For j = 1 To NumCols
67                    Result(i, j) = func_normsinv(Rnd())
68                Next j
69            Next i
70        ElseIf UseRnd And Not FlipToNormal Then
71            For i = 1 To NumRows
72                For j = 1 To NumCols
73                    Result(i, j) = Rnd()
74                Next j
75            Next i
76        ElseIf UseMersenne And FlipToNormal Then
77            For i = 1 To NumRows
78                For j = 1 To NumCols
79                    Result(i, j) = func_normsinv(genrand_real3())        '(0.0, 1.0) = [0.0000000001164..., 0.9999999998836...] (both 0.0 and 1.0 excluded)
80                Next j
81            Next i
82        ElseIf UseMersenne And Not FlipToNormal Then
83            For i = 1 To NumRows
84                For j = 1 To NumCols
85                    Result(i, j) = genrand_real1()        '[0.0, 1.0]   (both 0.0 and 1.0 included)
86                Next j
87            Next i
88        End If

89        If FlipToInteger Then
90            For i = 1 To NumRows
91                For j = 1 To NumCols
92                    Result(i, j) = CLng(Arg_1 * Result(i, j) + 0.5)
93                Next j
94            Next i
95        End If

96        sRandomVariable = Result

97        Exit Function
ErrHandler:
98        sRandomVariable = "#sRandomVariable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RndM
' Author    : Philip Swannell
' Date      : 27-Apr-2015
' Purpose   : A random number generator better than Rnd, that has period 16,777,216
' This is the Wichmann-Hill RNG that
' Copied from http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' See also http://www.pages.drexel.edu/~bdm25/excel-rng.pdf
' -----------------------------------------------------------------------------------------------------------------------
Private Function RndM(Optional ByVal Number As Long) As Double
          Dim dblRnd As Double
          Static blnInit As Boolean
          Static lngX As Long
          Static lngY As Long
          Static lngZ As Long
          ' if initialized and no input number given
1         If blnInit And Number = 0 Then
              ' lngX, lngY and lngZ will never be 0
2             lngX = (171 * lngX) Mod 30269
3             lngY = (172 * lngY) Mod 30307
4             lngZ = (170 * lngZ) Mod 30323
5         Else
              ' if no initialization, use Timer, otherwise ensure positive Number
6             If Number = 0 Then Number = Timer * 60 Else Number = Number And &H7FFFFFFF
7             lngX = (Number Mod 30269)
8             lngY = (Number Mod 30307)
9             lngZ = (Number Mod 30323)
              ' lngX, lngY and lngZ must be bigger than 0
10            If lngX > 0 Then Else lngX = 171
11            If lngY > 0 Then Else lngY = 172
12            If lngZ > 0 Then Else lngZ = 170
              ' mark initialization state
13            blnInit = True
14        End If
          ' generate a random number
15        dblRnd = CDbl(lngX) / 30269# + CDbl(lngY) / 30307# + CDbl(lngZ) / 30323#
          ' return a value between 0 and 1
16        RndM = dblRnd - Int(dblRnd)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestsRandomVariable
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : To test relative speed of the three RNGs we have to hand...
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestsRandomVariable()
          Const NumRows = 1000000
          Const NumCols = 1
          Dim Res1
          Dim Res2
          Dim res3
          Dim t1 As Double
          Dim t2 As Double
          Dim t3 As Double
          Dim t4 As Double

1         t1 = sElapsedTime()
2         Res1 = sRandomVariable(NumRows, NumCols, "Uniform", "VBA-Rnd")
3         t2 = sElapsedTime()
4         Res2 = sRandomVariable(NumRows, NumCols, "Uniform", "Wichmann-Hill")
5         t3 = sElapsedTime()
6         res3 = sRandomVariable(NumRows, NumCols, "Uniform", "Mersenne-Twister")
7         t4 = sElapsedTime()

8         ThrowIfError Res1
9         ThrowIfError Res2
10        ThrowIfError res3

11        Debug.Print "Miliseconds for " + Format$(NumRows * NumCols, "###,###") + " random variables:"
12        Debug.Print "         VBA-Rnd:", (t2 - t1) * 1000
13        Debug.Print "   Wichmann-Hill: ", (t3 - t2) * 1000
14        Debug.Print "Mersenne-Twister: ", (t4 - t3) * 1000
15        Debug.Print "Number of random variables generated per second"
16        Debug.Print "         VBA-Rnd:", Format$(NumRows * NumCols / (t2 - t1), "###,###")
17        Debug.Print "   Wichmann-Hill: ", Format$(NumRows * NumCols / (t3 - t2), "###,###")
18        Debug.Print "Mersenne-Twister: ", Format$(NumRows * NumCols / (t4 - t3), "###,###")
19        Debug.Print String(100, "=")
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sTimeSeriesVol
' Author    : Philip Swannell
' Date      : 15-Jun-2017
' Purpose   : Returns an estimate of the annualised volatility of time series data.
' Arguments
' TimeSeries: A 1-column array of data
' IsNormal  : TRUE if the time series has normal increments, FALSE if it has log-normal increments.
' Frequency : A string to describe the frequency of the data. Allowed values are: Annual, Monthly,
'             Weekly, Weekday, Daily.
' IgnoreNonNumbers: If TRUE, then non-numbers in TimeSeries are ignored. If FALSE, then non-numbers will cause
'             an error to be returned.
' UseMAD    : If TRUE the function returns  a robust estimate calculated from the Median Absolute
'             Deviation of the returns.
'             Sigma = k.MAD
'             k = 1/(InvNormal(0.75))
'             If FALSE or omitted the functions return is calculated via a call to the
'             Excel function STDEV.S
'
' Notes     : For a discussion of MAD estimation of sample standard deviation see
'             https://en.wikipedia.org/wiki/Median_absolute_deviation
' -----------------------------------------------------------------------------------------------------------------------
Function sTimeSeriesVol(ByVal TimeSeries, IsNormal As Boolean, Frequency As String, IgnoreNonNumbers As Boolean, Optional UseMAD As Boolean)
Attribute sTimeSeriesVol.VB_Description = "Returns an estimate of the annualised volatility of time series data."
Attribute sTimeSeriesVol.VB_ProcData.VB_Invoke_Func = " \n23"

          Dim ChooseVector
          Dim Deltas() As Double
          Dim i As Long
          Dim Multiplier As Double
          Dim N As Long
          Dim NumGood As Long

1         On Error GoTo ErrHandler
2         Select Case LCase$(Frequency)
              Case "annual", "year"
3                 Multiplier = 1
4             Case "monthly", "month"
5                 Multiplier = Sqr(12)
6             Case "weekly", "week"
7                 Multiplier = Sqr(365 / 7)
8             Case "weekday", "weekdays", "weekdaily"
9                 Multiplier = Sqr(365 * 5 / 7)
10            Case "daily", "day"
11                Multiplier = Sqr(365)
12            Case Else
13                Throw "Frequency must be a string. Allowed values: Annual, Monthly, Weekly, Weekday, Daily"
14        End Select

15        Force2DArrayR TimeSeries
16        If sNCols(TimeSeries) <> 1 Then Throw "TimeSeries must be a single column of data"
17        N = sNRows(TimeSeries)

18        ChooseVector = sArrayIsNumber(TimeSeries)
19        NumGood = sArrayCount(ChooseVector)

20        If Not IgnoreNonNumbers Then
21            If NumGood <> N Then
22                Throw "Non numbers found in TimeSeries. You may want to set IgnoreNonNumbers to TRUE"
23            End If
24        End If

25        If NumGood < 3 Then Throw "Not enough data in TimeSeries"

26        If NumGood < N Then
27            TimeSeries = sMChoose(TimeSeries, ChooseVector)
28        End If

29        ReDim Deltas(1 To NumGood - 1, 1 To 1)

30        If IsNormal Then
31            For i = 1 To NumGood - 1
32                Deltas(i, 1) = TimeSeries(i + 1, 1) - TimeSeries(i, 1)
33            Next i
34        Else
35            For i = 1 To NumGood - 1
36                Deltas(i, 1) = TimeSeries(i + 1, 1) / TimeSeries(i, 1)
37            Next i
38        End If
39        If UseMAD Then    'https://en.wikipedia.org/wiki/Median_absolute_deviation
              Dim MAD As Double
              Dim Median As Double
40            Median = Application.WorksheetFunction.Median(Deltas)
41            MAD = Application.WorksheetFunction.Median(sArrayAbs(sArraySubtract(Deltas, Median)))
42            sTimeSeriesVol = 1.4826022185056 * MAD * Multiplier ' 1/NORM.S.INV(0.75)
43        Else
44            sTimeSeriesVol = Application.WorksheetFunction.StDev_S(Deltas) * Multiplier
45        End If

46        Exit Function
ErrHandler:
47        sTimeSeriesVol = "#sTimeSeriesVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub TestNumberOfArguments()
          Dim Res1 As Variant
          Dim Res2
          'These work. In both cases adding one more argument would fail
1         Res1 = Application.Run("BERT.Call", "sum", 1, 2, 3, 4, 5, 6, 7, 8)
2         Res2 = Application.Run("R.CallFn", "sum", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
End Sub
