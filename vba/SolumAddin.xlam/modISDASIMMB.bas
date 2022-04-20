Attribute VB_Name = "modISDASIMMB"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMImputeMissingData
' Author     : Philip Swannell
' Date       : 29-Nov-2018
' Purpose    : Implements data imputation according to one of seven methodologies, for details of the methodologies see
'              "Notes on Calibration Data Preparation Methodology v2.pdf" by Philip Swannell dated 28 Nov 2018
' Parameters :
'  DataWithMissings      : Single column array of numbers or non-numbers (all non-numbers treated as "missings")
'  MedianFirstDifferences: A single column array of the medians. Should have one fewer row than argument DataWithMissings
'  Methodology           : String allowed values "None", "ISDA", "ISDA 2", "ISDA 3", "ISDA 4", "Linear", "Flat"
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMImputeMissingData(DataWithMissings, MedianFirstDifferences, Methodology As String)

          Dim ActualSum As Double
          Dim CRRes As Variant
          Dim DoThisChunk As Boolean
          Dim EN As Long
          Dim FirstDiffDWM
          Dim FirstDiffImputed
          Dim i As Long
          Dim Intercept As Double
          Dim j As Long
          Dim MAXGAP As Long
          Dim Res As Variant
          Dim Result
          Dim Slope As Double
          Dim TargetSum As Double
          Const AllowedMethods = "None,ISDA,ISDA 2,ISDA 3,ISDA 3b,ISDA 4,Linear,Flat"
          
1         On Error GoTo ErrHandler

2         Select Case Methodology
              Case "None" 'No imputation
3                 ISDASIMMImputeMissingData = DataWithMissings
4                 Exit Function
5             Case "ISDA", "ISDA 2", "ISDA 3", "ISDA 4", "ISDA 3b"
6                 If Methodology = "ISDA" Or Methodology = "ISDA 3b" Then
7                     MAXGAP = 10
8                 Else
9                     MAXGAP = 1000000#
10                End If
                  'FirstDiffDWM will contain the first differences of DataWithMissings
11                FirstDiffDWM = ThrowIfError(sDifference(DataWithMissings, 1, 1))
12                If Methodology = "ISDA" Or Methodology = "ISDA 2" Or Methodology = "ISDA 3" Or Methodology = "ISDA 3b" Then
                      'Calculate regression coefficients
13                    Res = ThrowIfError(LINESTWrap(FirstDiffDWM, MedianFirstDifferences, True, False))
14                    Slope = Res(1)
15                    Intercept = Res(2)
16                ElseIf Methodology = "ISDA 4" Then
17                    Slope = 1
18                    Intercept = 0
19                End If

                  'Functions ArrayAdd, ArrayMultiply do vector processing - like R, Julia or Python
20                FirstDiffImputed = sArrayAdd(sArrayMultiply(MedianFirstDifferences, Slope), Intercept)
          
21                Result = DataWithMissings
                  'NB understanding this function is easy once you understand what the function CountRepeats does!
                  'CRRes will have four columns and describes the "chunks of the Boolean vector that is the result of ArrayIsNumber(DataWithMissings)
                  'First column is the Boolean
                  'Second gives row number that chunk starts
                  'Third gives row number that chunk ends
                  'Fourth gives size of chunk (equal to 3rd-2nd+1)
22                CRRes = sCountRepeats(sArrayIsNumber(FirstDiffDWM), "CFTH")
23                For i = 1 To sNRows(CRRes)
24                    If CRRes(i, 1) = False Then 'i.e. this is a chunk of non-numeric data
25                        If CRRes(i, 4) <= MAXGAP + 1 Then '10 missing days data means 11 non-numeric 1-day differences. Note this line edited by PGS 20-Feb-19, previously read 'CRRes(i, 4) < MAXGAP + 1'
26                            DoThisChunk = True
27                            ActualSum = 0
28                            For j = CRRes(i, 2) To CRRes(i, 3)
29                                If IsNumber(FirstDiffImputed(j, 1)) Then
30                                    ActualSum = ActualSum + FirstDiffImputed(j, 1)
31                                Else
32                                    DoThisChunk = False
33                                    Exit For
34                                End If
35                            Next j
36                            If DoThisChunk Then
37                                On Error Resume Next
38                                TargetSum = DataWithMissings(CRRes(i, 3) + 1, 1) - DataWithMissings(CRRes(i, 2), 1)
39                                EN = Err.Number
40                                On Error GoTo ErrHandler
41                                If EN = 0 Then
42                                    For j = CRRes(i, 2) To CRRes(i, 3) - 1
43                                        Select Case Methodology
                                              Case "ISDA", "ISDA 2"
44                                                Result(j + 1, 1) = Result(j, 1) + FirstDiffImputed(j, 1) 'As far as I can tell, this is what ISDA do
45                                            Case "ISDA 3", "ISDA 4", "ISDA 3b"
46                                                Result(j + 1, 1) = Result(j, 1) + FirstDiffImputed(j, 1) + (TargetSum - ActualSum) / (CRRes(i, 4))
47                                        End Select
48                                    Next
49                                End If
50                            End If
51                        End If
52                    End If
53                Next

54                ISDASIMMImputeMissingData = Result

55            Case "Linear"
56                ISDASIMMImputeMissingData = ThrowIfError(LinearInfill(DataWithMissings))
57            Case "Flat"  'Flat interpolation of missing values
58                ISDASIMMImputeMissingData = ThrowIfError(FlatInfill(DataWithMissings))
59            Case Else
60                Throw "Unrecognised Methodology, Methodology must be one of: " + "'" + Replace(AllowedMethods, ",", "', '") + "'"
61        End Select

62        Exit Function
ErrHandler:
63        ISDASIMMImputeMissingData = "#ISDASIMMImputeMissingData (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LINESTWrap
' Author     : Philip Swannell
' Date       : 14-Nov-2018
' Purpose    : Wraps Excel worksheet function LINEST to ignore non-numeric values
' -----------------------------------------------------------------------------------------------------------------------
Private Function LINESTWrap(ByVal known_ys, ByVal known_xs, Optional b_const As Boolean = True, Optional stats As Boolean)

1         On Error GoTo ErrHandler
          
          Const SizeErr = "known_ys and known_xs must be 1-column arrays with the same number of rows"
          Dim ChooseVector As Variant

2         Force2DArrayR known_ys
3         Force2DArrayR known_xs
4         If sNCols(known_ys) <> 1 Then Throw SizeErr + ", but known_ys has " + CStr(sNCols(known_ys)) + " columns"
5         If sNCols(known_xs) <> 1 Then Throw SizeErr + ", but known_xs has " + CStr(sNCols(known_xs)) + " columns"
6         If sNRows(known_xs) <> sNRows(known_ys) Then Throw SizeErr + ", but the row-counts are " + Format$(sNRows(known_xs), "###,##0") + " and " + Format$(sNRows(known_ys), "###,##0")

7         ChooseVector = sArrayAnd(sArrayIsNumber(known_ys), sArrayIsNumber(known_xs))
8         If sArrayCount(ChooseVector) < 2 Then Throw "Insufficient data"
9         known_ys = sMChoose(known_ys, ChooseVector)
10        known_xs = sMChoose(known_xs, ChooseVector)
11        LINESTWrap = Application.WorksheetFunction.LinEst(known_ys, known_xs, b_const, stats)
          
12        Exit Function
ErrHandler:
13        LINESTWrap = "#LINESTWrap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'Linear interpolation of missing values - this algorithm will be faster than calling sInterp (though haven't tested that that's the case)
Private Function FlatInfill(DataWithMissings)
          Dim CRRes
          Dim i As Long
          Dim j As Long
          Dim NChunks As Long
          Dim Result
1         On Error GoTo ErrHandler
2         Result = DataWithMissings
3         CRRes = sCountRepeats(sArrayIsNumber(DataWithMissings), "CFTH")
4         NChunks = sNRows(CRRes)
5         For i = 1 To NChunks
6             If CRRes(i, 1) = False And i > 1 Then
7                 For j = CRRes(i, 2) To CRRes(i, 3)
8                     Result(j, 1) = Result(j - 1, 1)
9                 Next j
10            End If
11        Next i
12        FlatInfill = Result

13        Exit Function
ErrHandler:
14        FlatInfill = "#FlatInfill (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function LinearInfill(DataWithMissings As Variant)
          Dim CRRes
          Dim i As Long
          Dim j As Long
          Dim Result
1         On Error GoTo ErrHandler
2         Result = DataWithMissings
3         CRRes = sCountRepeats(sArrayIsNumber(DataWithMissings), "CFTH")
          Dim NChunks As Long
          Dim StepSize As Double
4         NChunks = sNRows(CRRes)
5         For i = 1 To NChunks
6             If CRRes(i, 1) = False And i > 1 And i < NChunks Then
7                 StepSize = (DataWithMissings(CRRes(i, 3) + 1) - DataWithMissings(CRRes(i, 2) - 1)) / (CRRes(i, 4) + 1)
8                 For j = CRRes(i, 2) To CRRes(i, 3)
9                     Result(j, 1) = Result(j - 1, 1) + StepSize
10                Next j
11            End If
12        Next i
13        LinearInfill = Result

14        Exit Function
ErrHandler:
15        Throw LinearInfill = "#LinearInfill (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMPrepareData
' Author     : Philip Swannell
' Date       : 28-May-2019
' Purpose    : Attempt to emulate ISDA\Satori's "Data Preparation Steps"
' Parameters :
'  BankTimeSeries : Bank raw data, one column per bank. One row per weekday.
'  IncludeVector  : Boolean vector to allow easy switching on and off of an individual bank's data
'  SuppressZeros    : Should values of zero be treated as non-numeric. Recommended TRUE.
'  StaleTestNumber: How many repeated values are indication of stale?
'                   Old Approach: When stale, all but the first of a sequence of repeats is set to non-numeric (ignored).
'                   New Approach (30-6-19): When stale, the first StaleTestNumber values are retained, the remainder set to non-numeric. This looks to be ISDA's approach, at least for JPY_2W with StaleTestNumber = 2
'  WhatToReturn :   Must be one of 'Levels', 'ReturnsFromLevels', or 'ReturnsFromReturns'
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMPrepareData(BankTimeSeries As Variant, ByVal IncludeVector As Variant, SuppressZeros As Boolean, StaleTestNumber As Long, _
        WhatToReturn As String, ByVal Lag As Long, Absolute As Boolean, ImputeAlgo As String, ExcludeZeroReturnsInDPrep As Boolean)

          Dim AnyFALSE As Boolean
          Dim anyTrue As Boolean
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim M As Long
          Dim NC As Long
          Dim NCOrig As Long
          Dim NR As Long
          Dim Stage1    ' Removed time series to be ignored, as set by IncludeVector
          Dim Stage2    ' Zero and Non-Numbers -> "NA"
          Dim Stage3    ' Stale -> "NA"
          Dim Stage4    ' "NA" -> Imputed, according to ImputeAlgo and using function ISDASIMMImputeMissingData
          'If WhatToReturn is "Levels" or "ReturnsFromLevels"
          Dim Stage5    ' Take Medians
          Dim Stage6    'Calculate returns
          'If WhatToReturn is "ReturnsFromReturns"
          Dim Stage5a   'Calculate returns for each bank
          Dim Stage5b   'calculate median returns
          
          Const AllowedImputeAlgos = "None,ISDA,ISDA 2,ISDA 3,ISDA 3b,ISDA 4,Linear,Flat"
          '  Const ImputeAlgo = "ISDA 3b"

1         On Error GoTo ErrHandler

2         ThrowIfError BankTimeSeries

3         Force2DArrayR BankTimeSeries, NR, NC
4         NCOrig = NC
5         Force2DArrayR IncludeVector

6         If sNRows(IncludeVector) = 1 Then
7             If sNCols(IncludeVector) = NC Then
8                 IncludeVector = sArrayTranspose(IncludeVector)
9             End If
10        End If
11        If sNRows(IncludeVector) <> NC Or sNCols(IncludeVector) <> 1 Then Throw "IncludeVector must be a one-column or one-row array with the same number of elements as BankTimeSeries has columns"

12        For i = 1 To NC
13            If Not VarType(IncludeVector(i, 1)) = vbBoolean Then Throw "IncludeVector must contain only Booleans"
14            If IncludeVector(i, 1) Then anyTrue = True
15            If Not IncludeVector(i, 1) Then AnyFALSE = True
16        Next
17        If Not anyTrue Then Throw "At least one element of IncludeVector must be TRUE"

18        If AnyFALSE Then
19            Stage1 = ThrowIfError(sRowMChoose(BankTimeSeries, sArrayTranspose(IncludeVector)))
              'Need to redefine!
20            NC = sNCols(Stage1)
21        Else
22            Stage1 = BankTimeSeries
23        End If

24        Stage2 = Stage1
25        For j = 1 To NC
26            For i = 1 To NR
27                If Not IsNumber(Stage2(i, j)) Then
28                    Stage2(i, j) = "NA"
29                ElseIf Stage2(i, j) = 0 Then
30                    If SuppressZeros Then
31                        Stage2(i, j) = "NA"
32                    End If
33                End If
34            Next i
35        Next j

36        If StaleTestNumber > 0 Then
37            Stage3 = Stage2
38            For j = 1 To NC
                  Dim CRRet
39                CRRet = sCountRepeats(sSubArray(Stage3, 1, j, , 1), "CFTH")
                  '      If ISDASIMM20PercentTest(CRRet, ApplyBlocksOf10Rule) Then
40                If True Then
41                    For k = 1 To sNRows(CRRet)
42                        If CRRet(k, 4) >= StaleTestNumber Then
43                            If IsNumber(CRRet(k, 1)) Then
                                  '                                 For M = CRRet(k, 2) + 1 To CRRet(k, 3)
44                                For M = CRRet(k, 2) + StaleTestNumber To CRRet(k, 3)
45                                    Stage3(M, j) = "NA"
46                                Next M
47                            End If
48                        End If
49                    Next k
50                End If
51            Next j
52        Else
53            Stage3 = Stage2
54        End If

55        Stage4 = Stage3
          Dim MedianFirstDiff
56        If LCase(ImputeAlgo) <> "none" Then

57            If Not IsNumber(sMatch(ImputeAlgo, sTokeniseString(AllowedImputeAlgos))) Then
58                Throw "Illegal ImputeAlgo"
59            End If

60            MedianFirstDiff = ThrowIfError(sDifference(Stage3, 1, 1))
61            MedianFirstDiff = ThrowIfError(sRowMedian(MedianFirstDiff, True))

              Dim AnyNonNumbers As Boolean
              Dim ImputeRet
62            For j = 1 To NC
63                AnyNonNumbers = False
64                For i = 1 To NR
65                    If Not IsNumber(Stage4(i, j)) Then
66                        AnyNonNumbers = True
67                        Exit For
68                    End If
69                Next i
70                If AnyNonNumbers Then
71                    ImputeRet = ISDASIMMImputeMissingData(sSubArray(Stage4, 1, j, , 1), MedianFirstDiff, ImputeAlgo)
72                    If VarType(ImputeRet) <> vbString Then
73                        For i = 1 To NR
74                            Stage4(i, j) = ImputeRet(i, 1)
75                        Next i
76                    End If
77                End If
78            Next j
79        End If

80        Select Case LCase(WhatToReturn)
              Case LCase("BankByBankLevels"), LCase("BankByBankLevelsPadded")
81                NonNumberToNA Stage4
82                If LCase(WhatToReturn) = LCase("BankByBankLevels") Then
83                    ISDASIMMPrepareData = Stage4
84                Else
85                    If NCOrig > sNCols(Stage4) Then
86                        ISDASIMMPrepareData = sArrayRange(Stage4, sReshape(CVErr(xlErrNA), sNRows(Stage4), NCOrig - sNCols(Stage4)))
87                    Else
88                        ISDASIMMPrepareData = Stage4
89                    End If
90                End If

91            Case LCase("BankByBankReturns"), LCase("BankByBankReturnsPadded")
92                Stage5a = ISDASIMMReturnsFromPrices(Stage4, Lag, Absolute, ExcludeZeroReturnsInDPrep)
93                NonNumberToNA Stage5a

94                If LCase(WhatToReturn) = LCase("BankByBankReturns") Then
95                    ISDASIMMPrepareData = Stage5a
96                Else
97                    If NCOrig > sNCols(Stage5a) Then
98                        ISDASIMMPrepareData = sArrayRange(Stage5a, sReshape(CVErr(xlErrNA), sNRows(Stage5a), NCOrig - sNCols(Stage5a)))
99                    Else
100                       ISDASIMMPrepareData = Stage5a
101                   End If
102               End If
103           Case LCase("Levels"), LCase("ReturnsFromLevels")
104               Stage5 = sReshape(0, NR, 1)
105               For i = 1 To NR
                      'SafeMedian2 returns #N/A! when median can't be calculated
106                   Stage5(i, 1) = SafeMedian2(sSubArray(Stage4, i, 1, 1))
107               Next i
108               If LCase(WhatToReturn) = "levels" Then
109                   ISDASIMMPrepareData = Stage5
110               Else
111                   Stage6 = ISDASIMMReturnsFromPrices(Stage5, Lag, Absolute, ExcludeZeroReturnsInDPrep)
112                   ISDASIMMPrepareData = Stage6
113               End If
114           Case LCase("ReturnsFromReturns"), LCase("BankByBankReturns")
115               Stage5a = ISDASIMMReturnsFromPrices(Stage4, Lag, Absolute, ExcludeZeroReturnsInDPrep)
116               Stage5b = sReshape(0, NR - Lag, 1)
117               For i = 1 To NR - Lag
118                   Stage5b(i, 1) = SafeMedian2(sSubArray(Stage5a, i, 1, 1))
119               Next i
120               ISDASIMMPrepareData = Stage5b
121           Case LCase("LevelsFromReturnsFrom1dReturns")
                  'calculate the 1-day returns for each bank, take the median return for each day, then flip back to a price series(aka level)
122               Lag = 1
123               Stage5a = ISDASIMMReturnsFromPrices(Stage4, Lag, Absolute, ExcludeZeroReturnsInDPrep)
124               Stage5b = sReshape(0, NR - Lag, 1)
125               For i = 1 To NR - Lag
126                   Stage5b(i, 1) = SafeMedian2(sSubArray(Stage5a, i, 1, 1))
127               Next i
                  Dim Initial
128               For i = 1 To sNRows(Stage4)
129                   If IsNumber(Stage4(i, 1)) Then
130                       Initial = Stage4(i, 1)
131                       Exit For
132                   End If
133               Next

134               ISDASIMMPrepareData = ISDASIMMPricesFromReturns(Lag, Initial, Stage5b, Absolute, True)

135           Case Else
136               Throw "WhatToReturn must be one of 'Levels', 'ReturnsFromLevels', 'ReturnsFromReturns' or 'LevelsFromReturnsFrom1dReturns" '
137       End Select

138       Exit Function
ErrHandler:
139       ISDASIMMPrepareData = "#ISDASIMMPrepareData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub NonNumberToNA(ByRef Data)
          Dim NR As Long, NC As Long, i As Long, j As Long
1         On Error GoTo ErrHandler
2         NR = sNRows(Data)
3         NC = sNCols(Data)
4         For i = 1 To NR
5             For j = 1 To NC
6                 If Not IsNumber(Data(i, j)) Then Data(i, j) = CVErr(xlErrNA)
7             Next j
8         Next i

9         Exit Sub
ErrHandler:
10        Throw "#NonNumberToNA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'See page 12, para 4 of "ISDA SIMM Calibration Data Preparation Methodology August 2019"
'This function no longer relevant for 2021
Function ISDASIMM20PercentTest(Data, ApplyTheRule As Boolean) As Boolean
1         On Error GoTo ErrHandler
          Dim i As Long
          Dim x As Long
          Dim NRows As Long
          Dim CRRet
          
2         If Not ApplyTheRule Then
3             ISDASIMM20PercentTest = True
4             Exit Function
5         End If
          
6         Select Case sNCols(Data)
              Case 1
7                 CRRet = sCountRepeats(Data, "CFTH")
8             Case 4
9                 CRRet = Data
10            Case Else
11                Throw "Data must be passed as either 1-col array of bank-submitted data or else a 4-column array of the result of such but processed via function sCountRepeats with second argument 'CFTH'"
12        End Select

13        For i = 1 To sNRows(CRRet)
14            NRows = NRows + CRRet(i, 4)
15            If CRRet(i, 4) >= 10 Then
16                If IsNumber(CRRet(i, 1)) Then
17                    x = x + CRRet(i, 4)
18                End If
19            End If
20        Next
21        ISDASIMM20PercentTest = x <= NRows / 5

22        Exit Function
ErrHandler:
23        Throw "#ISDASIMM20PercentTest (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMPricesFromReturns
' Author     : Philip Swannell
' Date       : 24-Apr-2020
' Purpose    : Back out the time series for "Prices" from that for returns
' Parameters :
'  Lag               : Typically 1 or 10 for the length of the lag. Or enter 0 to indicate that the "Returns" are in fact not returns, but prices and should be returned unchanged.
'  Initial           : What is/are the starting values? need as many as Lag
'  Returns           :
'  ReturnsAreAbsolute:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMPricesFromReturns(Lag As Long, ByVal Initial As Variant, ByVal Returns As Variant, ReturnsAreAbsolute As Boolean, Optional TreatNonNumbersAsZero As Boolean)
          Dim i As Long, Prices As Variant
          Dim NRI As Long, NCI As Long
          Dim NRR As Long, NCR As Long

1         On Error GoTo ErrHandler
2         ThrowIfError Returns
3         If Lag < 0 Then Throw "Lag must be positive, or zero to indicate that Returns are in fact Prices."

4         If Lag = 0 Then
5             ISDASIMMPricesFromReturns = Returns
6             Exit Function
7         End If

8         Force2DArrayR Initial, NRI, NCI
9         If NCI <> 1 Then Throw "Initial must be a one-column array"
10        Force2DArrayR Returns, NRR, NCR

11        Prices = sReshape(0, NRR + Lag, 1)
12        For i = 1 To Lag
13            If Not IsNumber(Initial(i, 1)) Then Throw "Non number found in returns at row " + CStr(i) + " of argument Initial"
14            Prices(i, 1) = Initial(i, 1)

15        Next i
16        For i = 1 To NRR
17            If Not IsNumber(Returns(i, 1)) Then
18                If TreatNonNumbersAsZero Then
19                    Returns(i, 1) = 0
20                Else
21                    Throw "Non number found in returns at row " + CStr(i) + " of argument Returns. Consider setting argument TreatNonNumbersAsZero to True"
22                End If
23            End If
24            If ReturnsAreAbsolute Then
25                Prices(i + Lag, 1) = Prices(i, 1) + Returns(i, 1)
26            Else
27                Prices(i + Lag, 1) = Prices(i, 1) * (1 + Returns(i, 1))
28            End If
29        Next i

30        ISDASIMMPricesFromReturns = Prices

31        Exit Function
ErrHandler:
32        ISDASIMMPricesFromReturns = "#ISDASIMMPricesFromReturns (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMReturnsFromPrices
' Author     : Philip Swannell
' Date       : 27-Apr-2020
' Purpose    : Much more simple-minded function than ISDASIMMReturnsFromTimeSeries
' Parameters :
'  Prices  : asset prices - 1 row per observation, 1 col per asset
'  Lag     : number of observations (for ISDA work, weekdays) that the return is calculated over
'  Absolute: a returns absolute or relative?
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMReturnsFromPrices(Prices, Lag As Long, Absolute As Boolean, ExcludeZeroReturns As Boolean)
          Dim i As Long
          Dim j As Long
          Dim Returns As Variant
          Dim NR As Long, NC As Long
          Dim NA As String

1         On Error GoTo ErrHandler
2         NA = "NA" 'Convenient when we take medians of the return from this function
3         Force2DArrayR Prices, NR, NC
4         Returns = sReshape(0, NR - Lag, NC)

5         If Absolute Then
6             For j = 1 To NC
7                 For i = 1 To NR - Lag
8                     If IsNumber(Prices(i, j)) And IsNumber(Prices(i + Lag, j)) Then
9                         Returns(i, j) = Prices(i + Lag, j) - Prices(i, j)
10                        If ExcludeZeroReturns Then
11                            If Returns(i, j) = 0 Then
12                                Returns(i, j) = "NA"
13                            End If
14                        End If
15                    Else
16                        Returns(i, j) = NA
17                    End If
18                Next i
19            Next j
20        Else
21            For j = 1 To NC
22                For i = 1 To NR - Lag
23                    If isNonZeroNumber(Prices(i, j)) And IsNumber(Prices(i + Lag, j)) Then
24                        Returns(i, j) = (Prices(i + Lag, j) / Prices(i, j)) - 1
25                        If ExcludeZeroReturns Then
26                            If Returns(i, j) = 0 Then
27                                Returns(i, j) = "NA"
28                            End If
29                        End If
30                    Else
31                        Returns(i, j) = NA
32                    End If
33                Next i
34            Next j
35        End If
36        ISDASIMMReturnsFromPrices = Returns

37        Exit Function
ErrHandler:
38        ISDASIMMReturnsFromPrices = "#ISDASIMMReturnsFromPrices (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function isNonZeroNumber(x As Variant) As Boolean
1         Select Case VarType(x)
              Case vbDouble, vbInteger, vbSingle, vbLong
2                 isNonZeroNumber = x <> 0
3         End Select
End Function
