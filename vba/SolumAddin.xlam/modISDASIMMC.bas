Attribute VB_Name = "modISDASIMMC"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMMatchDates
' Author     : Philip Swannell
' Date       : 30-Apr-2020
' Purpose    : Match the 4 key dates into a column of dates (which are typically weekdays only). Shared by ISDASIMMReturnsFromFile and ISDASIMMRiskWeightsFromReturns
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISDASIMMMatchDates(ByVal Dates As Variant, ThreeYStart As Long, ThreeYEnd As Long, _
          StressStart As Long, StressEnd As Long, ByRef StartRow As Long, ByRef EndRow As Long, ByRef StressStartRow As Long, ByRef StressEndRow As Long, Optional FileName As String, _
          Optional RecentStressStart As Long, Optional RecentStressEnd As Long, Optional ByRef RecentStressStartRow As Long, Optional ByRef RecentStressEndRow As Long)
          
          Dim FirstDate As Long
          Dim LastDate As Long
          Dim Res As Variant
          Dim ErrString
          
1         On Error GoTo ErrHandler

          Dim SumOfPeriods As Long
          'PGS 19 March 2021
          'Sanity check the dates...When using the function for real breaking these conditions is likely to be a bug.
          'Would be better (reduce scope for mistakes), if we passed arguments AssetClass, TheYear, StressPeriodCalcMethod down the call stack
          'and inferred the recent and stress periods from them, but making that change would be time consuming, as it would involve editing many of the workbooks
          
2         If StressStart = 0 And StressEnd = 0 Then
              'Calling to determine the stress periods...
3             If ThreeYStart >= ThreeYEnd Then Throw "ThreeYStart must be before ThreeYEnd"
4         Else
5             SumOfPeriods = ThreeYEnd - ThreeYStart + 1 + StressEnd - StressStart + 1
6             If SumOfPeriods < 1457 Or SumOfPeriods > 1465 Then
7                 Throw "Incorrect Dates. The total dates covered from ThreeYStart to ThreeYEnd plus StressStart to StressEnd should be 4 years (approx 1461 days) but it is " + CStr(SumOfPeriods) + " days"
8             End If
9             If ThreeYStart >= ThreeYEnd Then Throw "ThreeYStart must be before ThreeYEnd"
10            If StressStart >= StressEnd Then Throw "StressStart must be before StressEnd"
11            If RecentStressStart > 0 Or RecentStressEnd > 0 Then
12                If RecentStressStart > RecentStressEnd Then Throw "RecentStressStart must be before RecentStressEnd"
13                If RecentStressStart < ThreeYStart Or RecentStressStart > ThreeYEnd Then Throw "RecentStressStart must be between ThreeYStart and ThreeYEnd"
14                If RecentStressEnd < ThreeYStart Or RecentStressEnd > ThreeYEnd Then Throw "RecentStressEnd must be between ThreeYStart and ThreeYEnd"
15            End If
16            If StressEnd > ThreeYStart Then Throw "StressEnd must be before ThreeYStart"
17        End If

18        If FileName = "" Then
19            ErrString = " in dates"
20        Else
21            ErrString = " in dates column of '" + FileName + "'"
22        End If
          
23        LastDate = Dates(sNRows(Dates), 1)
24        FirstDate = Dates(1, 1)
25        If ThreeYStart < FirstDate Or ThreeYStart > LastDate Then Throw "ThreeYStart must be in the range " + Format$(FirstDate, "dd-mmm-yyyy") + " to " + Format$(LastDate, "dd-mmm-yyyy")
26        If ThreeYEnd < FirstDate Or ThreeYEnd > LastDate Then Throw "ThreeYEnd must be in the range " + Format$(FirstDate, "dd-mmm-yyyy") + " to " + Format$(LastDate, "dd-mmm-yyyy")

27        Res = sSearchSorted(ThreeYStart, Dates, True)
28        If Not IsNumber(Res) Then Throw "Cannot find " + Format$(ThreeYStart, "dd-mmm-yyyy") + ErrString

29        StartRow = Res
30        Res = sSearchSorted(ThreeYEnd, Dates, False)
31        If Not IsNumber(Res) Then Throw "Cannot find " + Format$(ThreeYEnd, "dd-mmm-yyyy") + ErrString

32        EndRow = Res
33        If StressStart = 0 Then
34            StressStartRow = 0
35        Else
36            If StressStart < FirstDate Or StressStart > LastDate Then Throw "StressStart must be in the range " + Format$(FirstDate, "dd-mmm-yyyy") + " to " + Format$(LastDate, "dd-mmm-yyyy")
37            Res = sSearchSorted(StressStart, Dates, True)
38            If Not IsNumber(Res) Then Throw "Cannot find " + Format$(StressStart, "dd-mmm-yyyy") + ErrString

39            StressStartRow = Res
40        End If
41        If StressEnd = 0 Then
42            StressEndRow = 0
43        Else
44            If StressEnd < FirstDate Or StressEnd > LastDate Then Throw "StressEnd must be in the range " + Format$(FirstDate, "dd-mmm-yyyy") + " to " + Format$(LastDate, "dd-mmm-yyyy")
45            Res = sSearchSorted(StressEnd, Dates, False)
46            If Not IsNumber(Res) Then Throw "Cannot find " + Format$(StressEnd, "dd-mmm-yyyy") + ErrString
47            StressEndRow = Res
48        End If

49        If RecentStressStart <> 0 Then
50            If RecentStressStart < ThreeYStart Or RecentStressStart > ThreeYEnd Then Throw "RecentStressStart must be in the range " + Format$(ThreeYStart, "dd-mmm-yyyy") + " to " + Format$(ThreeYEnd, "dd-mmm-yyyy") + " but it is " + Format$(RecentStressStart, "dd-mmm-yyyy") + " which is outside that range"
51            Res = sSearchSorted(RecentStressStart, Dates, True)
52            If Not IsNumber(Res) Then Throw "Cannot find " + Format$(RecentStressStart, "dd-mmm-yyyy") + ErrString
53            RecentStressStartRow = Res
54        End If

55        If RecentStressEnd <> 0 Then
56            If RecentStressEnd < ThreeYStart Or RecentStressEnd > ThreeYEnd Then Throw "RecentStressEnd must be in the range " + Format$(ThreeYStart, "dd-mmm-yyyy") + " to " + Format$(ThreeYEnd, "dd-mmm-yyyy") + " but it is " + Format$(RecentStressEnd, "dd-mmm-yyyy") + " which is outside that range"
57            Res = sSearchSorted(RecentStressEnd, Dates, False)
58            If Not IsNumber(Res) Then Throw "Cannot find " + Format$(RecentStressEnd, "dd-mmm-yyyy") + ErrString
59            RecentStressEndRow = Res
60        End If

61        Exit Function
ErrHandler:
62        Throw "#ISDASIMMMatchDates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMReturnsFromFile
' Author    : Philip Swannell
' Date      : 05-Jul-2017
' Purpose   : Produces a "returns" array with no headers and the same number of columns
'             as elements in input Headers
'             'Headers can be passed in as row array or column array or alternatively with one of the following syntaxes:
'1)             "RegExp:MyRegularExpression" and then all columns that match that MyRegularExpression are included
'2)             "FirstN:N" - take the first N - e.g. "FirstN:10" to take the first 10
'3)             "EveryNth:N" to take every Nth header e.g. "EveryNth:10" to take the 1st, 11th, 21st etc
'4)             "RandomN:N" to take a pseudo-random choice of N. We seed the generator to always yield the same choice for given N and number of series in the file
'5)             "MinimumDataCount:A"  Restricts to only those instruments with at least A returns available in the 3+1 period
'6)             "MinimumDataCount:A,B" for numbers A and B. Restricts to only those instruments with both at least A returns available in the 3+1 and B returns available in stress period
'7)             Concatenation of 1) and 6) or of 1) and 7) e.g. MinimumDataCount:500,125RegExp:^((?!derived).)*$
'               Unfortunately this argument has become VERY messy, would be better to split to two arguments...
'8)             "Count:A,B" applies no restrictions to columns returned (Like RegExp:.*) but arranges that argument DCSP (standing for DataCountStressPeriod) is populated as
'                     1-row array. This horrendous bodge is used from function ISDASIMMCountReturnsInFile

'
' Want only one period of returns, not the Three + One arrangement? Then enter StressStart and StressEnd as zero
'Allowed values for argument PostProcessing:
' 'Abs' (aka 'Absolute', 'Absolute Value'): function returns the absolute value of the returns
' -----------------------------------------------------------------------------------------------------------------------
'PGS 18-March2021. StressStarts and StressEnds may now be a two element array giving the start\end of the old-style disjoint stress period and the new-style StressBalance stress period
'we only need to know the new-style stress period when Headers is of the form MinimumDataCount:A,B
'It's possible that in future years I will have to improve this code to handle the case when there is more than one recent stress quarter, ie make the number of elements in StressStarts\StressEnds be up to 4
Function ISDASIMMReturnsFromFile(FileName As String, ByVal Headers As Variant, IsAbs As Boolean, ThreeYStart As Long, ThreeYEnd As Long, _
          StressStarts As Variant, StressEnds As Variant, DateFormat As String, AllowBadHeaders As Boolean, _
          Optional FileIsReturns As Boolean = False, Optional ByRef HeadersFound As Variant, _
          Optional WithTopRow As Boolean, Optional WithLeftCol As Boolean, Optional ReturnLag As String, _
          Optional ExcludeZeroReturns As Boolean, Optional ByRef NStressRows As Long, Optional PostProcessing As String, _
          Optional ReturnRounding As Variant = False, Optional DataCleaningRules As String, Optional CheckDates As Boolean = True, Optional ByRef DCSP)

          Dim DataInFile As Variant
          Dim Dates As Variant
          Dim EndRow As Long
          Dim i As Long
          Dim j As Long
          Dim MatchIDs
          Dim NumDates As Long
          Dim NumRowsToIgnore As Long
          Dim NumSeries As Long
          Dim NumSeriesInFile As Long
          Dim RegExp As String
          Dim Result As Variant
          Dim StartRow As Long
          Dim StressEndRow As Long
          Dim StressStartRow As Long
          Dim TimeSeries As Variant
          Dim TopRow As Variant
          Dim TopRowT As Variant
          Dim NRH As Long, NCH As Long
          Dim TmpN As Long
          Dim CalledFromISDASIMMCountReturnsInFile As Boolean

1         On Error GoTo ErrHandler
          'PGS 3 March 2018. This years's data files have an extra row of data at the top :-(
2         DataInFile = ThrowIfError(sFileShow(FileName, ",", True, True, False, False, DateFormat))
3         NumRowsToIgnore = -1
4         For i = 1 To SafeMin(100, sNRows(DataInFile))

              'Almost all the files have 'Date' as left element of the header row (typically second row) but \2018\IR Vega\IR_Vega_Consolidated_Implied_Vols.csv has 'ID' so allow that as an alternative...
              'Arrgh for 2019 files, instead of 'Date' or 'ID', the supplied files (Fx Delta) have empty string
              'Arrgh for 2020 files, first that I looked at had "date" rather than "Date"
              'Arrrgh 2021 - sometimes `Dates`
5             If UCase(DataInFile(i, 1)) = "DATE" Or UCase(DataInFile(i, 1)) = "DATES" Or _
                  UCase(DataInFile(i, 1)) = "ID" Or DataInFile(i, 1) = vbNullString Then
6                 NumRowsToIgnore = i - 1
7                 Exit For
8             End If
9         Next i
10        If NumRowsToIgnore = -1 Then Throw "Cannot find header row in file, whose first element must either 'Date' or 'ID' or the empty string"
11        Dates = sSubArray(DataInFile, 2 + NumRowsToIgnore, 1, , 1)
12        NumDates = sNRows(Dates)

13        For i = 1 To NumDates
14            If Not IsNumberOrDate(Dates(i, 1)) Then
15                Throw "Data in 'Date' column is inconsistent with input DateFormat of '" + DateFormat + "'. For example text '" + Dates(i, 1) + "' cannot be interpreted as a date. Is '" + DateFormat + "' correct?"
16            End If
17        Next i

          'Experimental 10 Feb 2020. Also checks that after filtering out weekends, remaining dates are consecutive weekdays.
18        If CheckDates Then
19            ISDASIMMFilterOutWeekends DataInFile, Dates, NumRowsToIgnore, NumDates, FileName
20        End If

21        Force2DArrayR Headers, NRH, NCH
22        If sNRows(Headers) = 1 Then
23            TmpN = NRH: NRH = NCH: NCH = TmpN
24            Headers = sArrayTranspose(Headers)
25        End If

26        For i = 1 To NRH
27            For j = 1 To NCH
28                If VarType(Headers(i, j)) <> vbString Then Throw "Headers must be provided as a string or array of strings"
29            Next j
30        Next i

31        TopRow = sSubArray(DataInFile, 1 + NumRowsToIgnore, 1, 1)
32        TopRowT = sArrayTranspose(TopRow)
33        NumSeriesInFile = sNRows(TopRowT) - 1

          Dim MDC As String

34        If sNRows(Headers) = 1 Then
35            If InStr(Headers(1, 1), "RegExp:") > 0 Then
36                RegExp = sStringBetweenStrings(Headers(1, 1), "RegExp:")
37            End If
38            If InStr(Headers(1, 1), "MinimumDataCount:") > 0 Then
39                MDC = sStringBetweenStrings(Headers(1, 1), "MinimumDataCount:", "RegExp:")
40            End If
41            If Headers(1, 1) = "ISDASIMMCountReturnsInFile" Then
42                MDC = "500,125"
43                CalledFromISDASIMMCountReturnsInFile = True
44            End If

              'AMENDING SYNTAX?
45            If RegExp <> vbNullString Then
                  Dim ChooseVector
46                ChooseVector = sIsRegMatch(RegExp, TopRowT, False)
47                ChooseVector(1, 1) = False
48                If sArrayCount(ChooseVector) = 0 Then Throw "No headers in file match regular expression: " + RegExp
49                Headers = sMChoose(TopRowT, ChooseVector)
50                Force2DArray Headers        'is this necessary?
51            End If
52            If MDC <> vbNullString Then
                  '33            ElseIf Left(Headers(1, 1), 17) = "MinimumDataCount:" Then    'Example syntax : 'MinimumDataCount:500,150' i.e. need 500 returns in 3+1 period AND ALSO 150 returns in stress period
53                If InStr(MDC, ",") = 0 Then MDC = MDC & ",0"    'if not specifying the minimum number of points in the stress period, set it to zero

                  Dim MinimumDataCount As Variant
                  Dim MinimumDataCountInStressPeriod As Variant
54                MinimumDataCount = sStringBetweenStrings(MDC, , ",")
55                MinimumDataCountInStressPeriod = sStringBetweenStrings(MDC, ",")

56                If Not IsNumeric(MinimumDataCount) Then Throw "Invalid Headers"
57                If Not IsNumeric(MinimumDataCountInStressPeriod) Then Throw "Invalid Headers"

58                MinimumDataCount = Val(MinimumDataCount)
59                MinimumDataCountInStressPeriod = Val(MinimumDataCountInStressPeriod)
60                If RegExp = vbNullString Then
61                    Headers = sSubArray(TopRowT, 2)    'because we don't yet know how to select series, and we have to lop off "Date"
62                End If
63            End If
64            If Left$(Headers(1, 1), 7) = "FirstN:" Then
                  Dim N As Variant
65                N = sArrayRight(Headers(1, 1), -7)
66                If Not IsNumeric(N) Then Throw "Invalid Headers"
67                N = CLng(N)
68                If N > NumSeriesInFile Then
69                    Headers = sSubArray(TopRowT, 2, 1, , 1)
70                Else
71                    Headers = sSubArray(TopRowT, 2, 1, N, 1)
72                End If
73            ElseIf Left$(Headers(1, 1), 9) = "EveryNth:" Then
74                N = sArrayRight(Headers(1, 1), -9)
75                If Not IsNumeric(N) Then Throw "Invalid Headers"
76                N = CLng(N)
77                Headers = sEveryNthElement(TopRowT, 2, CLng(N))
78            ElseIf Left$(Headers(1, 1), 8) = "RandomN:" Then
79                N = sArrayRight(Headers(1, 1), -8)
80                If Not IsNumeric(N) Then Throw "Invalid Headers"
81                N = CLng(N)
                  Dim Seed As Long
82                Seed = 100
83                ChooseVector = sArrayStack(False, RandomChooseVector(CLng(N), NumSeriesInFile, Seed))
84                Headers = sMChoose(TopRowT, ChooseVector)
85            End If
86        End If

87        HeadersFound = Headers

88        NumSeries = sNRows(Headers)
89        MatchIDs = sMatch(Headers, TopRowT)
90        Force2DArray MatchIDs

91        For i = 1 To sNRows(Headers)
92            If Not IsNumber(MatchIDs(i, 1)) Then
93                If AllowBadHeaders Then
94                    MatchIDs(i, 1) = 0
95                Else
96                    Throw "Cannot find header '" + CStr(Headers(i, 1)) + "' in file '" + FileName + "'. Consider setting AllowBadHeaders to TRUE"
97                End If
98            End If
99        Next i

          '18 March 2021. Deal with complexities of the StressBalance setup. Currently only handle one RecentStress Quarter
          Dim StressStart As Long, StressEnd As Long
          Dim RecentStressStart As Long, RecentStressEnd As Long
          Dim RecentStressStartRow As Long, RecentStressEndRow As Long
          Dim HaveRecentStress As Boolean
          
100       If sNRows(StressStarts) <> sNRows(StressEnds) Or sNCols(StressStarts) <> sNCols(StressEnds) Then
101           Throw "StressStarts and StressEnds must be arrays of the same size"
102       End If
103       If sNRows(StressEnds) > 1 Or sNCols(StressStarts) > 1 Then
104           StressStarts = sReshape(StressStarts, 1, sNRows(StressStarts) * sNCols(StressStarts))
105           StressEnds = sReshape(StressEnds, 1, sNRows(StressEnds) * sNCols(StressEnds))
106       End If

107       Select Case sNCols(StressStarts)
              Case 1
108               StressStart = StressStarts
109               StressEnd = StressEnds
110               HaveRecentStress = False
111               ISDASIMMMatchDates Dates, ThreeYStart, ThreeYEnd, StressStart, StressEnd, StartRow, EndRow, StressStartRow, StressEndRow, FileName
112           Case 2
113               StressStart = StressStarts(1, 1)
114               StressEnd = StressEnds(1, 1)
                    
115               If IsNumber(StressStarts(1, 2)) And IsNumber(StressEnds(1, 2)) Then
116                   RecentStressStart = StressStarts(1, 2)
117                   RecentStressEnd = StressEnds(1, 2)
118                   HaveRecentStress = True
119                   ISDASIMMMatchDates Dates, ThreeYStart, ThreeYEnd, StressStart, StressEnd, StartRow, EndRow, StressStartRow, StressEndRow, FileName, RecentStressStart, RecentStressEnd, RecentStressStartRow, RecentStressEndRow
120               Else
121                   HaveRecentStress = False
122                   ISDASIMMMatchDates Dates, ThreeYStart, ThreeYEnd, StressStart, StressEnd, StartRow, EndRow, StressStartRow, StressEndRow, FileName
123               End If
124           Case Else
125               Throw "Case of more than one recent stress quarter is not yet handled"
126       End Select

127       TimeSeries = sReshape(0, NumDates, NumSeries)

128       For j = 1 To NumSeries
129           If MatchIDs(j, 1) = 0 Then
130               For i = 1 To NumDates
131                   TimeSeries(i, j) = 1        'Fake data!
132               Next i
133           Else
134               For i = 1 To NumDates
135                   TimeSeries(i, j) = DataInFile(i + 1 + NumRowsToIgnore, MatchIDs(j, 1))
136               Next i
137           End If
138       Next j

139       If FileIsReturns Then
140           If StressStartRow = 0 And StressEndRow = 0 Then
141               Result = sSubArray(TimeSeries, StartRow, 1, EndRow - StartRow + 1)
142           Else
143               NStressRows = StressEndRow - StressStartRow + 1
144               Result = sArrayStack(sSubArray(TimeSeries, StressStartRow, 1, StressEndRow - StressStartRow + 1), sSubArray(TimeSeries, StartRow, 1, EndRow - StartRow + 1))
145           End If
146           ISDASIMMApplyDataCleaningRules Result, DataCleaningRules
147           Result = ISDASIMMApplyRounding(Result, ReturnRounding)
148       Else
149           Result = ThrowIfError(ISDASIMMReturnsFromTimeSeries(TimeSeries, IsAbs, True, StartRow, EndRow, StressStartRow, StressEndRow, ReturnLag, ExcludeZeroReturns, NStressRows, ReturnRounding, DataCleaningRules))
150       End If

151       Select Case LCase(Replace(PostProcessing, " ", ""))
              Case "", "none"
152           Case "abs", "absolute", "absolutevalue"
                  Dim NR As Long, NC As Long
153               NR = sNRows(Result): NC = sNCols(Result)
154               For i = 1 To NR
155                   For j = 1 To NC
156                       If IsNumber(Result(i, j)) Then
157                           Result(i, j) = Abs(Result(i, j))
158                       End If
159                   Next
160               Next
161           Case Else
162               Throw "Argument 'PostProcessing' not recognised. Allowed values 'Abs', 'None', 'omitted or null string"
163       End Select

164       If Not IsEmpty(MinimumDataCount) Then
              Dim DataCount
165           NR = sNRows(Result): NC = sNCols(Result)
166           DataCount = sReshape(0, 1, NC)
167           DCSP = sReshape(0, 1, NC)
168           For j = 1 To NC
169               For i = 1 To NStressRows
170                   If IsNumber(Result(i, j)) Then
171                       DCSP(1, j) = DCSP(1, j) + 1
172                       DataCount(1, j) = DataCount(1, j) + 1
173                   End If
174               Next
175               If HaveRecentStress Then
                      Dim IterateFrom As Long, IterateTo As Long
176                   IterateFrom = NStressRows + RecentStressStartRow - StartRow + 1
177                   IterateTo = NStressRows + RecentStressEndRow - StartRow + 1
178                   For i = IterateFrom To IterateTo
179                       If IsNumber(Result(i, j)) Then
180                           DCSP(1, j) = DCSP(1, j) + 1
                              'recentstress period should be inside the recent aka 3Y period so we don't increment DataCount here
181                       End If
182                   Next i
183               End If
184               For i = NStressRows + 1 To NR
185                   If IsNumber(Result(i, j)) Then DataCount(1, j) = DataCount(1, j) + 1
186               Next
187           Next
188           If Not CalledFromISDASIMMCountReturnsInFile Then
189               ChooseVector = sArrayGreaterThanOrEqual(DataCount, MinimumDataCount)
190               If MinimumDataCountInStressPeriod > 0 Then
191                   ChooseVector = sArrayAnd(ChooseVector, sArrayGreaterThanOrEqual(DCSP, MinimumDataCountInStressPeriod))
192               End If
193               Result = sRowMChoose(Result, ChooseVector)
194               HeadersFound = sMChoose(HeadersFound, sArrayTranspose(ChooseVector))
195           End If
196       End If

197       If AllowBadHeaders Then
198           For j = 1 To NumSeries
199               If MatchIDs(j, 1) = 0 Then
                      Dim tmp As String
200                   tmp = "#Header '" + CStr(Headers(j, 1)) + "' not found in file!"
201                   For i = 1 To sNRows(Result)
202                       Result(i, j) = tmp        'overwrite the returns of the fake data, although as it turns out this is not necessary since Kendal Tau is undefined if one of the sequences is constant...
203                   Next i
204               End If
205           Next j
206       End If

          Dim SelectedDates
207       If WithLeftCol Then
208           SelectedDates = sSubArray(Dates, StartRow, 1, EndRow - StartRow + 1, 1)
209           If StressStartRow <> 0 Or StressEndRow <> 0 Then
                  'NB Stess dates go at the top. Prior to 27 March 2017, they were at the bottom
210               SelectedDates = sArrayStack(sSubArray(Dates, StressStartRow, 1, StressEndRow - StressStartRow + 1, 1), SelectedDates)
211           End If
212       End If

213       If WithTopRow And WithLeftCol Then
214           ISDASIMMReturnsFromFile = sArraySquare("Date", sArrayTranspose(HeadersFound), SelectedDates, Result)
215       ElseIf WithTopRow Then
216           ISDASIMMReturnsFromFile = sArrayStack(sArrayTranspose(HeadersFound), Result)
217       ElseIf WithLeftCol Then
218           ISDASIMMReturnsFromFile = sArrayRange(SelectedDates, Result)
219       Else
220           ISDASIMMReturnsFromFile = Result
221       End If

222       Exit Function
ErrHandler:
223       ISDASIMMReturnsFromFile = "#ISDASIMMReturnsFromFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMWeekDays
' Author     : Philip Swannell
' Date       : 02-Feb-2022
' Purpose    : Returns sorted 1-column array of weekdays in range FromDate to ToDate
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMWeekDays(FromDate As Date, ToDate As Date)
          Dim AllDays() As Variant, ChooseVector() As Boolean, i As Long, NR As Long

1         On Error GoTo ErrHandler
2         NR = ToDate - FromDate + 1

3         ReDim AllDays(1 To NR, 1 To 1)
4         ReDim ChooseVector(1 To NR, 1 To 1)
5         For i = 1 To NR
6             AllDays(i, 1) = FromDate + i - 1
7             ChooseVector(i, 1) = AllDays(i, 1) Mod 7 > 1
8         Next i

9         ISDASIMMWeekDays = ThrowIfError(sMChoose(AllDays, ChooseVector))

10        Exit Function
ErrHandler:
11        Throw "#ISDASIMMWeekDays (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMFilterOutWeekends
' Author     : Philip Swannell
' Date       : 10-Feb-2020
' Purpose    : Arrgh. Finding that some files generated by ISDA have a small number of weekends in the dates column. This filters them out
'              see email correspondence 10 Feb 2020
' -----------------------------------------------------------------------------------------------------------------------
Sub ISDASIMMFilterOutWeekends(ByRef DataInFile, ByRef Dates, NumRowsToIgnore As Long, ByRef NumDates As Long, FileName As String)
          Dim ChooseVector
          Dim k As Long
          Dim i As Long
          Dim AnyWeekends As Boolean
1         On Error GoTo ErrHandler
2         ChooseVector = sReshape(True, sNRows(DataInFile), 1)
3         For i = 1 To sNRows(Dates)
4             If Dates(i, 1) Mod 7 <= 1 Then
5                 ChooseVector(i + NumRowsToIgnore + 1, 1) = False
6                 AnyWeekends = True
7             End If
8         Next i
9         If AnyWeekends Then
10            DataInFile = ThrowIfError(sMChoose(DataInFile, ChooseVector))
11            Dates = ThrowIfError(sMChoose(Dates, sSubArray(ChooseVector, NumRowsToIgnore + 2, 1)))
12        End If
13        NumDates = sNRows(Dates)

14        For i = 2 To sNRows(Dates)
15            k = IIf(Dates(i - 1, 1) Mod 7 = 6, 3, 1)
16            If Dates(i, 1) <> Dates(i - 1, 1) + k Then
17                Throw "Missing weekday in file '" + FileName + "'. Date " + Format$(Dates(i - 1, 1) + k, "dd-mmm-yyyy") + " not found or is out of order"
18            End If
19        Next

20        Exit Sub
ErrHandler:
21        Throw "#ISDASIMMFilterOutWeekends (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function ISDASIMMValidateRounding(ReturnRounding As Variant, ByRef ApplyRounding As Boolean, ByRef NumDigits As Long)
          Const ErrString = "ApplyRounding must be False or a positive integer"
1         On Error GoTo ErrHandler
2         If VarType(ReturnRounding) = vbBoolean Then
3             If ReturnRounding = False Then
4                 ApplyRounding = False
5             Else
6                 Throw ErrString
7             End If
8         ElseIf IsNumber(ReturnRounding) Then
9             If ReturnRounding > 0 And CInt(ReturnRounding) = ReturnRounding Then
10                ApplyRounding = True
11                NumDigits = ReturnRounding
12            Else
13                Throw ErrString
14            End If
15        Else
16            Throw ErrString
17        End If
18        Exit Function
ErrHandler:
19        Throw "#ISDASIMMValidateRounding (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMApplyRounding
' Author     : Philip Swannell
' Date       : 25-Feb-2020
' Purpose    : Applies rounding to numeric elements of argument Returns (so different from function sRound in how it handles non-numbers)
' Parameters :
'  Returns               :
'  ReturnRounding: Either False or positive integer to round to that many digits
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMApplyRounding(Returns As Variant, ReturnRounding As Variant)
          Dim ApplyRounding As Boolean
          Dim NumDigits As Long
          Dim NR As Long, NC As Long
          Dim i As Long
          Dim j As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR Returns, NR, NC
3         ISDASIMMValidateRounding ReturnRounding, ApplyRounding, NumDigits
4         If ApplyRounding Then
5             For i = 1 To NR
6                 For j = 1 To NC
7                     If IsNumber(Returns(i, j)) Then
8                         Returns(i, j) = Round(Returns(i, j), NumDigits)
9                     End If
10                Next
11            Next
12        End If
13        ISDASIMMApplyRounding = Returns
14        Exit Function
ErrHandler:
15        Throw "#ISDASIMMApplyRounding (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub TestISDASIMMApplyDataCleaningRules()
          Dim Data, DataCleaningRule As String
1         Data = sReshape(sIntegers(100), 10, 10)
2         DataCleaningRule = ">=10,<=30"
3         g Data
4         ISDASIMMApplyDataCleaningRules Data, DataCleaningRule
5         g Data
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMApplyDataCleaningRules
' Author     : Philip Swannell
' Date       : 28-Feb-2020
' Purpose    : Soft code some rules for rejecting obviously bad data
' Parameters :
'  Data             :
'  DataCleaningRules: comma delimited string, currently support only the string "equity" (faster version of ">0,<1000000000")
'  or tokens can be >N, >=N, <N, <=N for N some text representation of a number
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMApplyDataCleaningRules(ByRef Data As Variant, DataCleaningRules As String, Optional NumRowsToIgnore As Long, Optional NumColsToIgnore As Long)
          Dim Rules As Variant
          Dim Rule As String
          Dim NR As Long
          Dim NC As Long
          Dim i As Long, j As Long, k As Long
          Dim LowerBound As Double
          Dim upperbound As Double
          
          Const ErrString = "DataCleaningRules should be comma-delimited concatenation of allowed 'rules'. Allowed rules are of the form '>N', '>=N', '<N', '<=N' for N some text representation of a number. The rule 'Equity' is a (faster) abbreviation of '>0,<1000000000'. For No datacleaning use null string or 'None'"

1         If Len(DataCleaningRules) > 0 Then
2             Force2DArrayR Data, NR, NC
3             Rules = sTokeniseString(DataCleaningRules)
4             For k = 1 To sNRows(Rules)
5                 Rule = Rules(k, 1)
6                 If LCase(Rule) = "none" Then
                      'nothing to do
7                 ElseIf LCase(Rule) = "equity" Then
8                     For i = NumRowsToIgnore + 1 To NR
9                         For j = NumColsToIgnore + 1 To NC
10                            If IsNumber(Data(i, j)) Then
11                                If Data(i, j) <= 0 Or Data(i, j) > 1000000000# Then
12                                    Data(i, j) = "NA"
13                                End If
14                            End If
15                        Next
16                    Next
17                ElseIf Left(Rule, 2) = "<=" Then
18                    If Not IsNumeric(Mid(Rule, 3)) Then Throw "Unrecognised data cleaning rule: '" + Rule + "' " + ErrString
19                    upperbound = CDbl(Mid(Rule, 3))
20                    For i = NumRowsToIgnore + 1 To NR
21                        For j = NumColsToIgnore + 1 To NC
22                            If IsNumber(Data(i, j)) Then
23                                If Data(i, j) > upperbound Then ' i.e. If Not Data(i,j) <= UpperBound
24                                    Data(i, j) = "NA"
25                                End If
26                            End If
27                        Next
28                    Next
29                ElseIf Left(Rule, 1) = "<" Then
30                    If Not IsNumeric(Mid(Rule, 2)) Then Throw "Unrecognised data cleaning rule: '" + Rule + "' " + ErrString
31                    upperbound = CDbl(Mid(Rule, 2))
32                    For i = NumRowsToIgnore + 1 To NR
33                        For j = NumColsToIgnore + 1 To NC
34                            If IsNumber(Data(i, j)) Then
35                                If Data(i, j) >= upperbound Then ' i.e. If Not Data(i,j) < UpperBound
36                                    Data(i, j) = "NA"
37                                End If
38                            End If
39                        Next
40                    Next
41                ElseIf Left(Rule, 2) = ">=" Then
42                    If Not IsNumeric(Mid(Rule, 3)) Then Throw "Unrecognised data cleaning rule: '" + Rule + "' " + ErrString
43                    LowerBound = CDbl(Mid(Rule, 3))
44                    For i = NumRowsToIgnore + 1 To NR
45                        For j = NumColsToIgnore + 1 To NC
46                            If IsNumber(Data(i, j)) Then
47                                If Data(i, j) < LowerBound Then ' i.e. If Not Data(i,j) >= LowerBound
48                                    Data(i, j) = "NA"
49                                End If
50                            End If
51                        Next
52                    Next
53                ElseIf Left(Rule, 1) = ">" Then
54                    If Not IsNumeric(Mid(Rule, 2)) Then Throw "Unrecognised data cleaning rule: '" + Rule + "' " + ErrString
55                    LowerBound = CDbl(Mid(Rule, 2))
56                    For i = NumRowsToIgnore + 1 To NR
57                        For j = NumColsToIgnore + 1 To NC
58                            If IsNumber(Data(i, j)) Then
59                                If Data(i, j) <= LowerBound Then ' i.e. If Not Data(i,j) > LowerBound
60                                    Data(i, j) = "NA"
61                                End If
62                            End If
63                        Next
64                    Next
65                Else
66                    Throw "Unrecognised data cleaning rule: '" + Rule + "' " + ErrString
67                End If
68            Next k
69        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMReturnsFromTimeSeries
' Author    : Philip
' Date      : 07-Jul-2017
' Purpose   :
' THIS FUNCTION RETURNS THE TIME SERIES OF RETURNS IN TWO CHUNKS STACKED WITH STRESS PERIOD AT THE TOP (changed from bottom, 17 March 18)
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMReturnsFromTimeSeries(TimeSeries, IsAbsolute As Boolean, _
        Optional AllowNonNumbers As Boolean, _
        Optional StartRow As Long = 0, Optional ByVal EndRow As Long = -1, _
        Optional StressStartRow As Long, Optional ByVal StressEndRow As Long, _
        Optional ReturnLag As String, Optional ExcludeZeroReturns, Optional ByRef NumStressRows As Long, _
        Optional ReturnRounding As Variant = False, _
        Optional DataCleaningRules As String)

          Dim i As Long
          Dim j As Long
          Dim Lag As Long
          Dim Lagged As Variant
          Dim Naive As Boolean
          Dim NC As Long
          Dim NR As Long
          Dim NR2 As Long
          Dim NR3 As Long
          Dim Result
          Dim Returns

1         On Error GoTo ErrHandler
2         Force2DArrayR TimeSeries

3         ISDASIMMApplyDataCleaningRules TimeSeries, DataCleaningRules

4         NR = sNRows(TimeSeries): NC = sNCols(TimeSeries)

5         Select Case LCase$(ReturnLag)
              Case "1day"
6                 Lag = 1
7                 Naive = True
8             Case "10daynaive", "true"    'true for backward-compatibility argument ReturnLag was previously NaiveLag
9                 Lag = 10
10                Naive = True
11            Case "10dayenhanced", "false"
12                Lag = 10
13                Naive = False
14                Throw "ReturnLag '10DayEnhanced' should not be used since ISDA never adopted that method"
15            Case Else
16                Throw "ReturnLag not recognised. Allowed values: 1Day, 10DayNaive, 10DayEnhanced"
17        End Select

18        Result = sReshape(0, 1, NC)
          'Let negative offsets count back from the end
19        If EndRow < 0 Then EndRow = NR + 1 + EndRow
20        If StartRow < 0 Then StartRow = NR + 1 + StartRow
21        If StartRow = 0 Then StartRow = Lag + 1
22        If StressEndRow < 0 Then StressEndRow = NR + 1 + StressEndRow
23        If StressStartRow < 0 Then StressStartRow = NR + 1 + StressStartRow

24        If StartRow <= Lag Then Throw "StartRow in TimeSeries must be at least " + CStr(Lag + 1) + " but it's " + CStr(StartRow)
25        If StressStartRow <= Lag And StressStartRow <> 0 Then Throw "StressStartRow in TimeSeries must be at least " + CStr(Lag + 1) + " but it's " + CStr(StressStartRow)

26        NR2 = EndRow - StartRow + 1

27        If StressStartRow <> 0 Or StressEndRow <> 0 Then
28            NR3 = StressEndRow - StressStartRow + 1
29        Else
30            NR3 = 0
31        End If
32        NumStressRows = NR3

33        Returns = sReshape(0, NR2 + NR3, NC)

          Dim HaveGoodInputs As Boolean, ReadOffset As Long, WriteOffset As Long, LoopTo As Long, k As Long

34        For j = 1 To NC

35            For k = 1 To 2
36                If k = 1 Then
37                    ReadOffset = StressStartRow - 1
38                    WriteOffset = 0
39                    LoopTo = NR3
40                Else
41                    ReadOffset = StartRow - 1
42                    WriteOffset = NR3
43                    LoopTo = NR2
44                End If

45                For i = 1 To LoopTo
46                    Lagged = LookupWithLag(TimeSeries, ReadOffset + i, Lag, j, Naive)
47                    If IsAbsolute Then
48                        HaveGoodInputs = (IsNumber(TimeSeries(ReadOffset + i, j)) And IsNumber(Lagged))
49                    Else
50                        HaveGoodInputs = (IsNumber(TimeSeries(ReadOffset + i, j)) And isNonZeroNumber(Lagged))
51                    End If
52                    If HaveGoodInputs Then
53                        If IsAbsolute Then
54                            Returns(WriteOffset + i, j) = TimeSeries(ReadOffset + i, j) - Lagged
55                        Else
56                            Returns(WriteOffset + i, j) = (TimeSeries(ReadOffset + i, j) - Lagged) / Lagged
57                        End If
58                        If ExcludeZeroReturns Then If Returns(WriteOffset + i, j) = 0 Then Returns(WriteOffset + i, j) = vbNullString
59                    Else
60                        If AllowNonNumbers Then
61                            Returns(WriteOffset + i, j) = vbNullString
62                        Else
63                            Throw "Non numbers detected in TimeSeries. Consider setting AllowNonNumbers to TRUE"
64                        End If
65                    End If
66                Next i
67            Next k
68        Next j
69        Returns = ISDASIMMApplyRounding(Returns, ReturnRounding)

70        ISDASIMMReturnsFromTimeSeries = Returns
71        Exit Function
ErrHandler:
72        ISDASIMMReturnsFromTimeSeries = "#ISDASIMMReturnsFromTimeSeries (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function isNonZeroNumber(x As Variant)
1         Select Case VarType(x)
              Case vbDouble, vbInteger, vbSingle, vbLong
2                 isNonZeroNumber = x <> 0
3         End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LookupWithLag
' Author    : Philip
' Date      : 03-Jul-2017
' Purpose   : Sub of ISDASIMMReturnsFromTimeSeries
' -----------------------------------------------------------------------------------------------------------------------
Private Function LookupWithLag(TimeSeries, row As Long, Lag As Long, Column As Long, Naive As Boolean)

          Dim ExtraLag As Long
1         On Error GoTo ErrHandler

2         If Naive Then    'Emulate Satori's handling of lag - in which one bad day makes two missing return numbers
3             LookupWithLag = TimeSeries(row - Lag, Column)
4             Exit Function
5         End If

TryAgain:
6         If IsNumber(TimeSeries(row - Lag - ExtraLag, Column)) Then
7             LookupWithLag = TimeSeries(row - Lag - ExtraLag, Column)
8             Exit Function
9         Else
10            ExtraLag = ExtraLag + 1
11        End If
12        If ExtraLag = row - Lag Then
13            LookupWithLag = "#Cannot roll back to number!"
14            Exit Function
15        Else
16            GoTo TryAgain
17        End If

18        Exit Function
ErrHandler:
19        Throw "#LookupWithLag (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'TODO make higher level functions call this?
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRiskWeightFromReturns
' Author     : Philip Swannell
' Date       : 27-Apr-2020
' Purpose    : For when we have returns pre-calculated
' Parameters :
'  Dates             : 1-column array
'  Returns           : 1-column array of associated returns
'  ThreeYStart       : An index into Dates, not a date itself!
'  ThreeYEnd         : ditto
'  StressStart       : ditto
'  StressEnd         : ditto
'  ExcludeZeroReturns:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRiskWeightFromReturns(Dates As Variant, Returns As Variant, ThreeYStart As Long, ThreeYEnd As Long, _
        StressStart As Long, StressEnd As Long, ExcludeZeroReturns As Boolean)

          Dim RelevantReturns
          Dim ThreeYReturns, StressReturns
          Dim NR1 As Long, NC1 As Long, NR2 As Long, NC2 As Long
          Dim Percentile1, Percentile99
          Dim i As Long
          Dim StartRow As Long, EndRow As Long, StressStartRow As Long, StressEndRow As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR Dates, NR1, NC1
3         Force2DArrayR Returns, NR2, NC2

4         If sIsErrorString(Dates) Then Throw ("Dates is the error - " + Dates(1, 1))
5         If sIsErrorString(Returns) Then Throw ("Returns is the error - " + Returns(1, 1))

6         If NR1 <> NR2 Then Throw "Dates and Returns must have the same number of rows"
7         If NC1 <> 1 Then Throw "Dates must be a 1-column array"
8         If NC2 <> 1 Then Throw "Returns must be a 1-column array"

9         ISDASIMMMatchDates Dates, ThreeYStart, ThreeYEnd, StressStart, StressEnd, StartRow, EndRow, StressStartRow, StressEndRow

10        ThreeYReturns = ThrowIfError(sSubArray(Returns, StartRow, 1, EndRow - StartRow + 1))
11        StressReturns = ThrowIfError(sSubArray(Returns, StressStartRow, 1, StressEndRow - StressStartRow + 1))
12        RelevantReturns = sArrayStack(ThreeYReturns, StressReturns)

13        If ExcludeZeroReturns Then
14            For i = 1 To sNRows(RelevantReturns)
15                If IsNumber(RelevantReturns(i, 1)) Then
16                    If RelevantReturns(i, 1) = 0 Then
17                        RelevantReturns(i, 1) = "NA"
18                    End If
19                End If
20            Next i
21        End If

22        Percentile1 = ThrowIfError(sGeneralisedQuantile(RelevantReturns, 0.01, "CENTRAL", True))
23        Percentile99 = ThrowIfError(sGeneralisedQuantile(RelevantReturns, 0.99, "CENTRAL", True))

24        If Abs(Percentile99) > Abs(Percentile1) Then
25            ISDASIMMRiskWeightFromReturns = Abs(Percentile99)
26        Else
27            ISDASIMMRiskWeightFromReturns = Abs(Percentile1)
28        End If

29        Exit Function
ErrHandler:
30        ISDASIMMRiskWeightFromReturns = "#ISDASIMMRiskWeightFromReturns (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMRiskWeight
' Author    : Philip
' Date      : 29-Jun-2017
' Purpose   : For TimeSeriesData, returns a calculation of the ISDA SIMM Risk Weight
'             See Page 15 of "ISDA SIMM Calibration Methodology, 24 May 2017"
'             RiskWeight = Max(Abs(1% quantile of returns),Abs(99% quantile of returns))
'             TimeSeries can have more than one column in which case so does the return
'             IsAbsolute will need to be TRUE for credit spread and interest rates, FALSE for equity, commodity, Fx, volatilities
'   See also ISDASIMMRiskWeightsFromFile
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRiskWeight(TimeSeries, IsAbsolute As Boolean, _
        AllowNonNumbers As Boolean, _
        StartRow As Long, ByVal EndRow As Long, _
        StressStartRow As Long, ByVal StressEndRow As Long, _
        PercentileMethod As String, _
        ReturnLag As String, ExcludeZeroReturns As Boolean)

          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim P1 As Double
          Dim P2 As Double
          Dim Result
          Dim Returns
          Const Conf1 = 0.01
          Const Conf2 = 0.99
          Dim ThisCol As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR TimeSeries, NR, NC
3         Result = sReshape(0, 1, NC)

4         Returns = ThrowIfError(ISDASIMMReturnsFromTimeSeries(TimeSeries, IsAbsolute, AllowNonNumbers, StartRow, EndRow, StressStartRow, StressEndRow, ReturnLag, ExcludeZeroReturns))

5         For j = 1 To NC
6             ThisCol = sSubArray(Returns, 1, j, , 1)
7             P1 = PercentileWrap(ThisCol, Conf1, PercentileMethod)
8             P2 = PercentileWrap(ThisCol, Conf2, PercentileMethod)
9             Result(1, j) = IIf(Abs(P1) > Abs(P2), Abs(P1), Abs(P2))
10        Next j

11        ISDASIMMRiskWeight = Result
12        Exit Function
ErrHandler:
13        ISDASIMMRiskWeight = "#ISDASIMMRiskWeight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : PercentileWrap
' Author    : Philip Swannell
' Date      : 05-Jul-2017
' Purpose   : Put error handling around the call to Percentile_Exc or Percentile
' -----------------------------------------------------------------------------------------------------------------------
Private Function PercentileWrap(TheArray, k, PercentileMethod As String)
1         On Error GoTo ErrHandler
2         If UCase$(PercentileMethod) = "INC" Then
3             PercentileWrap = Application.WorksheetFunction.Percentile(TheArray, k)
4         ElseIf UCase$(PercentileMethod) = "EXC" Then
5             PercentileWrap = Application.WorksheetFunction.Percentile_Exc(TheArray, k)
6         Else
7             PercentileWrap = sGeneralisedQuantile(TheArray, CDbl(k), PercentileMethod, True)
8         End If
9         Exit Function
ErrHandler:
10        PercentileWrap = "#PercentileWrap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMRiskWeightsFromFile
' Author    : Philip Swannell
' Date      : 05-Jul-2017
' Purpose   : Get the RiskWeights for a number of assets in a file. Note that headers can have any number of rows and columns
' 18 April 2018 - have overloaded argument PercentileMethod to allow "Median" in which case we return the Median of each relevant column, not the RiskWeight...
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRiskWeightsFromFile(FileName As String, ByVal Headers As Variant, IsAbsolute As Boolean, ThreeYStart As Long, ThreeYEnd As Long, StressStarts As Variant, _
        StressEnds As Variant, DateFormat As String, AllowBadHeaders As Boolean, FileIsReturns As Boolean, PercentileMethod As String, _
        Optional WithHeaders As Boolean, Optional ReturnLag As String, Optional ExcludeZeroReturns As Boolean, Optional PostProcessing As Variant, _
        Optional ReturnRounding As Boolean, Optional DataCleaningRules As String, Optional CheckDates As Boolean = True)

          Dim HeadersAsCol As Variant
          Dim HeadersFound As Variant
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim P1 As Variant
          Dim P2 As Variant
          Dim Result As Variant
          Dim Returns As Variant
          Dim ReturnsNR As Long
          Dim RiskWeightsAsCol
          Dim ThisCol As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR Headers, NR, NC

3         HeadersAsCol = sReshape(Headers, NR * NC, 1)
4         Returns = ThrowIfError(ISDASIMMReturnsFromFile(FileName, HeadersAsCol, IsAbsolute, ThreeYStart, ThreeYEnd, StressStarts, StressEnds, DateFormat, AllowBadHeaders, FileIsReturns, HeadersFound, , , ReturnLag, ExcludeZeroReturns, , , ReturnRounding, DataCleaningRules, CheckDates))
5         ReturnsNR = sNRows(Returns)

6         If HeadersUseSpecialSyntax(Headers) Then
7             NC = sNRows(HeadersFound)        'Have to redefine to cope correctly with headers defined via Regular Expression
8             NR = 1
9         End If

10        RiskWeightsAsCol = sReshape("NA", NR * NC, 1)

11        For j = 1 To NC * NR
12            ThisCol = sSubArray(Returns, 1, j, , 1)
13            If ExcludeZeroReturns Then
14                If IsNumber(sMatch(0, ThisCol)) Then
15                    ThisCol = sMChoose(ThisCol, sArrayNot(sArrayEquals(ThisCol, 0)))
16                End If
17            End If
18            If LCase$(PercentileMethod) = "median" Then
19                For i = 1 To ReturnsNR
20                    If Not IsNumber(ThisCol(i, 1)) Then ThisCol(i, 1) = vbNullString
21                Next
22                RiskWeightsAsCol(j, 1) = SafeMedian(ThisCol)
23            Else
24                P1 = PercentileWrap(ThisCol, 0.99, PercentileMethod)
25                P2 = PercentileWrap(ThisCol, 0.01, PercentileMethod)
26                If IsNumber(P1) And IsNumber(P2) Then
27                    RiskWeightsAsCol(j, 1) = IIf(Abs(P1) > Abs(P2), Abs(P1), Abs(P2))
28                Else
29                    If sIsErrorString(ThisCol(1, 1)) Then
30                        RiskWeightsAsCol(j, 1) = ThisCol(1, 1)
31                    Else
32                        RiskWeightsAsCol(j, 1) = "#Cannot calculate percentiles. 99% percentile = '" + CStr(P1) + "' 1% percentile = '" + CStr(P2) + "'!"
33                    End If
34                End If
35            End If
36        Next j

37        Result = sReshape(RiskWeightsAsCol, NR, NC)
38        If WithHeaders Then
39            If NR = 1 Then
40                HeadersFound = sReshape(HeadersFound, 1, NC)
41                Result = sArrayStack(HeadersFound, Result)
42            ElseIf NC = 1 Then
43                HeadersFound = sReshape(HeadersFound, NR, 1)
44                Result = sArrayRange(HeadersFound, Result)
45            End If
46        End If

47        If Not IsEmpty(PostProcessing) Then
48            Result = DoPostProcessing(Result, PostProcessing)
49        End If

50        ISDASIMMRiskWeightsFromFile = Result

51        Exit Function
ErrHandler:
52        ISDASIMMRiskWeightsFromFile = "#ISDASIMMRiskWeightsFromFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HeadersUseSpecialSyntax
' Author    : Philip Swannell
' Date      : 09-Jul-2017
' Purpose   : Figure out if headers are being passed in explicitly or passed via a special
'             syntax giving instructions for choosing which to use given what's in the file
' -----------------------------------------------------------------------------------------------------------------------
Private Function HeadersUseSpecialSyntax(Headers) As Boolean
1         On Error GoTo ErrHandler

2         Force2DArrayR Headers
3         If sNRows(Headers) = 1 Then
4             If sNCols(Headers) = 1 Then
5                 If Left$(Headers(1, 1), 7) = "RegExp:" Then
6                     HeadersUseSpecialSyntax = True
7                 ElseIf Left$(Headers(1, 1), 7) = "FirstN:" Then
8                     HeadersUseSpecialSyntax = True
9                 ElseIf Left$(Headers(1, 1), 9) = "EveryNth:" Then
10                    HeadersUseSpecialSyntax = True
11                ElseIf Left$(Headers(1, 1), 9) = "RandomN:" Then
12                    HeadersUseSpecialSyntax = True
13                ElseIf Left$(Headers(1, 1), 17) = "MinimumDataCount:" Then
14                    HeadersUseSpecialSyntax = True
15                End If
16            End If
17        End If
18        Exit Function
ErrHandler:
19        Throw "#HeadersUseSpecialSyntax (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DoPostProcessing
' Author     : Philip Swannell
' Date       : 13-Apr-2019
' Purpose    : Common code to do post processing. Post processing inside the function ensures that error messages do not
'              get morphed to #VALUE!, which makes it quicker to understand what inputs are causing output errors.
' -----------------------------------------------------------------------------------------------------------------------
Private Function DoPostProcessing(ByVal x, ByVal PostProcessing As Variant)
1         On Error GoTo ErrHandler
          Const Unrecognised = "PostProcess not recognised. May be a number (to multiply), or one of the following strings: 'None', 'Median', 'Median*100', Median*10000"
          Dim i As Long
          Dim j As Long

2         If IsEmpty(PostProcessing) Or IsMissing(PostProcessing) Then
3             DoPostProcessing = x
4         ElseIf LCase(PostProcessing) = "none" Then
5             DoPostProcessing = x
6         ElseIf IsNumber(PostProcessing) Then
7             Force2DArrayR x
8             If PostProcessing <> 1 Then
                  Dim NC2 As Long
                  Dim NR2 As Long
9                 NR2 = sNRows(x)
10                NC2 = sNCols(x)
11                For i = 1 To NR2
12                    For j = 1 To NC2
13                        If IsNumber(x(i, j)) Then
14                            x(i, j) = x(i, j) * PostProcessing
15                        End If
16                    Next
17                Next
18            End If
19            DoPostProcessing = x
20        ElseIf VarType(PostProcessing) = vbString Then
21            Select Case LCase$(PostProcessing)
                  Case vbNullString
22                    DoPostProcessing = x
23                Case "median"
24                    DoPostProcessing = sArrayRange(Application.WorksheetFunction.Median(x), sNCols(x))
25                Case "median*100"
26                    DoPostProcessing = sArrayRange(Application.WorksheetFunction.Median(x) * 100, sNCols(x))
27                Case "median*10000"
28                    DoPostProcessing = sArrayRange(Application.WorksheetFunction.Median(x) * 10000, sNCols(x))
29                Case Else
30                    Throw Unrecognised
31            End Select
32        Else
33            Throw Unrecognised
34        End If

35        Exit Function
ErrHandler:
36        Throw "#DoPostProcessing (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRiskWeightsFromFiles
' Author     : Philip Swannell
' Date       : 21-Mar-2018
' Purpose    : Wrapper to ISDASIMMRiskWeightsFromFile to cope with many files at a time (specified via Folder and FileFilter)
'            : Return is array-range of returns from underlying function
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRiskWeightsFromFiles(Folder As String, FileFilter As String, ByVal Headers As Variant, IsAbsolute As Boolean, ThreeYStart As Long, _
        ThreeYEnd As Long, StressStart As Long, StressEnd As Long, DateFormat As String, AllowBadHeaders As Boolean, FileIsReturns As Boolean, _
        PercentileMethod As String, Optional WithHeaders As Boolean, Optional ReturnLag As String, Optional ExcludeZeroReturns As Boolean, Optional PostProcessing As String, _
        Optional ReturnRounding As Boolean, Optional DataCleaningRules As String)

          Dim Files As Variant
          Dim i As Long
          Dim NumFiles As Long
          Dim STK As clsStacker
          Dim ThisRes

1         On Error GoTo ErrHandler
2         Files = sDirList(Folder, False, False, "F", "F", FileFilter)
3         If sIsErrorString(Files) Then
4             Throw "Search for files in folder '" + Folder + "' that match the filter '" + FileFilter + "' yielded the error '" + Files + "'"
5         End If

6         NumFiles = sNRows(Files)
7         If NumFiles = 1 Then
8             ISDASIMMRiskWeightsFromFiles = ThrowIfError(ISDASIMMRiskWeightsFromFile(CStr(Files(1, 1)), Headers, IsAbsolute, ThreeYStart, ThreeYEnd, StressStart, StressEnd, DateFormat, AllowBadHeaders, FileIsReturns, PercentileMethod, WithHeaders, ReturnLag, ExcludeZeroReturns, , ReturnRounding, DataCleaningRules))
9         Else
10            Set STK = CreateStacker()
11            For i = 1 To NumFiles
12                ThisRes = ThrowIfError(ISDASIMMRiskWeightsFromFile(CStr(Files(i, 1)), Headers, IsAbsolute, ThreeYStart, ThreeYEnd, StressStart, StressEnd, DateFormat, AllowBadHeaders, FileIsReturns, PercentileMethod, WithHeaders, ReturnLag, ExcludeZeroReturns, , ReturnRounding, DataCleaningRules))
13                STK.Stack2D sArrayTranspose(ThisRes)
14            Next i
15            ISDASIMMRiskWeightsFromFiles = STK.ReportInTranspose
16        End If

17        ISDASIMMRiskWeightsFromFiles = DoPostProcessing(ISDASIMMRiskWeightsFromFiles, PostProcessing)

18        Exit Function
ErrHandler:
19        ISDASIMMRiskWeightsFromFiles = "#ISDASIMMRiskWeightsFromFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRowMediansFromFiles
' Author     : Philip Swannell
' Date       : 18-Apr-2018
' Purpose    : Calculates the row-by-row median of a file or a set of files. Files should be "regular" comma-delimited
'              each with the first header being "Date". The Date column of each file must be identical. The return is two columns
'              giving the date and the median over all the files of the numeric data on the corresponding row. Function designed to be used
'              as part of calculation of the "stress periods" for asset classes. But not yet plumbed in to the stress period workbooks.
' Parameters :
'  FileNames      : A column array of file names
'  HeaderRowNumber: The line number for the header row. Assumed to be the same for all the files
'  NewFile: If provided then data is written to file, with a header row "Date","Return", comma delimited. If not provided the function returns a two-column array without header row.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRowMediansFromFiles(ByVal FileNames As Variant, HeaderRowNumber As Long, Optional NewFile As String)
          Dim BlankArray
          Dim FileHeaders
          Dim FileName As String
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim LineFromFile As String
          Dim LineNum As Long
          Dim NumFiles As Long
          Dim NumSeries As Long
          Dim ParsedLineFromFile As Variant
          Dim STK As clsStacker
          Dim ThisDate As String
          Dim TS() As TextStream
          Dim WriteTo As Long

1         On Error GoTo ErrHandler

2         Set FSO = New Scripting.FileSystemObject
3         Force2DArrayR FileNames
4         NumFiles = sNRows(FileNames)

5         For i = 1 To NumFiles
6             FileName = FileNames(i, 1)
7             If Not sFileExists(FileName) Then Throw "Cannot find file '" + FileName + "'"
8             FileHeaders = ThrowIfError(sFileHeaders(FileName, ",", HeaderRowNumber))
9             If FileHeaders(1, 1) <> "Date" Then Throw "First element of header row in each file must be 'Date', but it is not for file'" + FileName + "'"
10            NumSeries = NumSeries + sNCols(FileHeaders) - 1
11        Next i

12        ReDim TS(1 To NumFiles)
13        For i = 1 To NumFiles
14            FileName = FileNames(i, 1)
15            Set TS(i) = FSO.GetFile(FileName).OpenAsTextStream(ForReading)
16            For j = 1 To HeaderRowNumber
17                TS(i).SkipLine
18            Next j
19        Next i

20        ReDim BlankArray(1 To NumSeries)
21        LineNum = HeaderRowNumber

          Dim BlankDataToStack
22        BlankDataToStack = sArrayRange(vbNullString, vbNullString)
23        Set STK = CreateStacker()
24        If NewFile <> "" Then
25            STK.Stack2D sArrayRange("Date", "Return")
26        End If

27        Do While Not TS(1).atEndOfStream
28            WriteTo = 0
29            LineNum = LineNum + 1
30            If LineNum Mod 250 = 1 Then
31                MessageLogWrite "ISDASIMMRowMediansFromFiles: processing line " + CStr(LineNum)
32            End If

33            For i = 1 To NumFiles
34                If TS(i).atEndOfStream Then Throw "All files must have the same number of lines"
35                LineFromFile = TS(i).ReadLine
36                ParsedLineFromFile = VBA.Split(LineFromFile, ",")
37                If i = 1 Then
38                    ThisDate = ParsedLineFromFile(LBound(ParsedLineFromFile))
39                ElseIf ThisDate <> ParsedLineFromFile(LBound(ParsedLineFromFile)) Then
40                    Throw "The 'Date' columns of all files must be identical, but they are not. At row number " + CStr(LineNum) + ", file '" + FileNames(i, 1) + "' reads '" + ParsedLineFromFile(LBound(ParsedLineFromFile)) + "', whereas file '" + FileNames(1, 1) + "' reads '" + ThisDate + "'"
41                End If
42                For j = LBound(ParsedLineFromFile) + 1 To UBound(ParsedLineFromFile)
43                    WriteTo = WriteTo + 1
44                    If (IsNumeric(ParsedLineFromFile(j))) Then
45                        BlankArray(WriteTo) = CDbl(ParsedLineFromFile(j))
46                    Else
47                        BlankArray(WriteTo) = vbNullString
48                    End If
49                Next j
50            Next i
51            BlankDataToStack(1, 1) = ThisDate
52            BlankDataToStack(1, 2) = SafeMedian(BlankArray)
53            STK.Stack2D BlankDataToStack
54        Loop
          Dim DataToWrite As Variant
55        DataToWrite = STK.Report

56        If NewFile <> "" Then
57            ThrowIfError sFileSave(NewFile, DataToWrite, ",")
58            ISDASIMMRowMediansFromFiles = NewFile
59        Else
60            ISDASIMMRowMediansFromFiles = DataToWrite
61        End If

62        Exit Function
ErrHandler:
63        ISDASIMMRowMediansFromFiles = "#ISDASIMMRowMediansFromFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RandomChooseVector
' Author    : Philip Swannell
' Date      : 09-Jul-2017
' Purpose   : Returns a vector of TRUE or FALSE length of vector = M. number of TRUEs = Min(N,M)
'             Because of seeding, always makes the same choice for given N and M
' -----------------------------------------------------------------------------------------------------------------------
Private Function RandomChooseVector(N As Long, M As Long, Optional Seed As Variant)
          Dim i As Long
          Dim Randoms
          Const GeneratorName = "Wichmann-Hill"

1         On Error GoTo ErrHandler
2         If N >= M Then
3             RandomChooseVector = sReshape(True, M, 1)
4         Else
5             If IsNumber(Seed) Then
6                 ThrowIfError sRandomSetSeed(GeneratorName, Seed)
7             End If

8             Randoms = sArrayRange(sIntegers(M), sRandomVariable(M, 1, "Uniform", GeneratorName))
9             Randoms = sSortedArray(Randoms, 2, , , True)
10            For i = 1 To M
11                Randoms(i, 2) = (i <= N)
12            Next
13            Randoms = sSortedArray(Randoms, 1, , , True)
14            RandomChooseVector = sSubArray(Randoms, 1, 2, , 1)
15        End If

16        Exit Function
ErrHandler:
17        RandomChooseVector = "#RandomChooseVector (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMedianOffDiagonal
' Author    : Philip Swannell
' Date      : 07-Jul-2017
' Purpose   : Returns the median of the elements of Matrix that are not on the diagonal.
' Arguments
' Matrix    : A square matrix of numbers.
' -----------------------------------------------------------------------------------------------------------------------
Function sMedianOffDiagonal(ByVal Matrix As Variant)
Attribute sMedianOffDiagonal.VB_Description = "Returns the median of the elements of Matrix that are not on the diagonal."
Attribute sMedianOffDiagonal.VB_ProcData.VB_Invoke_Func = " \n23"
          Dim i As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         ThrowIfError Matrix

3         Force2DArrayR Matrix, NR, NC
4         If NR <> NC Then Throw "Matrix must be square"
5         If NR = 1 Then Throw "Matrix must have at least two rows"
6         For i = 1 To NR
7             Matrix(i, i) = vbNullString    ' ignored by Median function
8         Next
9         sMedianOffDiagonal = Application.WorksheetFunction.Median(Matrix)

10        Exit Function
ErrHandler:
11        sMedianOffDiagonal = "#sMedianOffDiagonal (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SortTenures
' Author    : Philip Swannell
' Date      : 10-Jul-2017
' Purpose   : Sort a column array of tenure strings
' -----------------------------------------------------------------------------------------------------------------------
Function SortTenures(ByVal TenureStrings)
          Dim i As Long
          Dim N As Variant
          Dim NC As Long
          Dim NR As Long
          Dim TempArray

1         On Error GoTo ErrHandler
2         Force2DArrayR TenureStrings, NR, NC
3         If NC <> 1 Then Throw "TenureStrings must be a 1-column array"
4         TempArray = sReshape(0, NR, 2)
5         For i = 1 To NR
6             TempArray(i, 1) = TenureStrings(i, 1)
7             If VarType(TenureStrings(i, 1)) = vbString Then
8                 N = Left$(TenureStrings(i, 1), Len(TenureStrings(i, 1)) - 1)
9                 If IsNumeric(N) Then
10                    N = CDbl(N)
11                    Select Case UCase$(Right$(TenureStrings(i, 1), 1))
                          Case "Y"
12                            TempArray(i, 2) = N
13                        Case "M"
14                            TempArray(i, 2) = N / 12
15                        Case "W"
16                            TempArray(i, 2) = N / 365 * 7
17                        Case "D"
18                            TempArray(i, 2) = N / 365
19                    End Select
20                End If
21            End If
22        Next i

23        TempArray = sSortedArray(TempArray, 2, , , True)
24        SortTenures = sSubArray(TempArray, 1, 1, , 1)

25        Exit Function
ErrHandler:
26        SortTenures = "#SortTenures (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
