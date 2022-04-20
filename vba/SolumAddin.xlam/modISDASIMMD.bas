Attribute VB_Name = "modISDASIMMD"
' -----------------------------------------------------------------------------------------------------------------------
' Name: modISDASIMMD
' Kind: Module
' Purpose: 2020 work for ISDA - Global stress period calculation - not done in previous years and also code designed to
' make the calibration exercise more efficient  - i.e. use less of my time.
' Author: Philip Swannell
' Date: 09-Apr-2020
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Function DaysInMonth(y, M)
1         DaysInMonth = day(DateSerial(y, M + 1, 1) - 1)
End Function

Function ISDASIMMSlidingVolEndDates(StartDates, NumMonths As Long)
          Dim NR As Long, NC As Long, i As Long, j As Long, Res As Variant, CopyOfErr As String

1         On Error GoTo ErrHandler
2         Force2DArrayR StartDates, NR, NC

3         Res = sReshape(0, NR, NC)
4         For i = 1 To NR
5             For j = 1 To NC
6                 Res(i, j) = ISDASIMMSlidingVolEndDate(StartDates(i, j), NumMonths)
7             Next j
8         Next i
9         ISDASIMMSlidingVolEndDates = Res

10        Exit Function
ErrHandler:
11        CopyOfErr = Err.Description
12        CopyOfErr = "Failure at row " & CStr(i) & ", column " & CStr(j) & " " & CopyOfErr
13        ISDASIMMSlidingVolEndDates = "#ISDASIMMSlidingVolEndDates (line " & CStr(Erl) + "): " & CopyOfErr & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMSlidingVolEndDate
' Author     : Philip Swannell
' Date       : 28-Jan-2022
' Purpose    : Calculate the end of the "sliding vol period" from the start date. Slightly waffly rubric from ISDA:
'

'        The end day of a period of x quarters or y years
'        of length is defined to be one day before the start
'        day: for example, a 1-quarter period starting on
'        the 25th January would end on the 24th April. If
'        the start day is 1st of the month, the end day will
'        be the final day of a month (e.g. the 1-year
'        period starting on 1st February 2019 ends on 31st
'        January 2020). Thus, the length of each period
'        can vary depending on the start date.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMSlidingVolEndDate(StartDate, NumMonths As Long)

          Dim D As Long, M As Long, y As Long

1         On Error GoTo ErrHandler
2         D = day(StartDate)
3         M = Month(StartDate)
4         y = Year(StartDate)

5         If D < 28 Then
6             ISDASIMMSlidingVolEndDate = DateSerial(y, M + NumMonths, D) - 1
7         Else
8             ISDASIMMSlidingVolEndDate = DateSerial(y, M + NumMonths, Min(D - 1, DaysInMonth(y, M + NumMonths)))
9         End If

10        Exit Function
ErrHandler:
11        Throw "#ISDASIMMSlidingVolEndDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ISDASIMMRankedSlidingVolsCore(ByVal TheDates As Variant, ByVal TheReturns As Variant, FromDate As Long, _
          ToDate As Long, MonthsInWindow As Long, Optional WhatToReturn As String = "Ranks", Optional AllowWeekendVolStart As Boolean = True)

          Dim FromDates, ToDates
          Dim i As Long
          Dim FromPos, ToPos
          Dim N As Long
          Dim NRD As Long, NCD As Long
          Dim NRR As Long, NCR As Long
          Dim MChooseNeeded As Boolean, ChooseVector

1         On Error GoTo ErrHandler

2         Force2DArrayR TheDates, NRD, NCD
3         Force2DArrayR TheReturns, NRR, NCR

4         If NRD <> NRR Then Throw "TheDates and TheReturns must have the same number of rows"
5         If NCD <> 1 Then Throw "TheDates must have only a single column"
6         If NCR <> 1 Then Throw "TheReturns must have only a single column"

7         For i = 1 To NRD
8             If Not IsNumberOrDate(TheDates(i, 1)) Then
9                 Throw "TheDates must be dates or numbers, but '" + CStr(TheDates(i, 1)) + "' found at row " + CStr(i)
10            ElseIf i > 1 Then
11                If TheDates(i, 1) <= TheDates(i - 1, 1) Then
12                    Throw "Dates are not monotonic"
13                End If
14            End If
15        Next

16        For i = 1 To NRR
17            If Not IsNumber(TheReturns(i, 1)) Then
18                MChooseNeeded = True
19                Exit For
20            End If
21        Next

22        If MChooseNeeded Then
23            ChooseVector = sArrayIsNumber(TheReturns)
24            TheDates = sMChoose(TheDates, ChooseVector)
25            TheReturns = sMChoose(TheReturns, ChooseVector)
26            NRD = sNRows(TheDates)
27            NRR = NRD
28        End If

29        If AllowWeekendVolStart Then
30            FromDates = sArrayAdd(FromDate - 1, sIntegers(ToDate - FromDate + 1))
31        Else
32            FromDates = ISDASIMMWeekDays(CDate(FromDate), CDate(ToDate))
33        End If
34        N = sNRows(FromDates)

35        If LCase(WhatToReturn) = "fromdates" Then
36            ISDASIMMRankedSlidingVolsCore = FromDates
37            Exit Function
38        End If

39        ToDates = ISDASIMMSlidingVolEndDates(FromDates, MonthsInWindow)

          'Calculating FromPos and ToPos is a bit painful as sMatch doesn't quite have the behaviour we need and Excel MATCH function can't be called with array first argument from VBA
          'FromPos(i,1) gives the position (row number) in TheDates of the first date (reading top to bottom) that's greater than or equal to FromDates(i,1)
          'ToPos(i,1) gives the position (row number) in TheDates of the last date (reading top to bottom) that's less than or equal to ToDates(i,1)

40        FromPos = sReshape(0, N, 1)
41        For i = 1 To N
42            If i > 1 Then
43                FromPos(i, 1) = FromPos(i - 1, 1)
44            Else
45                FromPos(i, 1) = 1
46            End If
47            While (TheDates(FromPos(i, 1), 1) < FromDates(i, 1))
48                FromPos(i, 1) = FromPos(i, 1) + 1
49            Wend
50        Next i

51        ToPos = sReshape(0, N, 1)
52        For i = 1 To N
53            If i > 1 Then
54                ToPos(i, 1) = ToPos(i - 1, 1)
55            Else
56                ToPos(i, 1) = 1
57            End If
58            If ToPos(i, 1) < NRD Then
59                Do While (TheDates(ToPos(i, 1) + 1, 1) <= ToDates(i, 1))
60                    ToPos(i, 1) = ToPos(i, 1) + 1
61                    If ToPos(i, 1) = NRD Then
62                        Exit Do
63                    End If
64                Loop
65            End If
66        Next i

          Dim PartialSum, PartialSumSquares, Vols
          Dim EX2, EX

67        PartialSum = sPartialSum(TheReturns)
68        PartialSumSquares = sPartialSum(sArrayMultiply(TheReturns, TheReturns))

69        Vols = sReshape(0, N, 1)
70        For i = 1 To N
71            EX2 = PartialSumSquares(ToPos(i, 1), 1)
72            If FromPos(i, 1) > 1 Then
73                EX2 = EX2 - PartialSumSquares(FromPos(i, 1) - 1, 1)
74            End If
75            EX2 = EX2 / (ToPos(i, 1) - FromPos(i, 1) + 1)

76            EX = PartialSum(ToPos(i, 1), 1)
77            If FromPos(i, 1) > 1 Then
78                EX = EX - PartialSum(FromPos(i, 1) - 1, 1)
79            End If
80            EX = EX / (ToPos(i, 1) - FromPos(i, 1) + 1)
81            Vols(i, 1) = Sqr(EX2 - EX ^ 2)
82        Next i

          Dim Ranks

          'Old way used in 2021. Led to equal ranks for ties
          'Ranks = sMatch(Vols, sSortedArray(Vols, , , , False))
          'New way 2022
          Dim Headers
83        If LCase(WhatToReturn) = "datesandvols" Then
84            Headers = sArrayRange("From Date", "To Date", "Vol")
85            ISDASIMMRankedSlidingVolsCore = sArrayStack(Headers, sArrayRange(FromDates, ToDates, Vols))
              
86        ElseIf LCase(WhatToReturn) = "ranks" Then
87            Ranks = ISDASIMMRankVols(FromDates, Vols)
88            ISDASIMMRankedSlidingVolsCore = Ranks
89        ElseIf LCase(WhatToReturn) = "fromdates" Then
90            ISDASIMMRankedSlidingVolsCore = FromDates
91        ElseIf LCase(WhatToReturn) = "details" Then
92            Ranks = ISDASIMMRankVols(FromDates, Vols)
93            Headers = sArrayRange("Date" & IIf(MChooseNeeded, "(Filtered)", ""), "Return", "From Date", "To Date", "From Pos", "To Pos", "Vol", "Rank")
94            ISDASIMMRankedSlidingVolsCore = sArrayStack(Headers, sArrayRange(TheDates, TheReturns, FromDates, ToDates, FromPos, ToPos, Vols, Ranks))
95        Else
96            Throw "WhatToReturn not recognised"
97        End If
98        Exit Function
ErrHandler:

99        Throw "#ISDASIMMRankedSlidingVolsCore (line " & CStr(Erl) + "): " & Err.Description & "!"
100   End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRankVols
' Author     : Philip Swannell
' Date       : 28-Jan-2022
' Purpose    : Implements ISDA's vol ranking algorithm for which the rubric is:

'             "When computing rankings of periods in the
'              context of the global pseudo index calculation, if
'              the ranking of two dates ties, we rank the older
'              date ahead of the more recent date."
'
'              Unfortunately I find this hard to understand - what does "ahead" mean? So the Ascending2 argument to
'              the first call to sSortedArray in this function was chosen (to be True) so as to match ISDA's results
'              it means that in the case of ties, the later date has a higher rank, i.e. is treated as if it had the
'              lower vol.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRankVols(Dates, Vols)
          Dim tmp
1         On Error GoTo ErrHandler
2         tmp = sArrayRange(Dates, Vols)
3         tmp = sSortedArray(tmp, 2, 1, , False, True)
4         tmp = sArrayRange(tmp, sIntegers(sNRows(Dates)))
5         tmp = sSortedArray(tmp)
6         ISDASIMMRankVols = sSubArray(tmp, 1, 3, , 1)
7         Exit Function
ErrHandler:
8         Throw "#ISDASIMMRankVols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'Version for use in 2020 when, for each asset class we had a returns file with 1 col of dates and 1 col of returns
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRankedSlidingVols
' Author     : Philip Swannell
' Date       : 31-Mar-2020
' Purpose    : Used in workbook "ISDA SIMM 2020 Global Stress Period Calculation.xlsm"
' Parameters :
'  ReturnsFile : A "pseudo returns file" of 10-day returns for the asset class in question, typically each return is the median for all assets in the class
'                File should have two columns - dates and returns
'  FromDate    : The start of the first one-year window
'  ToDate      : The start of the last one-year window, so at least a year before the last date in the file
'  DateFormat  : The date format used in the returns file. ISDA use either Y-M-D or M/D/Y
'  HeaderRowNum: the number of the header row in the ReturnsFile
'  WhatToReturn: String. Allowed values "Ranks","FromDates","Details"
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRankedSlidingVols(ReturnsFile As String, FromDate As Long, ToDate As Long, DateFormat As String, _
          Optional HeaderRowNum As Long = 1, Optional WhatToReturn As String = "Ranks")

          Dim FileContents
          Dim DatesInFile
          Dim ReturnsInFile

1         On Error GoTo ErrHandler
2         FileContents = ThrowIfError(sFileShow(ReturnsFile, , True, True, , , DateFormat))
3         If sNCols(FileContents) <> 2 Then
4             Throw "ReturnsFile should have two columns but it has " + CStr(sNCols(FileContents))
5         End If

6         DatesInFile = sSubArray(FileContents, HeaderRowNum + 1, 1, , 1)

7         ReturnsInFile = sSubArray(FileContents, HeaderRowNum + 1, 2, , 1)

8         ISDASIMMRankedSlidingVols = ISDASIMMRankedSlidingVolsCore(DatesInFile, ReturnsInFile, FromDate, ToDate, 12, WhatToReturn)

9         Exit Function
ErrHandler:
10        ISDASIMMRankedSlidingVols = "#ISDASIMMRankedSlidingVols (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'version for use in 2021 when we had file with ~ 30 columns and for each asset class we take median of relevant columns...
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRankedSlidingVols2021
' Author     : Philip Swannell
' Date       : 31-Jan-2022
' Purpose    :
' Parameters :
'  ReturnsFile          :
'  FromDate             :
'  ToDate               :
'  MonthsInWindow       :
'  DateFormat           :
'  BucketingData        :
'  AssetClass           :
'  HeaderRowNum         :
'  CSAWeightsFile       :
'  IndividualWeightsFile:
'  WhatToReturn         :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRankedSlidingVols2021(ReturnsFile As String, FromDate As Long, ToDate As Long, MonthsInWindow As Long, _
          DateFormat As String, BucketingFileOrData As Variant, AssetClass As String, HeaderRowNum As Long, _
          CSAWeightsFile As String, IndividualWeightsFile As String, Optional WhatToReturn As String = "Ranks", _
          Optional AllowWeekendVolStart As Boolean = True)

          Dim FileContents As Variant
          Dim HeadersInFile As Variant
          Dim TheDates As Variant
          Dim TheReturns As Variant
          Dim MappedAssetClasses
          Dim ColumnChooser
          Dim RelevantColumns
          Dim RelevantHeaders
          Dim i As Long
          Dim FXWeights
          Dim BucketingData

1         On Error GoTo ErrHandler

2         If DateFormat = "Guess" Then
3             DateFormat = ThrowIfError(ISDASIMMGuessDateFormat(ReturnsFile))
4         End If

5         FileContents = ThrowIfError(sFileShow(ReturnsFile, , True, True, , , DateFormat))

6         TheDates = sSubArray(FileContents, HeaderRowNum + 1, 1, , 1)

7         Select Case AssetClass
              Case "IR", "FX", "EQ", "CRQ", "CRNQ", "CM"
8             Case Else
9                 Throw "AssetClass '" & AssetClass & "'  not recognised.Alled values are: 'IR', 'FX', 'EQ', 'CRQ', 'CRNQ' and 'CM'"
10        End Select

11        HeadersInFile = sArrayTranspose(sSubArray(FileContents, HeaderRowNum, 2, 1))

12        If VarType(BucketingFileOrData) Then
13            BucketingData = ThrowIfError(sCSVRead(BucketingFileOrData, , , , , , , , 2))
14        Else
15            BucketingData = BucketingFileOrData
16        End If

17        Force2DArrayR BucketingData
18        If sNCols(BucketingData) <> 2 Then Throw "BucketingData must have two columns to map series headers to asset classes"

19        MappedAssetClasses = sVLookup(HeadersInFile, BucketingData)

20        For i = 1 To sNRows(MappedAssetClasses)
21            If sIsErrorString(MappedAssetClasses(i, 1)) Then
22                Throw "BucketingData does not provide an asset class for '" + HeadersInFile(i, 1) + "'"
23            End If
24        Next

          Dim UnrecognisedData
25        UnrecognisedData = sCompareTwoArrays(sSubArray(BucketingData, 1, 1, , 1), HeadersInFile, "In1AndNotIn2")
26        If sNRows(UnrecognisedData) > 1 Then
27            Throw "BucketingData provides a bucket for '" & UnrecognisedData(2, 1) & "' but that does not appear as a header in the returns file"
28        End If

29        ColumnChooser = sArrayEquals(MappedAssetClasses, AssetClass)

30        If Not sAny(ColumnChooser) Then Throw "No time series found for AssetClass = '" + AssetClass + "'"
31        RelevantHeaders = sMChoose(HeadersInFile, ColumnChooser)
32        RelevantColumns = sMChoose(sIntegers(sNRows(MappedAssetClasses)), ColumnChooser)
33        RelevantColumns = sArrayAdd(1, RelevantColumns)

34        If sArrayCount(ColumnChooser) = 1 Then
35            TheReturns = sSubArray(FileContents, HeaderRowNum + 1, RelevantColumns(1, 1), , 1)
36        ElseIf sArraysIdentical(RelevantColumns, sArrayAdd(RelevantColumns(1, 1) - 1, sIntegers(sNRows(RelevantColumns)))) Then
37            TheReturns = sSubArray(FileContents, HeaderRowNum + 1, RelevantColumns(1, 1), , sNRows(RelevantColumns))
38        Else
39            TheReturns = ThrowIfError(sIndex(FileContents, , sArrayTranspose(RelevantColumns)))
40            TheReturns = sSubArray(TheReturns, HeaderRowNum + 1)
41        End If

42        If AssetClass = "FX" Then
43            TheReturns = sArrayAbs(TheReturns)
44            FXWeights = ThrowIfError(ISDASIMMFxTradingWeights2020(RelevantHeaders, CSAWeightsFile, IndividualWeightsFile))
45            TheReturns = Application.WorksheetFunction.MMult(TheReturns, FXWeights)
46        ElseIf sNCols(TheReturns) > 1 Then
47            TheReturns = ThrowIfError(sRowMedian(TheReturns, False))
48        End If

49        If LCase(WhatToReturn) = LCase("Returns") Then
50            ISDASIMMRankedSlidingVols2021 = sArrayRange(TheDates, TheReturns)
51        Else
52            ISDASIMMRankedSlidingVols2021 = ISDASIMMRankedSlidingVolsCore(TheDates, TheReturns, FromDate, ToDate, MonthsInWindow, WhatToReturn, AllowWeekendVolStart)
53        End If

54        Exit Function
ErrHandler:
55        ISDASIMMRankedSlidingVols2021 = "#ISDASIMMRankedSlidingVols2021 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMGuessDateFormat
' Author     : Philip Swannell
' Date       : 06-Apr-2020
' Purpose    : ISDA change the date format they use in different files from year to year, but always one of the two formats
'              M/D/Y or Y-M-D or very occasionaly D/M/Y and dates are always in the first column. So this function does the guessing...
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMGuessDateFormat(FileNameOrArray As Variant)

          Const Format1 = "M/D/Y"
          Const Format2 = "Y-M-D"
          Const Format3 = "D/M/Y"
          Const FirstRow = 3
          Const LastRow = 60
          Dim Format1Works As Boolean
          Dim Format2Works As Boolean
          Dim Format3Works As Boolean

          Dim Contents
          Dim Parsed1, Parsed2, Parsed3

1         On Error GoTo ErrHandler
2         If VarType(FileNameOrArray) = vbString Then
3             Contents = ThrowIfError(sCSVRead(FileNameOrArray, False, , , , , , , FirstRow, 1, LastRow - FirstRow + 1, 1))
4         ElseIf IsArray(FileNameOrArray) Then
5             Contents = FileNameOrArray
6         Else
7             Throw "FileNameOrArray must be a string (giving a file name) or an array (containing strings that represent dates in to-be-determined format)"
8         End If

9         Parsed1 = sParseDate(Contents, Format1, True)
10        Parsed2 = sParseDate(Contents, Format2, True)
11        Parsed3 = sParseDate(Contents, Format3, True)

12        Format1Works = sAll(sArrayIsNumber(Parsed1))
13        Format2Works = sAll(sArrayIsNumber(Parsed2))
14        Format3Works = sAll(sArrayIsNumber(Parsed3))

15        If Format1Works And (Not Format2Works) And (Not Format3Works) Then
16            ISDASIMMGuessDateFormat = Format1
17        ElseIf (Not Format1Works) And Format2Works And (Not Format3Works) Then
18            ISDASIMMGuessDateFormat = Format2
19        ElseIf (Not Format1Works) And (Not Format2Works) And Format3Works Then
20            ISDASIMMGuessDateFormat = Format3
21        Else
22            Throw "Cannot determine date format"
23        End If

24        Exit Function
ErrHandler:
25        ISDASIMMGuessDateFormat = "#ISDASIMMGuessDateFormat (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMTheYear
' Author     : Philip Swannell
' Date       : 06-Apr-2020
' Purpose    : Returns the year of the calibration round, based on my workbook-naming convention.
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMTheYear()
1         On Error GoTo ErrHandler
2         Application.Volatile
          Dim BookName As String
3         If TypeName(Application.Caller) = "Range" Then
4             BookName = Application.Caller.Parent.Parent.Name
5             If Left(UCase(BookName), 10) = "ISDA SIMM " Then
6                 ISDASIMMTheYear = CLng(Mid(BookName, 11, 4))
7                 Exit Function
8             End If
9         End If
10        Throw "Cannot determine year"
11        Exit Function
ErrHandler:
12        ISDASIMMTheYear = "#ISDASIMMTheYear (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMM3YDates
' Author     : Philip Swannell
' Date       : 07-Apr-2020
' Purpose    : Utility function - returns the strat and end date of the three year period.
' Parameters :
'  TheYear:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMM3YDates(TheYear As Long, Optional StressPeriodCalcMethod As String, Optional AssetClass As String)
1         ISDASIMM3YDates = sArrayStack(ISDASIMM3YStart(TheYear, StressPeriodCalcMethod, AssetClass), ISDASIMM3YEnd(TheYear, StressPeriodCalcMethod))
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMM3YStart
' Author     : Philip Swannell
' Date       : 13-Mar-2021
' Purpose    :
' Parameters :
'  TheYear            :
'  StressPeriodCalcMethod: Allowed strings as of March 2021 = "1+3"  & "StressBalance"
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMM3YStart(TheYear As Long, Optional ByVal StressPeriodCalcMethod As String, Optional AssetClass As String)
          Dim Res

1         On Error GoTo ErrHandler

2         If TheYear <= 2020 Then
3             StressPeriodCalcMethod = ""
4         ElseIf TheYear >= 2022 Then
5             If StressPeriodCalcMethod = "" Then
6                 StressPeriodCalcMethod = "StressBalance"
7             End If
8         End If

9         If StressPeriodCalcMethod = "" Then
10            Res = DateSerial(TheYear - 3, 1, 1)
11            While Res Mod 7 <= 1
12                Res = Res + 1
13            Wend
14            ISDASIMM3YStart = CLng(Res)
15            Exit Function
16        Else
17            If AssetClass = "" Then Throw "AssetClass must be provided when StressPeriodCalcMethod is provided"
              'Ensure that recent period and stress period sum to 48 months
              Dim StressDates As Variant
              Dim StressStart As Long
              Dim StressEnd As Long
              Dim StressMonths As Long
              Dim RecentMonths
              
18            StressDates = ThrowIfError(ISDASIMMStressDates(TheYear, AssetClass, StressPeriodCalcMethod))
19            StressStart = StressDates(1, 1)
20            StressEnd = StressDates(2, 1)
21            StressMonths = 12 * Year(StressEnd) + Month(StressEnd) - (12 * Year(StressStart) + Month(StressStart))
22            RecentMonths = 48 - StressMonths
23            Res = DateSerial(TheYear, 1 - RecentMonths, 1)
24            While Res Mod 7 <= 1
25                Res = Res + 1
26            Wend
27            ISDASIMM3YStart = CLng(Res)
28        End If
29        Exit Function
ErrHandler:
30        ISDASIMM3YStart = "#ISDASIMM3YStart (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMM3YEnd(TheYear As Long, Optional ByVal StressPeriodCalcMethod As String) 'StressPeriodCalcMethod not actually used
          Dim Res
1         On Error GoTo ErrHandler

2         Res = DateSerial(TheYear - 1, 12, 31)
3         While Res Mod 7 <= 1
4             Res = Res - 1
5         Wend
6         ISDASIMM3YEnd = CLng(Res)
7         Exit Function

8         Exit Function
ErrHandler:
9         ISDASIMM3YEnd = "#ISDASIMM3YEnd (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMMStressDates(TheYear As Long, ByVal AssetClass As String, Optional ByVal StressPeriodCalcMethod As String = "", Optional RecentPeriod As Boolean = False)
          Dim StressStart
          Dim StressEnd
          Dim Table As Range
          Dim Key As String
          Dim RowNum As Variant
          Dim NumMonthsInFirstPeriod As Long
          
1         On Error GoTo ErrHandler
          'Standardise
          
2         If TheYear <= 2020 Then
3             StressPeriodCalcMethod = ""
4         ElseIf TheYear >= 2022 Then '2022 use only StressBalance
5             If LCase(StressPeriodCalcMethod) = "stressbalance" Then
6                 StressPeriodCalcMethod = ""
7             ElseIf StressPeriodCalcMethod <> "" Then
8                 Throw "StressPeriodCalcMethod must be 'StressBalance' or be omitted"
9             End If
10        End If
          
11        AssetClass = ISDASIMMStandardiseAssetClass(AssetClass)
12        Select Case AssetClass
              Case "CCB"
13                AssetClass = "IR" 'Cross currency basis uses same stress dates as IR
14            Case "BC"
15                AssetClass = "CRQ" 'Base corr uses same stress dates as CRQ
16        End Select

17        If StressPeriodCalcMethod = "" Then
18            Key = TheYear & "-" & AssetClass
19        Else
20            Key = TheYear & "-" & AssetClass & "-" & StressPeriodCalcMethod
21        End If

22        Set Table = shISDASIMM.Range("ISDASIMMStressPeriods")
23        RowNum = sMatch(Key, Table.Columns(1).Value)
24        If Not IsNumber(RowNum) Then
25            Throw "Cannot find key '" + Key + "' in left column of range ISDASIMMStressPeriods on sheet " + shISDASIMM.Name + " of workbook " + ThisWorkbook.Name
26        End If
27        If RecentPeriod Then
28            StressStart = Table.Cells(RowNum, 2).Value
29            StressEnd = Table.Cells(RowNum, 3).Value
30            NumMonthsInFirstPeriod = 12 * Year(StressEnd) + Month(StressEnd) - 12 * Year(StressStart) - Month(StressStart)
31            If NumMonthsInFirstPeriod >= 12 Then
                  Dim ErrMsg As String
32                ErrMsg = "No recent stress period for AssetClass = " + AssetClass + ", Year = " + CStr(TheYear) + ", StressPeriodCalcMethod = " & StressPeriodCalcMethod
33                ISDASIMMStressDates = sArrayStack(ErrMsg, ErrMsg) 'Turns out to be convenient to return two error messages - for design of workbook Correlation Generator
34                Exit Function
35            End If
36            StressStart = Table.Cells(RowNum, 4).Value
37            StressEnd = Table.Cells(RowNum, 5).Value
38        Else
39            StressStart = Table.Cells(RowNum, 2).Value
40            StressEnd = Table.Cells(RowNum, 3).Value
41        End If
42        ISDASIMMStressDates = sArrayStack(StressStart, StressEnd)
43        Exit Function

44        Exit Function
ErrHandler:
45        ISDASIMMStressDates = "#ISDASIMMStressDates (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMResultsFileName
' Author     : Philip Swannell
' Date       : 13-Mar-2021
' Purpose    : "Guess" the file name thast ISDA use for their results files
' Parameters :
'  Folder             :
'  FirstPart          :
'  ReturnLag          :
'  StressPeriodCalcMethod:
'  extension          :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMResultsFileName(Folder As String, FirstPart As String, ReturnLag As String, StressPeriodCalcMethod As String, Optional Extension As String = ".csv")
          Dim FileName As String

1         On Error GoTo ErrHandler
2         FileName = FirstPart

3         Select Case StressPeriodCalcMethod
              Case "1+3"
4                 FileName = FileName + "_recent_0"
5             Case "StressBalance"
6                 FileName = FileName + "_recent_1"
7             Case ""
                  'This case applies for 2020 and maybe for 2022, we shall see...
8             Case Else
9                 Throw "StressPeriodCalcMethod not recognised. Allowed values are '1+3' and 'StressBalance'"
10        End Select

11        Select Case ReturnLag
              Case "10DayNaive", "10Day"
12                FileName = FileName + "-10d"
13            Case "1Day"
14                FileName = FileName + "-1d"
15            Case ""
                  'Sometimes ReturnLag indicator does not appear in the file name
16            Case Else
17                Throw "ReturnLag not recognised"
18        End Select

19        FileName = FileName + Extension

20        ISDASIMMResultsFileName = CoreJoinPath(Folder, FileName)

21        Exit Function
ErrHandler:
22        ISDASIMMResultsFileName = "#ISDASIMMResultsFileName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMMMakeID(ByVal AssetClass, ByVal Parameter, Optional ByVal Label1 = "", Optional ByVal Label2 = "", Optional ByVal Label3 = "", Optional Lag = 0)

1         On Error GoTo ErrHandler
2         If VarType(AssetClass) < vbArray Then
3             If VarType(Parameter) < vbArray Then
4                 If VarType(Label1) < vbArray Then
5                     If VarType(Label2) < vbArray Then
6                         If VarType(Label3) < vbArray Then
7                             If VarType(Lag) < vbArray Then
8                                 ISDASIMMMakeID = ISDASIMMMakeID_Core(AssetClass, Parameter, Label1, Label2, Label3, CLng(Lag))
9                                 Exit Function
10                            End If
11                        End If
12                    End If
13                End If
14            End If
15        End If

16        ISDASIMMMakeID = Broadcast(FuncIDISDASIMMMakeID, AssetClass, Parameter, Label1, Label2, Label3, Lag)

17        Exit Function
ErrHandler:
18        Throw "#ISDASIMMMakeID (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRoundingRule2022
' Author     : Philip Swannell
' Date       : 22-Mar-2021
' Purpose    : For use in the Summary workbook, gets the 2022-style rounding method from the ID of a result.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRoundingRule2022(ByVal ID As Variant)
          Dim NR As Long, NC As Long, Result() As Variant, i As Long, j As Long
1         On Error GoTo ErrHandler
2         If VarType(ID) < vbArray Then
3             ISDASIMMRoundingRule2022 = ISDASIMMRoundingRule2022_Core(CStr(ID))
4             Exit Function
5         Else
6             Force2DArrayR ID, NR, NC
7             ReDim Result(1 To NR, 1 To NC)
8             For i = 1 To NR
9                 For j = 1 To NC
10                    Result(i, j) = ISDASIMMRoundingRule2022_Core(CStr(ID(i, j)))
11                Next j
12            Next i
13            ISDASIMMRoundingRule2022 = Result
14        End If
15        Exit Function
ErrHandler:
16        ISDASIMMRoundingRule2022 = "#ISDASIMMRoundingRule2022 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRoundingRule2022_Core
' Author     : Philip Swannell
' Date       : 10-Feb-2022
' Purpose    : See document emailed by Paola and copied to
'  C:\Users\phili\Solum Financial Limited\Shared Documents - Documents\ISDA\2022\DocumentsReceived\Calibration Rounding Rules_20211214.pdf
' Parameters :
'  ID:  an ID - see function ISDASIMMMakeID
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISDASIMMRoundingRule2022_Core(ID As String)
          Dim parts() As String
          Dim AssetClass As String
          Dim Parameter As String

1         On Error GoTo ErrHandler
2         If InStr(ID, "-") = 0 Then Throw "ID '" & ID & "' not recognised"

3         If InStr(ID, "-") = 0 Then
4             Throw "ID '" & ID & "' not recognised"
5         End If

6         parts = VBA.Split(ID, "-")
7         AssetClass = parts(0)
8         Parameter = parts(1)

9         Select Case LCase(AssetClass & "-" & Parameter)
              Case "bc-drw"
10                ISDASIMMRoundingRule2022_Core = "A"
11            Case "bc-inter family corr"
12                ISDASIMMRoundingRule2022_Core = "E"
13            Case "bc-stress period"
14                ISDASIMMRoundingRule2022_Core = "Date"
15            Case "ccb-corr"
16                ISDASIMMRoundingRule2022_Core = "E"
17            Case "ccb-drw"
18                ISDASIMMRoundingRule2022_Core = "A"
19            Case "cm-corr"
20                ISDASIMMRoundingRule2022_Core = "E"
21            Case "cm-drw"
22                ISDASIMMRoundingRule2022_Core = "A"
23            Case "cm-hvr"
24                ISDASIMMRoundingRule2022_Core = "A"
25            Case "cm-stress period"
26                ISDASIMMRoundingRule2022_Core = "Date"
27            Case "cm-vrw"
28                ISDASIMMRoundingRule2022_Core = "A"
29            Case "crnq-drw"
30                ISDASIMMRoundingRule2022_Core = "C"
31            Case "crnq-inter corr"
32                ISDASIMMRoundingRule2022_Core = "E"
33            Case "crnq-intra corr"
34                ISDASIMMRoundingRule2022_Core = "E"
35            Case "crnq-stress period"
36                ISDASIMMRoundingRule2022_Core = "Date"
37            Case "crq-drw"
38                ISDASIMMRoundingRule2022_Core = "A"
39            Case "crq-inter corr"
40                ISDASIMMRoundingRule2022_Core = "E"
41            Case "crq-intra corr"
42                ISDASIMMRoundingRule2022_Core = "E"
43            Case "crq-stress period"
44                ISDASIMMRoundingRule2022_Core = "Date"
45            Case "crq-vrw"
46                ISDASIMMRoundingRule2022_Core = "A"
47            Case "eq-corr"
48                ISDASIMMRoundingRule2022_Core = "E"
49            Case "eq-drw"
50                ISDASIMMRoundingRule2022_Core = "A"
51            Case "eq-hvr"
52                ISDASIMMRoundingRule2022_Core = "A"
53            Case "eq-stress period"
54                ISDASIMMRoundingRule2022_Core = "Date"
55            Case "eq-vrw"
56                ISDASIMMRoundingRule2022_Core = "A"
57            Case "fx-corr"
58                ISDASIMMRoundingRule2022_Core = "E"
59            Case "fx-drw"
60                ISDASIMMRoundingRule2022_Core = "A"
61            Case "fx-hvr"
62                ISDASIMMRoundingRule2022_Core = "A"
63            Case "fx-stress period"
64                ISDASIMMRoundingRule2022_Core = "Date"
65            Case "fx-vrw"
66                ISDASIMMRoundingRule2022_Core = "A"
67            Case "ir-drw"
68                ISDASIMMRoundingRule2022_Core = "A"
69            Case "ir-hvr"
70                ISDASIMMRoundingRule2022_Core = "A"
71            Case "ir-inflation corr"
72                ISDASIMMRoundingRule2022_Core = "E"
73            Case "ir-inter corr"
74                ISDASIMMRoundingRule2022_Core = "E"
75            Case "ir-intra corr"
76                ISDASIMMRoundingRule2022_Core = "E"
77            Case "ir-stress period"
78                ISDASIMMRoundingRule2022_Core = "Date"
79            Case "ir-sub curve corr"
80                ISDASIMMRoundingRule2022_Core = "F"
81            Case "ir-vrw"
82                ISDASIMMRoundingRule2022_Core = "A"
83            Case "xa-corr"
84                ISDASIMMRoundingRule2022_Core = "E"
85            Case "xa-stress period"
86                ISDASIMMRoundingRule2022_Core = "Date"
87            Case Else
88                Throw "ID '" & ID & "' not recognised"
89        End Select

90        Exit Function
ErrHandler:
91        ISDASIMMRoundingRule2022_Core = "#ISDASIMMRoundingRule2022_Core (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function ISDASIMMApplyRounding2022_Core(InputValue As Variant, ByVal RuleOrID As String)
          Dim Rule As String
          Dim Res

1         On Error GoTo ErrHandler
2         If LCase(RuleOrID) = "date" Then
3             Rule = "Date"
4         ElseIf Len(RuleOrID) > 1 Then
5             Rule = ThrowIfError(ISDASIMMRoundingRule2022_Core(RuleOrID))
6         Else
7             Rule = RuleOrID
8         End If

9         If IsEmpty(InputValue) Then
10            ISDASIMMApplyRounding2022_Core = "NA"
11            Exit Function
12        ElseIf Not IsNumberOrDate(InputValue) Then
13            ISDASIMMApplyRounding2022_Core = InputValue
14            Exit Function
15        End If

16        Select Case UCase(Rule)

              Case "DATE"
17                Res = CDate(CLng(InputValue))
18            Case "A"
                  '• If the number is less than 10, round to two significant figures (e.g. 9.445 =>9.5)
                  '• If the number is equal or greater than 10, then round to the nearest integer (e.g. 10.45 => 10)
                  '• Halves are rounded up (e.g. 4.55 => 4.6
19                If InputValue < 10 Then
20                    Res = sRoundSF(InputValue, 2, 0)
21                Else
22                    Res = sRound(InputValue, 0, 0)
23                End If
24            Case "B"
                  '• For all numbers, round to the nearest integer (e.g. 10.23 => 10)
                  '• Halves are rounded up (e.g. 10.5 => 11
25                Res = sRound(InputValue, 0, 0)
26            Case "C"
                  '• For all numbers, round to two significant figures (e.g. 1045=> 1000)
                  '• Halves are rounded up (e.g. 1050 => 1100)
27                Res = sRoundSF(InputValue, 2, 0)
28            Case "D"
                  '• For all numbers, round to the nearest 10-1 digit (e.g. 10.2354 => 10.2)
                  '• Halves are rounded up (e.g. 10.55 => 10.6)
29                Res = sRound(InputValue, 1, 0)
30            Case "E"
                  '• For all numbers, round to the nearest 10-2 digit (e.g. 10.2354 => 10.24)
                  '• Halves are rounded up (e.g. 10.555 => 10.56)
31                Res = sRound(InputValue, 2, 0)
32            Case "F"
                  '• For all numbers, round to the nearest 10-3 digit (e.g. 10.2354 => 10.235)
                  '• Halves are rounded up (e.g. 10.5555 => 10.556
33                Res = sRound(InputValue, 3, 0)
34            Case "G"
                  '• For all numbers, round to two significant figures (e.g. 1045=> 1000, 9.461=> 9.5)
                  '• Halves are rounded down (e.g. 1050 => 1000)
35                Res = sRoundSF(InputValue, 2, 1)
36            Case Else
37                Throw "Unrecognised Rule '" & Rule & "'"
38        End Select

39        ISDASIMMApplyRounding2022_Core = Res

40        Exit Function
ErrHandler:
41        ISDASIMMApplyRounding2022_Core = "#ISDASIMMApplyRounding2022_Core (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMApplyRounding2022
' Author     : Philip Swannell
' Date       : 10-Feb-2022
' Purpose    : For use from the Summary workbook, using the better-defined rounding methods for 2022
' Parameters :
'  InputValue  :
'  RoundingRule:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMApplyRounding2022(ByVal InputValue, ByVal RoundingRule)
1         On Error GoTo ErrHandler
2         If VarType(InputValue) < vbArray And VarType(RoundingRule) < vbArray Then
3             ISDASIMMApplyRounding2022 = ISDASIMMApplyRounding2022_Core(InputValue, CStr(RoundingRule))
4             Exit Function
5         Else
6             ISDASIMMApplyRounding2022 = Broadcast(FuncIdISDASIMMApplyRounding2022, InputValue, RoundingRule)
7         End If
8         Exit Function
ErrHandler:
9         ISDASIMMApplyRounding2022 = "#ISDASIMMApplyRounding2022 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRoundingMethod
' Author     : Philip Swannell
' Date       : 22-Mar-2021
' Purpose    : For use in the Summary workbook, gets the rounding method from the ID of a result.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRoundingMethod(ByVal ID As Variant)
          Dim NR As Long, NC As Long, Result() As Variant, i As Long, j As Long
1         On Error GoTo ErrHandler
2         If VarType(ID) < vbArray Then
3             ISDASIMMRoundingMethod = ISDASIMMRoundingMethod_Core(CStr(ID))
4             Exit Function
5         Else
6             Force2DArrayR ID, NR, NC
7             ReDim Result(1 To NR, 1 To NC)
8             For i = 1 To NR
9                 For j = 1 To NC
10                    Result(i, j) = ISDASIMMRoundingMethod_Core(CStr(ID(i, j)))
11                Next j
12            Next i
13            ISDASIMMRoundingMethod = Result
14        End If
15        Exit Function
ErrHandler:
16        ISDASIMMRoundingMethod = "#ISDASIMMRoundingMethod (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRoundingMethod_Core
' Author     : Philip Swannell
' Date       : 22-Mar-2021
' Purpose    : Return a description of the rounding method from the unique id
' Parameters :
'  ID:
' -----------------------------------------------------------------------------------------------------------------------
Private Function ISDASIMMRoundingMethod_Core(ByVal ID As String)
          Dim AssetClass As String, Parameter As String, Labels As String
          Dim parts
          Dim rm As String, Lag As String
          Dim IDReconstructed As String

1         On Error GoTo ErrHandler
2         parts = VBA.Split(ID, "-")
3         AssetClass = parts(0)
4         Parameter = parts(1)
5         If UBound(parts) >= 2 Then
              'Mmmm some IDs don't contain labels
6             If parts(2) <> "10d" And parts(2) <> "1d" Then
7                 Labels = parts(2)
8             End If
9         End If
10        If Right(ID, 3) = "10d" Then
11            Lag = "10d"
12        ElseIf Right(ID, 2) = "1d" Then
13            Lag = "1d"
14        End If

15        IDReconstructed = AssetClass & "-" & Parameter & IIf(Labels <> "", "-" & Labels, "") & IIf(Lag <> "", "-" & Lag, "")

16        If IDReconstructed <> ID Then
17            Throw "Parsing Failure"
18        End If
        
          'The "rules" below replicate the hard-wired values used in the summary workbook sheet prior _
           to writing this function. But can probably be simplified to make more use of 2sf...
        
19        If Parameter = "stress period" Then
20            rm = "Date"
21        ElseIf InStr(Parameter, "corr") > 0 Then
              'all the styles of correlation are 2dp
22            If LCase(Parameter) = "sub curve corr" Then
23                rm = "3dp"
24            Else
25                rm = "2dp"
26            End If
27        ElseIf Parameter = "drw" Then
              'conventions for Delta Risk Weight vary by asset class
28            Select Case AssetClass
                  Case "ir"
29                    If Lag = "10d" Then
30                        rm = "0dp"
31                    ElseIf Lag = "1d" Then
32                        rm = "2sf"
33                    End If
34                Case "eq"
35                    If Lag = "10d" Then
36                        rm = "0dp"
37                    ElseIf Lag = "1d" Then
38                        rm = "2sf"
39                    End If
40                Case "crq"
41                    rm = "0dp"
42                Case "fx"
43                    If Lag = "10d" Then
44                        rm = "1dp"
45                    ElseIf Lag = "1d" Then
46                        rm = "2dp"
47                    End If
48                Case "ccb", "cm"
49                    rm = "1dp"
50                Case "bc"
51                    rm = "2sf"
52                Case "crnq"
53                    Select Case Labels
                          Case "1"
54                            rm = "2sf"
55                        Case "2"
56                            rm = "2sf"
57                    End Select
58            End Select
59        ElseIf Parameter = "vrw" Then
              'Vega Risk Weights 2dp except for FX
60            Select Case AssetClass
                  Case "fx"
61                    rm = "3dp"
62                Case "eq", "crq"
63                    If Lag = "10d" Then
64                        rm = "2dp"
65                    ElseIf Lag = "1d" Then
66                        rm = "2sf"
67                    End If
68                Case Else
69                    rm = "2dp"
70            End Select
71        ElseIf Parameter = "hvr" Then
              'Historical Volatility Ratio 2dp except for FX
72            Select Case AssetClass
                  Case "fx"
73                    rm = "3dp"
74                Case Else
75                    rm = "2dp"
76            End Select
77        End If

78        If rm = "" Then
79            rm = "Unknown rounding method for ID = '" & ID & "'"
80        End If

81        ISDASIMMRoundingMethod_Core = rm

82        Exit Function
ErrHandler:
83        ISDASIMMRoundingMethod_Core = "#ISDASIMMRoundingMethod_Core (line " & CStr(Erl) + "): " & Err.Description & "!"
84    End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMDataFolder
' Author     : Philip Swannell
' Date       : 14-Mar-2021
' Purpose    : Returns either a Folder to which ISDA write data or a file name to which isda write results
' This function tries to duplicate ISDA's "operating procedures" that change quite a bit year to year (groan)
' Parameters :
'  TheYear               :
'  SubPath1              :
'  SubPath2              :
'  PartialFileName       :
'  ReturnLag             :
'  StressPeriodCalcMethod:
'  Extension             :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMDataFolder(TheYear As Long, Optional SubPath1 As String, Optional SubPath2 As String, _
          Optional PartialFileName As String, Optional ReturnLag As String, Optional ByVal StressPeriodCalcMethod As String, _
          Optional Extension As String = ".csv")

1         On Error GoTo ErrHandler
2         Application.Volatile
          Dim BaseFolder As String
          Dim Result As String

3         If TheYear <= 2020 Then
4             StressPeriodCalcMethod = ""
5         ElseIf TheYear >= 2022 Then
              'It's a mistake to be passing in the no-longer used 1+3
6             If StressPeriodCalcMethod = "1+3" Then Throw "StressPeriodCalcMethod of 1+3 is not supported for Year >= 2022"
              '... and we don't want function ISDASIMMResultsFile (which does not receive TheYear as an argument) to create _
               the extra layer in the folder structure that was necessary in 2021
            '  StressPeriodCalcMethod = ""
7         End If

8         Select Case TheYear
              Case 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027
                  Dim Roman As String
9                 Roman = Application.WorksheetFunction.Roman(TheYear - 2015)
10                BaseFolder = sJoinPath(sEnvironmentVariable("OneDriveConsumer"), "ISDA SIMM\Solum Validation C-" & Roman & " " & TheYear)
11            Case Else
12                Throw "No base folder defined for TheYear = " & CStr(TheYear)
13        End Select
14        If Not sFolderExists(BaseFolder) Then Throw "Cannot find folder '" + BaseFolder + "'"

15        Result = CoreJoinPath(BaseFolder, SubPath1, SubPath2)
16        If PartialFileName <> "" Then
17            Result = ThrowIfError(ISDASIMMResultsFileName(Result, PartialFileName, ReturnLag, StressPeriodCalcMethod, Extension))
18        End If

          'Arrgh Xiaowei is not consistent in the file names for result file!
          'See C:\ISDA SIMM\SolumWorking\2021\AnalysisOfXiaoweisFileNamingConvention.xlsm
19        If TheYear >= 2021 Then
20            If LCase(Right(Result, 4)) = ".csv" Then
                  Dim Result2
21                Result2 = LCase(Result)
22                Result2 = Replace(Result2, "_recent_0-10d", "-10d_recent_0")
23                Result2 = Replace(Result2, "_recent_1-10d", "-10d_recent_1")
24                Result2 = Replace(Result2, "_recent_0-1d", "-1d_recent_0")
25                Result2 = Replace(Result2, "_recent_1-1d", "-1d_recent_1")
26                If LCase(Result2) <> LCase(Result) Then
27                    If Not sFileExists(Result) Then
28                        If sFileExists(Result2) Then
29                            Result = Result2
30                        End If
31                    End If
32                End If
33            End If
34        End If

35        ISDASIMMDataFolder = Result

36        Exit Function
ErrHandler:
37        ISDASIMMDataFolder = "#ISDASIMMDataFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMSaveResults
' Author     : Philip Swannell
' Date       : 07-Apr-2020
' Purpose    : Utility function to be called from ISDA SIMM workbooks. Saves data for both Solum's calculations and
'              ISDA's in a standardised format with quite a lot of validation of the input data. Unusually for a
'              spreadsheet function it can post a message box. Function saves a file containing the Data passed in
'              (but with the labels in the first column morphed). MsgBox is posted only if file already exists with different data.
' Parameters :
'  TheYear   : The year of the run, e.g. calibration IV takes 2020
'  AssetClass: One of "IR", "FX", "EQ", "CRQ", "CRNQ" or "CM" - or their pseudonyms such as "Equity" for "EQ"
'  Parameter : Example "DRW" or "DeltaRiskWeight" - see allowed values in second token of return from ISDASIMMValidIDs
'  StressPeriodCalcMethod. New (Arrgh) for 2021, One of "1+3" or "StressBalance"
'  Data      : Three column array with header row. Header row must read {"Element", "Solum", "ISDA"}
' -----------------------------------------------------------------------------------------------------------------------
'NB this function must have the same signature as ISDASIMMResultsFileNameSolum except with the added Data argument. Code in workbook "ISDASIMM ???? Control.xlsm" relies on that.
Function ISDASIMMSaveResults(TheYear As Long, AssetClass As String, Parameter As String, _
          StressPeriodCalcMethod As String, ByVal Data As Variant)
          
          Dim FileName As String
          Dim i As Long, j As Long
          Dim Label1 As String
          Dim Lag As Long
          Dim MatchRes
          Dim OrigData, NR As Long, NC As Long
          
1         On Error GoTo ErrHandler
2         Force2DArrayR Data, NR, NC
3         OrigData = Data
4         If sNCols(Data) <> 3 Then
5             Throw "Data must have 3 columns"
6         End If
7         If Data(1, 1) <> "Element" Then
8             Throw "Element 1,1 of Data must read 'Element'"
9         ElseIf Data(1, 2) <> "Solum" Then
10            Throw "Element 1,2 of Data must read 'Solum'"
11        ElseIf Data(1, 3) <> "ISDA" Then
12            Throw "Element 1,3 of Data must read 'ISDA'"
13        End If
          
14        For i = 2 To sNRows(Data)
15            If LCase(Data(i, 1)) = "10d" Then
16                Label1 = ""
17                Lag = 10
18            ElseIf LCase(Data(i, 1)) = "1d" Then
19                Label1 = ""
20                Lag = 1
21            ElseIf InStr(Data(i, 1), "-") = 0 Then
22                Label1 = Data(i, 1)
23                Lag = 10
24            Else
25                If Right(Data(i, 1), 4) = "-10d" Then
26                    Label1 = Left(Data(i, 1), Len(Data(i, 1)) - 4)
27                    Lag = 10
28                ElseIf Right(Data(i, 1), 3) = "-1d" Then
29                    Label1 = Left(Data(i, 1), Len(Data(i, 1)) - 3)
30                    Lag = 1
31                Else
32                    Label1 = Data(i, 1)
33                    Lag = 10
34                End If
35            End If
36            Data(i, 1) = ThrowIfError(ISDASIMMMakeID_Core(AssetClass, Parameter, Label1, , , Lag))
37        Next i
          Dim AllValidIDs
38        AllValidIDs = ISDASIMMValidIDs(True)
        
39        MatchRes = sMatch(sSubArray(Data, 1, 1, , 1), AllValidIDs, True)
40        For i = 2 To sNRows(Data)
41            If Not IsNumber(MatchRes(i, 1)) Then
42                Throw "Invalid ID '" + OrigData(i, 1) + "' at Data(" + CStr(i) + ",1). Invalid because it 'expands to' '" + Data(i, 1) + "' which does not match any element of ISDASIMMValidIDs()"
43            End If
44        Next i

          Dim IDsToSave
45        IDsToSave = sSubArray(Data, 2, 1, , 1)
46        If sNRows(IDsToSave) <> sNRows(sRemoveDuplicates(IDsToSave)) Then
              Dim tmp
47            tmp = sCountDistinctItems(IDsToSave)
48            Throw "Duplicate IDs found in data to save, for example '" + tmp(1, 1) + "' appears " + CStr(tmp(1, 2)) + " times."
49        End If

50        For i = 2 To NR
51            For j = 2 To NC
52                If Not IsNumber(Data(i, j)) Then
53                    If VarType(Data(i, j)) <> vbString Then
54                        Throw "Invalid values in Data - must be numbers or the string 'NA'"
55                    ElseIf Data(i, j) <> "NA" Then
56                        Throw "Invalid values in Data - must be numbers or the string 'NA'"
57                    End If
58                End If
59            Next j
60        Next i

61        FileName = ThrowIfError(ISDASIMMResultsFileNameSolum(TheYear, AssetClass, Parameter, StressPeriodCalcMethod))
62        ThrowIfError sCreateFolder(sSplitPath(FileName, False))
        
          Dim FileExists, oldFileContents
63        FileExists = sFileExists(FileName)
          Dim DoSave As VbMsgBoxResult
          Dim Prompt As String
          Dim FileName2 As String
64        FileName2 = FileName & "_Backup_" & Format$(Now, "yyyy-mm-dd-hh-mm-ss")

65        If Not FileExists Then
66            ISDASIMMSaveResults = ThrowIfError(sFileSave(FileName, Data, Chr(9)))
67            ThrowIfError sFileSave(FileName2, Data, Chr(9))
68        Else
69            oldFileContents = sFileShow(FileName, Chr(9), True)
70            If sArraysNearlyIdentical(Data, oldFileContents, , 0.000000000001) Then
                  'Save again - updates the timestamp.
71                ISDASIMMSaveResults = ThrowIfError(sFileSave(FileName, Data, Chr(9)))
72            Else
73                Prompt = "File '" + FileName + "' already exists with different contents." + vbLf + vbLf + _
                      "Overwrite?"
74                DoSave = MsgBoxPlus(Prompt, vbYesNo + vbQuestion + vbDefaultButton1, "ISDA SIMM Save Results", "Overwrite", "Cancel", , , , , , 30, vbYes)
75                If DoSave = vbYes Then
76                    ISDASIMMSaveResults = ThrowIfError(sFileSave(FileName, Data, Chr(9)))
77                    ThrowIfError sFileSave(FileName2, Data, Chr(9))
78                Else
79                    ISDASIMMSaveResults = "#Saving '" + FileName + "' ABORTED BY USER!"
80                End If
81            End If
82        End If

83        Exit Function
ErrHandler:
84        ISDASIMMSaveResults = "#ISDASIMMSaveResults (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMMValidIDs(Sorted As Boolean)
          Static SortedResult
1         On Error GoTo ErrHandler
2         If Sorted Then
3             If IsEmpty(SortedResult) Or sIsErrorString(SortedResult) Then
                  'Note that need to have UseExcelSortMethod passed as False, for compatibility with sMatch
4                 SortedResult = sSortedArray(shISDASIMM.Range("ValidIDs").Value, , , , , , , , False)
5             End If
6             ISDASIMMValidIDs = SortedResult
7         Else
8             ISDASIMMValidIDs = shISDASIMM.Range("ValidIDs").Value
9         End If
10        Exit Function
ErrHandler:
11        Throw "#ISDASIMMValidIDs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMWorkingFolder
' Author     : Philip Swannell
' Date       : 20-Mar-2021
' Purpose    : Return as location for saving intermediate working files
' Parameters :
'  TheYear               :
'  AssetClass            :
'  StressPeriodCalcMethod:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMWorkingFolder(TheYear As Long, AssetClass As String, StressPeriodCalcMethod As String)
1         On Error GoTo ErrHandler
2         Select Case StressPeriodCalcMethod
              Case "1+3", "StressBalance"
3             Case Else
4                 Throw "StressPeriodCalcMethod not recognised. Allowed values are '1+3' and 'StressBalance'"
5         End Select
6         ISDASIMMWorkingFolder = ThrowIfError(sJoinPath(ISDASIMMDataFolder(TheYear), "..", "SolumWorking", TheYear, StressPeriodCalcMethod, ThrowIfError(ISDASIMMStandardiseAssetClass(AssetClass))))
7         Exit Function
ErrHandler:
8         ISDASIMMWorkingFolder = "#ISDASIMMWorkingFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMResultsFileNameSolum
' Author     : Philip Swannell
' Date       : 14-Mar-2021
' Purpose    : The location of the file which we generate containing a results for both Solum and ISDA, and picked up by workbook "ISDA SIMM ???? Workbook Summary.xlsm"
' See also ISDASIMMResultsFileName that gives the location of where ISDA write results
' Parameters :
'  TheYear               :
'  AssetClass            :
'  Parameter             :
'  StressPeriodCalcMethod:
' -----------------------------------------------------------------------------------------------------------------------
'NB this function must have the same signature as ISDASIMMSaveResults except without the final Data argument. Code in workbook "ISDASIMM ???? Control.xlsm" relies on that.
Function ISDASIMMResultsFileNameSolum(TheYear As Long, AssetClass As String, Parameter As String, Optional ByVal StressPeriodCalcMethod As String)
          Dim Res As String, SubPath1 As String
1         Application.Volatile
2         On Error GoTo ErrHandler

          'Prior to 2021 we used "1+3" as the only StressPeriodCalcmethod, in 2021 we ran with two alternatives, _
           "1+3" and "StressBalance", and in 2022 we run with "StressBalance" only.
          
3         If TheYear < 2021 Then
4             StressPeriodCalcMethod = ""
5         End If

6         If TheYear >= 2022 Then
7             If StressPeriodCalcMethod = "1+3" Then Throw "StressPeriodCalcMethod of 1+3 is not supported for Year >= 2022"
8             StressPeriodCalcMethod = ""
9         End If

10        SubPath1 = "SolumResults"
11        Select Case StressPeriodCalcMethod
              Case "1+3", "StressBalance", ""
12            Case Else
13                Throw "StressPeriodCalcMethod not recognised. Allowed values are '1+3' and 'StressBalance' or omitted"
14        End Select

15        Res = ThrowIfError(sJoinPath(ISDASIMMDataFolder(TheYear), "..", "SolumResults", TheYear, StressPeriodCalcMethod))
16        ISDASIMMResultsFileNameSolum = CoreJoinPath(Res, ThrowIfError(ISDASIMMMakeID_Core(AssetClass, Parameter, , , , 0)) & ".csv")

17        Exit Function
ErrHandler:
18        Throw "#ISDASIMMResultsFileNameSolum (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMGetResults
' Author     : Philip Swannell
' Date       : 08-Apr-2020
' Purpose    : Grabs the contents of files written by calls to ISDASIMMSaveResults
' Parameters :
'  TheYear:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMGetResults(TheYear As Long, ResultsSubFolder As String)
          Dim Folder As String
          Dim AllFiles As Variant
          Dim STK As clsStacker
          Dim ThisChunk
          Dim ChunkGood As Boolean
          Dim ValidLabels
          Dim i As Long
          Dim DirListRet
          Static LastDirListRet
          Static LastTheYear As Long
          Static LastResultsSubFolder As String
          Static LastReturn As Variant

1         On Error GoTo ErrHandler
2         Folder = ThrowIfError(sJoinPath(ISDASIMMDataFolder(TheYear), "..", "SolumResults", TheYear, ResultsSubFolder))
3         DirListRet = ThrowIfError(sDirList(Folder, False, False, "FSCM#", , "*.csv"))
          'Memoise
4         If TheYear = LastTheYear Then
5             If ResultsSubFolder = LastResultsSubFolder Then
6                 If Not IsEmpty(LastReturn) Then
7                     If sArraysIdentical(DirListRet, LastDirListRet) Then
8                         ISDASIMMGetResults = LastReturn
9                         Exit Function
10                    End If
11                End If
12            End If
13        End If

14        AllFiles = sSubArray(DirListRet, 1, 1, , 1)

15        ValidLabels = ISDASIMMValidIDs(True)

16        Set STK = CreateStacker()
17        STK.Stack2D sArrayRange("Element", "Solum", "ISDA")

18        For i = 1 To sNRows(AllFiles)
19            ChunkGood = False
20            ThisChunk = sFileShow(CStr(AllFiles(i, 1)), vbTab, True)
21            If sNCols(ThisChunk) = 3 Then
22                If ThisChunk(1, 2) = "Solum" Then
23                    If ThisChunk(1, 3) = "ISDA" Then
24                        If sAll(sArrayIsNumber(sMatch(sSubArray(ThisChunk, 2, 1, , 1), ValidLabels, True))) Then
25                            ChunkGood = True
26                        End If
27                    End If
28                End If
29            End If
30            If ChunkGood Then
31                STK.Stack2D sSubArray(ThisChunk, 2)
32            Else
33                Throw "Bad file contents in " + CStr(AllFiles(i, 1)) 'Make error report more informative if necessary
34            End If
35        Next

36        LastReturn = STK.Report
37        LastDirListRet = DirListRet
38        LastTheYear = TheYear
39        LastResultsSubFolder = ResultsSubFolder

40        ISDASIMMGetResults = LastReturn

41        Exit Function
ErrHandler:
42        ISDASIMMGetResults = "#ISDASIMMGetResults (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMMatchStartDate
' Author     : Philip Swannell
' Date       : 08-Apr-2020
' Purpose    : ISDA is not consistent on how they adjust stress periods for weekends, these functions allow us to nudge
'       stress period dates to match theirs if the business days (i.e. weekdays) cogvered by the stress period are identical
' Parameters :
'  SolumStartDate:
'  ISDAStartDate :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMMatchStartDate(SolumStartDate As Long, ISDAStartDate As Long)

1         On Error GoTo ErrHandler
2         If FollWeekDay(SolumStartDate) = FollWeekDay(ISDAStartDate) Then
3             ISDASIMMMatchStartDate = ISDAStartDate
4         Else
5             ISDASIMMMatchStartDate = SolumStartDate
6         End If
7         Exit Function
ErrHandler:
8         ISDASIMMMatchStartDate = "#ISDASIMMMatchStartDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMMMatchEndDate(SolumEndDate As Long, ISDAEndDate As Long)
1         On Error GoTo ErrHandler
2         If PrevWeekDay(SolumEndDate) = PrevWeekDay(ISDAEndDate) Then
3             ISDASIMMMatchEndDate = ISDAEndDate
4         Else
5             ISDASIMMMatchEndDate = SolumEndDate
6         End If
7         Exit Function
ErrHandler:
8         ISDASIMMMatchEndDate = "#ISDASIMMMatchEndDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PrevWeekDay
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Returns InputDate if that's a weekday or preceding Friday otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Private Function PrevWeekDay(InputDate)
1         On Error GoTo ErrHandler
2         Select Case InputDate Mod 7
              Case 0 'Sat
3                 PrevWeekDay = InputDate - 1
4             Case 1 'Sun
5                 PrevWeekDay = InputDate - 2
6             Case Else
7                 PrevWeekDay = InputDate
8         End Select
9         Exit Function
ErrHandler:
10        Throw "#PrevWeekDay (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FollWeekDay
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Returns InputDate if thats a weekday or following Monday if weekend
' -----------------------------------------------------------------------------------------------------------------------
Private Function FollWeekDay(InputDate)
1         On Error GoTo ErrHandler
2         Select Case InputDate Mod 7
              Case 0 'Sat
3                 FollWeekDay = InputDate + 2
4             Case 1 'Sun
5                 FollWeekDay = InputDate + 1
6             Case Else
7                 FollWeekDay = InputDate
8         End Select
9         Exit Function
ErrHandler:
10        Throw "#FollWeekDay (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMMStandardiseAssetClass(AssetClass As String)
1         On Error GoTo ErrHandler
2         Select Case UCase(Replace(AssetClass, " ", ""))
              Case "CROSSCURRENCYBASIS", "CCB"
3                 ISDASIMMStandardiseAssetClass = "CCB"
4             Case "IR", "INTERESTRATE"
5                 ISDASIMMStandardiseAssetClass = "IR"
6             Case "FX", "FOREIGNEXCHANGE"
7                 ISDASIMMStandardiseAssetClass = "FX"
8             Case "EQ", "EQUITY"
9                 ISDASIMMStandardiseAssetClass = "EQ"
10            Case "CRQ", "CREDITQUALIFYING"
11                ISDASIMMStandardiseAssetClass = "CRQ"
12            Case "BC", "BASECORR", "BASECORRELATION"
13                ISDASIMMStandardiseAssetClass = "BC"
14            Case "CRNQ", "CREDITNON-QUALIFYING", "CREDITNONQUALIFYING"
15                ISDASIMMStandardiseAssetClass = "CRNQ"
16            Case "CM", "COMMODITY"
17                ISDASIMMStandardiseAssetClass = "CM"
18            Case "CROSSRISKCLASS", "CROSSASSET", "XA"
19                ISDASIMMStandardiseAssetClass = "XA"
20            Case Else
21                Throw "AssetClass not recognised. Recognised short-forms are: 'BC', 'CCB', 'CM', 'CRNQ', 'CRQ', 'EQ', 'FX', 'IR', 'XA'"
22        End Select

23        Exit Function
ErrHandler:
24        Throw "#ISDASIMMStandardiseAssetClass (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMCompareCorrelations
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Used for asset classes CM, EQ and CQ to compare Solums results (ISDA SIMM YYYY Correlation Generator.xlsm) with the contents of 1 or 2 files generated by ISDA
' Can optionaly make a call to ISDASIMMSaveResults
' Code attempts to deal with idiosyncrasies  of ISDA's data representations - mainly via function SortAndCompleteCorrelationMatrix
' Parameters :
'  TheYear            : The year of the run
'  AssetClass         : Allowed: CM, EQ CRQ
'  SolumCorrelations  : Array (or Range) with headers in order, none missing 1 to N
'  ISDAInterbucketFile: Contains the on-diagonal elements. Has top+left headers but may not be in order and some may be missing (e.g. 16 for commodity asset class)
'  ISDAIntraBucketFile: Contains on-diagonal elements. NB may be two cols index, corr or may be two-dimensional
'  DoSave             :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMCompareCorrelations(TheYear As Long, AssetClass As String, SolumCorrelations As Variant, ISDAInterbucketFile As String, ByVal ISDAIntraBucketFile As String, _
          DoSave As Boolean, StressPeriodCalcMethod As String, Optional ByRef SaveResult)
          Dim N As Long
          Dim i As Long, j As Long
          Const ErrString1 = "Top and left headers of SolumCorrelations should be in-order integers but they are not"
          Const ErrString2 = "Non-number found in SolumCorrelations"
          Const ErrString3 = "ISDAIntraBucketFile expected to have integers in the left column, but does not"
          ' Const ErrString4 = "ISDAIntraBucketFile expected to have numbers in the right column, but does not"
          Const ErrString5 = "On-diagonal elements of ISDAIntrabucketFile should be numbers but they are not"
          
          Dim IncDiagonal As Boolean
          Dim DataToSave
          Dim ISDACorrelations
          Dim ISDADiagonals
          Dim ISDAMelted
          Dim SolumMelted
          Dim Parameter As String

1         On Error GoTo ErrHandler
2         Force2DArrayR SolumCorrelations

3         Select Case ISDASIMMStandardiseAssetClass(AssetClass)
              Case "EQ"
4                 Parameter = "Correlations"
5                 IncDiagonal = True
6             Case "CRQ"
7                 Parameter = "Inter-bucket correlation"
8                 ISDAIntraBucketFile = "" 'Not needed in this case
9                 IncDiagonal = False
10            Case "CM"
11                Parameter = "Correlations"
12                IncDiagonal = True
13            Case Else
14                Throw "AssetClass not recognised"
15        End Select

16        N = sNRows(SolumCorrelations) - 1

17        If sNCols(SolumCorrelations) <> (N + 1) Then Throw "SolumCorrelations must be square"

18        For i = 1 To N
19            If SolumCorrelations(1, i + 1) <> i Then Throw ErrString1
20            If SolumCorrelations(i + 1, 1) <> i Then Throw ErrString1
21        Next

22        For i = 2 To N + 1
23            For j = 2 To N + 1
24                If Not IsNumber(SolumCorrelations(i, j)) Then Throw ErrString2
25            Next
26        Next
27        For i = 2 To N + 1
28            For j = 2 To i - 1
29                If SolumCorrelations(i, j) <> SolumCorrelations(j, i) Then
30                    Throw "SolumCorrelations must be symmetric, but elements " + CStr(i) + "," + CStr(j) + " <> " + CStr(j) + "," + CStr(i)
31                End If
32            Next
33        Next

34        SolumMelted = ThrowIfError(ISDASIMMMeltLabelledArray(SolumCorrelations, True, True, IncDiagonal))

          'ISDA's results not available. Set to zero!
35        If sFileExists(ISDAInterbucketFile) Then
36            ISDACorrelations = ThrowIfError(sFileShow(ISDAInterbucketFile, , True))
37        Else
38            ISDACorrelations = SolumCorrelations
39            For i = 2 To N + 1
40                For j = 2 To N + 1
41                    ISDACorrelations(i, j) = "NA"
42                Next
43            Next
44        End If

          'Arrgh rows and columns may be out of order and some may be missing - e.g. bucket 16 for commodity
45        ISDACorrelations = SortAndCompleteCorrelationMatrix(ISDACorrelations, N, "ISDAInterbucketFile")

          'ISDA's results not available. Set to zero!
46        If IncDiagonal Then
47            If sFileExists(ISDAIntraBucketFile) Then

48                ISDADiagonals = ThrowIfError(sFileShow(ISDAIntraBucketFile, , True))
                  'see comments in code as to why necessary...
49                ISDADiagonals = FixIntraBucketContents(ISDADiagonals, N)
50            Else
51                ISDADiagonals = SolumCorrelations
52                For i = 2 To N + 1
53                    For j = 2 To N + 1
54                        ISDADiagonals(i, j) = "NA"
55                    Next
56                Next
57            End If

58            If sNCols(ISDADiagonals) = 2 Then

59                For i = 1 To N
60                    ISDACorrelations(i + 1, i + 1) = 0
61                Next i

62                For i = 1 To N
                      Dim tmp
63                    tmp = sVLookup(i, ISDADiagonals)
64                    If IsNumber(tmp) Then
65                        ISDACorrelations(i + 1, i + 1) = tmp
66                    Else
67                        If i = 16 And ISDASIMMStandardiseAssetClass(AssetClass) = "CM" Then
68                            ISDACorrelations(i + 1, i + 1) = 0
69                        Else
70                            Throw "Cannot find intrabucket correlation for bucket " & CStr(i) & " in the ISDAIntraBucket file"
71                        End If
72                    End If
73                Next i
74            ElseIf sNRows(ISDADiagonals) = sNCols(ISDADiagonals) Then
75                ISDADiagonals = SortAndCompleteCorrelationMatrix(ISDADiagonals, N, "ISDAIntraBucketFile")
76                For i = 1 To N
77                    If Not IsNumberOrNA(ISDADiagonals(i + 1, i + 1)) Then Throw ErrString5
78                    ISDACorrelations(i + 1, i + 1) = ISDADiagonals(i + 1, i + 1)
79                Next i
80            Else
81                Throw "ISDAIntraBucketFile has an unexpected number of rows and columns"
82            End If
83        End If

84        ISDAMelted = ThrowIfError(ISDASIMMMeltLabelledArray(ISDACorrelations, True, True, IncDiagonal))

85        If Not sArraysIdentical(sSubArray(SolumMelted, 1, 1, , 1), sSubArray(ISDAMelted, 1, 1, , 1)) Then
86            Throw "Unexpected error - left columns of arrays SolumMelted and ISDAMelted do not match"
87        End If

88        DataToSave = sArrayRange(SolumMelted, sSubArray(ISDAMelted, 1, 2, , 1))
89        DataToSave = sArrayStack(sArrayRange("Element", "Solum", "ISDA"), DataToSave)

          Dim DataToReturn
90        If DoSave Then
91            SaveResult = ISDASIMMSaveResults(TheYear, AssetClass, Parameter, StressPeriodCalcMethod, DataToSave)
92            DataToReturn = sArrayStack("", "", "TheYear", TheYear, "", "", "AssetClass", AssetClass, "", "", "Parameter", Parameter, "", "", "SaveResults", SaveResult)
93            DataToReturn = sReshape(DataToReturn, 4, 4)
94            DataToSave = sArrayRange(DataToSave, sArrayStack("Difference", sArraySubtract(sSubArray(DataToSave, 2, 2, , 1), sSubArray(DataToSave, 2, 3, , 1))))
95            DataToReturn = sArrayStack(DataToReturn, DataToSave)
96            ISDASIMMCompareCorrelations = DataToReturn
97        Else
98            DataToSave = sArrayRange(DataToSave, sArrayStack("Difference", sArraySubtract(sSubArray(DataToSave, 2, 2, , 1), sSubArray(DataToSave, 2, 3, , 1))))
99            ISDASIMMCompareCorrelations = DataToSave
100       End If

101       Exit Function
ErrHandler:
102       ISDASIMMCompareCorrelations = "#ISDASIMMCompareCorrelations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function IsNumberOrNA(x As Variant) As Boolean
1         On Error GoTo ErrHandler
2         If IsNumber(x) Then
3             IsNumberOrNA = True
4         ElseIf VarType(x) = vbString Then
5             IsNumberOrNA = x = "NA"
6         Else
7             IsNumberOrNA = False
8         End If
9         Exit Function
ErrHandler:
10        Throw "#IsNumberOrNA (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SortAndCompleteCorrelationMatrix
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Takes a correlation matrix with headers, where the headers are a sub-set of 1...N and may not be in order
'              returns a correlation matrix with headers in order, and "missing" rows\cols filled with zeros.
'              New for 2022 - deletes rows\cols where headers exceed N
' Parameters :
'  M           : correlation matrix with headers as described above
'  N           :
'  NameOfMatrix: used for error message generation
' -----------------------------------------------------------------------------------------------------------------------
Function SortAndCompleteCorrelationMatrix(M As Variant, N As Long, NameOfMatrix As String)
          'M has headers
          Dim i As Long
          Dim TopHeaders, LeftHeaders
          Dim OldM11 As Variant
          Dim ErrString1 As String, ErrString2 As String, ErrString3 As String, ErrString4 As String
1         ErrString1 = "Invalid (non integer) headers found in " + NameOfMatrix
2         ErrString2 = "Invalid (out-of-range) headers found in " + NameOfMatrix
3         ErrString3 = "Top headers of " + NameOfMatrix + " do not match left headers"
4         ErrString4 = "Duplicated headers found in " + NameOfMatrix

5         On Error GoTo ErrHandler
          
6         Force2DArrayR M

          'Arrgh new for 2022 am receiving a matrix (in the case of Equity) that has too many rows/columns
          'Delete those with index bigger than N

          Dim ChooseVector
          Dim NeedToFilter As Boolean
          'Delete rows
7         ChooseVector = sReshape(True, sNRows(M), 1)
8         For i = 2 To sNRows(M)
9             If IsNumber(M(i, 1)) Then
10                If M(i, 1) > N Then
11                    ChooseVector(i, 1) = False
12                    NeedToFilter = True
13                End If
14            End If
15        Next i
16        If NeedToFilter Then
17            M = sMChoose(M, ChooseVector)
18        End If

          'Delete columns
19        NeedToFilter = False
20        ChooseVector = sReshape(True, 1, sNCols(M))
21        For i = 2 To sNCols(M)
22            If IsNumber(M(1, i)) Then
23                If M(1, i) > N Then
24                    ChooseVector(1, i) = False
25                    NeedToFilter = True
26                End If
27            End If
28        Next i
29        If NeedToFilter Then
30            M = sRowMChoose(M, ChooseVector)
31        End If
          
32        For i = 2 To sNRows(M)
33            If Not IsNumber(M(i, 1)) Then Throw ErrString1
34            If Not IsNumber(M(1, i)) Then Throw ErrString1
35            If M(i, 1) < 1 Then Throw ErrString2
36            If M(i, 1) > N Then Throw ErrString2
37            If M(i, 1) <> CInt(M(i, 1)) Then Throw ErrString1
38        Next i
39        OldM11 = M(1, 1)
40        M(1, 1) = 0

41        TopHeaders = sArrayTranspose(sSubArray(M, 1, 2, 1))
42        LeftHeaders = sSubArray(M, 2, 1, , 1)
43        If Not sArraysIdentical(sSortedArray(TopHeaders), sSortedArray(LeftHeaders)) Then Throw ErrString3
44        If sNRows(LeftHeaders) <> sNRows(sRemoveDuplicates(LeftHeaders)) Then Throw ErrString4

          Dim MissingIndices, NumMissing As Long

45        MissingIndices = sCompareTwoArrays(sIntegers(N), sSubArray(M, 2, 1, , 1), "In1AndNotIn2")
46        If sNRows(MissingIndices) > 1 Then
47            MissingIndices = sDrop(MissingIndices, 1)
48            NumMissing = sNRows(MissingIndices)

49            M = sArrayStack(M, sArrayRange(MissingIndices, sReshape(0, NumMissing, sNCols(M) - 1)))
50            M = sArrayRange(M, sArrayStack(sArrayTranspose(MissingIndices), sReshape(0, sNRows(M) - 1, NumMissing)))
51        End If

52        M = sSortedArray(M)
53        M = sArrayTranspose(sSortedArray(sArrayTranspose(M)))
54        M(1, 1) = OldM11

55        SortAndCompleteCorrelationMatrix = M

56        Exit Function
ErrHandler:
57        Throw "#SortAndCompleteCorrelationMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FixIntraBucketContents
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Files such as C:\ISDA SIMM\Solum Validation C-V 2020\EQ_delta\9_correlations\eq_delta-intra-bucket-10d.csv
'              are mal-formed in that an extraneous column appears as the second column with a header "RW". columns "should have" headers 1 to N (though don't assume in order or none missing...)
' -----------------------------------------------------------------------------------------------------------------------
Private Function FixIntraBucketContents(ByVal FileContents, N As Long)
          Dim HeadersGood As Variant
          Dim i As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR FileContents

3         If sNCols(FileContents) <= 2 Then
4             FixIntraBucketContents = FileContents
5             Exit Function
6         Else
7             HeadersGood = sReshape(False, 1, sNCols(FileContents))
8             HeadersGood(1, 1) = True
9             For i = 2 To sNCols(FileContents)
10                If IsNumber(FileContents(1, i)) Then
11                    If FileContents(1, i) >= 1 Then
12                        If FileContents(1, i) <= N Then
13                            If FileContents(1, i) = CInt(FileContents(1, i)) Then
14                                HeadersGood(1, i) = True
15                            End If
16                        End If
17                    End If
18                End If
19            Next i
20            FileContents = ThrowIfError(sRowMChoose(FileContents, HeadersGood))
21            HeadersGood = sReshape(False, sNRows(FileContents), 1)
22            HeadersGood(1, 1) = True
23            For i = 1 To sNRows(FileContents)
24                If IsNumber(FileContents(i, 1)) Then
25                    If FileContents(i, 1) >= 1 Then
26                        If FileContents(i, 1) <= N Then
27                            If FileContents(i, 1) = CInt(FileContents(i, 1)) Then
28                                HeadersGood(i, 1) = True
29                            End If
30                        End If
31                    End If
32                End If
33            Next i
34            FileContents = ThrowIfError(sMChoose(FileContents, HeadersGood))

35            FixIntraBucketContents = FileContents
36        End If
37        Exit Function
ErrHandler:
38        Throw "#FixIntraBucketContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMRiskWeightsFromIndividualRiskWeights
' Author     : Philip Swannell
' Date       : 05-May-2020
' Purpose    : Designed for use as part of the Data Preparation Validation" work for ISDA - workbook
'              "ISDA SIMM YYYY Analyse Data Series.xlsm" in that project we have the individual risk weights readily to hand,
'              and need to do the last stage of the calculation to risk weights for the SIMM, which is based around taking
'              medians but with complications that vary by asset class.
' Parameters :
'  AssetClass           : misnomer compared to most other functions in this module - really asset class + risk type
'  SeriesNames          : a colum array of the names of the series. In some cases assumes standard format of Underlying_Tenor or (for IR_vega) Currency_TimeToExercise_SwapTenor
'  IndividualRiskWeights: The risk weigths for each time series, calculated with appropriate 3Y and stress periods etc etc.
'  CountsIn4Y           : Same dimensions as SeriesName. For each series how many valid data points are there in the 4Y period - i.e. union of recent and stress
'  CountsInStress       : Same dimensions as SeriesName. For each series how many valid data points are the stress period, which may now be disjoint (StressBalance approach)
'  BucketingFile        : Simple csv file giving the allocation of series to "buckets"
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMRiskWeightsFromIndividualRiskWeights(AssetClass As String, ByVal SeriesNames As Variant, ByVal IndividualRiskWeights As Variant, _
          Optional CountsIn4Y As Variant, Optional CountsInStress As Variant, _
          Optional BucketingFile As String)

          Const Supported = "CM_vega,CRNQ_delta,CRQ_base,CRQ_delta,CRQ_vega,EQ_vega,FX_vega,IR_basis,IR_vega,IR_xccy"

1         On Error GoTo ErrHandler
2         If InStr("," + Supported + ",", "," + AssetClass + ",") = 0 Then
3             Throw "AssetClass not recognised, supported values - " + Supported
4         End If

          'Because in practice these are results from call to sColumnFromTable
5         ThrowIfError SeriesNames
6         ThrowIfError IndividualRiskWeights
7         ThrowIfError CountsIn4Y
8         ThrowIfError CountsInStress

          Dim NR_SN As Long, NC_SN As Long
          Dim NR_IRW As Long, NC_IRW As Long

9         Force2DArrayR SeriesNames, NR_SN, NC_SN
10        Force2DArrayR IndividualRiskWeights, NR_IRW, NC_IRW

11        If NR_SN * NC_SN <> NR_IRW * NC_IRW Then
12            Throw "SeriesNames and IndividualRiskWeights must have the same number of elements"
13        End If

14        If NC_SN <> 1 Then
15            SeriesNames = sReshape(SeriesNames, , 1)
16            NC_SN = 1
17            NR_SN = sNRows(SeriesNames)
18        End If

19        If NC_IRW <> 1 Then
20            IndividualRiskWeights = sReshape(IndividualRiskWeights, , 1)
21            NC_IRW = 1
22            NR_IRW = sNRows(IndividualRiskWeights)
23        End If

24        If Not IsMissing(CountsIn4Y) Then
              Dim NR_C4Y As Long, NC_C4Y As Long
25            Force2DArrayR CountsIn4Y, NR_C4Y, NC_C4Y
26            If NC_C4Y <> 1 Then
27                CountsIn4Y = sReshape(CountsIn4Y, , 1)
28                NC_C4Y = 1
29                NR_C4Y = sNRows(CountsIn4Y)
30            End If
31            If NR_C4Y <> NR_SN Then Throw ("CountsIn4Y (if provided) must have the same number of elements as IndividualRiskWeights")
32        End If

33        If Not IsMissing(CountsInStress) Then
              Dim NR_CS As Long, NC_CS As Long
34            Force2DArrayR CountsInStress, NR_CS, NC_CS
35            If NC_CS <> 1 Then
36                CountsInStress = sReshape(CountsInStress, , 1)
37                NC_CS = 1
38                NR_CS = sNRows(CountsInStress)
39            End If
40            If NR_CS <> NR_SN Then Throw ("CountsInStress (if provided) must have the same number of elements as IndividualRiskWeights")
41        End If

          Dim ChooseVector

          'Update 3 May 2021
          'All risk classes for which we do data preparation checks now use 500,125 rule - see email from Xiaowei Yan 13 April 2021:

          'I apologize that we didn't let you know earlier - we applied the 500d/125d rule to all the risk classes in C-VI, except for CM Delta correlation and XA Delta.

42        Select Case AssetClass
              Case "CM_vega"
43                If IsMissing(CountsIn4Y) Then Throw "Argument 'CountsIn4Y' must be provided for AssetClass = " & AssetClass
44                If IsMissing(CountsInStress) Then Throw "Argument 'CountsInStress' must be provided for AssetClass = " & AssetClass
45                ChooseVector = sArrayAnd(sArrayGreaterThanOrEqual(CountsIn4Y, 500), _
                      sArrayGreaterThanOrEqual(CountsInStress, 125))
46                IndividualRiskWeights = sMChoose(IndividualRiskWeights, ChooseVector)
47                ISDASIMMRiskWeightsFromIndividualRiskWeights = ThrowIfError(sColumnMedian(IndividualRiskWeights))
48            Case "CRQ_base"
49                If IsMissing(CountsIn4Y) Then Throw "Argument 'CountsIn4Y' must be provided for AssetClass = " & AssetClass
50                If IsMissing(CountsInStress) Then Throw "Argument 'CountsInStress' must be provided for AssetClass = " & AssetClass
51                ChooseVector = sArrayAnd(sArrayGreaterThanOrEqual(CountsIn4Y, 500), _
                      sArrayGreaterThanOrEqual(CountsInStress, 125))
52                IndividualRiskWeights = sMChoose(IndividualRiskWeights, ChooseVector)
53                ISDASIMMRiskWeightsFromIndividualRiskWeights = sArrayMultiply(100, sColumnMedian(IndividualRiskWeights))
54            Case "CRQ_vega", "EQ_vega", "FX_vega"
55                If IsMissing(CountsIn4Y) Then Throw "Argument 'CountsIn4Y' must be provided for AssetClass = " & AssetClass
56                If IsMissing(CountsInStress) Then Throw "Argument 'CountsInStress' must be provided for AssetClass = " & AssetClass
57                ChooseVector = sArrayAnd(sArrayGreaterThanOrEqual(CountsIn4Y, 500), _
                      sArrayGreaterThanOrEqual(CountsInStress, 125))
58                SeriesNames = sMChoose(SeriesNames, ChooseVector)
59                IndividualRiskWeights = sMChoose(IndividualRiskWeights, ChooseVector)
                  Dim Tenors, Assets, LookupTable, LookupValues, ReshapedIRW, MedianOverTenors
60                LookupTable = sArrayRange(SeriesNames, IndividualRiskWeights)
61                Tenors = sRemoveDuplicates(sStringBetweenStrings(SeriesNames, "_"))
62                Assets = sRemoveDuplicates(sStringBetweenStrings(SeriesNames, , "_"))
63                LookupValues = sArrayConcatenate(Assets, "_", sArrayTranspose(Tenors))
64                ReshapedIRW = sVLookup(LookupValues, LookupTable)
65                MedianOverTenors = sRowMedian(ReshapedIRW, True)
66                ISDASIMMRiskWeightsFromIndividualRiskWeights = ThrowIfError(sColumnMedian(MedianOverTenors, True))
67            Case "CRNQ_delta", "CRQ_delta" 'THESE RISK CLASSES NOT RELEVANT FOR 2021 DATA PREPARATION
                  'Need MinimumDataCount:500,125
                  'Changing the logic here? Then change method IncludeIRWInRWCalc in workbook "ISDA SIMM YYYY Analyse Data Series.xlsm"
68                If Not sFileExists(BucketingFile) Then Throw "File '" + BucketingFile + "' not found"
69                If IsMissing(CountsIn4Y) Then Throw "Argument 'CountsIn4Y' must be provided for AssetClass = " & AssetClass
70                If IsMissing(CountsInStress) Then Throw "Argument 'CountsInStress' must be provided for AssetClass = " & AssetClass

71                ChooseVector = sArrayAnd(sArrayGreaterThanOrEqual(CountsIn4Y, 500), _
                      sArrayGreaterThanOrEqual(CountsInStress, 125))
72                SeriesNames = sMChoose(SeriesNames, ChooseVector)
73                IndividualRiskWeights = sMChoose(IndividualRiskWeights, ChooseVector)

                  Dim BucketFileContents, IndividualBuckets, BucketList, RiskWeights, i As Long
74                BucketFileContents = ThrowIfError(sFileShow(BucketingFile, , True))

75                IndividualBuckets = sVLookup(SeriesNames, BucketFileContents)
76                BucketList = sRemoveDuplicates(IndividualBuckets, True)
77                RiskWeights = sReshape(0, sNRows(BucketList), 1)

78                For i = 1 To sNRows(BucketList)
79                    ChooseVector = sArrayEquals(IndividualBuckets, BucketList(i, 1))
80                    RiskWeights(i, 1) = sColumnMedian(sMChoose(IndividualRiskWeights, ChooseVector), True)(1, 1) * 10000
81                Next i
82                ISDASIMMRiskWeightsFromIndividualRiskWeights = sArrayRange(BucketList, RiskWeights)
83            Case "IR_vega"
                  'Logic follows of workbook "ISDA SIMM 2020 IR Vega Risk Weight.xlsm"

                  Dim Expiries, Currencies, NumCurrencies As Long, NumExpiries As Long, NumTenors As Long, Headers
                  Dim OneMatrix, Medians1, Medians2

84                Expiries = sTokeniseString("2W,1M,3M,6M,1Y,2Y,3Y,5Y,10Y,15Y,20Y,30Y")
85                Currencies = sTokeniseString("USD,EUR,JPY,GBP,CAD,AUD,SEK")
86                Tenors = sArrayTranspose(sTokeniseString("6M,1Y,2Y,3Y,4Y,5Y,6Y,7Y,8Y,9Y,10Y,15Y,20Y,30Y"))
87                NumCurrencies = sNRows(Currencies)
88                NumExpiries = sNRows(Expiries)
89                NumTenors = sNCols(Tenors)
                  
90                OneMatrix = sArrayConcatenate(Expiries, "_", Tenors)
91                Headers = sArrayConcatenate(sGroupReshape(Currencies, NumExpiries), "_", sReshape(OneMatrix, NumExpiries * NumTenors, NumTenors))

92                LookupTable = sArrayRange(SeriesNames, IndividualRiskWeights)
93                ReshapedIRW = sVLookup(Headers, LookupTable)
94                Medians1 = sRowMedian(ReshapedIRW, True)
95                Medians2 = sColumnMedianByChunks(Medians1, NumExpiries, True)

96                ISDASIMMRiskWeightsFromIndividualRiskWeights = sColumnMedian(Medians2, True)
97            Case "IR_xccy" 'THIS RISK CLASS NOT RELEVANT FOR 2021 DATA PREPARATION
                  'Logic follows of workbook "ISDA SIMM 2020 Cross-Currency Basis delta risk weight.xlsm"
                  Dim CurrencyPairs

98                CurrencyPairs = sRemoveDuplicates(sStringBetweenStrings(SeriesNames, "", "_"))
99                Tenors = sArrayTranspose(SortTenures(sRemoveDuplicates(sStringBetweenStrings(SeriesNames, "_"))))
100               LookupTable = sArrayRange(SeriesNames, IndividualRiskWeights)
101               Headers = sArrayConcatenate(CurrencyPairs, "_", Tenors)
102               ReshapedIRW = sVLookup(Headers, LookupTable)
103               ISDASIMMRiskWeightsFromIndividualRiskWeights = SafeMedian(sColumnMedian(ReshapedIRW, True)) * 10000

104           Case Else
105               Throw "Case AssetClass = " & AssetClass + " not coded yet"
106       End Select

107       Exit Function
ErrHandler:
108       ISDASIMMRiskWeightsFromIndividualRiskWeights = "#ISDASIMMRiskWeightsFromIndividualRiskWeights (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMWeightedQuantile
' Author     : Philip Swannell
' Date       : 02-Jul-2020
' Purpose    : Generalisation of quantiles to the weighted case.
' Parameters :
'  DataValues     :Column array, non-numbers ignored
'  Weights        : Column array, same size as DataValues
'  PercentileLevel:
'  CalcStyle      : 'CENTRAL', 'INC' or 'EXC'
' Follows algorithm documented by Martin Baxter, and emailed to PGS 2 July 2020
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMWeightedQuantile(ByVal DataValues, ByVal Weights, ByVal PercentileLevel, Optional CalcStyle = "CENTRAL")
          Dim ChooseVector As Variant
          Dim NRV As Long, NCV As Long
          Dim NRW As Long, NCW As Long
          Dim i As Long
          Dim AnyExcluded As Boolean
          Dim AnyIncluded As Boolean

1         On Error GoTo ErrHandler
2         Force2DArrayR DataValues, NRV, NCV
3         If NCV <> 1 Then Throw "DataValues must have one column"

4         Force2DArrayR Weights, NRW, NCW
5         If NCW <> 1 Then Throw "Weights must have one column"

6         If NRV <> NRW Then Throw "DataValues and Weights must have the same number of rows"

7         ChooseVector = sReshape(False, NRV, 1)

8         For i = 1 To NRV
9             If IsNumber(DataValues(i, 1)) Then
10                If IsNumber(Weights(i, 1)) Then
11                    If Weights(i, 1) > 0 Then
12                        ChooseVector(i, 1) = True
13                        AnyIncluded = True
14                    End If
15                End If
16            End If
17            If Not ChooseVector(i, 1) Then AnyExcluded = True
18        Next i

19        If Not AnyIncluded Then Throw "At least one element of DataValues must be a number with a positive weight"

20        If AnyExcluded Then
21            Weights = sMChoose(Weights, ChooseVector)
22            DataValues = sMChoose(DataValues, ChooseVector)
23            NRV = sNRows(Weights)
24            NRW = NRV
25        End If

          Dim S As Double
          Dim c As Double
          Dim SortedVW
          Dim Ps
          Dim PartialSumW
          
26        Select Case UCase(CalcStyle)
              Case "CENTRAL"
27                c = 0.5
28            Case "INC"
29                c = 1
30            Case "EXC"
31                c = 0
32            Case Else
33                Throw "CalcStyle must be either 'CENTRAL', 'INC' or 'EXC'"
34        End Select

35        S = sSumOfNums(Weights)
36        SortedVW = Application.WorksheetFunction.Sort(sArrayRange(DataValues, Weights))

37        DataValues = sSubArray(SortedVW, 1, 1, , 1)
38        Weights = sSubArray(SortedVW, 1, 2, , 1)

39        PartialSumW = sPartialSum(Weights)
          
40        Ps = sReshape(0, NRW, 1)
          
41        For i = 1 To NRW
42            Ps(i, 1) = (PartialSumW(i, 1) - c * Weights(i, 1)) / (S + (1 - 2 * c) * Weights(i, 1))
43        Next

44        ISDASIMMWeightedQuantile = sInterp(Ps, DataValues, PercentileLevel, "Linear", "NN")

45        Exit Function
ErrHandler:
46        ISDASIMMWeightedQuantile = "#ISDASIMMWeightedQuantile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
