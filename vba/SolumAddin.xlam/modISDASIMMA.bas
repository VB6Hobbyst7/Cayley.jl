Attribute VB_Name = "modISDASIMMA"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatCorrelations
' Author    : Philip Swannell
' Date      : 09-Jul-2017
' Purpose   : Common code for formatting a correlation range
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatCorrelations(RangeWithHeaders As Range, Optional Name As String)

1         On Error GoTo ErrHandler
2         With RangeWithHeaders
3             If Len(Name) > 0 Then
4                 .Parent.Names.Add Name & "WithHeaders", .Offset(0)
5             End If
6             AddGreyBorders .Offset(0), True
7             AddGreyBorders .Cells(1, 1), True
8             .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1).NumberFormat = "0.0%"
9             .Columns(1).AutoFit
10            .HorizontalAlignment = xlHAlignCenter

11            With .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1)
12                AutoFitColumns .Offset(0), 1, , 8
13                If Len(Name) > 0 Then
14                    .Parent.Names.Add Name, .Offset(0)
15                End If
16                AddGreyBorders .Offset(0), True
17                .FormatConditions.AddColorScale ColorScaleType:=3
18                .FormatConditions(.FormatConditions.Count).SetFirstPriority
19                .FormatConditions(1).ColorScaleCriteria(1).Type = _
                      xlConditionValueNumber
20                .FormatConditions(1).ColorScaleCriteria(1).Value = -1
21                With .FormatConditions(1).ColorScaleCriteria(1).FormatColor
22                    .Color = 255
23                    .TintAndShade = 0
24                End With
25                .FormatConditions(1).ColorScaleCriteria(2).Type = _
                      xlConditionValueNumber
26                .FormatConditions(1).ColorScaleCriteria(2).Value = 0
27                With .FormatConditions(1).ColorScaleCriteria(2).FormatColor
28                    .ThemeColor = xlThemeColorDark1
29                    .TintAndShade = 0
30                End With
31                .FormatConditions(1).ColorScaleCriteria(3).Type = _
                      xlConditionValueNumber
32                .FormatConditions(1).ColorScaleCriteria(3).Value = 1
33                With .FormatConditions(1).ColorScaleCriteria(3).FormatColor
34                    .Color = 16711680
35                    .TintAndShade = 0
36                End With
37            End With
38        End With
39        Exit Sub
ErrHandler:
40        Throw "#FormatCorrelations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMAverageTopHalf
' Author     : Philip Swannell
' Date       : 23-Apr-2019
' Purpose    : Implements ISDA's method of taking the average of the biggest half of a set of data, see email from XYAn@isda.org
'              to PGS, dated 23 April 2019. If there are an even number of inputs take the average of the top half, if there are
'              an odd number, the median is included in the average with a weight of .5
' Parameters :
'  InputData:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMAverageTopHalf(InputData As Variant, Optional ByRef retSimpleAverage As Double)
          Dim ArrayToOrder As Variant
          Dim ChooseVector As Variant
          Dim N As Long
          Dim SortedArray As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR InputData

3         If sNCols(InputData) > 1 Then
4             ArrayToOrder = sReshape(InputData, sNRows(InputData) * sNCols(InputData), 1)
5         Else
6             ArrayToOrder = InputData
7         End If
          
8         SortedArray = sSortedArray(ArrayToOrder, , , , False)
9         ChooseVector = sArrayIsNumber(SortedArray)
10        Select Case sArrayCount(ChooseVector)
              Case 0
11                Throw "No numeric data found in InputData"
12            Case sNRows(ArrayToOrder)
                  'Nothing to do
13            Case Else
14                SortedArray = sMChoose(SortedArray, ChooseVector)
15        End Select

16        retSimpleAverage = Application.WorksheetFunction.Average(SortedArray)

17        N = sNRows(SortedArray)
18        If N Mod 2 = 0 Then
19            ISDASIMMAverageTopHalf = Application.WorksheetFunction.Average(sSubArray(SortedArray, 1, 1, N / 2))
20        Else
21            SortedArray((N + 1) / 2, 1) = SortedArray((N + 1) / 2, 1) / 2
22            ISDASIMMAverageTopHalf = Application.WorksheetFunction.sum(sSubArray(SortedArray, 1, 1, (N + 1) / 2)) / (N / 2)
23        End If

24        Exit Function
ErrHandler:
25        ISDASIMMAverageTopHalf = "#ISDASIMMAverageTopHalf (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ISDASIMMBucketData(BucketNumbers, FileNames, HeaderRowNumber As Long)
          Dim i As Long
          Dim STK As clsStacker
          Dim TheseHeaders As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayRMulti BucketNumbers, FileNames

3         Set STK = CreateStacker()
4         For i = 1 To sNRows(FileNames)
5             TheseHeaders = ThrowIfError(sFileHeaders(CStr(FileNames(i, 1)), ",", HeaderRowNumber))
              '6             If TheseHeaders(1, 1) <> "Date" Then Throw "Assetion failed - first element of headers in file '" + CStr(FileNames(i, 1)) + "' should be 'Date' but it isn't so"
6             TheseHeaders = sSubArray(TheseHeaders, 1, 2)
7             TheseHeaders = sArrayStack(TheseHeaders, sReshape(BucketNumbers(i, 1), 1, sNCols(TheseHeaders)))
8             STK.Stack2D sArrayTranspose(TheseHeaders)
9         Next

10        ISDASIMMBucketData = STK.Report

11        Exit Function
ErrHandler:
12        ISDASIMMBucketData = "#ISDASIMMBucketData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMCorrelations
' Author    : Philip Swannell
' Date      : 05-Jul-2017
' Purpose   : Returns correlations with files as inputs
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMCorrelations(FileName As String, ByVal Headers As Variant, IsAbsolute As Boolean, DateFormat As String, ThreeYStart As Long, ThreeYEnd As Long, _
        StressStarts As Variant, StressEnds As Variant, AllowBadHeaders As Boolean, Optional FileIsReturns As Boolean, Optional WithHeaders As Boolean, _
        Optional ReturnLag As String, Optional ExcludeZeroReturns As Boolean, Optional PostProcessing As String, Optional ReturnRounding As Variant = False, _
        Optional DataCleaningRules As String, Optional CheckDates As Boolean = True)

          Dim HeadersFound
          Dim Result
          Dim Returns As Variant

1         On Error GoTo ErrHandler
2         Returns = ThrowIfError(ISDASIMMReturnsFromFile(FileName, Headers, IsAbsolute, ThreeYStart, ThreeYEnd, StressStarts, StressEnds, DateFormat, AllowBadHeaders, FileIsReturns, HeadersFound, , , ReturnLag, ExcludeZeroReturns, , , ReturnRounding, DataCleaningRules, CheckDates))
3         Result = ThrowIfError(sKendallTau(Returns, , True))
4         If WithHeaders Then
5             Result = sArraySquare(vbNullString, sArrayTranspose(HeadersFound), HeadersFound, Result)
6         End If

7         Select Case LCase$(PostProcessing)
              Case vbNullString
                  'Nothing to do
8             Case LCase$("sMedianOffDiagonal"), LCase$("MedianOffDiagonal")
9                 If WithHeaders Then
10                    Result = sMedianOffDiagonal(sSubArray(Result, 2, 2))
11                Else
12                    Result = sMedianOffDiagonal(Result)
13                End If
14            Case Else
15                Throw "PostProcessing not recognised. Allowed values: omitted or 'MedianOffDiagonal'"
16        End Select

17        ISDASIMMCorrelations = Result

18        Exit Function
ErrHandler:
19        ISDASIMMCorrelations = "#ISDASIMMCorrelations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMCountReturnsInFile
' Author     : Philip Swannell
' Date       : 08-Apr-2019
' Purpose    : For a given file of returns (or asset prices) produce a report of the headers in the file together with the
'              number of returns in the stress period and in the three year period. If bucketing file is provided then an
'              extra column is provided giving the bucket of each asset, with the report sorted on the "bucket" column.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMCountReturnsInFile(FileName As String, IsAbsolute As Boolean, ThreeYStart As Long, ThreeYEnd As Long, StressStarts As Variant, _
        StressEnds As Variant, DateFormat As String, AllowBadHeaders As Boolean, FileIsReturns As Boolean, PercentileMethod As String, _
        Optional ReturnLag As String, Optional ExcludeZeroReturns As Boolean, Optional BucketingFile As String, Optional CheckDates As Boolean = True)
        
          Dim Returns As Variant
          Const Headers = "ISDASIMMCountReturnsInFile"
          Dim HeadersFound
          Dim i As Long
          Dim NumHeaders As Long
          Dim NumStressRows As Long
          Dim Result
          Dim ResultHeaders
          Dim DataCountInStressPeriod

1         On Error GoTo ErrHandler

2         Returns = ThrowIfError(ISDASIMMReturnsFromFile(FileName, Headers, IsAbsolute, ThreeYStart, ThreeYEnd, StressStarts, StressEnds, DateFormat, AllowBadHeaders, _
              FileIsReturns, HeadersFound, False, False, ReturnLag, ExcludeZeroReturns, NumStressRows, , , , CheckDates, DataCountInStressPeriod))
3         NumHeaders = sNRows(HeadersFound)
4         Result = sReshape(0, NumHeaders, 3)

5         For i = 1 To NumHeaders
6             Result(i, 1) = HeadersFound(i, 1)
7             Result(i, 2) = DataCountInStressPeriod(1, i)
8             Result(i, 3) = sArrayCount(sArrayIsNumber(sSubArray(Returns, 1, i, , 1)))
9         Next i

10        ResultHeaders = sArrayRange("Asset", "CountReturnsStress", "CountReturnsFOURyears")
11        If BucketingFile <> vbNullString Then
              Dim BucketingFileContents
              Dim Buckets
12            BucketingFileContents = sFileShow(BucketingFile, ",", True)
13            Buckets = sVLookup(HeadersFound, BucketingFileContents)
14            Result = sArrayRange(Buckets, Result)
15            Result = sSortedArray(Result) 'sort on bucket
16            ResultHeaders = sArrayRange("Bucket", ResultHeaders)
17        End If
18        Result = sArrayStack(ResultHeaders, Result)

19        ISDASIMMCountReturnsInFile = Result

20        Exit Function
ErrHandler:
21        ISDASIMMCountReturnsInFile = "#ISDASIMMCountReturnsInFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMCrossAssetCorrelations
' Author    : Philip Swannell
' Date      : 05-Jul-2017
' Purpose   : Returns cross-correlations, with files as inputs
'             Note abbreviated some argument names so that Ctrl+Shift+A still works in Excel
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMCrossAssetCorrelations(FileName1 As String, ByVal Headers1 As Variant, IsAbsolute1 As Boolean, File1DateFormat As String, File1IsReturns As Boolean, ReturnLag1 As String, ExclZR1 As Boolean, _
        FileName2 As String, ByVal Headers2 As Variant, IsAbsolute2 As Boolean, File2DateFormat As String, File2IsReturns As Boolean, ReturnLag2 As String, ExclZR2 As Boolean, _
        ThreeYStart As Long, ThreeYEnd As Long, StressStart As Long, StressEnd As Long, AllowBadHeaders As Boolean, WithHeaders As Boolean)
          Dim HeadersFound1 As Variant
          Dim HeadersFound2 As Variant
          Dim Result As Variant
          Dim Returns1 As Variant
          Dim Returns2 As Variant

1         On Error GoTo ErrHandler
2         Returns1 = ThrowIfError(ISDASIMMReturnsFromFile(FileName1, Headers1, IsAbsolute1, ThreeYStart, ThreeYEnd, StressStart, StressEnd, File1DateFormat, AllowBadHeaders, File1IsReturns, HeadersFound1, False, False, ReturnLag1, ExclZR1))
3         Returns2 = ThrowIfError(ISDASIMMReturnsFromFile(FileName2, Headers2, IsAbsolute2, ThreeYStart, ThreeYEnd, StressStart, StressEnd, File2DateFormat, AllowBadHeaders, File2IsReturns, HeadersFound2, False, False, ReturnLag2, ExclZR2))
4         Result = ThrowIfError(sKendallTau(Returns1, Returns2, True))

5         If WithHeaders Then
6             ISDASIMMCrossAssetCorrelations = sArraySquare(vbNullString, sArrayTranspose(HeadersFound2), HeadersFound1, Result)
7         Else
8             ISDASIMMCrossAssetCorrelations = Result
9         End If

10        Exit Function
ErrHandler:
11        ISDASIMMCrossAssetCorrelations = "#ISDASIMMCrossAssetCorrelations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMCrossAssetCorrelations2
' Author     : Philip Swannell
' Date       : 24-Mar-2020
' Purpose    : Version with additional parameters for rounding and data cleaning
' Parameters :
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMCrossAssetCorrelations2(File1 As String, ByVal Headers1 As Variant, IsAbs1 As Boolean, DateFormat1 As String, IsReturns1 As Boolean, ReturnLag1 As String, ExclZR1 As Boolean, Rounding1 As Variant, DCRule1 As String, _
        File2 As String, ByVal Headers2 As Variant, IsAbs2 As Boolean, DateFormat2 As String, IsReturns2 As Boolean, ReturnLag2 As String, ExclZR2 As Boolean, Rounding2 As Variant, DCRule2 As String, _
        ThreeYStart As Long, ThreeYEnd As Long, StressStart As Long, StressEnd As Long, AllowBadHeaders As Boolean, WithHeaders As Boolean)
          Dim HeadersFound1 As Variant
          Dim HeadersFound2 As Variant
          Dim Result As Variant
          Dim Returns1 As Variant
          Dim Returns2 As Variant

1         On Error GoTo ErrHandler
2         Returns1 = ThrowIfError(ISDASIMMReturnsFromFile(File1, Headers1, IsAbs1, ThreeYStart, ThreeYEnd, StressStart, StressEnd, DateFormat1, AllowBadHeaders, IsReturns1, HeadersFound1, False, False, ReturnLag1, ExclZR1, , , Rounding1, DCRule1))
3         Returns2 = ThrowIfError(ISDASIMMReturnsFromFile(File2, Headers2, IsAbs2, ThreeYStart, ThreeYEnd, StressStart, StressEnd, DateFormat2, AllowBadHeaders, IsReturns2, HeadersFound2, False, False, ReturnLag2, ExclZR2, , , Rounding2, DCRule2))
4         Result = ThrowIfError(sKendallTau(Returns1, Returns2, True))

5         If WithHeaders Then
6             ISDASIMMCrossAssetCorrelations2 = sArraySquare(vbNullString, sArrayTranspose(HeadersFound2), HeadersFound1, Result)
7         Else
8             ISDASIMMCrossAssetCorrelations2 = Result
9         End If

10        Exit Function
ErrHandler:
11        ISDASIMMCrossAssetCorrelations2 = "#ISDASIMMCrossAssetCorrelations2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMDeltaInterBucketCorrelations
' Author    : Philip
' Date      : 04-Jul-2017
' Purpose   : Implements our calculation of the interest rate correlations, one bucket to another.
'             See fourth paragraph of section 3.5 of ISDA SIMM Calibration Methodology, 24 May 2017
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMDeltaInterBucketCorrelations(DataFile As String, DateFormat As String, ThreeYStart As Long, ThreeYEnd As Long, _
        StressStart As Long, StressEnd As Long, ReturnLag As String, ExcludeZeroReturns As Boolean, HeaderRowNumber As Long, Tenors As String, _
        ReturnRounding As Variant)

          Dim AllCcys As Variant
          Dim AllTenors As Variant
          Dim Ccy As String
          Dim Ccy1 As String
          Dim Ccy2 As String
          Dim Dates As Variant
          Dim EndRow As Long
          Dim Expression As String
          Dim HeadersForCcy As Variant
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim MatchIDs As Variant
          Dim NCcys As Long
          Dim Result
          Dim Returns As Variant
          Dim ReturnsForCcy
          Dim StartRow As Long
          Dim StressEndRow As Long
          Dim StressStartRow As Long
          Dim TheData
          Dim TimeSeries As Variant
          Dim TopRow As Variant

1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             ISDASIMMDeltaInterBucketCorrelations = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         TheData = ThrowIfError(sFileShow(DataFile, ",", True, True, False, False, DateFormat))
7         Dates = sSubArray(TheData, HeaderRowNumber + 1, 1, , 1)
          Dim NumDates As Long
          'Arrgh. Sometimes working with files that contain small number of weekends _
           This code filters them out and also checks that after filtering the dates that remain of consecutive weekdays.
8         ISDASIMMFilterOutWeekends TheData, Dates, HeaderRowNumber - 1, NumDates, DataFile

9         TopRow = sSubArray(TheData, HeaderRowNumber, 2, 1)
10        AllCcys = sRemoveDuplicates(sArrayLeft(sArrayTranspose(TopRow), 3), True)
11        AllTenors = sTokeniseString(Tenors)

12        StartRow = ThrowIfError(sMatch(ThreeYStart, Dates, True))
13        EndRow = ThrowIfError(sMatch(ThreeYEnd, Dates, True))
14        StressStartRow = ThrowIfError(sMatch(StressStart, Dates, True))
15        StressEndRow = ThrowIfError(sMatch(StressEnd, Dates, True))
16        TimeSeries = sSubArray(TheData, HeaderRowNumber + 1, 2)

17        Returns = ThrowIfError(ISDASIMMReturnsFromTimeSeries(TimeSeries, True, True, StartRow, EndRow, StressStartRow, StressEndRow, ReturnLag, ExcludeZeroReturns, , ReturnRounding))

18        NCcys = sNRows(AllCcys)

          'For each currency, save the returns data to R. 1 column for each tenor
19        For k = 1 To NCcys
20            Ccy = AllCcys(k, 1)

21            HeadersForCcy = sArrayConcatenate(Ccy, "_", AllTenors)
22            MatchIDs = sMatch(HeadersForCcy, sArrayTranspose(TopRow))
23            ReturnsForCcy = sReshape(0, sNRows(Returns), sNRows(MatchIDs))

24            For j = 1 To sNRows(MatchIDs)
25                If IsNumber(MatchIDs(j, 1)) Then
26                    For i = 1 To sNRows(Returns)
27                        ReturnsForCcy(i, j) = Returns(i, MatchIDs(j, 1))
28                    Next i
29                Else
30                    For i = 1 To sNRows(Returns)
31                        ReturnsForCcy(i, j) = "NA"
32                    Next i
33                End If
34            Next j
35            ThrowIfError SaveArrayToR(ReturnsForCcy, "ReturnsFor" & Ccy, 2, False, False, True)
36        Next k

37        Result = sReshape(0, NCcys, NCcys)
38        For i = 2 To NCcys
39            Ccy1 = AllCcys(i, 1)
40            For j = 1 To i - 1
41                Ccy2 = AllCcys(j, 1)
                  ' Expression = "median(sin(cor(ReturnsFor" & Ccy1 & ",ReturnsFor" & Ccy2 & ",""pairwise.complete.obs"",""kendall"")*pi/2),na.rm=TRUE)"
42                Expression = "median(sin(KendallTau(ReturnsFor" & Ccy1 & ",ReturnsFor" & Ccy2 & ")*pi/2),na.rm=TRUE)"
43                Result(i, j) = ThrowIfError(sExecuteRCode(Expression))
44                Result(j, i) = Result(i, j)
45            Next j
46        Next i

47        For i = 1 To NCcys
48            Result(i, i) = "NA"
49        Next

50        Result = sArraySquare(vbNullString, sArrayTranspose(AllCcys), AllCcys, Result)

51        ISDASIMMDeltaInterBucketCorrelations = Result

52        Exit Function
ErrHandler:
53        ISDASIMMDeltaInterBucketCorrelations = "#ISDASIMMDeltaInterBucketCorrelations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMDeltaIntraBucketCorrelations
' Author    : Philip
' Date      : 04-Jul-2017
' Purpose   : Implements our calculation of the interest rate correlations. See section 3.5
'             of ISDA SIMM Calibration Methodology, 24 May 2017
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMDeltaIntraBucketCorrelations(DataFile As String, DateFormat As String, ThreeYStart As Long, ThreeYEnd As Long, StressStart As Long, StressEnd As Long, _
        ReturnLag As String, ExcludeZeroReturns As Boolean, HeaderRowNumber As Long, Tenors As String, Optional ReturnFormat As String, Optional ReturnRounding As Variant = False)

          Dim AllCcys As Variant
          Dim AllTenors As Variant
          Dim Ccy As String
          Dim Dates As Variant
          Dim EndRow As Long
          Dim HeadersForCcy As Variant
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim MatchIDs As Variant
          Dim Melt As Boolean
          Dim NT As Long
          Dim Returns As Variant
          Dim StartRow As Long
          Dim STK As clsStacker
          Dim StressEndRow As Long
          Dim StressStartRow As Long
          Dim TheData
          Dim TimeSeries As Variant
          Dim TopRow As Variant

1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             ISDASIMMDeltaIntraBucketCorrelations = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         Select Case LCase$(ReturnFormat)
              Case "stackedmatrices"
7                 Melt = False
8             Case "melted"
9                 Melt = True
10            Case Else
11                Throw "ReturnFormat must be 'StackedMatrices' or 'Melted'"
12        End Select

13        TheData = ThrowIfError(sFileShow(DataFile, ",", True, True, False, False, DateFormat))
14        Dates = sSubArray(TheData, HeaderRowNumber + 1, 1, , 1)

          Dim NumDates As Long
          'Arrgh. Sometimes working with files that contain small number of weekends _
           This code filters them out and also checks that after filtering the dates that remain are consecutive weekdays.
15        ISDASIMMFilterOutWeekends TheData, Dates, HeaderRowNumber - 1, NumDates, DataFile

16        TopRow = sSubArray(TheData, HeaderRowNumber, 2, 1)
17        AllCcys = sRemoveDuplicates(sArrayLeft(sArrayTranspose(TopRow), 3), True)
18        AllTenors = sTokeniseString(Tenors)
19        NT = sNRows(AllTenors)
20        Dates = sSubArray(TheData, HeaderRowNumber + 1, 1, , 1)
21        StartRow = ThrowIfError(sMatch(ThreeYStart, Dates, True))
22        EndRow = ThrowIfError(sMatch(ThreeYEnd, Dates, True))
23        StressStartRow = ThrowIfError(sMatch(StressStart, Dates, True))
24        StressEndRow = ThrowIfError(sMatch(StressEnd, Dates, True))
25        TimeSeries = sSubArray(TheData, HeaderRowNumber + 1, 2)

26        Returns = ThrowIfError(ISDASIMMReturnsFromTimeSeries(TimeSeries, True, True, StartRow, EndRow, StressStartRow, StressEndRow, ReturnLag, ExcludeZeroReturns, , ReturnRounding))

27        Set STK = CreateStacker()
28        For k = 1 To sNRows(AllCcys)
29            Ccy = AllCcys(k, 1)
30            HeadersForCcy = sArrayConcatenate(Ccy, "_", AllTenors)
31            MatchIDs = sMatch(HeadersForCcy, sArrayTranspose(TopRow))
32            ISDASIMMDeltaIntraBucketCorrelations = MatchIDs

              Dim CcyCorrMatrix
              Dim ReturnsForCcy

33            ReturnsForCcy = sReshape(0, sNRows(Returns), sNRows(MatchIDs))

34            For j = 1 To sNRows(MatchIDs)
35                If IsNumber(MatchIDs(j, 1)) Then
36                    For i = 1 To sNRows(Returns)
37                        ReturnsForCcy(i, j) = Returns(i, MatchIDs(j, 1))
38                    Next i
39                Else
40                    For i = 1 To sNRows(Returns)
41                        ReturnsForCcy(i, j) = "NA"
42                    Next i
43                End If
44            Next j

45            CcyCorrMatrix = ThrowIfError(sKendallTau(ReturnsForCcy, , True))

              Dim ThisChunk As Variant
46            ThisChunk = sArraySquare(vbNullString, sArrayTranspose(HeadersForCcy), HeadersForCcy, CcyCorrMatrix)

47            If Melt Then
48                ThisChunk = ISDASIMMMeltLabelledArray(ThisChunk, True, False, True)
                  Dim ChooseVector
49                ChooseVector = sArrayIsNumber(sSubArray(ThisChunk, 1, 2, , 1))
50                If sArrayCount(ChooseVector) > 0 Then
51                    If sArrayCount(ChooseVector) < sNRows(ChooseVector) Then
52                        ThisChunk = sMChoose(ThisChunk, ChooseVector)
53                    End If
54                    STK.Stack2D ThisChunk
55                End If
56            Else
57                STK.Stack2D ThisChunk
58            End If

59        Next k

60        ISDASIMMDeltaIntraBucketCorrelations = STK.Report

61        Exit Function
ErrHandler:
62        ISDASIMMDeltaIntraBucketCorrelations = "#ISDASIMMDeltaIntraBucketCorrelations (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMFxTradingWeights2020
' Author     : Philip Swannell
' Date       : 12-Feb-2020
' Purpose    : Calculate weights used for the calculation of the fx "pseudo index" from the 2020 calibration onwards. See email from Xiaowei Yan to PGS 11 Feb 2020
' Parameters :
'  Pairs                :
'  CSAWeightsFile       :
'  IndividualWeightsFile:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMFxTradingWeights2020(ByVal Pairs, Optional CSAWeightsFile As String, Optional IndividualWeightsFile As String)
          Dim LookupRes1
          Dim LookupRes2
          Dim IndividualWeights As Variant
          Dim Result
          Dim CSAWeights As Variant

1         On Error GoTo ErrHandler

2         CSAWeights = ThrowIfError(sFileShow(CSAWeightsFile, ",", True))
3         IndividualWeights = ThrowIfError(sFileShow(IndividualWeightsFile, ",", True))

4         LookupRes1 = ThrowIfError(sVLookup(sArrayLeft(Pairs, 3), CSAWeights, "Weight", "Instrument"))
5         LookupRes2 = ThrowIfError(sVLookup(sArrayRight(Pairs, 3), IndividualWeights, "Weight", "Instrument"))
6         Result = sArrayMultiply(LookupRes1, LookupRes2)

7         Result = sArrayIf(sArrayIsNumber(Result), Result, 0)

8         ISDASIMMFxTradingWeights2020 = Result

9         Exit Function
ErrHandler:
10        ISDASIMMFxTradingWeights2020 = "#ISDASIMMFxTradingWeights2020 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ISDASIMMFxTradingWeights
' Author    : Philip
' Date      : 07-Jul-2017
' Purpose   : Encapsulate (our understanding of) calculation of trading weights for fx time series
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMFxTradingWeights(Optional ByVal Pairs, Optional WeightsFile As String, Optional PairWeightsFile As String)
          Dim LookupRes1
          Dim LookupRes2
          Dim PairWeightsData As Variant
          Dim Result
          Dim TwoColReturn As Boolean
          Dim WeightsData As Variant

1         On Error GoTo ErrHandler

2         WeightsData = ThrowIfError(sFileShow(WeightsFile, ",", True))
3         PairWeightsData = ThrowIfError(sFileShow(PairWeightsFile, ",", True))

4         If IsMissing(Pairs) Then
5             Pairs = sSubArray(PairWeightsData, 2, 1, , 1)
6             TwoColReturn = True
7         End If

8         LookupRes1 = ThrowIfError(sVLookup(sArrayLeft(Pairs, 3), WeightsData, "Weight", "Instrument"))
9         LookupRes2 = ThrowIfError(sVLookup(Pairs, PairWeightsData, "Weight", "Instrument"))
10        Result = sArrayMultiply(LookupRes1, LookupRes2)

11        Result = sArrayIf(sArrayIsNumber(Result), Result, 0)
12        If TwoColReturn Then
13            Result = sArrayRange(Pairs, Result)
14        End If

15        ISDASIMMFxTradingWeights = Result

16        Exit Function
ErrHandler:
17        ISDASIMMFxTradingWeights = "#ISDASIMMFxTradingWeights (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMMeltCorrelationMatrix
' Author     : Philip Swannell
' Date       : 23-Mar-2018
' Purpose    : Turn a symetric matrix into a two-column array listing its elements with each symmetric pair of
'              off-diagonals only listed once. Use to include matrices along with singletons in a single column comparison of all ISDA
'              calibration data versus Solum calibration data.
' Parameters :
'  CorrMatrixNoHeaders:  A symmetric matrix
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMMeltCorrelationMatrix(CorrMatrixNoHeaders As Variant, WithLabels As Boolean)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR CorrMatrixNoHeaders, NR, NC

3         If NR <> NC Then Throw "CorrMatrixNoHeaders must be a square matrix"

4         Result = sReshape(vbNullString, NR + (NR * (NR - 1) / 2), 2)

5         For i = 1 To NR
6             For j = 1 To i
7                 k = k + 1
8                 Result(k, 1) = "(" + CStr(i) + "," + CStr(j) + ")"
9                 Result(k, 2) = CorrMatrixNoHeaders(i, j)
10            Next
11        Next

12        If Not WithLabels Then
13            Result = sSubArray(Result, 1, 2)
14        End If

15        ISDASIMMMeltCorrelationMatrix = Result

16        Exit Function
ErrHandler:
17        ISDASIMMMeltCorrelationMatrix = "#ISDASIMMMeltCorrelationMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMMeltLabelledArray
' Author     : Philip Swannell
' Date       : 10-Apr-2018
' Purpose    : Turn an array with labels in the left column and top row a two-column array listing its elements.
'              Use to include such data along with singletons in a single column comparison of all ISDA
'              calibration data versus Solum calibration data.
' Parameters :
'  M:  A symmetric matrix
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMMeltLabelledArray(ByVal M As Variant, Optional RowByRow As Boolean = True, Optional IsSymmetric As Boolean, Optional IncDiagonal As Boolean)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long
          Dim NumOut As Long
          Dim Result As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR M

3         NR = sNRows(M) - 1
4         NC = sNCols(M) - 1
5         If IsSymmetric Then If NC <> NR Then Throw "Symmetric matrix must be square"

6         If IsSymmetric Then If Not sArraysNearlyIdentical(M, sArrayTranspose(M)) Then Throw "Matrix is not symmetric"

7         If IsSymmetric Then
8             If IncDiagonal Then
9                 NumOut = (NR + 1) * NR / 2
10            Else
11                NumOut = (NR) * (NR - 1) / 2
12            End If
13        Else
14            NumOut = NR * NC
15        End If

16        Result = sReshape(vbNullString, NumOut, 2)
17        k = 0

18        If RowByRow Then
19            For i = 1 To NR
20                For j = IIf(IsSymmetric, IIf(IncDiagonal, i, i + 1), 1) To NC
21                    k = k + 1
22                    Result(k, 1) = CStr(M(i + 1, 1)) + "," + CStr(M(1, j + 1))
23                    Result(k, 2) = M(i + 1, j + 1)
24                Next
25            Next
26        Else
27            For j = 1 To NC
28                For i = IIf(IsSymmetric, IIf(IncDiagonal, j, j + 1), 1) To NR
29                    k = k + 1
30                    Result(k, 1) = CStr(M(i + 1, 1)) + "," + CStr(M(1, j + 1))
31                    Result(k, 2) = M(i + 1, j + 1)
32                Next
33            Next
34        End If

35        ISDASIMMMeltLabelledArray = Result

36        Exit Function
ErrHandler:
37        ISDASIMMMeltLabelledArray = "#ISDASIMMMeltLabelledArray (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMNumFilesMatchingFilter
' Author     : Philip Swannell
' Date       : 21-Mar-2018
' Purpose    : The number of files in a folder (no recursion) that match a file filter
' Parameters :
'  Folder    : Folder - with or without training backslash
'  FileFilter: See sDirList, supports both Regular Expressions and wild cards
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMNumFilesMatchingFilter(Folder As String, FileFilter As String)
          Dim Res
1         On Error GoTo ErrHandler
2         Res = sDirList(Folder, False, False, "F", "F", FileFilter)
3         If sIsErrorString(Res) Then
4             ISDASIMMNumFilesMatchingFilter = 0
5         Else
6             ISDASIMMNumFilesMatchingFilter = sNRows(Res)
7         End If

8         Exit Function
ErrHandler:
9         ISDASIMMNumFilesMatchingFilter = "#ISDASIMMNumFilesMatchingFilter (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMNumSeriesInFiles
' Author     : Philip Swannell
' Date       : 21-Mar-2018
' Purpose    : Returns the total number of time series in all files in a folder that match the FileFilter
'              the number of time series in a file is taken to be one minus the number of columns in the file
'              since by convention the first column is a "Date" column. The number of columns is read by
'              reading the header row as HeaderRowNumber - asumes fixed size of file preamble
' Parameters :
'  Folder         :
'  FileFilter     :
'  HeaderRowNumber:
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMNumSeriesInFiles(Folder As String, FileFilter As String, HeaderRowNumber As Long)
          Dim Files As Variant
          Dim i As Long
          Dim TheseHeaders As Variant
          Dim Total As Long
1         On Error GoTo ErrHandler
2         Files = sDirList(Folder, False, False, "F", "F", FileFilter)
3         If sIsErrorString(Files) Then
4             ISDASIMMNumSeriesInFiles = 0
5             Exit Function
6         Else
7             For i = 1 To sNRows(Files)
8                 TheseHeaders = sFileHeaders(CStr(Files(i, 1)), ",", HeaderRowNumber)
9                 If Not sIsErrorString(TheseHeaders) Then
10                    Total = Total + sNCols(TheseHeaders) - 1
11                End If
12            Next
13        End If
14        ISDASIMMNumSeriesInFiles = Total

15        Exit Function
ErrHandler:
16        ISDASIMMNumSeriesInFiles = "#ISDASIMMNumSeriesInFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
