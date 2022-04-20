Attribute VB_Name = "modTable"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modTable
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : Methods relating to the sheet called Table.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private Const chNHR = "Neither Trade nor Fx headroom"
Private Const chTHR = "Trade headroom (for BarChart)"
Private Const chFx = "Fx headroom (for BubbleChart)"
Private Const chBoth = "Both Trade and Fx headroom"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuTableSheet
' Author    : Philip Swannell
' Date      : 29-Sep-2015
' Purpose   : Attached to the Menu... button on the sheet Tabel
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuTableSheet()
          Const chRun = "&Run Table..."
          Const chExport = "E&xport Table..."
          '     Const chAddOrRemove = "&Add or Remove Banks..."
          Const chShow = "&Unhide Columns"
          Const chHide = "&Hide Columns"
          Const chBubbleCharts = "Nine Bubble charts..."
          Const FidRun = 156
          Const FidExport = 3
          Const FidShow = 137
          Const FidHide = 138
          Const FidBubbleCharts = 13230
          Dim chOpenOtherBooks As Variant
          Dim EnableFlags As Variant
          Dim enbOpenOtherBooks As Variant
          Dim FaceIDs As Variant
          Dim FidOpenOtherBooks As Variant
          Dim LinesBookIsOpen As Boolean
          Dim MarketBookIsOpen As Boolean
          Dim OBAO As Boolean
          Dim Res
          Dim TheChoices As Variant
          Dim TradesBookIsOpen As Boolean

1         On Error GoTo ErrHandler
2         RunThisAtTopOfCallStack

3         OBAO = OtherBooksAreOpen(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen)
4         If OBAO Then
5             chOpenOtherBooks = CreateMissing()
6             FidOpenOtherBooks = CreateMissing()
7             enbOpenOtherBooks = CreateMissing()
8         Else
9             chOpenOtherBooks = NameForOpenOthers(MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen, False)
10            FidOpenOtherBooks = 23
11            enbOpenOtherBooks = True
12        End If

13        TheChoices = sArrayStack(chOpenOtherBooks, chRun, chExport, "--" & chShow, chHide, "--" & chBubbleCharts)
14        EnableFlags = sArrayStack(enbOpenOtherBooks, sReshape(OBAO, 5, 1))
15        FaceIDs = sArrayStack(FidOpenOtherBooks, FidRun, FidExport, FidShow, FidHide, FidBubbleCharts)

16        Res = ShowCommandBarPopup(TheChoices, FaceIDs, EnableFlags, , ChooseAnchorObject())

17        Select Case Res
              Case Unembellish(CStr(chOpenOtherBooks))
18                OpenOtherBooks
19            Case Unembellish(chRun)
20                RunTable
21            Case Unembellish(chShow)
                  Dim SPH As clsSheetProtectionHandler
                  Dim SUH As clsScreenUpdateHandler
22                Set SUH = CreateScreenUpdateHandler()
23                Set SPH = CreateSheetProtectionHandler(shTable)
24                shTable.UsedRange.Columns.Hidden = False
25                GroupingButtonDoAllOnSheet shTable, True
26            Case Unembellish(chHide)
27                GroupingButtonDoAllOnSheet shTable, False
28            Case Unembellish(chExport)
29                ExportTable
30            Case "#Cancel!"
31            Case Unembellish(chBubbleCharts)
32                NineBubbleCharts
33            Case Else
34                Throw "Unrecognised choice in menu: " & CStr(Res)
35        End Select
36        Exit Sub
ErrHandler:
37        SomethingWentWrong "#MenuTableSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ChooseBanksToRun
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : Allow the user to select which of the banks listed on the Table sheet should be run.
' -----------------------------------------------------------------------------------------------------------------------
Function ChooseBanksToRun()

          Static PreviousBanks As Variant

          Dim AllBanks
          Dim BanksToRun
          Dim LinesBook As Workbook
          Dim TheTable As Range
          Dim TheTableNoHeaders As Range
          Dim Title As String
          Dim TopText As String

1         On Error GoTo ErrHandler

2         Set LinesBook = OpenLinesWorkbook(True, False)

3         Title = "Run Table"
4         TopText = "Select banks to run"

5         Set TheTable = RangeFromSheet(shTable, "TheTable")
6         With TheTable
7             Set TheTableNoHeaders = .offset(1).Resize(.Rows.Count - 1)
8         End With

9         AllBanks = GetColumnFromLinesBook("CPTY_PARENT", LinesBook)
10        AllBanks = AnnotateBankNames(AllBanks, True, LinesBook)

11        TopText = "Select banks to run"

12        BanksToRun = ShowMultipleChoiceDialog(AllBanks, PreviousBanks, Title, TopText, , , "Next >", , False)
13        If sArraysIdentical(BanksToRun, "#User Cancel!") Then GoTo EarlyExit

14        PreviousBanks = BanksToRun

15        ChooseBanksToRun = AnnotateBankNames(BanksToRun, False, LinesBook)
EarlyExit:
16        Set LinesBook = Nothing

17        Exit Function
ErrHandler:
18        Throw "#ChooseBanksToRun (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExportTable
' Author    : Philip Swannell
' Date      : 25-Jul-2016
' Purpose   : Export the contents of the sheet Table to a tab-delimited text file. Idea
'             is that Airbus will pick up such files.
' -----------------------------------------------------------------------------------------------------------------------
Sub ExportTable(Optional FileName As String)
          Dim Data
          Dim i As Long
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         If FileName = "" Then
3             FileName = GetSaveAsFilenameWrap("CayleyTableFiles", _
                  "ResultsByCounterpartyParent_" & Format(Date, "yyyy-mm-dd") & ".csv", _
                  "csv Files (*.csv),*.cvs", , "Save Results by Parent Counterparty", "Save")
4             If FileName = "False" Then Exit Sub
5         End If

6         Data = RangeFromSheet(shTable, "TheTable").Value 'Use Value not Value2 for file to have Date and DateTime formatting
7         NR = sNRows(Data): NC = sNCols(Data)

          'Headers on the sheet have line breaks to display prettily, but we don't want those in the file
8         For i = 1 To NC
9             Data(1, i) = Replace(Replace(Data(1, i), vbLf, ""), " ", "")
10        Next i

11        ThrowIfError sFileSaveCSV(FileName, Data)

12        Exit Sub
ErrHandler:
13        Throw "#ExportTable (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SoakTest
' Author    : Philip Swannell
' Date      : 11-Jan-2017
' Purpose   : Try to replicate out-of-resources-of-some-kind problems that Guy reports.
' -----------------------------------------------------------------------------------------------------------------------
Sub SoakTest()
          Const NumRuns = 20
          Const DoTradeHeadroom = True
          Const DoFxHeadroom = True
          Dim BanksToRun As Variant
          Dim et As Double
          Dim i As Long
          Dim st As Double
          Dim TheTable As Range
          Dim TheTableNoHeaders As Range

1         On Error GoTo ErrHandler
2         Set TheTable = RangeFromSheet(shTable, "TheTable")
3         With TheTable
4             Set TheTableNoHeaders = .offset(1).Resize(.Rows.Count - 1)
5         End With
6         BanksToRun = TheTableNoHeaders.Columns(1).Value

7         For i = 1 To NumRuns
8             st = sElapsedTime()
9             RunTable DoTradeHeadroom, DoFxHeadroom, True, BanksToRun
10            et = sElapsedTime()
11            MessageLogWrite "Soak Test Run " & CStr(i) & " took " & Format(et - st, "0.0") & " seconds"
12        Next i
13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#SoakTest (line " & CStr(Erl) & "): " & Err.Description & "!"
15    End Sub

Sub TestRunTable()
1         On Error GoTo ErrHandler
2         RunTable False, False, False, sArrayStack("Foo", "Bar")

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestRunTable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunTable
' Author    : Philip Swannell
' Date      : 26-May-2015
' Purpose   : Wrapped by both buttons on the sheet "Table"
' -----------------------------------------------------------------------------------------------------------------------
Sub RunTable(Optional DoTradeHeadroom, Optional DoFxHeadroom, Optional SilentMode As Boolean, _
          Optional BanksToRun As Variant)
          Dim Choice As String
          Dim D1 As Dictionary
          Dim D2 As Dictionary
          Dim D3 As Dictionary
          Dim D4 As Dictionary
          Dim i As Long
          Dim LinesScaleFactor As Double
          Dim Message As String
          Dim Prompt
          Dim PromptArray
          Dim PromptArrayForDialog
          Dim SUH As Object
          Dim TheTable As Range
          Dim TheTableNoHeaders As Range
          Dim TradesScaleFactor As Double
          
          Dim Options As Variant
          Static CurrentChoice As String
          Static EraseHeadroom As Boolean
          Static HaveRunBefore As Boolean
          Dim ShowBackButton As Boolean
          Dim UnrecognisedBanks As Variant

1         On Error GoTo ErrHandler

2         SyncBanksInCayleyWithBanksInLinesBook

          'By default EraseHeadroom should be TRUE
3         If Not HaveRunBefore Then EraseHeadroom = True
4         HaveRunBefore = True

5         If IsEmpty(BanksToRun) Or IsMissing(BanksToRun) Then
6             ShowBackButton = True
GoBack:
7             BanksToRun = ChooseBanksToRun()
8         End If
9         If IsEmpty(BanksToRun) Then Exit Sub

10        Options = sArrayStack(chNHR, chTHR, chFx, chBoth)
11        If VarType(DoTradeHeadroom) <> vbBoolean Or VarType(DoFxHeadroom) <> vbBoolean Then
GoBack2:
              Dim ButtonClicked As String
12            Choice = ShowOptionButtonDialog(Options, "Run Table", _
                  "What headroom calculations do you want to do?", _
                  CurrentChoice, , , "Erase currently displayed data", EraseHeadroom, , _
                  IIf(ShowBackButton, "< Back", "Next >"), IIf(ShowBackButton, "Next >", ""), _
                  "Cancel", ButtonClicked)
13            If ButtonClicked = "< Back" Then
14                DoTradeHeadroom = Empty: DoFxHeadroom = Empty
15                GoTo GoBack
16            End If
17            Select Case Choice
                  Case chNHR
18                    CurrentChoice = chNHR
19                    DoTradeHeadroom = False
20                    DoFxHeadroom = False
21                Case chTHR
22                    CurrentChoice = chTHR
23                    DoTradeHeadroom = True
24                    DoFxHeadroom = False
25                Case chFx
26                    CurrentChoice = chFx
27                    DoTradeHeadroom = False
28                    DoFxHeadroom = True
29                Case chBoth
30                    CurrentChoice = chBoth
31                    DoTradeHeadroom = True
32                    DoFxHeadroom = True
33                Case Else
34                    Exit Sub
35            End Select
36        End If

          Dim SPH1 As Object
          Dim SPH2 As Object
37        Set SPH1 = CreateSheetProtectionHandler(shCreditUsage)
38        Set SPH2 = CreateSheetProtectionHandler(shTable)
39        TradesScaleFactor = RangeFromSheet(shCreditUsage, "TradesScaleFactor", True, False, False, False, False).Value
40        LinesScaleFactor = RangeFromSheet(shCreditUsage, "LinesScaleFactor", True, False, False, False, False).Value
41        Set TheTable = RangeFromSheet(shTable, "TheTable")
42        With TheTable
43            Set TheTableNoHeaders = .offset(1).Resize(.Rows.Count - 1)
44        End With

45        UnrecognisedBanks = sCompareTwoArrays(BanksToRun, TheTableNoHeaders.Columns(1).Value, "In1AndNotIn2")
46        If sNRows(UnrecognisedBanks) >= 2 Then
47            Throw "The following banks are not recognised, since they do not appear in the Lines workbook: " + sConcatenateStrings(sSubArray(UnrecognisedBanks, 2), ", ")
48        End If

49        If Not SilentMode Then

50            PromptArray = sArrayStack("Number of banks", sNRows(BanksToRun), _
                  "Headroom calcs", Trim(sStringBetweenStrings(Choice, , "(")), _
                  "FilterBy2", RangeFromSheet(shCreditUsage, "FilterBy2").Value, _
                  "Filter2Value", RangeFromSheet(shCreditUsage, "Filter2Value").Value, _
                  "IncludeAssetClasses", RangeFromSheet(shCreditUsage, "IncludeAssetClasses").Value, _
                  "PortfolioAgeing", RangeFromSheet(shCreditUsage, "PortfolioAgeing").Value, _
                  "FxShock", RangeFromSheet(shCreditUsage, "FxShock").Value, _
                  "FxVolShock", RangeFromSheet(shCreditUsage, "FxVolShock").Value, _
                  "NumMCPaths", Format(RangeFromSheet(shCreditUsage, "NumMCPaths"), "###,###"), _
                  "NumObservations", Format(RangeFromSheet(shCreditUsage, "NumObservations"), "###,###"), _
                  "CurrenciesToInclude", RangeFromSheet(shConfig, "CurrenciesToInclude"))

51            If TradesScaleFactor <> 1 Or LinesScaleFactor <> 1 Then
52                PromptArray = sArrayStack(PromptArray, _
                      "", "", _
                      "Morphing:", "", _
                      "TradesScaleFactor", TradesScaleFactor, _
                      "LinesScaleFactor", LinesScaleFactor)
53            End If

54            PromptArray = sArrayStack(PromptArray, "Erase currently displayed data", EraseHeadroom)
55            PromptArray = sReshape(PromptArray, sNRows(PromptArray) / 2, 2)
56            PromptArrayForDialog = CleanUpPromptArray(PromptArray)

57            Prompt = "Update Results by Bank with the following inputs:" & vbLf & _
                  sConcatenateStrings(sJustifyArrayOfStrings(PromptArrayForDialog, "Calibri", 11, " " & vbTab), vbLf)

58            Select Case MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion + vbDefaultButton2, _
                  "Run Table", "< &Back", "Yes, Update", "Cancel", , 800)
                  Case vbYes
59                    GoTo GoBack2
60                Case vbNo
                      'bash on
61                Case Else
62                    Exit Sub
63            End Select
64        End If

          Dim TimeStart
65        TimeStart = Now()
66        MessageLogWrite "Run table core starting at " & _
              Format(TimeStart, "dd-mmm-yyyy hh:mm:ss") & vbLf & "Prompt was:" & _
              vbLf & sConcatenateStrings(sJustifyArrayOfStrings(PromptArray, "Courier New", 11, " "), vbLf)

67        Set SUH = CreateScreenUpdateHandler()

68        JuliaLaunchForCayley
69        OpenOtherBooks
70        BuildModelsInJulia False, RangeFromSheet(shCreditUsage, "FxShock"), RangeFromSheet(shCreditUsage, "FxVolShock")

          ' When using the HW model, there is a significant speed up in headroom solving (cacheing of
          ' the "cube" of trade values) that works best when successive banks being solved for
          ' have the same BaseCurrency. So we re-order the banks by BaseCurrency.
          Dim ArrayToSort
          Dim colBCs
          Dim colIndices
          Dim N As Long
          Dim RowToProcess As Long
          Dim ThisBank
71        N = sNRows(BanksToRun)
72        colBCs = sReshape(0, N, 1)
73        For i = 1 To N
74            colBCs(i, 1) = ThrowIfError(LookupCounterpartyInfo(BanksToRun(i, 1), "Base Currency"))(1, 1)
75        Next i

76        colIndices = sMatch(BanksToRun, TheTable.Columns(1).Value)
77        ArrayToSort = sArrayRange(colBCs, BanksToRun, colIndices)
78        ArrayToSort = sSortedArray(ArrayToSort)

          'Pain that we have to store a separate list of potentially valid keys into the collection - See suggestions at http://stackoverflow.com/questions/5702362/vba-collection-list-of-keys
          Dim BankLimitsHeaders
          Dim CopyOffset As Long
          Dim LimitsHeaders
          Dim MethodologyHeaders
79        LimitsHeaders = sTokeniseString("1Y Limit,2Y Limit,3Y Limit,4Y Limit,5Y Limit,7Y Limit,10Y Limit")
80        MethodologyHeaders = sTokeniseString("Methodology,Confidence %,Volatility Input,Base Currency,Product Credit Limits,Notional Cap")
81        BankLimitsHeaders = sTokeniseString("THR Bank 1Y,THR Bank 3Y")

          Dim m0 As Double
          Dim m1 As Double
          Dim t0 As Double

82        With RangeFromSheet(shTable, "TheTable")
83            For i = 1 To sNRows(ArrayToSort)
84                t0 = sElapsedTime()
85                If gDebugMode Then m0 = sExcelWorkingSetSize()
86                RowToProcess = ArrayToSort(i, 3)
87                ThisBank = .Cells(RowToProcess, 1).Value
88                If ThisBank <> ArrayToSort(i, 2) Then Throw "Assertion failed - mismatch in bank names"
89                PrepareForCalculation ThisBank, False, False, True
90                StatusBarWrap CStr(i) & "/" & CStr(N) & "    " & ThisBank
91                Set D1 = New Dictionary
92                Set D2 = New Dictionary
93                Set D3 = New Dictionary
                  
                  'Can use .Add method rather than DictAdd since we know that we are adding rather than overwriting the contents of the dictionary
94                D1.Add "CounterpartyLongName", LookupCounterpartyInfo(ThisBank, "CPTY LONG NAME")
95                D1.Add "CounterpartyVeryShortName", LookupCounterpartyInfo(ThisBank, "Very short name")
96                D1.Add "Limits", LookupCounterpartyInfo(ThisBank, LimitsHeaders, Empty)
                  Dim MethodologyArray
97                MethodologyArray = sArrayTranspose(LookupCounterpartyInfo(ThisBank, MethodologyHeaders, Empty))
98                MethodologyArray(1, 5) = Replace(MethodologyArray(1, 5), "Calculation", "Calc")
99                D1.Add "Methodology", MethodologyArray
100               D1.Add "BankLimits", sArrayTranspose(LookupCounterpartyInfo(ThisBank, BankLimitsHeaders))
101               D1.Add "AirbusTHR3Y", LookupCounterpartyInfo(ThisBank, "Airbus THR 3Y")

102               D1.Add "TradeSolveResult", Empty
103               D1.Add "FxSolveResult", Empty
104               D1.Add "ProfileResult", Empty

105               If DoTradeHeadroom Then
106                   RunCreditUsageSheet "Solve1to5", False, True, False, D2
107                   D2.Add "THR3YMinBkAirbus", sArrayMin(LookupCounterpartyInfo(ThisBank, "Airbus THR 3Y"), SafeIndex(D2.Item("TradeHeadroom"), 3, 1))
108               End If
109               If DoFxHeadroom Then
110                   RunCreditUsageSheet "SolveFx", False, True, False, D3
111               End If
112               If Not (DoTradeHeadroom Or DoFxHeadroom) Then
113                   RunCreditUsageSheet "Standard", False, True, False, D4
114               End If

115               D1.Add "TimeStamp", Now()
116               D1.Add "ProcessTime", sElapsedTime() - t0

117               If EraseHeadroom Then
118                   .Rows(RowToProcess).offset(, 1).Resize(, .Columns.Count - 1).ClearContents
119               End If

120               CopyOffset = .Cells(RowToProcess, 1).Row - shTable.Range("TradeSolveResult").Row
121               DictionaryToSheet D1, shTable, CopyOffset
122               If DoTradeHeadroom Then
123                   If Not DictGet(D2, "Success") Then
124                       TrimDictionary D2, "TradeSolveResult"
125                   End If
126                   DictionaryToSheet D2, shTable, CopyOffset
127               End If
128               If DoFxHeadroom Then
129                   If Not DictGet(D3, "Success") Then
130                       TrimDictionary D3, "FxSolveResult"
131                   End If
132                   DictionaryToSheet D3, shTable, CopyOffset
133               End If
134               If Not (DoTradeHeadroom Or DoFxHeadroom) Then
135                   If DictGet(D4, "Success") Then
136                       TrimDictionary D3, "ProfileResult"
137                   End If
138                   DictionaryToSheet D4, shTable, CopyOffset
139               End If

140               shTable.Calculate
141               RefreshScreen
142               If gDebugMode Then m1 = sExcelWorkingSetSize()
143               Message = "i: " & CStr(i) & " Bank: " & ThisBank
144               If gDebugMode Then Message = Message & " WorkingSetSize: " & Format(m1, "###,##0") + _
                      " DeltaWorkingSetSize: " & Format(m1 - m0, "###,##0")
145               MessageLogWrite Message
146           Next i
147       End With

148       FormatTable False
149       shBubbleChart.Calculate
150       shBarChart.Calculate

          Dim TimeEnd
151       TimeEnd = Now()

152       MessageLogWrite "Run table core finishing at " & Format(TimeEnd, "dd-mmm-yyyy hh:mm:ss") & _
              " TimeElapsed = " & CStr((TimeEnd - TimeStart) * 24 * 60 * 60) & " seconds"
            
153       StatusBarWrap False
154       Exit Sub
ErrHandler:
155       Throw "#RunTable (line " & CStr(Erl) & "): " & Err.Description & "!"
156       StatusBarWrap False
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MenuBarChart
' Author    : Philip Swannell
' Date      : 04-Dec-2016
' Purpose   : Attached to Menu button on sheet BarChart
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuBarChart()
          Dim chChooseSeries As String
          Dim chSort0 As String
          Dim chSort1 As String
          Dim chSort2 As String
          Dim chSort3 As String
          Dim chSort4 As String
          Dim Fid0 As Long
          Dim Fid1 As Long
          Dim Fid2 As Long
          Dim Fid3 As Long
          Dim Fid4 As Long
          Dim Fid5 As Long
          Dim Res

          Dim Choices
          Dim FIDs
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         chChooseSeries = "&Choose series to show..."
4         With RangeFromSheet(shBarChart, "DataForChart")
5             chSort0 = "&" & .Cells(1, 1).Value
6             chSort1 = "&" & .Cells(1, 2).Value
7             chSort2 = "&" & .Cells(1, 3).Value
8             chSort3 = "&" & .Cells(1, 4).Value
9             chSort4 = "&" & .Cells(1, 5).Value
10        End With

11        With RangeFromSheet(shBarChart, "SortBy")
12            If .Value = chSort0 Then
13                Fid0 = 1087
14            ElseIf .Value = chSort1 Then
15                Fid1 = 1087
16            ElseIf .Value = chSort2 Then
17                Fid2 = 1087
18            ElseIf .Value = chSort3 Then
19                Fid3 = 1087
20            ElseIf .Value = chSort4 Then
21                Fid4 = 1087
22            End If
23        End With
24        Fid5 = 420

25        Choices = sArrayStack(sArrayRange("&Sort by", chSort0), _
              sArrayRange("", chSort1), _
              sArrayRange("", chSort2), _
              sArrayRange("", chSort3), _
              sArrayRange("", chSort4), _
              "--" & chChooseSeries)

26        FIDs = sArrayStack(Fid0, Fid1, Fid2, Fid3, Fid4, Fid5)

27        Res = ShowCommandBarPopup(Choices, FIDs, , , ChooseAnchorObject())

28        Select Case Res
              Case "#Cancel!"
                  'Nothing to do
29            Case Unembellish(chChooseSeries)
30                FixBarChart True
31            Case Unembellish(chSort0), Unembellish(chSort1), Unembellish(chSort2), _
                  Unembellish(chSort3), Unembellish(chSort4)
32                Set SPH = CreateSheetProtectionHandler(shBarChart)
33                RangeFromSheet(shBarChart, "SortBy").Value = Res
34            Case Else
35                Throw "Unrecognised choice in menu:" & CStr(Res)

36        End Select
37        Exit Sub
ErrHandler:
38        SomethingWentWrong "#MenuBarChart (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FixBarChart
' Author    : Philip Swannell
' Date      : 04-Dec-2016
' Purpose   : Sets the source data for the bar chart. If ChangeSeries is TRUE then posts
'             dialog asking the user what series they want to display.
' -----------------------------------------------------------------------------------------------------------------------
Sub FixBarChart(ChangeSeries As Boolean)
          Dim cht As Chart
          Dim DataRange As Range
          Dim Headers
          Dim i As Long
          Dim InitialChoices As Variant
          Dim NumRowsToShow As Long
          Dim SeriesToShow
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set cht = shBarChart.ChartObjects(1).Chart

3         Headers = sSubArray(sArrayTranspose(RangeFromSheet(shBarChart, "DataForChart").Rows(1)), 2)

4         InitialChoices = CreateMissing()
5         For i = cht.SeriesCollection.Count To 1 Step -1
6             InitialChoices = sArrayStack(InitialChoices, cht.SeriesCollection(i).Name)
7         Next i

8         If ChangeSeries Then
9             SeriesToShow = ShowMultipleChoiceDialog(Headers, InitialChoices, , "Select series to show", , , , , False)
10            If sIsErrorString(SeriesToShow) Then Exit Sub
11            If sArraysIdentical(sSortedArray(SeriesToShow), sSortedArray(InitialChoices)) Then
12                Exit Sub
13            End If
14        Else
15            SeriesToShow = InitialChoices
16        End If

17        Set SPH = CreateSheetProtectionHandler(shBarChart)

18        NumRowsToShow = 1
19        With RangeFromSheet(shBarChart, "NumBanksForChart")
20            If IsNumber(.Value) Then
21                If .Value >= 1 Then
22                    NumRowsToShow = .Value
23                End If
24            End If
25        End With

          'DataRange does not include Header row or xValues range (i.e. the bank names) and its height is set to only include good data
26        Set DataRange = RangeFromSheet(shBarChart, "DataForChart")
27        Set DataRange = DataRange.offset(1, 1).Resize(NumRowsToShow, DataRange.Columns.Count - 1)

          'Code below uses SeriesCollection rather than FullSeriesCollection for compatibility with Excel 2010
28        For i = cht.SeriesCollection.Count To 1 Step -1
29            cht.SeriesCollection(i).Delete
30        Next i

          Dim Color As Long
31        For i = 1 To sNRows(Headers)
32            If IsNumber(sMatch(Headers(i, 1), SeriesToShow)) Then
33                With cht.SeriesCollection.NewSeries
34                    .xValues = "=" & shBarChart.Name & "!" & DataRange.Columns(0).Address
35                    .Values = "=" & shBarChart.Name & "!" & DataRange.Columns(i).Address
36                    .Name = "=" & shBarChart.Name & "!" & DataRange.Cells(0, i).Address
                      'Various web sites help with choosing maximally distictive colours, values below are from "ColorBrewer"
                      'http://colorbrewer2.org/?type=qualitative&scheme=Paired&n=12#type=qualitative&scheme=Set1&n=4
37                    Color = Choose(i, RGB(228, 26, 28), RGB(55, 126, 184), RGB(77, 175, 74), RGB(152, 78, 163))
                      ' or these ones are the colors that Excel line charts choose if you let them...
38                    Color = Choose(i, RGB(91, 155, 213), RGB(237, 125, 49), RGB(165, 165, 165), RGB(255, 192, 0))
                      'or these that Guy likes
39                    Color = Choose(i, RGB(217, 225, 242), RGB(180, 198, 231), RGB(142, 169, 219), RGB(48, 84, 150))

40                    With .Format.Fill
41                        .Visible = msoTrue
42                        .ForeColor.RGB = Color
43                        .Transparency = 0
44                        .Solid
45                    End With
46                    With .Format.Line
47                        .ForeColor.RGB = Color
48                        .Visible = msoTrue
49                        .Transparency = 0
50                    End With
51                End With
52            End If
53        Next i

54        Exit Sub
ErrHandler:
55        Throw "#FixBarChart (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeIndex
' Author    : Philip Swannell
' Date      : 17-Nov-2016
' Purpose   : Use instead of direct indexing to get an error string rather than
'            raising an error in the event that R or C are out of bounds.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SafeIndex(AnArray, R As Long, c As Long)
1         On Error GoTo ErrHandler
2         SafeIndex = AnArray(R, c)
3         Exit Function
ErrHandler:
4         SafeIndex = "#SafeIndex (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub TrimDictionary(D As Dictionary, KeepThisKey As String)
          Dim k As Variant

1         On Error GoTo ErrHandler
2         For Each k In D.Keys
3             If k <> KeepThisKey Then
4                 D.Remove (k)
5             End If
6         Next

7         Exit Sub
ErrHandler:
8         Throw "#TrimDictionary (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DictionaryToSheet
' Author    : Philip Swannell
' Date      : 30-Sep-2016
' Purpose   : Paste the contents of the D to the sheet Table. Assumes
'             there is a correspondence between the names on the sheet and the keys in the
'             Results collection.
' -----------------------------------------------------------------------------------------------------------------------
Sub DictionaryToSheet(Dict As Dictionary, Sh As Worksheet, CopyOffset As Long)
          Dim c As Range
          Dim Data As Variant
          Dim HAln()
          Dim HH As Long
          Dim i As Long
          Dim ItemName As String
          Dim Target As Range
          Dim x As Variant
1         On Error GoTo ErrHandler

2         HH = GetHedgeHorizon()

3         For Each x In Dict.Keys
4             ItemName = CStr(x)

5             Select Case ItemName
                  Case "BaseUSD", "PVBase", "PVUSD", "Success"
                      'skip these ones
6                 Case Else
7                     Set Target = RangeFromSheet(Sh, ItemName)
8                     Data = Dict(ItemName)

9                     Select Case ItemName
                          'The HedgeHorizon (which determines the height of these arrays) might be anything from 5 to 10 _
                           but the worksheet has 10 columns.
                          Case "MinHeadroomOverFirstN", "MinHeadroomOverFirstNUSD", "MaxPFEByYear", "TradeHeadroom"
10                            If sNRows(Data) < 10 Then
11                                Data = sArrayStack(Data, sReshape(Empty, 10 - sNRows(Data), 1))
12                            End If
13                    End Select

14                    If sNRows(Data) <> 1 Or sNCols(Data) <> 1 Then
15                        If sNRows(Data) = Target.Rows.Count And sNCols(Data) = Target.Columns.Count Then
                              'ok
16                        ElseIf sNRows(Data) = Target.Columns.Count And sNCols(Data) = Target.Rows.Count Then
17                            Data = sArrayTranspose(Data)
18                        Else
19                            Throw "Data for " & ItemName & " has unexpected number of rows or columns, expecting " _
                                  & CStr(Target.Rows.Count) & "," & CStr(Target.Columns.Count) & _
                                  " but got " & CStr(sNRows(Data)) & "," & CStr(sNCols(Data))
20                        End If
21                    End If
22                    Target.Value = sArrayExcelString(Data)

                      'Pasting in the values messes with the horizontal alignment which looks unattractive since we _
                       do a screen refresh after processing each bank, so code below ensures horizontal alignment _
                       remains unchanged.
23                    If CopyOffset <> 0 Then
24                        ReDim HAln(1 To Target.Cells.Count): i = 1
25                        For Each c In Target.offset(CopyOffset).Cells
26                            HAln(i) = c.HorizontalAlignment
27                            i = i + 1
28                        Next c
29                        With Target.offset(CopyOffset)
30                            .Value = sArrayExcelString(Data)
31                            i = 1
32                            For Each c In Target.offset(CopyOffset).Cells
33                                c.HorizontalAlignment = HAln(i)
34                                i = i + 1
35                            Next c
36                        End With
37                    End If
38            End Select
39        Next

40        Exit Sub
ErrHandler:
41        Throw "#DictionaryToSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddOrRemoveBanks
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : Maintenance of the list of banks for which we do static headroom analysis
'             and are included in Scenario Analysis
' -----------------------------------------------------------------------------------------------------------------------
Sub AddOrRemoveBanks()

          Dim AllBanks
          Dim CurrentBanks
          Dim LinesBook As Workbook
          Dim LinesWorkbook As String
          Dim NewBanks
          Dim TheTable As Range
          Dim TheTableNoHeaders As Range
          Dim TopText As String

1         On Error GoTo ErrHandler

2         Set LinesBook = OpenLinesWorkbook(True, False)

3         Set TheTable = RangeFromSheet(shTable, "TheTable")
4         With TheTable
5             Set TheTableNoHeaders = .offset(1).Resize(.Rows.Count - 1)
6         End With

7         AllBanks = sSortedArray(GetColumnFromLinesBook("CPTY_PARENT", LinesBook))
8         CurrentBanks = TheTableNoHeaders.Columns(1).Value

9         AllBanks = AnnotateBankNames(AllBanks, True, LinesBook)
10        CurrentBanks = AnnotateBankNames(CurrentBanks, True, LinesBook)

11        LinesWorkbook = FileFromConfig("LinesWorkbook")

12        TopText = "Please select the banks to be included for:" & vbLf & _
              "a) Credit Headroom analysis using the Table sheet" & vbLf & _
              "b) Dynamic Hedging Analysis using the Scenario sheet." & vbLf & vbLf & _
              "The ""master list"" of banks is held on the Lines workbook" & vbLf & _
              LinesWorkbook & vbLf & vbLf

13        NewBanks = ShowMultipleChoiceDialog(AllBanks, CurrentBanks, "Add or Remove Banks", TopText)
14        If sArraysIdentical(NewBanks, "#User Cancel!") Then GoTo EarlyExit
15        If sNRows(NewBanks) < 2 Then Throw "At least two banks must be selected", True

16        NewBanks = AnnotateBankNames(NewBanks, False, LinesBook)
17        AmendBanksInRange RangeFromSheet(shTable, "TheTable"), NewBanks
18        RangeFromSheet(shTable, "TheFilters").ClearContents
19        FormatTable False
20        AmendBanksInRange RangeFromSheet(shWhoHasLines, "TheTable"), NewBanks
21        RangeFromSheet(shTable, "TheFilters").ClearContents        'has the effect of updating the message above the filters

EarlyExit:
22        Set LinesBook = Nothing
23        Exit Sub
ErrHandler:
24        SomethingWentWrong "#AddOrRemoveBanks (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AnnotateBankNames
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : The CPTY_PARENT strings are a bit unfriendly so this method appends
'             the more understandable CPTY LONG NAMEs
' -----------------------------------------------------------------------------------------------------------------------
Function AnnotateBankNames(TheBanks, Annotate As Boolean, LinesBook As Workbook, Optional ForCommandBar = False)
          Dim AllLongNames
          Dim AllNames
          Dim AllPrettyNames
          Dim Res
          Static HaveSavedToRegistry As Boolean

1         On Error GoTo ErrHandler
2         AllNames = GetColumnFromLinesBook("CPTY_PARENT", LinesBook)
3         AllLongNames = GetColumnFromLinesBook("CPTY LONG NAME", LinesBook)

4         If Not HaveSavedToRegistry Then
              'Somewhat shameful hack in section below - necessary so we can run a version of AnnotateBankNames
              'from SCRiPTUtils which does not know where to find the Lines book
5             SaveSetting "Cayley", "DataFromLinesBook", "AllNames", sMakeArrayString(AllNames)
6             SaveSetting "Cayley", "DataFromLinesBook", "AllLongNames", sMakeArrayString(AllLongNames)
7             HaveSavedToRegistry = True
8         End If

9         If ForCommandBar Then
10            AllPrettyNames = sJustifyArrayOfStrings(sArrayRange(AllNames, AllLongNames), "Segoe UI", 9, "           " & vbTab)
11        Else
12            AllPrettyNames = sJustifyArrayOfStrings(sArrayRange(AllNames, AllLongNames), "Tahoma", 8, " " & vbTab)
13        End If

14        If Annotate Then
15            Res = sVLookup(TheBanks, sArrayRange(AllNames, AllPrettyNames))
16        Else
17            Res = sVLookup(TheBanks, sArrayRange(AllPrettyNames, AllNames))
18        End If

19        Res = sArrayIf(sArrayEquals(Res, "#Not found!"), TheBanks, Res)

20        AnnotateBankNames = Res
21        Exit Function
ErrHandler:
22        Throw "#AnnotateBankNames (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AmendBanksInRange
' Author    : Philip Swannell
' Date      : 23-Sep-2016
' Purpose   : Provides user interface for adding and removing banks from the ranges TheTableWithHeaders on sheets
'             Table and WhoHasLines.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AmendBanksInRange(TheTableWithHeaders As Range, NewBanks)
          Dim BanksToAdd
          Dim CurrentBanks
          Dim i As Long
          Dim MatchIDs
          Dim RangeToDelete As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TheTableNoHeaders As Range

1         On Error GoTo ErrHandler
2         With TheTableWithHeaders
3             Set TheTableNoHeaders = .offset(1).Resize(.Rows.Count - 1)
4         End With

5         If TheTableWithHeaders.Rows.Count < 3 Then Throw "Assertion Failed: existing range must have at least two rows"

6         CurrentBanks = TheTableNoHeaders.Columns(1).Value
7         If sArraysIdentical(sSortedArray(CurrentBanks), sSortedArray(NewBanks)) Then Exit Sub

8         Set SPH = CreateSheetProtectionHandler(TheTableWithHeaders.Parent)
9         Set SUH = CreateScreenUpdateHandler()

10        BanksToAdd = sCompareTwoArrays(NewBanks, CurrentBanks, "In1AndNotIn2")

          'Clear the "Index" column to the left
11        If TheTableWithHeaders.Parent Is shWhoHasLines Then
12            TheTableWithHeaders.Columns(-1).Clear
13        End If

          'Add banks
14        If sNRows(BanksToAdd) > 1 Then
15            BanksToAdd = sDrop(BanksToAdd, 1)
16            TheTableWithHeaders.Rows(3).Resize(sNRows(BanksToAdd)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
17            TheTableWithHeaders.Cells(3, 1).Resize(sNRows(BanksToAdd)).Value = BanksToAdd
18        End If
          'Delete banks
19        MatchIDs = sMatch(TheTableNoHeaders.Columns(1), NewBanks)
20        For i = 1 To sNRows(MatchIDs)
21            If Not IsNumber(MatchIDs(i, 1)) Then
22                If RangeToDelete Is Nothing Then
23                    Set RangeToDelete = TheTableWithHeaders.Rows(i + 1)
24                Else
25                    Set RangeToDelete = Application.Union(RangeToDelete, TheTableWithHeaders.Rows(i + 1))
26                End If
27            End If
28        Next i
29        If Not RangeToDelete Is Nothing Then
30            RangeToDelete.Delete Shift:=xlUp
31        End If

32        AddGreyBorders TheTableWithHeaders, True
          'Making the left border of the N+1 column grey is useful given that we have more-less buttons on the table sheet
33        With TheTableWithHeaders.Resize(TheTableWithHeaders.Rows.Count + 100).Columns(TheTableWithHeaders.Columns.Count + 1).Borders(xlEdgeLeft)
34            .LineStyle = xlNone
35        End With

36        With TheTableWithHeaders.Columns(TheTableWithHeaders.Columns.Count + 1).Borders(xlEdgeLeft)
37            .LineStyle = xlContinuous
38            .Weight = xlThin
39            .ThemeColor = 1
40            .TintAndShade = -0.249946592608417
41        End With

          'Put the "Index" column to the left back in
42        If TheTableWithHeaders.Parent Is shWhoHasLines Then
43            TheTableNoHeaders.Columns(-1).FormulaArray = "=sIntegers(" & CStr(TheTableNoHeaders.Rows.Count) & ")"
44        End If

          'Sort the table
45        TheTableWithHeaders.Parent.Sort.SortFields.Clear
46        TheTableWithHeaders.Parent.Sort.SortFields.Add key:=TheTableNoHeaders.Columns(1) _
              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
47        With TheTableWithHeaders.Parent.Sort
48            .SetRange TheTableWithHeaders
49            .Header = xlYes
50            .MatchCase = False
51            .Orientation = xlTopToBottom
52            .SortMethod = xlPinYin
53            .Apply
54        End With

55        If TheTableWithHeaders.Parent Is shTable Then
56            FormatTable False
57        End If

58        ResetSortButtons TheTableWithHeaders.Rows(0), False, False

59        Exit Sub
ErrHandler:
60        Throw "#AmendBanksInRange (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResetTableButtons
' Author    : Philip Swannell
' Date      : 07-Nov-2016
' Purpose   : Ensures the sort and grouping buttons on the Table sheet are in good order.
'             Call in release cleanup. Relies on cell merging in rows 4 and 5 to determine
'             which columns the grouping buttons should span.
' -----------------------------------------------------------------------------------------------------------------------
Sub ResetTableButtons()
          Dim b As Button
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim ThisBlock As Range

1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shTable)
3         Set SUH = CreateScreenUpdateHandler()

4         shTable.Activate
5         ActiveWindow.Zoom = 100
6         shTable.UsedRange.Columns.Hidden = False
7         For Each b In shTable.Buttons
8             If Len(b.Caption) = 1 Or InStr("GroupingButton", b.OnAction) > 0 Then
9                 b.Delete
10            End If
11        Next b

12        AddSortButtons RangeFromSheet(shTable, "TheTable").Rows(0)

13        Application.GoTo RangeFromSheet(shTable, "TheTable").Rows(-2)

14        With RangeFromSheet(shTable, "TheTable").Rows(-2)
15            Set ThisBlock = .Cells(1, 0)
16            While (ThisBlock.Column + ThisBlock.Columns.Count - 1) < (.Column + .Columns.Count - 1)
17                Set ThisBlock = ThisBlock.offset(0, ThisBlock.Columns.Count).Resize(1, ThisBlock.Cells(-1, ThisBlock.Columns.Count + 1).MergeArea.Columns.Count)
18                AddGroupingButtonToRange ThisBlock, False
19            Wend
20        End With
21        Application.GoTo shTable.Cells(1, 1)
22        Exit Sub
ErrHandler:
23        Throw "#ResetTableButtons (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : FormatTable
' Author    : Philip Swannell
' Date      : 03-Oct-2016
' Purpose   : Apply cell formatting to the range TheTable on sheet Table
' -----------------------------------------------------------------------------------------------------------------------
Sub FormatTable(ClearFilters As Boolean)
          Dim Headers
          Dim i As Long
1         On Error GoTo ErrHandler

2         Dim SUH As clsScreenUpdateHandler: Set SUH = CreateScreenUpdateHandler()
3         Dim SPH As clsSheetProtectionHandler: Set SPH = CreateSheetProtectionHandler(shTable)

4         With RangeFromSheet(shTable, "TheFilters")
5             .Locked = False
6             CayleyFormatAsInput .offset(0)
7             .NumberFormat = "@"
8             .HorizontalAlignment = xlHAlignCenter
9             If ClearFilters Then .ClearContents
10        End With

11        With RangeFromSheet(shTable, "TheTable")
12            shTable.Names.Add "TheTableNoHeaders", .offset(1).Resize(.Rows.Count - 1)
13            With .Cells(-2, .Columns.Count + 1)
14                .Value = " <-click to expand or shift+click to expand all"
15                .Font.Color = g_Col_GreyText
16            End With
17            With .Cells(-1, .Columns.Count + 1)
18                .Value = " <-double-click to filter"
19                .Font.Color = g_Col_GreyText
20            End With

21            With .Cells(0, .Columns.Count + 1)
22                .Value = " <-click to sort"
23                .Font.Color = g_Col_GreyText
24            End With

25            .Rows(-3).EntireRow.Hidden = True
26            Headers = sArrayTranspose(.Rows(1).Value)
27            For i = 1 To sNRows(Headers)
28                Headers(i, 1) = Replace(Replace(Replace(Headers(i, 1), vbLf, ""), vbCr, ""), " ", "")
29            Next i
30            .ClearFormats
31            With .Rows(1)
32                .Font.Bold = True
33                .VerticalAlignment = xlVAlignCenter
34                .HorizontalAlignment = xlHAlignCenter
35                .WrapText = True
36            End With

37            With GetSubTable(.offset(0), Headers, "TheBank", "Short Name", True)
38                AddGreyBorders .offset(0), True
39                AddGreyBorders .offset(-4).Resize(2)
40            End With

41            With GetSubTable(.offset(0), Headers, "Process Time", "Process Time", True)
42                .NumberFormat = "0.0"
43                .HorizontalAlignment = xlHAlignCenter
44            End With

45            With GetSubTable(.offset(0), Headers, "Time Stamp", "Time Stamp", True)
46                .NumberFormat = "dd-mmm hh:mm"
47                .HorizontalAlignment = xlHAlignCenter
48            End With

49            With GetSubTable(.offset(0), Headers, "EUR PV", "Time Stamp", True)
50                AddGreyBorders .offset(0), True
51                AddGreyBorders .offset(-4).Resize(2)
52                .Columns(1).NumberFormat = "#,##0;[Red]-#,##0"
53            End With

54            With GetSubTable(.offset(0), Headers, "Num Trades", "FxVolShock", True)
55                .HorizontalAlignment = xlHAlignCenter
56            End With

57            With GetSubTable(.offset(0), Headers, "EURUSD3YVol", "EURUSD3YVol", True)
58                .NumberFormat = "0.00%"
59            End With

60            With GetSubTable(.offset(0), Headers, "1Y Limit", "10Y Limit", True)
61                AddGreyBorders .offset(0), True
62                AddGreyBorders .offset(-4).Resize(2)
63                .NumberFormat = "#,##0;[Red]-#,##0"
64            End With
65            With GetSubTable(.offset(0), Headers, "Methodology", "Notional Cap", True)
66                .Columns(6).NumberFormat = "#,##0;[Red]-#,##0"
67                AddGreyBorders .offset(0), True
68                AddGreyBorders .offset(-4).Resize(2)
69            End With

70            With GetSubTable(.offset(0), Headers, "PFEPercentile", "PFEPercentile", True)
71                .HorizontalAlignment = xlHAlignCenter
72            End With

73            With GetSubTable(.offset(0), Headers, "Base Currency", "Base Currency", True)
74                .HorizontalAlignment = xlHAlignCenter
75            End With

76            With GetSubTable(.offset(0), Headers, "HR 1Y", "HR 10Y", True)
77                AddGreyBorders .offset(0), True
78                AddGreyBorders .offset(-4).Resize(2)
79                .NumberFormat = "#,##0;[Red]-#,##0"
80            End With

81            With GetSubTable(.offset(0), Headers, "HR 1Y USD", "HR 10Y USD", True)
82                AddGreyBorders .offset(0), True
83                AddGreyBorders .offset(-4).Resize(2)
84                .NumberFormat = "#,##0;[Red]-#,##0"
85            End With

86            With GetSubTable(.offset(0), Headers, "Max PFE 1Y", "Max PFE 10Y", True)
87                AddGreyBorders .offset(0), True
88                AddGreyBorders .offset(-4).Resize(2)
89                .NumberFormat = "#,##0;[Red]-#,##0"
90            End With

91            With GetSubTable(.offset(0), Headers, "Bank THR 1Y", "Bank THR 3Y", True)
92                AddGreyBorders .offset(0), True
93                AddGreyBorders .offset(-4).Resize(2)
94                .NumberFormat = "#,##0;[Red]-#,##0"
95            End With

96            With GetSubTable(.offset(0), Headers, "THR 1Y", "THR 10Y", True)
97                AddGreyBorders .offset(0), True
98                AddGreyBorders .offset(-4).Resize(2)
99                .NumberFormat = "#,##0;[Red]-#,##0"
100           End With

101           With GetSubTable(.offset(0), Headers, "FxHR", "FxHR", True)
102               AddGreyBorders .offset(0), True
103               AddGreyBorders .offset(-4).Resize(2)
104               .NumberFormat = "0.000"
105               .HorizontalAlignment = xlHAlignCenter
106           End With

107           With GetSubTable(.offset(0), Headers, "Airbus THR 3Y", "Net THR 3Y", True)
108               AddGreyBorders .offset(0), True
109               AddGreyBorders .offset(-4).Resize(2)
110               .NumberFormat = "#,##0;[Red]-#,##0"

111               With .offset(-4, .Columns.Count).Resize(.Rows.Count + 4).Borders(xlEdgeLeft)
112                   .LineStyle = xlContinuous
113                   .Weight = xlThin
114                   .ThemeColor = 1
115                   .TintAndShade = -0.249946592608417
116               End With

117           End With

118       End With

119       Exit Sub
ErrHandler:
120       Throw "#FormatTable (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetHeaderNumber
' Author    : Philip Swannell
' Date      : 03-Oct-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetHeaderNumber(Headers, Header As String) As Long
          Dim CleanHeader As String
          Dim MatchRes As Variant
1         On Error GoTo ErrHandler
2         CleanHeader = Replace(Replace(Replace(Header, vbLf, ""), vbCr, ""), " ", "")
3         MatchRes = sMatch(CleanHeader, Headers)
4         If Not IsNumber(MatchRes) Then Throw "Cannot find header titled '" & Header & "' in top row of range TheTable on sheet " & shTable.Name
5         GetHeaderNumber = MatchRes
6         Exit Function
ErrHandler:
7         Throw "#GetHeaderNumber (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetSubTable
' Author    : Philip Swannell
' Date      : 03-Oct-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetSubTable(TheTable As Range, Headers, FirstHeader As String, LastHeader As String, WithTopRow As Boolean) As Range
          Dim Col1 As Long
          Dim Col2 As Long
          Dim Res As Range
1         On Error GoTo ErrHandler
2         Col1 = GetHeaderNumber(Headers, FirstHeader)
3         Col2 = GetHeaderNumber(Headers, LastHeader)

4         Set Res = TheTable.Columns(Col1).Resize(, Col2 - Col1 + 1)

5         If Not WithTopRow Then
6             Set Res = Res.offset(1).Resize(Res.Rows.Count - 1)
7         End If

8         Set GetSubTable = Res

9         Exit Function
ErrHandler:
10        Throw "#GetSubTable (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SyncBanksInCayleyWithBanksInLinesBook
' Author     : Philip Swannell
' Date       : 07-Apr-2022
' Purpose    : We store the list of banks twice in this workbook, on sheet Table and on sheet WhoHasLines, this method
'              brings those sheets into sync with the list of banks in the lines workbook.
' -----------------------------------------------------------------------------------------------------------------------
Sub SyncBanksInCayleyWithBanksInLinesBook()

          Dim LinesBook As Workbook
          Dim LinesBookBanks

1         On Error GoTo ErrHandler
2         Set LinesBook = OpenLinesWorkbook(True, False)

3         LinesBookBanks = GetColumnFromLinesBook("CPTY_PARENT", LinesBook)
4         AmendBanksInRange RangeFromSheet(shTable, "TheTable"), LinesBookBanks
5         AmendBanksInRange RangeFromSheet(shWhoHasLines, "TheTable"), LinesBookBanks

6         Set LinesBook = Nothing
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#SyncBanksInCayleyWithBanksInLinesBook (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub


